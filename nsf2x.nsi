!define NAME "nsf2x"
!ifndef VERSION
    !define VERSION "X.X.X"
!endif
!ifndef PUBLISHER
    !define PUBLISHER "root@localhost"
!endif
!define UNINSTKEY "${NAME}-${VERSION}"

!define DEFAULTNORMALDESTINATON "$ProgramFiles\${NAME}-${VERSION}"
!define DEFAULTPORTABLEDESTINATON "$LocalAppdata\Programs\${NAME}-${VERSION}"

Name "${NAME}"
Outfile "${NAME}-${VERSION}-setup.exe"
RequestExecutionlevel highest
SetCompressor LZMA

; Keep NSIS v3.0 Happy
Unicode true
ManifestDPIAware true

Var NormalDestDir
Var LocalDestDir
Var InstallAllUsers
Var InstallAllUsersCtrl
Var InstallShortcuts
Var InstallShortcutsCtrl

!include LogicLib.nsh
!include FileFunc.nsh
!include MUI2.nsh
!include nsDialogs.nsh

!insertmacro MUI_PAGE_WELCOME
!define MUI_LICENSEPAGE_TEXT_BOTTOM "The source code for NSF2X is freely redistributable under the terms of the GNU General Public License (GPL) as published by the Free Software Foundation."
!define MUI_LICENSEPAGE_BUTTON "Next >"
!insertmacro MUI_PAGE_LICENSE "LICENSE"
Page Custom OptionsPageCreate OptionsPageLeave
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_SHOWREADME $INSTDIR\README.txt
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE English

Function .onInit
StrCpy $NormalDestDir "${DEFAULTNORMALDESTINATON}"
StrCpy $LocalDestDir "${DEFAULTPORTABLEDESTINATON}"

${GetParameters} $9

ClearErrors
${GetOptions} $9 "/?" $8
${IfNot} ${Errors}
    MessageBox MB_ICONINFORMATION|MB_SETFOREGROUND "\
      /ALL : Extract application for all users$\n\
      /USER : Extract application for current user$\n\
      /SHORTCUT : Install desktop shortcut$\n\
      /S : Silent install$\n\
      /D=%directory% : Specify destination directory$\n"
    Quit
${EndIf}

ClearErrors
${GetOptions} $9 "/ALL" $8
${IfNot} ${Errors}
    StrCpy $0 $NormalDestDir
    ${If} ${Silent}
        Call RequireAdmin
    ${EndIf}
    SetShellVarContext all
    StrCpy $InstallAllUsers ${BST_CHECKED}
${Else}
    SetShellVarContext current
    StrCpy $0 $LocalDestDir
    StrCpy $InstallAllUsers ${BST_UNCHECKED}
${EndIf}

${GetOptions} $9 "/USER" $8
${IfNot} ${Errors}
    SetShellVarContext current
    StrCpy $0 $LocalDestDir
    StrCpy $InstallAllUsers ${BST_UNCHECKED}
${Else}
    StrCpy $0 $NormalDestDir
    ${If} ${Silent}
        Call RequireAdmin
    ${EndIf}
    SetShellVarContext all
    StrCpy $InstallAllUsers ${BST_CHECKED}
${EndIf}

${GetOptions} $9 "/SHORTCUT" $8
${IfNot} ${Errors}
    StrCpy $InstallShortCuts ${BST_CHECKED}
${Else}
    StrCpy $InstallShortCuts ${BST_UNCHECKED}
${EndIf}

${If} $InstDir == ""
    ; User did not use /D to specify a directory, 
    ; we need to set a default based on the install mode
    StrCpy $InstDir $0
${EndIf}
Call SetModeDestinationFromInstdir
FunctionEnd

Function RequireAdmin
UserInfo::GetAccountType
Pop $8
${If} $8 != "admin"
    MessageBox MB_ICONSTOP "You need administrator rights to install ${NAME}"
    SetErrorLevel 740 ;ERROR_ELEVATION_REQUIRED
    Abort
${EndIf}
FunctionEnd

Function SetModeDestinationFromInstdir
${If} $InstallAllUsers == ${BST_CHECKED}
    StrCpy $NormalDestDir $InstDir
${Else}
    StrCpy $LocalDestDir $InstDir
${EndIf}
FunctionEnd

Function OptionsPageCreate
!insertmacro MUI_HEADER_TEXT "Install Mode" "Choose how you want to install ${NAME}."

Push $0
nsDialogs::Create 1018
Pop $0
${If} $0 == error
    Abort
${EndIf} 

Call SetModeDestinationFromInstdir ; If the user clicks BACK on the directory page we will remember their mode specific directory

${NSD_CreateCheckBox} 0 0 100% 12u "Install for all users (requires administrator privileges)"
Pop $InstallAllUsersCtrl

UserInfo::GetAccountType
Pop $8
${If} $8 != "admin"
    ${NSD_SetState} $InstallAllUsersCtrl ${BST_UNCHECKED}
${Else}
    ${NSD_SetState} $InstallAllUsersCtrl ${BST_CHECKED}
${EndIf}

${NSD_CreateCheckBox} 0 20 100% 12u "Create desktop shortcut"
Pop $InstallShortcutsCtrl
${NSD_SetState} $InstallShortcutsCtrl ${BST_CHECKED}

nsDialogs::Show
FunctionEnd

Function OptionsPageLeave
${NSD_GetState} $InstallAllUsersCtrl $InstallAllUsers
${NSD_GetState} $InstallShortcutsCtrl $InstallShortcuts

${If} $InstallAllUsers  == ${BST_CHECKED}
    StrCpy $InstDir $NormalDestDir
    Call RequireAdmin
    SetShellVarContext all
${Else}
    StrCpy $InstDir $LocalDestDir
    SetShellVarContext current
${EndIf}

; Check to see if already installed
ReadRegStr $R0 SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "UninstallString"
${If} $R0 != "" 
    MessageBox MB_OKCANCEL|MB_TOPMOST "NSF2X version  ${VERSION} is already installed. Launch the uninstaller?"  IDOK Ok IDCANCEL Cancel
    Cancel:
    Quit

    Ok:
    ; Use SilentMode for the uninstall so that we can wait on the termination
    ${If} $InstallAllUsers  == ${BST_CHECKED}
        ExecWait "$R0 /S /ALL" 
    ${Else}
        ExecWait "$R0 /S"
    ${EndIf}
${EndIf}
FunctionEnd

Section
SetOutPath "$InstDir"
File /r dist\*

WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "DisplayName" "${NAME}"
WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "DisplayVersion" "${VERSION}"
WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "Publisher" "${PUBLISHER}"
${If} $InstallAllUsers  == ${BST_CHECKED}
    WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "UninstallString" '"$InstDir\uninstall.exe" /ALL'
${Else}
    WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "UninstallString" '"$InstDir\uninstall.exe"'
${EndIf}

${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
IntFmt $0 "0x%08X" $0
WriteRegDWORD SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "EstimatedSize" "$0"

WriteUninstaller "$InstDir\uninstall.exe"
SectionEnd

Section "Shortcuts"
CreateDirectory "$SMPROGRAMS\${NAME}-${VERSION}"
CreateShortCut "$SMPROGRAMS\${NAME}-${VERSION}\nsf2x.lnk" "$InstDir\nsf2x.exe" "" "" 0
CreateShortCut "$SMPROGRAMS\${NAME}-${VERSION}\README.lnk" "$InstDir\README.txt" "" "" 0

${If} $InstallAllUsers  == ${BST_CHECKED}
    CreateShortCut "$SMPROGRAMS\${NAME}-${VERSION}\uninstall.lnk" "$InstDir\uninstall.exe" "/ALL" "" 0
${Else}
    CreateShortCut "$SMPROGRAMS\${NAME}-${VERSION}\uninstall.lnk" "$InstDir\uninstall.exe" "" "" 0
${EndIf}

${If} $InstallShortCuts == ${BST_CHECKED}
    Delete "$DESKTOP\${NAME}.lnk"
    CreateShortCut "$DESKTOP\${NAME}.lnk" "$InstDir\nsf2x.exe" "" "" 0
${Endif}
SectionEnd

Section un.onInit
${GetParameters} $9

ClearErrors
${GetOptions} $9 "/?" $8
${IfNot} ${Errors}
    MessageBox MB_ICONINFORMATION|MB_SETFOREGROUND "\
      /ALL : remove application for all users$\n\
      /LOCAL : remove application for the current user$\n\
      /S : Silent$\n"
    Quit
${EndIf}

ClearErrors
${GetOptions} $9 "/ALL" $8
${IfNot} ${Errors}
    SetShellVarContext all
${Else}
    SetShellVarContext current
${EndIf}

${GetOptions} $9 "/USER" $8
${IfNot} ${Errors}
    SetShellVarContext current
${Else}
    SetShellVarContext all
${EndIf}
SectionEnd

Section Uninstall
DeleteRegKey SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}"

RMDir /r "$SMPROGRAMS\${NAME}-${VERSION}"
RMDir "$SMPROGRAMS\${NAME}-${VERSION}"
Delete "$InstDir\uninstall.exe"
RMDir /r "$InstDir"
RMDir "$InstDir"

Delete "$DESKTOP\${NAME}.lnk"
SectionEnd