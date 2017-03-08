!define NAME "nsf2x"
!ifndef VERSION
    !define VERSION "X.X.X"
!endif
!ifndef PUBLISHER
    !define PUBLISHER "root@localhost"
!endif
!ifndef BITNESS
    !define BITNESS "x86"
!endif
!define UNINSTKEY "${NAME}-${VERSION}-${BITNESS}"

!define DEFAULTNORMALDESTINATON "$ProgramFiles\${NAME}-${VERSION}-${BITNESS}"
!define DEFAULTPORTABLEDESTINATON "$LocalAppdata\Programs\${NAME}-${VERSION}-${BITNESS}"

!define CLSID_IConverterSession "{4E3A7680-B77A-11D0-9DA5-00C04FD65685}"
!define CLSID_IMimeMessage "{9EADBD1A-447B-4240-A9DD-73FE7C53A981}"

# Alternative locations of different versions of ClickToRun Office
# If new values are added here, the logic of the ForEach loop 
# below needs to be changed (NSIS doesn't have arrays or lists)
!define OfficeXXClickToRun "Software\Microsoft\Office\ClickToRun\Registry\MACHINE\Software\Classes"
!define Office15ClickToRun "Software\Microsoft\Office\15.0\ClickToRun\Registry\MACHINE\Software\Classes"
!define Office16ClickToRun "Software\Microsoft\Office\16.0\ClickToRun\Registry\MACHINE\Software\Classes"

Name "${NAME}"
Outfile "${NAME}-${VERSION}-${BITNESS}-setup.exe"
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
Var ClickToRun
Var OfficeClickToRun
Var ClickToRunDestination
Var i

!include LogicLib.nsh
!include FileFunc.nsh
!include MUI2.nsh
!include nsDialogs.nsh
!include registry.nsh

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

Function CheckClickToRun
StrCpy $ClickToRun ""
StrCpy $ClickToRunDestination ""

# No arrays or lists in NSIS, so use a bit of messy logic to choose
# potential registry locations of ClickToRun installs

${ForEach} $i 1 3 + 1
    ${If} $i == "1"
        StrCpy $OfficeClickToRun ${OfficeXXClickToRun}
    ${ElseIf} $i == "2"
        StrCpy $OfficeClickToRun ${Office16ClickToRun}
    ${ElseIf} $i == "3"
        StrCpy $OfficeClickToRun ${Office15ClickToRun}
    ${EndIf}
    
    ReadRegStr $R1 HKLM "$OfficeClickToRun\Wow6432Node\CLSID\${CLSID_IConverterSession}" ""
    ${If} $R1 != ""
        ; Have Click To Run with different bitness
        StrCpy $ClickToRun "$OfficeClickToRun\Wow6432Node\CLSID"
        StrCpy $ClickToRunDestination "Software\Classes\Wow6432Node\CLSID"
        ${Break}
    ${EndIf} 

    ReadRegStr $R2 HKLM "$OfficeClickToRun\CLSID\${CLSID_IConverterSession}" ""
    ${If} $R2 != ""
        ; Have Click To Run with same bitness
        StrCpy $ClickToRun "$OfficeClickToRun\CLSID"
        StrCpy $ClickToRunDestination "Software\Classes\CLSID"
        ${Break}
    ${EndIf}
${Next}

# Check if registry is already patched.
ReadRegStr $R0 HKLM "$ClickToRunDestination\${CLSID_IConverterSession}" ""
${If} $R0 != ""
    StrCpy $ClickToRun ""
    StrCpy $ClickToRunDestination ""
${EndIf}

${If} $ClickToRun != ""
    MessageBox MB_YESNOCANCEL|MB_TOPMOST|MB_ICONEXCLAMATION \
    "You appear to have a $\"Click To Run$\" version of Outlook installed. This will interfere with \
    the conversion to Outlook PST files. You have three choices$\r$\n\
    $\r$\n\
    1. Allow NSF2X to patch the registry. In this case the NSF2X installer must run with administrator \
    privileges. This fix will also cause issues if you are running multiple versions of Outlook. \
    After patching, your $\"Click To Run$\" Outlook client will work correctly, but older versions \
    of Outlook might fail in unexplained manners. See$\r$\n\
    $\r$\n\
    https://blogs.msdn.microsoft.com/stephen_griffin/2014/04/21/outlook-2013-click-to-run-and-com-interfaces/$\r$\n\
    $\r$\n\
    for more information$\r$\n\
    2. Continue the installation knowing that the conversion to PST will not be possible$\r$\n\
    3. Cancel the installation of NSF2X$\r$\n\
    $\r$\n\
    Do you wish to let NSF2X patch the registry ?" IDYES Yes IDNO No
    ; Only get here if the cancel button was pressed
    Quit
    
    Yes:
    ${registry::CopyKey} "HKLM\$ClickToRun\${CLSID_IConverterSession}" "HKLM\$ClickToRunDestination\${CLSID_IConverterSession}" $R0
    ${If} $R0 != 0
        MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IConverterSession$\" registry value. Aborting"
        Quit
    ${EndIf}
    ${registry::CopyKey} "HKLM\$ClickToRun\${CLSID_IMimeMessage}" "HKLM\$ClickToRunDestination\${CLSID_IMimeMessage}" $R0
    ${If} $R0 != 0
        MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IMimeMessage$\" registry value. Aborting"
        Quit
    ${EndIf}

    Goto Finished
    No:
    ; Do nothing. PST conversion won't work 
    StrCpy $ClickToRun ""
    StrCpy $ClickToRunDestination ""
    Finished:
${EndIf}
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
    MessageBox MB_OKCANCEL|MB_TOPMOST "NSF2X version ${VERSION} is already installed. Launch the uninstaller?"  IDOK Ok IDCANCEL Cancel
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

Call CheckClickToRun
FunctionEnd

Section "Install"
SetOutPath "$InstDir"
File /r dist\*

WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "DisplayName" "${NAME}"
WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "DisplayVersion" "${VERSION}"
WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "Publisher" "${PUBLISHER}"
WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "ClickToRun" $ClickToRun
WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "ClickToRunDestination" $ClickToRunDestination
${If} $InstallAllUsers  == ${BST_CHECKED}
    WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "UninstallString" '"$InstDir\uninstall.exe" /ALL'
${Else}
    WriteRegStr SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "UninstallString" '"$InstDir\uninstall.exe" /USER'
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
    CreateShortCut "$SMPROGRAMS\${NAME}-${VERSION}\uninstall.lnk" "$InstDir\uninstall.exe" "/USER" "" 0
${EndIf}

${If} $InstallShortCuts == ${BST_CHECKED}
    Delete "$DESKTOP\${NAME}.lnk"
    CreateShortCut "$DESKTOP\${NAME}.lnk" "$InstDir\nsf2x.exe" "" "" 0
${Endif}
SectionEnd

Function un.onInit
${GetParameters} $9

ClearErrors
${GetOptions} $9 "/?" $8
${IfNot} ${Errors}
    MessageBox MB_ICONINFORMATION|MB_SETFOREGROUND "\
      /ALL : remove application for all users$\n\
      /USER : remove application for the current user$\n\
      /S : Silent$\n"
    Quit
${EndIf}

ReadRegStr $R0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "UninstallString"
${If} $R0 != ""
    SetShellVarContext all
${Else}
    SetShellVarContext current
${EndIf}
    
ClearErrors
${GetOptions} $9 "/ALL" $8
${IfNot} ${Errors}
    SetShellVarContext all
${EndIf}

${GetOptions} $9 "/USER" $8
${IfNot} ${Errors}
    SetShellVarContext current
${EndIf}
FunctionEnd

Section "Uninstall"
ReadRegStr $R0 SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "ClickToRun"
${If} $R0 != ""
    MessageBox MB_YESNOCANCEL "Do you wish to let NSF2X remove the modifications it made to the registry to support $\"Click To Run$\" ?" IDYES Yes2 IDNO No2
    Quit
    Yes2:
    ReadRegStr $R1 SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}" "ClickToRunDestination"
    ${If} $R1 != ""
        DeleteRegKey HKLM "$R1\${CLSID_IConverterSession}"
        DeleteRegKey HKLM "$R1\${CLSID_IMimeMessage}"   
    ${EndIf}
    No2:
${EndIf}
DeleteRegKey SHCTX "Software\Microsoft\Windows\CurrentVersion\Uninstall\${UNINSTKEY}"

RMDir /r "$SMPROGRAMS\${NAME}-${VERSION}"
RMDir "$SMPROGRAMS\${NAME}-${VERSION}"
Delete "$InstDir\uninstall.exe"
RMDir /r "$InstDir"
RMDir "$InstDir"

Delete "$DESKTOP\${NAME}.lnk"
SectionEnd