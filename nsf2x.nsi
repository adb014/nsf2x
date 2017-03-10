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

; Keep NSIS v3.0 Happy
Unicode true
ManifestDPIAware true

Name "${NAME}"
Outfile "${NAME}-${VERSION}-${BITNESS}-setup.exe"
RequestExecutionlevel highest
SetCompressor LZMA

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

!include "nsf2x_lang.nsi"

Function .onInit
StrCpy $NormalDestDir "${DEFAULTNORMALDESTINATON}"
StrCpy $LocalDestDir "${DEFAULTPORTABLEDESTINATON}"

${GetParameters} $9

ClearErrors
${GetOptions} $9 "/?" $8
${IfNot} ${Errors}
    MessageBox MB_ICONINFORMATION|MB_SETFOREGROUND "$(COMMANDLINE_HELP)"
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
    MessageBox MB_YESNOCANCEL|MB_TOPMOST|MB_ICONEXCLAMATION "$(HAVE_CLICKTORUN)" IDYES Yes IDNO No
    ; Only get here if the cancel button was pressed
    Quit
    
    Yes:
    ${registry::CopyKey} "HKLM\$ClickToRun\${CLSID_IConverterSession}" "HKLM\$ClickToRunDestination\${CLSID_IConverterSession}" $R0
    ${If} $R0 != 0
        MessageBox MB_OK|MB_ICONEXCLAMATION "$(FAIL_COPY_ICONV)"
        Quit
    ${EndIf}
    ${registry::CopyKey} "HKLM\$ClickToRun\${CLSID_IMimeMessage}" "HKLM\$ClickToRunDestination\${CLSID_IMimeMessage}" $R0
    ${If} $R0 != 0
        MessageBox MB_OK|MB_ICONEXCLAMATION "$(FAIL_COPY_MIME)"
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
    MessageBox MB_ICONSTOP "$(ADMIN_RIGHTS)"
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
!insertmacro MUI_HEADER_TEXT "$(INSTALL_MODE)" "$(INSTALL_CHOOSE)"

Push $0
nsDialogs::Create 1018
Pop $0
${If} $0 == error
    Abort
${EndIf} 

Call SetModeDestinationFromInstdir ; If the user clicks BACK on the directory page we will remember their mode specific directory

${NSD_CreateCheckBox} 0 0 100% 12u "$(INSTALL_ADMIN)"
Pop $InstallAllUsersCtrl

UserInfo::GetAccountType
Pop $8
${If} $8 != "admin"
    ${NSD_SetState} $InstallAllUsersCtrl ${BST_UNCHECKED}
${Else}
    ${NSD_SetState} $InstallAllUsersCtrl ${BST_CHECKED}
${EndIf}

${NSD_CreateCheckBox} 0 20 100% 12u "$(INSTALL_SHORTCUT)"
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
    MessageBox MB_OKCANCEL|MB_TOPMOST "$(ALREADY_INSTALLED)" IDOK Ok IDCANCEL Cancel
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
    MessageBox MB_ICONINFORMATION|MB_SETFOREGROUND "$(COMMANDLINE_HELP2)"
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
    MessageBox MB_YESNOCANCEL "$(UNINSTALL_REGISTRY)" IDYES Yes2 IDNO No2
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