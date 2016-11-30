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

!define CLSID_IConverterSession "{4E3A7680-B77A-11D0-9DA5-00C04FD65685}"
!define CLSID_IMimeMessage "{9EADBD1A-447B-4240-A9DD-73FE7C53A981}"

!define Office15ClickToRun "Software\Microsoft\Office\15.0\ClickToRun\Registry\MACHINE\Software\Classes"
!define Office16ClickToRun "Software\Microsoft\Office\16.0\ClickToRun\Registry\MACHINE\Software\Classes"

!define TestReg "Software\Classes"

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
Var ClickToRun

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

Function CheckClickToRun
StrCpy $ClickToRun ""

ReadRegStr $R0 HKLM "${TestReg}\CLSID\${CLSID_IConverterSession}" ""
${If} $R0 != ""
    MessageBox MB_OK "Have CLSID_IConverterSession - $R0"
${EndIf}

ReadRegStr $R1 HKLM "${TestReg}\Wow6432Node\CLSID\${CLSID_IConverterSession}" ""
${If} $R1 != ""
    MessageBox MB_OK "Have Wow6432Node CLSID_IConverterSession - $R0"
${EndIf}

ReadRegStr $R1 HKLM "${Office16ClickToRun}\CLSID\${CLSID_IConverterSession}" ""
${If} $R1 != ""
    ; Have Click To Run Outlook 2016
    StrCpy $ClickToRun "1"
    Goto FoundClickToRun
${EndIf}

ReadRegStr $R2 HKLM "${Office16ClickToRun}\Wow6432Node\CLSID\${CLSID_IConverterSession}" ""
${If} $R2 != ""
    ; Have Click To Run Outlook 2016 32bit on 64bit
    StrCpy $ClickToRun "2"
    Goto FoundClickToRun
${EndIf}

ReadRegStr $R3 HKLM "${Office15ClickToRun}\CLSID\${CLSID_IConverterSession}" ""
${If} $R3 != ""
    ; Have Click To Run Outlook 2013
    StrCpy $ClickToRun "3"
    Goto FoundClickToRun
${EndIf}

ReadRegStr $R4 HKLM "${Office15ClickToRun}\Wow6432Node\CLSID\${CLSID_IConverterSession}" ""
${If} $R4 != ""
    ; Have Click To Run Outlook 2013 32bit on 64bit
    StrCpy $ClickToRun "4"
    Goto FoundClickToRun
${EndIf}

FoundClickToRun:
ReadRegStr $R0 HKLM "Software\Classes\CLSID" "${CLSID_IConverterSession}"
${If} $R0 != ""
    ${If} $ClickToRun == "1"
    ${OrIf} $ClickToRun == "3"
        ; Registry is already patched, so don't need to do anything
        StrCpy $ClickToRun ""
    ${EndIf}
${EndIf}

ReadRegStr $R0 HKLM "Software\Classes\Wow6432Node\CLSID" "${CLSID_IConverterSession}"
${If} $R0 != ""
    ${If} $ClickToRun == "2"
    ${OrIf} $ClickToRun == "4"
        ; Registry is already patched, so don't need to do anything
        StrCpy $ClickToRun ""
    ${EndIf}
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
    ${If} $ClickToRun == "1"
        ; ClickToRun Outlook 2016        
        ${registry::CopyKey} "HKLM\${Office16ClickToRun}\CLSID\${CLSID_IConverterSession}" "HKLM\Software\Classes\CLSID\${CLSID_IConverterSession}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IConverterSession$\" registry value. Aborting"
            Quit
        ${EndIf}
        ${registry::CopyKey} "HKLM\${Office16ClickToRun}\CLSID\${CLSID_IMimeMessage}" "HKLM\Software\Classes\CLSID\${CLSID_IMimeMessage}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IMimeMessage$\" registry value. Aborting"
            Quit
        ${EndIf}
    ${ElseIf} $ClickToRun == "2"
        ; ClickToRun Outlook 2016 32bit on 64 bit
        ${registry::CopyKey} "HKLM\${Office16ClickToRun}\Wow6432Node\CLSID\${CLSID_IConverterSession}" "HKLM\Software\Classes\Wow6432Node\CLSID\${CLSID_IConverterSession}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IConverterSession$\" registry value. Aborting"
            Quit
        ${EndIf}
        ${registry::CopyKey} "HKLM\${Office16ClickToRun}\Wow6432Node\CLSID\${CLSID_IMimeMessage}" "HKLM\Software\Classes\Wow6432Node\CLSID\${CLSID_IMimeMessage}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IMimeMessage$\" registry value. Aborting"
            Quit
        ${EndIf}
    ${ElseIf} $ClickToRun == "3"
        ; ClickToRun Outlook 2013
        ${registry::CopyKey} "HKLM\${Office15ClickToRun}\CLSID\${CLSID_IConverterSession}" "HKLM\Software\Classes\CLSID\${CLSID_IConverterSession}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IConverterSession$\" registry value. Aborting"
            Quit
        ${EndIf}
        ${registry::CopyKey} "HKLM\${Office15ClickToRun}\CLSID\${CLSID_IMimeMessage}" "HKLM\Software\Classes\CLSID\${CLSID_IMimeMessage}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IMimeMessage$\" registry value. Aborting"
            Quit
        ${EndIf}
    ${ElseIf} $ClickToRun == "4"
        ; ClickToRun Outlook 2013 32bit on 64 bit
        ${registry::CopyKey} "HKLM\${Office15ClickToRun}\Wow6432Node\CLSID\${CLSID_IConverterSession}" "HKLM\Software\Classes\Wow6432Node\CLSID\${CLSID_IConverterSession}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IConverterSession$\" registry value. Aborting"
            Quit
        ${EndIf}
        ${registry::CopyKey} "HKLM\${Office15ClickToRun}\Wow6432Node\CLSID\${CLSID_IMimeMessage}" "HKLM\Software\Classes\Wow6432Node\CLSID\${CLSID_IMimeMessage}" $R0
        ${If} $R0 != 0
            MessageBox MB_OK|MB_ICONEXCLAMATION "Failed to Copy $\"CLSID_IMimeMessage$\" registry value. Aborting"
            Quit
        ${EndIf}
    ${EndIf}
    Goto Finished
    No:
    ; Do nothing. PST conversion won't work 
    StrCpy $ClickToRun ""
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
    CreateShortCut "$SMPROGRAMS\${NAME}-${VERSION}\uninstall.lnk" "$InstDir\uninstall.exe" "" "" 0
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
    
    ${If} $R0 == "1"
    ${OrIf} $R0 == "3"
        DeleteRegKey HKLM "Software\Classes\CLSID\${CLSID_IConverterSession}"
        DeleteRegKey HKLM "Software\Classes\CLSID\${CLSID_IMimeMessage}"
    ${Else}
        DeleteRegKey HKLM "Software\Classes\Wow6432Node\CLSID\${CLSID_IConverterSession}"
        DeleteRegKey HKLM "Software\Classes\Wow6432Node\CLSID\${CLSID_IMimeMessage}"    
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