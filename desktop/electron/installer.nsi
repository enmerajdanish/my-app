!include "MUI2.nsh"
!include "x64.nsh"
Name "${PRODUCT_NAME}"
InstallDir "$LOCALAPPDATA\Programs\Hamrah Mohasel"
InstallDirRegKey HKCU "Software\HamrahMohasel" "InstallDir"
RequestExecutionLevel user
ShowInstDetails show
ShowUninstDetails show
!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\Hamrah Mohasel.exe"
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "Farsi"
Function .onInit
 ${IfNot} ${RunningX64}
  MessageBox MB_ICONSTOP "این برنامه به ویندوز ۶۴ بیتی نیاز دارد."
  Abort
 ${EndIf}
FunctionEnd
Function .onVerifyInstDir
 IfFileExists "$INSTDIR\Hamrah Mohasel.exe" allowed
 IfFileExists "$INSTDIR\*.*" 0 allowed
 Abort
 allowed:
FunctionEnd
Section "Install"
 SetShellVarContext current
 InitPluginsDir
 SetOutPath "$PLUGINSDIR"
 SetCompress off
 File /oname=app.7z "${APP_64}"
 SetCompress auto
 SetOutPath "$INSTDIR"
 Nsis7z::Extract "$PLUGINSDIR\app.7z"
 IfFileExists "$INSTDIR\Hamrah Mohasel.exe" +3 0
 MessageBox MB_ICONSTOP "استخراج فایل‌های برنامه موفق نشد."
 Abort
 WriteUninstaller "$INSTDIR\Uninstall.exe"
 CreateShortcut "$DESKTOP\Hamrah Mohasel.lnk" "$INSTDIR\Hamrah Mohasel.exe"
 CreateDirectory "$SMPROGRAMS\Hamrah Mohasel"
 CreateShortcut "$SMPROGRAMS\Hamrah Mohasel\Hamrah Mohasel.lnk" "$INSTDIR\Hamrah Mohasel.exe"
 CreateShortcut "$SMPROGRAMS\Hamrah Mohasel\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
 WriteRegStr HKCU "Software\HamrahMohasel" "InstallDir" "$INSTDIR"
 WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HamrahMohasel" "DisplayName" "${PRODUCT_NAME}"
 WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HamrahMohasel" "DisplayVersion" "${VERSION}"
 WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HamrahMohasel" "UninstallString" '$\"$INSTDIR\Uninstall.exe$\"'
 WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HamrahMohasel" "DisplayIcon" "$INSTDIR\Hamrah Mohasel.exe"
 WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HamrahMohasel" "NoModify" 1
 WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HamrahMohasel" "NoRepair" 1
SectionEnd
Section "Uninstall"
 SetShellVarContext current
 Delete "$DESKTOP\Hamrah Mohasel.lnk"
 RMDir /r "$SMPROGRAMS\Hamrah Mohasel"
 DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HamrahMohasel"
 DeleteRegKey HKCU "Software\HamrahMohasel"
 RMDir /r "$INSTDIR"
SectionEnd
