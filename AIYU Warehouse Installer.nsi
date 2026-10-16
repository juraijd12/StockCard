!define APPNAME "AIYU Warehouse"
!define VERSION "1.0.0"
!define PUBLISHER "AIYU"
!define EXE "AIYU Warehouse.exe"

Name "${APPNAME}"
OutFile "dist\AIYU Warehouse Setup.exe"
InstallDir "$LOCALAPPDATA\${APPNAME}"
RequestExecutionLevel user
SetCompressor /SOLID lzma

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Section "${APPNAME}" SecMain
  SetOutPath "$INSTDIR"
  File /r "dist\AIYU Warehouse\*.*"
  SetOutPath "$INSTDIR\AIYU Scanner"
  File /r "dist\AIYU Scanner\*.*"
  SetOutPath "$INSTDIR\AIYU Database Builder"
  File /r "dist\AIYU Database Builder\*.*"
  SetOutPath "$INSTDIR"
  CreateDirectory "$INSTDIR\05_Exports"
  CreateDirectory "$INSTDIR\04_System_Data"
  CreateDirectory "$INSTDIR\TemplateForms"
  CreateDirectory "$INSTDIR\asset"
  CreateShortcut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\${EXE}"
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\${EXE}"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\${APPNAME}.lnk"
  Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
  RMDir "$SMPROGRAMS\${APPNAME}"
  RMDir /r "$INSTDIR"
SectionEnd
