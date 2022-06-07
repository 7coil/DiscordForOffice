; Includes
!include "MUI2.nsh"
!include "logiclib.nsh"

!define NAME "DiscordForOffice"
!define APPFILE "discordforoffice.exe"
!define VERSION "7.0.0"
!define SLUG "${NAME} Version ${VERSION}"

Name "${NAME}"
OutFile "${NAME} Setup.exe"
InstallDir "$PROGRAMFILES\${NAME}"
InstallDirRegKey HKCU "Software\${NAME}" ""
RequestExecutionLevel admin

; !define MUI_ICON "assets\captura.ico"
!define MUI_HEADERIMAGE
!define MUI_WELCOMEFINISHPAGE_BITMAP "assets\MUI_WELCOMEFINISHPAGE_BITMAP.bmp"
!define MUI_HEADERIMAGE_BITMAP "assets\MUI_HEADERIMAGE_BITMAP.bmp"
!define MUI_ABORTWARNING
!define MUI_WELCOMEPAGE_TITLE "${SLUG} Setup"


; Installer pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "../LICENCE"
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!define MUI_FINISHPAGE_TEXT "To complete the installation, restart Microsoft Office."
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

; Set UI language
!insertmacro MUI_LANGUAGE "English"


Section
    DetailPrint "Hello, World!"
SectionEnd