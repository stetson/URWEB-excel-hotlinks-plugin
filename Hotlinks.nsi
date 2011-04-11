; NSIS Excel Add-In Installer Script
; Include
!include MUI.nsh
!include LogicLib.nsh

; General
Name "Hotlinks Plugin"
OutFile "Hotlinks_Setup.exe"
InstallDir "$APPDATA\Microsoft\AddIns"
InstallDirRegKey HKCU "Software\Hotlinks Plugin" "InstallDir" ; Overrides InstallDir


; Languages
!insertmacro MUI_LANGUAGE "English"

; Installer Section
Section "-Install"
SetOutPath "$INSTDIR"

; ADD FILES HERE
File "hotlinks.xla"

; Check Installed Excel Version
ReadRegStr $1 HKCR "Excel.Application\CurVer" ""

${If} $1 == 'Excel.Application.8' ; Excel 95
StrCpy $2 "8.0"
${ElseIf} $1 == 'Excel.Application.9' ; Excel 2000
StrCpy $2 "9.0"
${ElseIf} $1 == 'Excel.Application.10' ; Excel XP
StrCpy $2 "10.0"
${ElseIf} $1 == 'Excel.Application.11' ; Excel 2003
StrCpy $2 "11.0"
${ElseIf} $1 == 'Excel.Application.12' ; Excel 2007
StrCpy $2 "12.0"
${ElseIf} $1 == 'Excel.Application.14' ; Excel 2010
StrCpy $2 "14.0"
${Else}
Abort "An appropriate version of Excel is not installed ($1).$\nNSIS Test setup will be canceled."
${EndIf}

; Find available "OPEN" key
StrCpy $3 ""
loop:
ReadRegStr $4 HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${If} $4 == ""
; Available OPEN key found
${Else}
IntOp $3 $3 + 1
Goto loop
${EndIf}

; Write install data to registry
WriteRegStr HKCU "Software\Hotlinks Plugin" "InstallDir" $INSTDIR
; Install Directory
WriteRegStr HKCU "Software\Hotlinks Plugin" "ExcelCurVer" $2
; Current Excel Version

; Write key to install AddIn in Excel Addin Manager
WriteRegStr HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3" '"$INSTDIR\hotlinks.xla"'

; Write keys to uninstall
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Hotlinks Plugin" "DisplayName" "Hotlinks Plugin"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Hotlinks Plugin" "UninstallString" '"$INSTDIR\hotlinks_uninstall.exe"'
WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Hotlinks Plugin" "NoModify" 1
WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Hotlinks Plugin" "NoRepair" 1

; Create uninstaller
WriteUninstaller "$INSTDIR\hotlinks_uninstall.exe"
SectionEnd

; Uninstaller Section
Section "Uninstall"
; ADD FILES HERE...
Delete "$INSTDIR\hotlinks.xla"
Delete "$INSTDIR\hotlinks_uninstall.exe"

RMDir "$INSTDIR"

; Find AddIn Manager Key and Delete
; AddIn Manager key name and location may have changed since installation depending on actions taken by user in AddIn Manager.
; Need to search for the target AddIn key and delete if found.
ReadRegStr $2 HKCU "Software\Hotlinks Plugin" "ExcelCurVer"
StrCpy $3 ""

loop:
ReadRegStr $4 HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${If} $4 == '"$INSTDIR\hotlinks.xla"'
; Found Key
DeleteRegValue HKCU "Software\Microsoft\Office\$2\Excel\Options" "OPEN$3"
${ElseIf} $4 == ""
; Blank Key Found. Addin is no longer installed in AddIn Manager.
; Need to delete Addin Manager Reference.
DeleteRegValue HKCU "Software\Microsoft\Office\$2\Excel\Add-in Manager" "$INSTDIR\hotlinks.xla"
${Else}
IntOp $3 $3 + 1
Goto loop
${EndIf}

DeleteRegKey HKCU "Software\Hotlinks Plugin"
DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Hotlinks Plugin"
SectionEnd