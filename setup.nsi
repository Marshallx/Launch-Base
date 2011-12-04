;Var RA2
;Var TFD

SetCompressor /SOLID lzma
Name "Launch Base"
Caption "Launch Base 0.99.270"
SubCaption 0 " "
SubCaption 1 ": Shortcut Options"
BrandingText "http://marshall.strategy-x.com"
CRCCheck on
Icon "Resource\lb_icon.ico"
WindowIcon on
;CheckBitmap "check.bmp"
InstallColors FFFF99 000000
ShowInstDetails show
InstProgressFlags smooth colored
OutFile "LB_Setup.exe"
AutoCloseWindow true
InstallDirRegKey HKLM "SOFTWARE\Westwood\Red Alert 2" "InstallPath"
ComponentText "Welcome to the Launch Base setup program!$\nPlease select which shortcuts you would like."
UninstallText "WARNING! By uninstalling Launch Base you will also permanently delete any mods, plugins or tools that are installed in Launch Base, including any user-generated content. It is strongly recommended that you uninstall each mod via Launch Base before uninstalling Launch Base itself."
UninstallIcon "Resource\lb_uninstall.ico"

;PageEx license
;LicenseText "Launch Base: Information" "Continue"
;LicenseData lbinfo.txt
;PageExEnd

Page components

PageEx directory
DirText "Please select the directory where you would like to install Launch Base. Please note that all Launch Base mods, plugins and tools must later be installed here too, so be sure to choose a drive that has sufficient space."
PageExEnd

Page instfiles

Function PrereqCheck
    System::Call 'kernel32::CreateMutexA(i 0, i 0, t "YRLBMUTEXERM1") i .r1 ?e'
    Pop $R0
    StrCmp $R0 0 +3
    MessageBox MB_OK|MB_ICONSTOP "Either Launch Base itself or another installer is already running."
    Abort
FunctionEnd

Function .onInit
    Call PrereqCheck
    ReadRegStr $R1 HKLM "SOFTWARE\Marshallx Industries\YR Launch Base" "InstallPath"
    IfFileExists "$R1\LaunchBase.exe" 0 +3
    StrCpy $INSTDIR "$R1"
    Goto EndInit
    ReadRegStr $R1 HKLM "SOFTWARE\Westwood\Red Alert 2" "InstallPath"
    StrCmp $R1 "" +4 0
    IfFileExists $R1 0 +4
    StrCpy $R1 $R1 -8
    StrCpy $INSTDIR "$R1\LaunchBase"
    Goto EndInit
    ReadRegStr $R1 HKLM "SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade" "r2_folder"
    StrCmp $R1 "" +5 0
    ReadRegStr $R2 HKLM "SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade" "r2_executable"
    IfFileExists "$R1\$R2" 0
    StrCpy $INSTDIR "$R1\LaunchBase"
    Goto EndInit
    StrCpy $INSTDIR "C:\Program Files\LaunchBase"
    EndInit:
FunctionEnd

Section "-main" 0
SetOutPath $INSTDIR
File "comdlg32.ocx"
File "LaunchBase.exe"
File "md5.dll"
File "mscomctl.ocx"
File "msinet.ocx"
;File "scrrun.dll"
File "splash.bmp"
IfFileExists "$INSTDIR\LaunchBase.ini" +2 0
File /oname=$INSTDIR\LaunchBase.ini Setup1.ini
WriteINIStr "$INSTDIR\LaunchBase.ini" URL 0 "Launch Base Website,http://marshall.strategy-x.com/LaunchBase"
WriteINIStr "$INSTDIR\LaunchBase.ini" URL 1 "Renegade Projects' Forum,http://forums.renegadeprojects.com"
SetOutPath "$INSTDIR\Help"
File "Help\*.txt"
SetOutPath "$INSTDIR\Resource"
File "Resource\btna0.bmp"
File "Resource\btna1.bmp"
File "Resource\btna2.bmp"
File "Resource\btna3.bmp"
File "Resource\btnb0r.bmp"
File "Resource\btnb0y.bmp"
File "Resource\btnb1.bmp"
File "Resource\btnb2.bmp"
File "Resource\btnb3.bmp"
File "Resource\btnb4.bmp"
File "Resource\ea_wwlogo.bik"
File "Resource\eva_dl.wav"
File "Resource\fabanner.bmp"
File "Resource\flac.exe"
File "Resource\gunzip.exe"
File "Resource\ipb_icon.ico"
File "Resource\md5deep.exe"
File "Resource\nobanner.bmp"
File "Resource\oggdec.exe"
File "Resource\ra2banner.bmp"
File "Resource\tar.exe"
File "Resource\yrbanner.bmp"
File "Resource\yrpm.csf"
SetOutPath "$INSTDIR\Skins"
SetOutPath "$INSTDIR\Skins\ren_glass"
File "Skins\ren_glass\btna0.bmp"
File "Skins\ren_glass\btna1.bmp"
File "Skins\ren_glass\btna2.bmp"
File "Skins\ren_glass\btna3.bmp"
File "Skins\ren_glass\btnb0r.bmp"
File "Skins\ren_glass\btnb0y.bmp"
File "Skins\ren_glass\btnb1.bmp"
File "Skins\ren_glass\btnb2.bmp"
File "Skins\ren_glass\btnb3.bmp"
File "Skins\ren_glass\btnb4.bmp"
File "Skins\ren_glass\skin.ini"
File "Skins\ren_glass\tab0.bmp"
File "Skins\ren_glass\tab1.bmp"
File "Skins\ren_glass\tab2.bmp"
File "Skins\ren_glass\tab3.bmp"
File "Skins\ren_glass\tab4.bmp"
WriteRegStr HKLM "SOFTWARE\Marshallx Industries\YR Launch Base" "InstallPath" "$INSTDIR"
ReadRegDWORD $0 HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings" "SyncMode5"
StrCmp $0 "3" +2 0
WriteRegDWORD HKCU "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings" "SyncMode5" 0x00000003
Delete "$INSTDIR\ModCat.lbd"
;Uninstaller
WriteUninstaller "$INSTDIR\Uninstall.exe"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "DisplayName" "Command &&& Conquer Red Alert 2 - Yuri's Revenge - Launch Base"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "UninstallString" "$INSTDIR\Uninstall.exe"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "DisplayIcon" "$INSTDIR\LaunchBase.exe"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "Publisher" "Marshallx Industries"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "HelpLink" "http://marshall.cannis.net/LaunchBase"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "URLUpdateInfo" "http://marshall.strategy-x.com/LaunchBase"
WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "URLInfoAbout" "http://marshall.strategy-x.com/LaunchBase"
WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "NoModify" 1
WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base" "NoRepair" 1
SectionEnd

Section "Start Menu Shortcut" 1
CreateShortCut "$SMPROGRAMS\Launch Base.lnk" "$INSTDIR\LaunchBase.exe" "" "$INSTDIR\LaunchBase.exe" 0
SectionEnd

Section "Desktop Shortcut" 2
CreateShortCut "$DESKTOP\Launch Base.lnk" "$INSTDIR\LaunchBase.exe" "" "$INSTDIR\LaunchBase.exe" 0
SectionEnd

Section "Uninstall"
    ReadRegStr $INSTDIR HKLM "SOFTWARE\Marshallx Industries\YR Launch Base" "InstallPath"
    IfFileExists "$INSTDIR\LaunchBase.exe" 0 UninstNoLB
    ReadINIStr $R1 "$INSTDIR\LaunchBase.ini" "Mod" "Name"
    StrCmp $R1 "" 0 UninstModActive
    ReadINIStr $R1 "$INSTDIR\LaunchBase.ini" "ActivePlugins" "0"
    StrCmp $R1 "" 0 UninstModActive
;    RMDIr /r $INSTDIR
    Push "$INSTDIR" ; File to delete, supports wildcards
    Call un.SendToRecycleBin
    Pop $R0 ; Return code
    DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Launch Base"
    DeleteRegKey HKLM "SOFTWARE\Marshallx Industries\YR Launch Base"
    MessageBox MB_OK|MB_ICONEXCLAMATION "Launch Base uninstallation complete."
    Goto UninstEnd
UninstModActive:
    MessageBox MB_OK|MB_ICONEXCLAMATION "There are one or more mod files still active in the game.$\nPlease run Launch Base and deactivate any mods or plugins before uninstalling the program." IDOK UninstEnd
UninstNoLB:
    MessageBox MB_OK|MB_ICONEXCLAMATION "Launch Base install path could not be found - no files have been removed." IDOK UninstEnd
UninstEnd:
SectionEnd

Function .onInstSuccess
MessageBox MB_YESNO|MB_ICONINFORMATION "Launch Base has been successfully installed.$\nWould you like to run Launch Base now?" IDNO +2 IDYES 0
Exec "$INSTDIR\LaunchBase.exe"
FunctionEnd

!ifndef FO_DELETE
!define FO_DELETE 0x3
!endif
!ifndef FOF_SILENT
!define FOF_SILENT 0x4
!endif
!ifndef FOF_NOCONFIRMATION
!define FOF_NOCONFIRMATION 0x10
!endif
!ifndef FOF_ALLOWUNDO
!define FOF_ALLOWUNDO 0x40
!endif
 
Function un.SendToRecycleBin
Exch $R0
Push $R1
Push $R2
 
System::Alloc 28
Pop $R1
System::Call "*$R1(i $HWNDPARENT, i ${FO_DELETE}, t '$R0', t '', i ${FOF_ALLOWUNDO}|${FOF_SILENT}|${FOF_NOCONFIRMATION}, i 0, i 0, t '')"
System::Call "shell32::SHFileOperationA(i R1)i.R2"
System::Free $R1
StrCpy $R0 $R2
Pop $R2
Pop $R1
Exch $R0
FunctionEnd

; eof
