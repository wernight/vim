!include UpgradeDLL.nsi
!include AddSharedDLL.nsi
!include un.RemoveSharedDLL.nsi

!define MUI_PRODUCT "V.I.M." ;Define your own software name here
!define MUI_VERSION "1.41" ;Define your own software version here

!include "MUI.nsh"

;--------------------------------
;Configuration

	;Do A CRC Check
	CRCCheck On

	;Output File Name
	OutFile "V.I.M.-v1.41-Install.exe"

	;The Default Installation Directory
	InstallDir "$PROGRAMFILES\VIM"

	;Remember install folder
	InstallDirRegKey HKCU "Software\ALC-WBC\${MUI_PRODUCT}" ""

;--------------------------------
;Modern UI Configuration

	!define MUI_WELCOMEPAGE
	!define MUI_LICENSEPAGE
	!define MUI_DIRECTORYPAGE
	!define MUI_FINISHPAGE
	!define MUI_FINISHPAGE_RUN "$INSTDIR\VIM32.EXE"
	
	!define MUI_ABORTWARNING
 
	!define MUI_UNINSTALLER
	!define MUI_UNCONFIRMPAGE

;--------------------------------
;Languages
 
	!insertmacro MUI_LANGUAGE "French"

;--------------------------------
;Data

	;License Data
	LicenseData /LANG=${LANG_FRENCH} "Licence.txt"

;--------------------------------
;Installer Sections

Section "V.I.M."
	;Install Files
	SetOutPath $INSTDIR
	File "..\Content.wri"
	File "..\Historique.wri"
	File "..\Légal.wri"
	File "..\Licence.wri"
	File "..\VIM32.EXE"
	File "${NSISDIR}\Contrib\Icons\modern-uninstall.ico"
	;VRB
	SetOutPath $INSTDIR\VRB
	File "..\VRB\Anglais.vrb"
	File "..\VRB\2° All.vrb"
	File "..\VRB\2° Ang.vrb"
	File "..\VRB\3° Ang StepIn.vrb"
	File "..\VRB\3° Ang.vrb"
	File "..\VRB\4° All.vrb"
	File "..\VRB\4° Ang 1-20.vrb"
	File "..\VRB\4° Ang 20-40.vrb"
	File "..\VRB\4° Ang 40-60.vrb"
	File "..\VRB\4° Ang 60-80.vrb"
	File "..\VRB\4° Ang 80-97.vrb"
	File "..\VRB\4° Ang StepIn.vrb"
	File "..\VRB\5° All.vrb"
	File "..\VRB\Danois - par Delphine Courtial.vrb"

	; Write the uninstall keys for Windows
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\V.I.M." "DisplayName" "V.I.M."
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\V.I.M." "UninstallString" "$INSTDIR\Uninst.exe"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "DisplayIcon" "$INSTDIR\Uninst.exe"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "DisplayVersion" "${MUI_VERSION}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "Publisher" "ALC-WBC"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "HelpLink" "http://www.alc-wbc.com/"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "URLInfoAbout" "http://www.alc-wbc.com/"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "URLUpdateInfo" "http://www.alc-wbc.com/"
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "NoModify" 0x00000001
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${MUI_PRODUCT}" "NoRepair" 0x00000001
	WriteUninstaller "Uninst.exe"
SectionEnd

Section "VB Runtime DLLs"
	!insertmacro UpgradeDLL VBRun60sp5\Asycfilt.dll $SYSDIR\Asycfilt.dll
	!insertmacro UpgradeDLL VBRun60sp5\Comcat.dll $SYSDIR\Comcat.dll
	!insertmacro UpgradeDLL VBRun60sp5\Msvbvm60.dll $SYSDIR\Msvbvm60.dll
	!insertmacro UpgradeDLL VBRun60sp5\Oleaut32.dll $SYSDIR\Oleaut32.dll
	!insertmacro UpgradeDLL VBRun60sp5\Olepro32.dll $SYSDIR\Olepro32.dll
	!define UPGRADEDLL_NOREGISTER
		!insertmacro UpgradeDLL VBRun60sp5\Stdole2.tlb $SYSDIR\Stdole2.tlb
	!undef UPGRADEDLL_NOREGISTER
	!insertmacro UpgradeDLL VBRun60sp5-Plus\VB6FR.dll $SYSDIR\VB6FR.dll
	!insertmacro UpgradeDLL VBRun60sp5-Plus\MSCOMCTL.OCX $SYSDIR\MSCOMCTL.OCX
	!insertmacro UpgradeDLL VBRun60sp5-Plus\COMDLG32.OCX $SYSDIR\COMDLG32.OCX
	!insertmacro UpgradeDLL THREED32.OCX $SYSDIR\THREED32.OCX
	;Skip shared count increasing if already done once for this application
	IfFileExists $INSTDIR\VIM32.exe skipAddShared
		Push $SYSDIR\Asycfilt.dll
		Call AddSharedDLL
		Push $SYSDIR\Comcat.dll
		Call AddSharedDLL
		Push $SYSDIR\Msvbvm60.dll
		Call AddSharedDLL
		Push $SYSDIR\Oleaut32.dll
		Call AddSharedDLL
		Push $SYSDIR\Olepro32.dll
		Call AddSharedDLL
		Push $SYSDIR\Stdole2.tlb
		Call AddSharedDLL
		Push $SYSDIR\VB6FR.dll
		Call AddSharedDLL
		Push $SYSDIR\MSCOMCTL.OCX
		Call AddSharedDLL
		Push $SYSDIR\COMDLG32.OCX
		Call AddSharedDLL
		Push $SYSDIR\THREED32.OCX
		Call AddSharedDLL
	skipAddShared:
SectionEnd

Section "Shortcuts"
	;Add Shortcuts
	CreateDirectory "$SMPROGRAMS\V.I.M."
	CreateShortCut "$SMPROGRAMS\V.I.M.\Verbes Irrégulier Multilingues.lnk" "$INSTDIR\VIM32.EXE" "" "$INSTDIR\VIM32.EXE" 0
	CreateShortCut "$SMPROGRAMS\V.I.M.\Content.lnk" "$INSTDIR\Content.wri" "" "$INSTDIR\Content.wri" 0
	CreateShortCut "$SMPROGRAMS\V.I.M.\Historique.lnk" "$INSTDIR\Historique.wri" "" "$INSTDIR\Historique.wri" 0
	CreateShortCut "$SMPROGRAMS\V.I.M.\Legal.lnk" "$INSTDIR\Légal.wri" "" "$INSTDIR\Légal.wri" 0
	CreateShortCut "$SMPROGRAMS\V.I.M.\Licence.lnk" "$INSTDIR\Licence.wri" "" "$INSTDIR\Licence.wri" 0
	CreateShortCut "$SMPROGRAMS\V.I.M.\Site web.lnk" "http://www.alc-wbc.com/"
	CreateShortCut "$SMPROGRAMS\V.I.M.\Uninstall.lnk" "$INSTDIR\Uninst.exe" "" "$INSTDIR\modern-uninstall.ico" 0
SectionEnd

Section Uninstall
	;Delete Files
	Delete "$INSTDIR\Content.wri"
	Delete "$INSTDIR\Historique.wri"
	Delete "$INSTDIR\Légal.wri"
	Delete "$INSTDIR\Licence.wri"
	Delete "$INSTDIR\VIM32.EXE"
	Delete "$INSTDIR\VRB\Anglais.vrb"
	Delete "$INSTDIR\VRB\2° All.vrb"
	Delete "$INSTDIR\VRB\2° Ang.vrb"
	Delete "$INSTDIR\VRB\3° Ang StepIn.vrb"
	Delete "$INSTDIR\VRB\3° Ang.vrb"
	Delete "$INSTDIR\VRB\4° All.vrb"
	Delete "$INSTDIR\VRB\4° Ang 1-20.vrb"
	Delete "$INSTDIR\VRB\4° Ang 20-40.vrb"
	Delete "$INSTDIR\VRB\4° Ang 40-60.vrb"
	Delete "$INSTDIR\VRB\4° Ang 60-80.vrb"
	Delete "$INSTDIR\VRB\4° Ang 80-97.vrb"
	Delete "$INSTDIR\VRB\4° Ang StepIn.vrb"
	Delete "$INSTDIR\VRB\5° All.vrb"
	Delete "$INSTDIR\VRB\Danois - par Delphine Courtial.vrb"
	Delete "$INSTDIR\modern-uninstall.ico"

	;Delete VB DLLs
	Push $SYSDIR\Asycfilt.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Comcat.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Msvbvm60.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Oleaut32.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Olepro32.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\Stdole2.tlb
	Call un.RemoveSharedDLL
	Push $SYSDIR\VB6FR.dll
	Call un.RemoveSharedDLL
	Push $SYSDIR\MSCOMCTL.OCX
	Call un.RemoveSharedDLL
	Push $SYSDIR\COMDLG32.OCX
	Call un.RemoveSharedDLL
	Push $SYSDIR\THREED32.OCX
	Call un.RemoveSharedDLL

	;Delete Start Menu Shortcuts
	Delete "$SMPROGRAMS\V.I.M.\*.*"
	RmDir "$SMPROGRAMS\V.I.M."

	;Delete Uninstaller And Unistall Registry Entries
	Delete "$INSTDIR\Uninst.exe"
	DeleteRegKey HKEY_LOCAL_MACHINE "SOFTWARE\V.I.M."
	DeleteRegKey HKEY_LOCAL_MACHINE "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\V.I.M."
	RMDir "$INSTDIR\VRB"
	RMDir "$INSTDIR"

	;Delete saved install path
	DeleteRegKey /ifempty HKCU "Software\ALC-WBC\${MUI_PRODUCT}"

	;Display the Finish header
	!insertmacro MUI_UNFINISHHEADER
SectionEnd