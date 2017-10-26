#Region ### START Library section ###

#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <GUICtrlPic.au3>
#include <GuiButton.au3>
#include <Misc.au3>
#include <File.au3>
#include <Process.au3>
#include <String.au3>
#include <GUIConstants.au3>
#include <ProgressConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <Clipboard.au3>
#include <AutoItConstants.au3>
#include <BlockInputEx.au3>
#include <Date.au3>
#include <OOoCalc.au3>
#include <OOoCalcConstants.au3>

#EndRegion ### START Library section ###

#Region ### START Variables section ###

Global _
		$sMsg, $iError, $iExtended, $g_idError_Message, $oIE, $iPID, $iSavePosStartingServer = 0, $progressbarLoadingScreen, _
		$gui_main, $gui_webbrowser, $gui_loading_screen, _
		$btn_tabula, $btn_pdfbearbeiten, $btn_mendeley, $btn_exit_main, $btn_openoffice, $btn_wizard, $btn_home, $btn_exit_embedded, _
		$path_mainDir, $path_iniFile, $path_configDir, $path_installFilesDir, $path_pdfextractorInternalDir, $path_mendeleyExe = "\PDF_Extractor_InstallFiles\Mendeley\Mendeley Desktop\MendeleyDesktop.exe", _
		$path_MendeleyPDFData, $exe_pdfbearbeiten = "\PDF_Extractor_InstallFiles\PDFBearbeiten\pdfbearbeiten.exe"


#EndRegion ### START Variables section ###

#Region ### START System parameters section ###

Opt("WinTitleMatchMode", 2)
Opt("GUIOnEventMode", 1)

#EndRegion ### START System parameters section ###

#Region ### START Running section ###

_Start_Loading_Screen()
_Enter_Ini_Details()
_Start_GUI_Main()

#EndRegion ### START Running section ###

#Region ### START Functions section ###

Func _Start_GUI_Main()

	#cs - - - Path of zipped installing files - - -
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\7za.exe"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\Mendeley.7z"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\Tabula-Win-1.1.1.7z"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\OpenOffice.7z"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\PDFBearbeiten.7z"
	#ce - - - Path of zipped installing files - - -

	Local _
			$path_mendeleyPic = $path_installFilesDir & "\ICON_Mendeley.png", _
			$path_nitropdfPic = $path_installFilesDir & "\ICON_PDF.png", _
			$path_tabulaPic = $path_installFilesDir & "\LOGO_Tabula.png", _
			$path_ubaPic = $path_installFilesDir & "\LOGO_UBA.png", _
			$path_wizardPic = $path_installFilesDir & "\ICON_Wizard.png", _
			$path_openofficescalcPic = $path_installFilesDir & "\ICON_OpenOffice.png"

	If Not FileExists($path_mainDir & $path_mendeleyPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\Documents\Pictures\ICON_Mendeley.png", $path_mainDir & $path_mendeleyPic)

	EndIf

	If Not FileExists($path_mainDir & $path_nitropdfPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\Documents\Pictures\ICON_PDF.png", $path_mainDir & $path_nitropdfPic)

	EndIf

	If Not FileExists($path_mainDir & $path_tabulaPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\Documents\Pictures\LOGO_Tabula.png", $path_mainDir & $path_tabulaPic)

	EndIf

	If Not FileExists($path_mainDir & $path_ubaPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\Documents\Pictures\LOGO_UBA.png", $path_mainDir & $path_ubaPic)

	EndIf

	If Not FileExists($path_mainDir & $path_wizardPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\Documents\Pictures\ICON_Wizard.png", $path_mainDir & $path_wizardPic)

	EndIf

	If Not FileExists($path_mainDir & $path_openofficescalcPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\Documents\Pictures\ICON_OpenOffice.png", $path_mainDir & $path_openofficescalcPic)

	EndIf

	$gui_main = GUICreate("EDV_Ecotox_Database", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close") ;

	$lb_header = GUICtrlCreateLabel("Software for extracting and converting tables from a PDF", 10, 10, 700, 30)
	GUICtrlSetFont(-1, 16, 600, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0x000000)

	GUISetState(@SW_SHOW)

	_GUICtrlPic_Create($path_mainDir & $path_ubaPic, @DesktopWidth / 1.2, @DesktopHeight / 1201, @DesktopWidth / 6, @DesktopHeight / 7)
	_GUICtrlPic_Create($path_mainDir & $path_mendeleyPic, @DesktopWidth / 20, @DesktopHeight / 10, @DesktopWidth / 5, @DesktopHeight / 3, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)
	_GUICtrlPic_Create($path_mainDir & $path_nitropdfPic, @DesktopWidth / 3, @DesktopHeight / 10, @DesktopWidth / 5, @DesktopHeight / 3, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)
	_GUICtrlPic_Create($path_mainDir & $path_tabulaPic, @DesktopWidth / 20, @DesktopHeight / 1.85, @DesktopWidth / 5, @DesktopHeight / 3, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)
	_GUICtrlPic_Create($path_mainDir & $path_openofficescalcPic, @DesktopWidth / 3, @DesktopHeight / 1.8, @DesktopWidth / 5, @DesktopHeight / 3, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)

	$btn_tabula = GUICtrlCreateButton("Tabula", (@DesktopWidth / 20), (@DesktopHeight / 1.124), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btn_pdfbearbeiten = GUICtrlCreateButton("PDF-Editor", (@DesktopWidth / 3), (@DesktopHeight / 2.3), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btn_mendeley = GUICtrlCreateButton("Mendeley", (@DesktopWidth / 20), (@DesktopHeight / 2.3), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btn_wizard = GUICtrlCreateButton("Assisstent", (@DesktopWidth / 1.2), (@DesktopHeight / 6), (@DesktopWidth / 6), (@DesktopHeight / 25), $BS_ICON)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btn_openoffice = GUICtrlCreateButton("OpenOffice", (@DesktopWidth / 3), (@DesktopHeight / 1.124), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btn_exit_main = GUICtrlCreateButton("Exit", (@DesktopWidth / 1.113), (@DesktopHeight / 1.075), (@DesktopWidth / 10), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	While 1
		Sleep(10)
	WEnd

EndFunc   ;==>_Start_GUI_Main

Func _Start_Embedded_Browser()

	$oIE = _IECreateEmbedded()
	$gui_webbrowser = GUICreate("Embedded Web-Browser", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")
	GUICtrlCreateObj($oIE, 10, 40, (@DesktopWidth - 20), (@DesktopHeight - 50))
	$btn_home = GUICtrlCreateButton("Home", 10, 5, 100, 30)
	GUICtrlSetOnEvent(-1, "_On_Button")
	$btn_exit_embedded = GUICtrlCreateButton("Exit", 1810, 5, 100, 30)
	GUICtrlSetOnEvent(-1, "_On_Button")
	$g_idError_Message = GUICtrlCreateLabel("", 100, 500, 500, 30)
	GUICtrlSetColor(-1, 0xff0000)
	GUISetState(@SW_SHOW)
	_IENavigate($oIE, "http://127.0.0.1:8080")
	_IEAction($oIE, "stop")
	_IELinkClickByText($oIE, "My Files")

	Return $oIE

EndFunc   ;==>_Start_Embedded_Browser

Func _Start_Loading_Screen()

	;_BlockinputEx(1)

	$path_configDir = @LocalAppDataDir & "\PDF_Extractor_Config"
	$path_installFilesDir = "\PDF_Extractor_InstallFiles"


	If (_Get_Ini(_FileListToArray($path_configDir, Default, Default, True)) = False) Then

		If Not FileExists($path_configDir) Then

			;_BlockinputEx(0)

			DirCreate($path_configDir)
			$path_iniFile = _WinAPI_GetTempFileName($path_configDir, "CFG")
			$path_mainDir = FileSelectFolder("Choose the installation directory", "C:\", $WS_POPUP)
			IniWrite($path_iniFile, " - - INFORMATION - - ", "EXPLANATION -", "- This file contains parameters for the App 'PDF_Extractor'")
			IniWrite($path_iniFile, "Paths", "Main_Directory", $path_mainDir)
			IniWrite($path_iniFile, "Trigger", "Main_Directory", "1")

			;_BlockinputEx(1)

		Else

			;_BlockinputEx(0)

			$path_iniFile = _WinAPI_GetTempFileName($path_configDir, "CFG")
			$path_mainDir = FileSelectFolder("Choose the installation directory", "C:\", $WS_POPUP)
			IniWrite($path_iniFile, " - - INFORMATION - - ", "EXPLANATION -", "- This file contains parameters for the App 'PDF_Extractor'")
			IniWrite($path_iniFile, "Paths", "Main_Directory", $path_mainDir)
			IniWrite($path_iniFile, "Trigger", "Main_Directory", "1")

			;_BlockinputEx(1)

		EndIf

	EndIf

	$path_iniFile = _Get_Ini(_FileListToArray($path_configDir, Default, Default, True))
	;_BlockinputEx(0)
	$path_mainDir = IniRead($path_iniFile, "Paths", "Main_Directory", "Can't read key 'Main_Directory' from section 'Paths' in ini-file.")
	;_BlockinputEx(1)

	;==> Values for altering the progressbar are triggered in '_Start_File_Install()' GUICtrlSetData($progressbarLoadingScreen, $i)
	$gui_loading_screen = GUICreate("Starting program ...", 300, 40, 100, 200)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")
	$progressbarLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
	GUICtrlSetColor(-1, 32250) ; not working with Windows XP Style
	GUISetState(@SW_SHOW)
	;==> Values for altering the progressbar are triggered in '_Start_File_Install()' with GUICtrlSetData($progressbarLoadingScreen, $i)

	_Start_File_Install()

	GUIDelete($gui_loading_screen)
	_Process_Close_Tree($iPID)

	;_BlockinputEx(0)

EndFunc   ;==>_Start_Loading_Screen

Func _Start_File_Install()

	Local _
			$path_7zipExe = $path_installFilesDir & "\7za.exe", _
			$path_mendeleyDir = $path_installFilesDir & "\Mendeley", _
			$path_tabulaDir = $path_installFilesDir & "\Tabula-Win-1.1.1", _
			$path_openofficeDir = $path_installFilesDir & "\OpenOffice", _
			$path_pdfbearbeitenDir = $path_installFilesDir & "\PDFBearbeiten", _
			$path_mendeleyZipped = $path_installFilesDir & "\Mendeley.7z", _
			$path_tabulaZipped = $path_installFilesDir & "\Tabula-Win-1.1.1.7z", _
			$path_openofficeZipped = $path_installFilesDir & "\OpenOffice.7z", _
			$path_pdfbearbeitenZipped = $path_installFilesDir & "\PDFBearbeiten.7z", _
			$path_pdfextractorDir = $path_mainDir & "\PDF_Extractor"

	$path_pdfextractorInternalDir = $path_pdfextractorDir & "\Internal"

	#Region ### START Directory in personal documents (default path of installing files) ###

	If Not FileExists($path_mainDir & $path_installFilesDir) Then

		DirCreate($path_mainDir & $path_installFilesDir)

	EndIf

	GUICtrlSetData($progressbarLoadingScreen, 5)
	Sleep(100)

	If Not FileExists($path_pdfextractorDir) Then

		DirCreate($path_pdfextractorDir)

	EndIf

	If Not FileExists($path_pdfextractorInternalDir) Then

		DirCreate($path_pdfextractorInternalDir)

	EndIf

	GUICtrlSetData($progressbarLoadingScreen, 10)
	Sleep(100)

	#EndRegion ### START Directory in personal documents (default path of installing files) ###

	#Region ### START Zip installing files (7zip) and set up saving directory ###

	#cs - - - Path of zipped installing files - - -
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\7za.exe"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\Mendeley.7z"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\Tabula-Win-1.1.1.7z"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\OpenOffice.7z"
		"P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\PDFBearbeiten.7z"
	#ce - - - Path of zipped installing files - - -

	If Not FileExists($path_mainDir & $path_7zipExe) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\7za.exe", $path_mainDir & $path_7zipExe)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 15)
	Sleep(100)

	If Not FileExists($path_mainDir & $path_mendeleyDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\Mendeley.7z", $path_mainDir & $path_mendeleyZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 20)
	Sleep(100)

	If Not FileExists($path_mainDir & $path_tabulaDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\Tabula-Win-1.1.1.7z", $path_mainDir & $path_tabulaZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 25)
	Sleep(100)

	If Not FileExists($path_mainDir & $path_openofficeDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\OpenOffice.7z", $path_mainDir & $path_openofficeZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 30)
	Sleep(100)

	If Not FileExists($path_mainDir & $path_pdfbearbeitenDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\PDF_To_Database\External_Software\PDFBearbeiten.7z", $path_mainDir & $path_pdfbearbeitenZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 35)
	Sleep(100)

	#EndRegion ### START Zip installing files (7zip) and set up saving directory ###

	#Region ### START Unzip installing files with portable 7zip ###

	If Not FileExists($path_mainDir & $path_mendeleyDir) Then

		RunWait($path_mainDir & $path_7zipExe & ' x ' & $path_mainDir & $path_mendeleyZipped & ' -o' & $path_mainDir & $path_installFilesDir, "", @SW_HIDE)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 50)
	Sleep(100)

	If Not FileExists($path_mainDir & $path_tabulaDir) Then

		RunWait($path_mainDir & $path_7zipExe & ' x ' & $path_mainDir & $path_tabulaZipped & ' -o' & $path_mainDir & $path_installFilesDir, "", @SW_HIDE)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 60)
	Sleep(100)

	If Not FileExists($path_mainDir & $path_openofficeDir) Then

		RunWait($path_mainDir & $path_7zipExe & ' x ' & $path_mainDir & $path_openofficeZipped & ' -o' & $path_mainDir & $path_installFilesDir, "", @SW_HIDE)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 70)
	Sleep(100)

	If Not FileExists($path_mainDir & $path_pdfbearbeitenDir) Then

		RunWait($path_mainDir & $path_7zipExe & ' x ' & $path_mainDir & $path_pdfbearbeitenZipped & ' -o' & $path_mainDir & $path_installFilesDir, "", @SW_HIDE)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 80)
	Sleep(100)

	#EndRegion ### START Unzip installing files with portable 7zip ###

	#Region ### START Delete zipped installing files ###

	If FileExists($path_mainDir & $path_7zipExe) Then

		FileDelete($path_mainDir & $path_7zipExe)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 85)
	Sleep(100)

	If FileExists($path_mainDir & $path_mendeleyZipped) Then

		FileDelete($path_mainDir & $path_mendeleyZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 90)
	Sleep(100)

	If FileExists($path_mainDir & $path_openofficeZipped) Then

		FileDelete($path_mainDir & $path_openofficeZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 95)
	Sleep(100)

	If FileExists($path_mainDir & $path_pdfbearbeitenZipped) Then

		FileDelete($path_mainDir & $path_pdfbearbeitenZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 97)
	Sleep(100)

	If FileExists($path_mainDir & $path_tabulaZipped) Then

		FileDelete($path_mainDir & $path_tabulaZipped)

	EndIf
	GUICtrlSetData($progressbarLoadingScreen, 100)
	Sleep(100)

	#EndRegion ### START Delete zipped installing files ###

EndFunc   ;==>_Start_File_Install

Func _Start_Mendeley_with_AutoImport()

	;_BlockinputEx(1)

	Run($path_mainDir & $path_mendeleyExe, "", @SW_SHOW)

	$path_MendeleyBackup = IniRead($path_iniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file.")
	_ClipBoard_SetData($path_MendeleyBackup)
	WinWaitActive("Mendeley Desktop")
	Sleep(3000)
	Send("{ALT}")
	Sleep(200)
	Send("{LEFT}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{RIGHT}")
	Sleep(200)
	Send("{ENTER}")
	Sleep(200)
	Send("{UP}")
	Sleep(200)
	Send("{ENTER}")
	Sleep(1000)
	Send("^v")
	Sleep(500)
	Send("{ENTER}")
	WinWaitActive("Restore Backup")
	Sleep(500)
	Send("{TAB}{SPACE}")
	Sleep(500)
	Send("{TAB}{TAB}{SPACE}")
	Sleep(500)
	Send("{ENTER}")
	WinWaitActive("Welcome to Mendeley Desktop")
	Sleep(500)
	Send("{ENTER}")

	WinWaitActive("Mendeley Desktop")
	Sleep(3000)
	$path_MendeleyPDFData = IniRead($path_iniFile, "Paths", "Mendeley_PDFData", "Can't read key 'Mendeley_PDFData' from section 'Paths' in ini-file.")
	_ClipBoard_SetData($path_MendeleyPDFData)
	Send("{ALT}")
	Sleep(200)
	Send("{LEFT}")
	Sleep(200)
	Send("{LEFT}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{UP}")
	Sleep(200)
	Send("{ENTER}")
	Sleep(500)
	Send("{RIGHT}")
	Sleep(200)
	Send("{RIGHT}")
	Sleep(200)
	Send("{TAB}")
	Sleep(200)
	Send("{SPACE}")
	Sleep(500)
	Send("{TAB}")
	Sleep(200)
	Send("^v")
	Sleep(200)
	Send("{ENTER}")
	Sleep(500)
	Send("{LEFT}")
	Sleep(200)
	Send("{ENTER}")

	;_BlockinputEx(0)

EndFunc   ;==>_Start_Mendeley_with_AutoImport

Func _Start_Mendeley_Create_Backup()

	;_BlockinputEx(1)

	Run($path_mainDir & $path_mendeleyExe, "", @SW_SHOW)

	$path_MendeleyBackup = IniRead($path_iniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file.")
	$path_MendeleyBackupArchive = IniRead($path_iniFile, "Paths", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Paths' in ini-file.")

	_ClipBoard_SetData($path_MendeleyBackup)

	$time = _Date_Time_SystemTimeToFileTime(_Date_Time_GetSystemTime())
	$time = _Date_Time_FileTimeToStr($time)
	$time = StringReplace($time, ":", "_")
	$time = StringReplace($time, "/", "_")
	$time = StringReplace($time, " ", "_")

	FileMove($path_MendeleyBackup, $path_MendeleyBackupArchive & "\" & "Archive_" & $time & ".zip", 1)
	Sleep(1000)

	WinWaitActive("Mendeley Desktop")
	Sleep(3000)
	Send("{ALT}")
	Sleep(200)
	Send("{LEFT}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{DOWN}")
	Sleep(200)
	Send("{ENTER}")
	Sleep(500)
	Send("^v")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Send("{ESC}")
	Sleep(500)
	Send("!{F4}")

	;_BlockinputEx(0)

EndFunc   ;==>_Start_Mendeley_Create_Backup

Func _Process_Close_Tree($sPID)

	If IsString($sPID) Then $sPID = ProcessExists($sPID)
	If Not $sPID Then Return SetError(1, 0, 0)

	Return Run(@ComSpec & " /c taskkill /F /PID " & $sPID & " /T", @SystemDir, @SW_HIDE)

EndFunc   ;==>_Process_Close_Tree

Func _Embed_External_App()
	$hGUI = GUICreate("Test", 800, 600, -1, -1, BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU, $WS_CLIPCHILDREN))
	$PID = Run("C:\Windows\System32\cmd.exe", "", @SW_HIDE)
	$hWnd = 0
	$stPID = DllStructCreate("int")
	Do
		$WinList = WinList()
		For $i = 1 To $WinList[0][0]
			If $WinList[$i][0] <> "" Then
				DllCall("user32.dll", "int", "GetWindowThreadProcessId", "hwnd", $WinList[$i][1], "ptr", DllStructGetPtr($stPID))
				If DllStructGetData($stPID, 1) = $PID Then
					$hWnd = $WinList[$i][1]
					ExitLoop
				EndIf
			EndIf
		Next
		Sleep(100)
	Until $hWnd <> 0
	$stPID = 0
	If $hWnd <> 0 Then
		$nExStyle = DllCall("user32.dll", "int", "GetWindowLong", "hwnd", $hWnd, "int", -20)
		$nExStyle = $nExStyle[0]
		DllCall("user32.dll", "int", "SetWindowLong", "hwnd", $hWnd, "int", -20, "int", BitOR($nExStyle, $WS_EX_MDICHILD))
		DllCall("user32.dll", "int", "SetParent", "hwnd", $hWnd, "hwnd", $hGUI)
		WinSetState($hWnd, "", @SW_SHOW)
		WinMove($hWnd, "", 0, 0, 600, 400)
	EndIf
	GUISetState()
	While 1
		$msg = GUIGetMsg()
		If $msg = -3 Then ExitLoop
	WEnd
EndFunc   ;==>_Embed_External_App

Func _Enter_Ini_Details()

	If (IniRead($path_iniFile, "Trigger", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Trigger' in ini-file.") = "1") And _
			(IniRead($path_iniFile, "Trigger", "Mendeley_PDFData", "Can't read key 'Mendeley_PDFData' from section 'Trigger' in ini-file.") = "1") And _
			(IniRead($path_iniFile, "Trigger", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Trigger' in ini-file.") = "1") Then

		MsgBox($MB_SYSTEMMODAL, "Mendeley backup file", "The current used Mendeley backup file is stored at: " & IniRead($path_iniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file."))
		MsgBox($MB_SYSTEMMODAL, "Mendeley backup archive folder", "The current used Mendeley backup archive folder is stored at: " & IniRead($path_iniFile, "Paths", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Paths' in ini-file."))
		MsgBox($MB_SYSTEMMODAL, "Mendeley PDFA data folder", "The current used Mendeley PDF data folder is stored at: " & IniRead($path_iniFile, "Paths", "Mendeley_PDFData", "Can't read key 'Mendeley_PDFdata' from section 'Paths' in ini-file."))

	Else

		Sleep(100)
		IniWrite($path_iniFile, "Paths", "Mendeley_Backup", FileOpenDialog("Open the mendeley backup file", "\\gruppende\IV2.2\Int\WRMG\PDF_Extract_Files\", "All (*.zip)"))
		IniWrite($path_iniFile, "Paths", "Mendeley_Backup_Archive", FileSelectFolder("Select the mendeley backup archive folder ", "\\gruppende\IV2.2\Int\WRMG\PDF_Extract_Files\"))
		IniWrite($path_iniFile, "Paths", "Mendeley_PDFData", FileSelectFolder("Select the PDF data folder", "\\gruppende\IV2.2\Int\WRMG\PDF_Extract_Files\"))
		IniWrite($path_iniFile, "Trigger", "Mendeley_Backup", "1")
		IniWrite($path_iniFile, "Trigger", "Mendeley_Backup_Archive", "1")
		IniWrite($path_iniFile, "Trigger", "Mendeley_PDFData", "1")

	EndIf

EndFunc   ;==>_Enter_Ini_Details

Func _Get_Ini($fileList) ;==> The 'fileList' needs full paths

	If UBound($fileList) > 0 Then

		For $element In $fileList

			If StringRegExp($element, "CFG") Then

				Return $element

			EndIf

		Next

	Else

		Return False

	EndIf

EndFunc   ;==>_Get_Ini

Func _Handoff_PDF_From_Mendeley_To_Internal()

	Local $trigger = True, $path_pdfFromInternal

	While ($trigger = True)

		Sleep(50)

		If WinActive("Data") = True Then

			;_BlockinputEx(1)

			Sleep(500)

			Send("^c")

			Sleep(500)

			ShellExecute($path_pdfextractorInternalDir)
			Sleep(1000)
			WinActivate("Internal")

			Sleep(500)

			Send("^v")

			Sleep(1000)

			WinClose("Internal")
			WinClose("Data")

			$trigger = False

			Sleep(500)

			Send("!{F4}")

			$path_pdfFromInternal = _FileListToArray($path_mainDir & "\PDF_Extractor\Internal", Default, Default, True)

			;_BlockinputEx(0)

		EndIf

	WEnd

	Return $path_pdfFromInternal[1]

EndFunc   ;==>_Handoff_PDF_From_Mendeley_To_Internal

Func _CheckError($sMsg, $iError, $iExtended)

	If $iError Then
		$sMsg = "Error using " & $sMsg & " button (" & $iExtended & ")"
	Else
		$sMsg = ""
	EndIf
	GUICtrlSetData($g_idError_Message, $sMsg)

EndFunc   ;==>_CheckError

Func _Start_PDFeditor_with_file($path_pdfFile)

	_ClipBoard_SetData($path_pdfFile)

	Run($path_mainDir & $exe_pdfbearbeiten, "", @SW_SHOW)

	Sleep(2000)

	Send("{ALT}{ENTER}{ENTER}")
	Sleep(1000)
	Send("^v{ENTER}")

EndFunc   ;==>_Start_PDFeditor_with_file

Func _Start_Tabula_with_file($path_pdfFile)

	Local $path_tabulaDir = $path_installFilesDir & "\Tabula-Win-1.1.1", _
		$oObject

	;_BlockinputEx(1)

	FileChangeDir($path_mainDir & $path_tabulaDir)
	$iPID = Run(@ComSpec & " /k tabula.exe", "", @SW_HIDE) ; Execute the Tabula-Win-1.1.1 software (/k means 'keep' (without it does not executed))

	$gui_loading_screen = GUICreate("Starting server ...", 300, 40, 100, 200)
	$progressbarLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")
	GUICtrlSetColor(-1, 32250) ; not working with Windows XP Style
	GUISetState(@SW_SHOW)

	For $i = $iSavePosStartingServer To 100

		GUICtrlSetData($progressbarLoadingScreen, $i)

		Sleep(200)

	Next

	GUIDelete($gui_loading_screen)

	$oIE = _Start_Embedded_Browser()

	sleep(2000)

	$oObject = _IEGetObjByName($oIE, "files[]")
	_IEAction($oObject, "click")

	_ClipBoard_SetData($path_pdfFile)

	Local $path_csvFile = (StringReplace($path_pdfFile, ".pdf", ".csv"))

	MsgBox($MB_SYSTEMMODAL, "TEST", $path_csvFile)

	While NOT FileExists($path_csvFile)

		sleep(50)

	WEnd

	return $path_csvFile

	;_BlockinputEx(0)

EndFunc   ;==>_Start_Tabula_with_file

Func _Start_table_calculator_with_csv($path_csvFile)

	Local $exe_scalc = "\PDF_Extractor_InstallFiles\OpenOffice\program\scalc.exe"

	Run($path_mainDir & $exe_scalc, "", @SW_SHOW)

	sleep(3000)

	MsgBox($MB_SYSTEMMODAL, "TEST", $path_csvFile)

	_OOoCalc_BookOpen($path_csvFile)

	WinClose("Embedded Web-Browser")
	WinClose("Download beendet")

EndFunc

Func _On_Button()

	Local _
			$exe_scalc = "\PDF_Extractor_InstallFiles\OpenOffice\program\scalc.exe", _
			$path_tabulaDir = $path_installFilesDir & "\Tabula-Win-1.1.1", _
			$path_pdfFromInternal

	Switch @GUI_CtrlId ;Check which button sent the message

		Case $btn_tabula

			;_BlockinputEx(1)

			FileChangeDir($path_mainDir & $path_tabulaDir)
			$iPID = Run(@ComSpec & " /k tabula.exe", "", @SW_HIDE) ; Execute the Tabula-Win-1.1.1 software (/k means 'keep' (without it does not executed))

			$gui_loading_screen = GUICreate("Starting server ...", 300, 40, 100, 200)
			$progressbarLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
			GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")
			GUICtrlSetColor(-1, 32250) ; not working with Windows XP Style
			GUISetState(@SW_SHOW)

			For $i = $iSavePosStartingServer To 100

				GUICtrlSetData($progressbarLoadingScreen, $i)

				Sleep(200)

			Next

			GUIDelete($gui_loading_screen)

			_Start_Embedded_Browser()

			;_BlockinputEx(0)

		Case $btn_wizard

			_Start_Mendeley_with_AutoImport()

			$path_pdfFromInternal = _Handoff_PDF_From_Mendeley_To_Internal()

			_Start_PDFeditor_with_file($path_pdfFromInternal)

			WinWaitClose("PDF Bearbeiten")

			$path_csvFromInternal = _Start_Tabula_with_file($path_pdfFromInternal)

			_Start_table_calculator_with_csv($path_csvFromInternal)

		Case $btn_pdfbearbeiten

			Run($path_mainDir & $exe_pdfbearbeiten, "", @SW_SHOW)

		Case $btn_openoffice

			Run($path_mainDir & $exe_scalc, "", @SW_SHOW)

		Case $btn_mendeley

			_Start_Mendeley_with_AutoImport()

			WinWaitClose("Mendeley Desktop")

			If MsgBox(4, "Any changes?", "Do you want to create a new backup file?") = 6 Then

				_Start_Mendeley_Create_Backup()

			EndIf

		Case $btn_exit_main

			_Process_Close_Tree($iPID)
			GUIDelete($gui_main)
			Exit

		Case $btn_home

			_IENavigate($oIE, "http://127.0.0.1:8080")
			_IEAction($oIE, "stop")
			_CheckError("Home", @error, @extended)

		Case $btn_exit_embedded

			GUIDelete($gui_webbrowser)

	EndSwitch

EndFunc   ;==>_On_Button

Func _On_Close()

	Switch @GUI_WinHandle

		Case $gui_main

			GUIDelete($gui_main)
			Exit

		Case $gui_webbrowser

			GUIDelete($gui_webbrowser)

		Case $gui_loading_screen

			GUIDelete($gui_loading_screen)
			Exit

	EndSwitch

EndFunc   ;==>_On_Close

#EndRegion ### START Functions section ###
