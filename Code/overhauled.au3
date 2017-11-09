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
#include <ColorConstants.au3>

#Region ### START Variables section ###

Global $guiLoadingScreen, _
		$guiProgressLoadingScreen, _
		$i, _
		$iPID, _
		$pathIniFile, _
		$pathMainDir, _
		$pathConfigDir = @MyDocumentsDir & "\Table_Extractor_Config"

#EndRegion ### START Variables section ###

#Region ### START System parameters section ###

Opt("WinTitleMatchMode", 2)
Opt("GUIOnEventMode", 1)

#EndRegion ### START System parameters section ###

#Region ### START Running section ###

_Start_Loading_Screen()
;_Enter_Ini_Details()
;_Start_GUI_Main()

#EndRegion ### START Running section ###

Func _Start_Loading_Screen()

	If (FileFindNextFile(FileFindFirstFile($pathConfigDir & "\Table_Extractor.ini")) <> "Table_Extractor.ini") Then

		If Not FileExists($pathConfigDir) Then

			DirCreate($pathConfigDir)
			$pathIniFile = $pathConfigDir & "\Table_Extractor.ini"
			$pathMainDir = FileSelectFolder("Choose the installation directory", "C:\", $WS_POPUP)
			IniWrite($pathIniFile, "Information", "Explanation", "This file contains parameters for the Software 'Table_Extractor'")
			IniWrite($pathIniFile, "Paths", "MainDir", $pathMainDir)
			IniWrite($pathIniFile, "Trigger", "MainDir", "1")

		Else

			$pathIniFile = $pathConfigDir & "\Table_Extractor.ini"
			$pathMainDir = FileSelectFolder("Choose the installation directory", "C:\", $WS_POPUP)
			IniWrite($pathIniFile, "Information", "Explanation", "This file contains parameters for the Software 'Table_Extractor'")
			IniWrite($pathIniFile, "Paths", "MainDir", $pathMainDir)
			IniWrite($pathIniFile, "Trigger", "MainDir", "1")

		EndIf

	EndIf

	$pathIniFile = $pathConfigDir & "\Table_Extractor.ini"
	$pathMainDir = IniRead($pathIniFile, "Paths", "MainDir", "Can't read key 'MainDir' from section 'Paths' in " & ($pathConfigDir & "\Table_Extractor.ini"))

	;==> Values for altering the progressbar are triggered in '_Start_File_Install()' GUICtrlSetData($guiProgressLoadingScreen, $i)
	$guiLoadingScreen = GUICreate("Starting Table Extractor", 300, 40, 100, 200)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close", $guiLoadingScreen)
	$guiProgressLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
	GUICtrlSetColor($guiProgressLoadingScreen, $COLOR_GREEN) ; not working with Windows XP Style
	GUISetState(@SW_SHOW, $guiLoadingScreen)

	_Start_File_Install()

	GUIDelete($guiLoadingScreen)

	;_Process_Close_Tree($iPID)

EndFunc   ;==>_Start_Loading_Screen

Func _Start_File_Install()

	Local _
			$pathInstallFilesDir = $pathMainDir & "\InstallFiles", _
			$pathTableExtractorInternalDir = $pathMainDir & "\Internal", _
			$pathExe7zip = $pathInstallFilesDir & "\7za.exe", _
			$pathMendeleyDir = $pathInstallFilesDir & "\Mendeley", _
			$pathTabulaDir = $pathInstallFilesDir & "\Tabula-Win-1.1.1", _
			$pathOpenofficeDir = $pathInstallFilesDir & "\OpenOffice", _
			$pathMendeleyZipped = $pathInstallFilesDir & "\Mendeley.7z", _
			$pathTabulaZipped = $pathInstallFilesDir & "\Tabula-Win-1.1.1.7z", _
			$pathOpenofficeZipped = $pathInstallFilesDir & "\OpenOffice.7z"

	#Region ### START Directory in personal documents (default path of installing files) ###

	If Not FileExists($pathInstallFilesDir) Then

		DirCreate($pathInstallFilesDir)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 5)

	If Not FileExists($pathTableExtractorInternalDir) Then

		DirCreate($pathTableExtractorInternalDir)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 10)

	#EndRegion ### START Directory in personal documents (default path of installing files) ###

	#Region ### START Zip installing files (7zip) and set up saving directory ###

	#cs - - - Path of zipped installing files - - -
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\7za.exe"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Mendeley.7z"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Tabula-Win-1.1.1.7z"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\OpenOffice.7z"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\PDFBearbeiten.7z"
	#ce - - - Path of zipped installing files - - -

	If Not FileExists($pathExe7zip) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\7za.exe", $pathExe7zip)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 15)

	If Not FileExists($pathMendeleyDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Mendeley.7z", $pathMendeleyZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 20)

	If Not FileExists($pathTabulaDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Tabula-Win-1.1.1.7z", $pathTabulaZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 25)

	If Not FileExists($pathOpenofficeDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\OpenOffice.7z", $pathOpenofficeZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 30)

	#EndRegion ### START Zip installing files (7zip) and set up saving directory ###

	#Region ### START Unzip installing files with portable 7zip ###

	If Not FileExists($pathMendeleyDir) Then

		RunWait($pathExe7zip & ' x ' & $pathMendeleyZipped & ' -o' & $pathInstallFilesDir, "", @SW_HIDE)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 50)
	Sleep(100)

	If Not FileExists($pathTabulaDir) Then

		RunWait($pathExe7zip & ' x ' & $pathTabulaZipped & ' -o' & $pathInstallFilesDir, "", @SW_HIDE)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 60)
	Sleep(100)

	If Not FileExists($pathOpenofficeDir) Then

		RunWait($pathExe7zip & ' x ' & $pathOpenofficeZipped & ' -o' & $pathInstallFilesDir, "", @SW_HIDE)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 70)
	Sleep(100)

	#EndRegion ### START Unzip installing files with portable 7zip ###

	#Region ### START Delete zipped installing files ###

	If FileExists($pathExe7zip) Then

		FileDelete($pathExe7zip)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 85)
	Sleep(100)

	If FileExists($pathMendeleyZipped) Then

		FileDelete($pathMendeleyZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 90)
	Sleep(100)

	If FileExists($pathTabulaZipped) Then

		FileDelete($pathTabulaZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 95)
	Sleep(100)

	If FileExists($pathOpenofficeZipped) Then

		FileDelete($pathOpenofficeZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 100)
	Sleep(100)

	#EndRegion ### START Delete zipped installing files ###

EndFunc   ;==>_Start_File_Install

#cs
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
	$guiProgressLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")
	GUICtrlSetColor(-1, 32250) ; not working with Windows XP Style
	GUISetState(@SW_SHOW)

	For $i = $iSavePosStartingServer To 100

	GUICtrlSetData($guiProgressLoadingScreen, $i)

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
#ce

Func _On_Close()

	Switch @GUI_WinHandle

		;Case $gui_main

		;GUIDelete($gui_main)
		;Exit

		;Case $gui_webbrowser

		;GUIDelete($gui_webbrowser)

		Case $guiLoadingScreen

			GUIDelete($guiLoadingScreen)
			Exit

	EndSwitch

EndFunc   ;==>_On_Close

Func _Process_Close_Tree($sPID)

	If IsString($sPID) Then $sPID = ProcessExists($sPID)
	If Not $sPID Then Return SetError(1, 0, 0)

	Return Run(@ComSpec & " /c taskkill /F /PID " & $sPID & " /T", @SystemDir, @SW_HIDE)

EndFunc   ;==>_Process_Close_Tree
