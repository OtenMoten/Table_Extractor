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
#include <ColorConstants.au3>

#EndRegion ### START Library section ###

#Region ### START Variables section ###

; mostly only declaration ==> initializing in Func '_Start_Loading_Screen'
Global $guiLoadingScreen, $guiProgressLoadingScreen, $guiMain, $guiEmbeddedBrowser, $guiControlPanel, _
		$lbHeader, _
		$btnTabula, $btnMendeley, $btnOpenoffice, $btnConsensus, $btnWizard, $btnHome, $btnExitEmbedded, $btnExit, _
		$btnExitControlPanel, $btnUnpivot, $btnMerge, $btnKillSpaces, $btnConvert, $btnSave, _
		$oIE, $oObject, $i, $iSavePos, _
		$iPID, _ ;$iPID for calling the function
		$pathIniFile, _
		$pathMainDir, _
		$pathInstallFilesDir, _
		$pathInternalDir, _
		$pathTabulaDir, _
		$pathConsensusDir, _
		$pathPDFfromInternal, _
		$pathCSVfromInternal, _
		$pathConfigDir = @MyDocumentsDir & "\Table_Extractor_Config", _
		$pathMendeleyBackup, _
		$pathMendeleyBackupArchiveDir, _
		$pathMendeleyPDFdataDir, _
		$pathMendeleyExe, _
		$pathScalcExe, _
		$currentTablePath, _
		$pathLayouts

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

Func _Enter_Ini_Details()

	If (IniRead($pathIniFile, "Trigger", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Trigger' in ini-file.") = "1") And _
			(IniRead($pathIniFile, "Trigger", "Mendeley_PDFdata", "Can't read key 'Mendeley_PDFdata' from section 'Trigger' in ini-file.") = "1") And _
			(IniRead($pathIniFile, "Trigger", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Trigger' in ini-file.") = "1") And _
			(IniRead($pathIniFile, "Trigger", "Layouts", "Can't read key 'Layouts' from section 'Trigger' in ini-file.") = "1") Then


		;MsgBox($MB_SYSTEMMODAL, "Mendeley backup file", "The current used Mendeley backup file is stored at: " & IniRead($pathIniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file."))
		;MsgBox($MB_SYSTEMMODAL, "Mendeley backup archive folder", "The current used Mendeley backup archive folder is stored at: " & IniRead($pathIniFile, "Paths", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Paths' in ini-file."))
		;MsgBox($MB_SYSTEMMODAL, "Mendeley PDFA data folder", "The current used Mendeley PDF data folder is stored at: " & IniRead($pathIniFile, "Paths", "Mendeley_PDFdata", "Can't read key 'Mendeley_PDFdata' from section 'Paths' in ini-file."))

	Else

		Sleep(100)
		IniWrite($pathIniFile, "Paths", "Mendeley_Backup", FileOpenDialog("Open the Mendeley backup file", "\\gruppende\IV2.2\Int\WRMG\Table_Extractor\", "All (*.zip)"))
		IniWrite($pathIniFile, "Paths", "Mendeley_Backup_Archive", FileSelectFolder("Select the Mendeley backup archive folder ", "\\gruppende\IV2.2\Int\WRMG\Table_Extractor\"))
		IniWrite($pathIniFile, "Paths", "Mendeley_PDFdata", FileSelectFolder("Select the PDF data folder", "\\gruppende\IV2.2\Int\WRMG\Table_Extractor\"))
		IniWrite($pathIniFile, "Paths", "Layouts", FileSelectFolder("Select the layout folder", "\\gruppende\IV2.2\Int\WRMG\Table_Extractor\"))
		IniWrite($pathIniFile, "Trigger", "Mendeley_Backup", "1")
		IniWrite($pathIniFile, "Trigger", "Mendeley_Backup_Archive", "1")
		IniWrite($pathIniFile, "Trigger", "Mendeley_PDFdata", "1")
		IniWrite($pathIniFile, "Trigger", "Layouts", "1")

	EndIf

EndFunc   ;==>_Enter_Ini_Details

Func _Send_PDF_From_Mendeley_To_InternalDir()

	Local $foldernameMendeleyPDFdataDir, _
			$arrayPathMendeleyPDFdataDir, _
			$nameOfPDF, _
			$arrayNameOfPDF, _
			$trigger = True

	$arrayPathMendeleyPDFdataDir = StringSplit($pathMendeleyPDFdataDir, "\", 2)

	For $i = 0 To (UBound($arrayPathMendeleyPDFdataDir) - 1) Step +1

		If ($i) = (UBound($arrayPathMendeleyPDFdataDir) - 1) Then

			$foldernameMendeleyPDFdataDir = $arrayPathMendeleyPDFdataDir[$i]

		EndIf

	Next

	While ($trigger = True)

		Sleep(50)

		If WinActive($foldernameMendeleyPDFdataDir) = True Then

			Sleep(1000)

			Send("^c")

			Sleep(1000)

			$arrayPathMendeleyPDF= StringSplit(ClipGet(), "\", 2)

			For $i = 0 To (UBound($arrayPathMendeleyPDF) - 1) Step +1

				If ($i) = (UBound($arrayPathMendeleyPDF) - 1) Then

					$nameOfPDF = $arrayPathMendeleyPDF[$i]

				EndIf

			Next

			Sleep(1000)

			FileCopy($pathMendeleyPDFdataDir & "\" & $nameOfPDF, $pathInternalDir, 1)

			While _WinAPI_FileInUse($pathInternalDir & "\" & $nameOfPDF) = True

				Sleep(1000)

			WEnd

			$trigger = False

			Sleep(500)

			WinClose($foldernameMendeleyPDFdataDir)

		EndIf

	WEnd

	WinWaitActive("Mendeley Desktop")

	; Getting the XML (meta data) of the user-selected PDF
	Send("^e")
	ClipPut($pathInternalDir & "\metadata.xml")
	Sleep(500)
	ClipGet()
	Sleep(500)
	Send("^v")
	Sleep(500)
	Send("{TAB}")
	Sleep(500)
	Send("{DOWN 3}")
	Sleep(500)
	Send("{ENTER 2}")

	WinWaitActive("Mendeley Desktop")
	Send("!{F4}")

	DirRemove($pathInternalDir & "\metadata.Data", 1)

	Return $pathInternalDir & "\" & $nameOfPDF

EndFunc   ;==>_Handoff_PDF_From_Mendeley_To_Internal

Func _On_Button()

	Switch @GUI_CtrlId ;Check which button sent the message

		Case $btnTabula

			FileChangeDir($pathTabulaDir)
			$iPID = Run(@ComSpec & " /k tabula.exe", "", @SW_HIDE) ; Execute the Tabula-Win-1.2.0 software (/k means 'keep' (without it does not executed))
			FileChangeDir($pathMainDir)

			$guiLoadingScreen = GUICreate("Starting server ...", 302, 65, 100, 200, $WS_BORDER)
			$guiProgressLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
			GUICtrlSetColor(-1, 32250) ; not working with Windows XP Style
			GUISetState(@SW_SHOW)

			For $i = $iSavePos To 100

				GUICtrlSetData($guiProgressLoadingScreen, $i)

				Sleep(200)

			Next

			GUIDelete($guiProgressLoadingScreen)
			GUIDelete($guiLoadingScreen)

			_Start_Embedded_Browser()

		Case $btnWizard

			_Start_Mendeley_with_AutoImport()

			$pathPDFfromInternal = _Send_PDF_From_Mendeley_To_InternalDir()

			$pathCSVfromInternal = _Start_Tabula_with_file($pathPDFfromInternal)

			_Start_Table_Calculator_with_CSV($pathCSVfromInternal)


		Case $btnOpenoffice

			$pathScalcExe = $pathInstallFilesDir & "\OpenOffice\program\scalc.exe"

			Run($pathScalcExe, "", @SW_SHOW)

			_Start_Control_Panel()

		Case $btnMendeley

			_Start_Mendeley_with_AutoImport()

			WinWaitClose("Mendeley Desktop")

			If MsgBox(4, "Any changes?", "Do you want to create a new backup file?") = 6 Then

				_Start_Mendeley_Create_Backup()

			EndIf

		Case $btnConsensus

			_Start_Consensus()

		Case $btnExit

			_Process_Tree_Close($iPID)
			GUIDelete($guiMain)
			Exit

		Case $btnHome

			_IENavigate($oIE, "http://127.0.0.1:8080")

		Case $btnExitEmbedded

			_Process_Tree_Close($iPID)
			GUIDelete($guiEmbeddedBrowser)

		Case $btnUnpivot

			MsgBox($MB_SYSTEMMODAL, "btnUnpivot", "Successful")

		Case $btnMerge

			MsgBox($MB_SYSTEMMODAL, "btnMerge", "Successful")

		Case $btnKillSpaces

			Local $OOobject

			;$OOobject = _OOoCalc_BookOpen("C:\Test_Extraktor\Internal\Boulard_et_al._2018__polare_AM.ods", False)
			;MsgBox($MB_SYSTEMMODAL, "OOobject handle is: ", $OOobject + @error )
			;_OOoCalc_ColumnDelete($OOobject, 1)

			Local $fileName = "C:\Test_Extraktor\Internal\Boulard_et_al._2018__polare_AM.ods"
			Local $OOobject = _OOoCalc_BookOpen($fileName)
			MsgBox($MB_SYSTEMMODAL, "OOobject", $OOobject + @error)
			;_OOoCalc_RowDelete($OOobject, 1)

		Case $btnConvert

			MsgBox($MB_SYSTEMMODAL, "btnConvert", "Successful")

		Case $btnSave

			WinActivate("Calc")
			WinWaitActive("Calc")

			Sleep(500)
			Send("+^s")
			WinWaitActive("Speichern unter")
			Sleep(500)
			Send("{TAB}")
			Sleep(500)
			Send("{DOWN}")
			Sleep(500)
			Send("{UP 7}")
			Sleep(500)
			Send("{ENTER}")
			Send("{ENTER}")

			WinWaitActive("OpenOffice")
			Send("{ENTER}")
			WinWaitActive("OpenOffice")
			Send("!{F4}")

			GUIDelete($guiControlPanel)
			Sleep(500)

			WinWaitActive("OpenOffice")
			WinWaitClose("OpenOffice")
			Sleep(500)

			ShellExecute($pathConsensusDir & "\Consensus.jar", IniRead($pathIniFile, "Paths", "Layouts", "Can't read path 'Layouts' from section 'Paths' in ini-file.") & "%" & $pathInternalDir, $pathConsensusDir)


		Case $btnExitControlPanel

			GUIDelete($guiControlPanel)

	EndSwitch

EndFunc   ;==>_On_Button

Func _On_Close()

	Switch @GUI_WinHandle

		Case $guiMain

			_Process_Tree_Close($iPID)
			GUIDelete($guiMain)
			Exit

		Case $guiEmbeddedBrowser

			_Process_Tree_Close($iPID)
			GUIDelete($guiEmbeddedBrowser)

		Case $guiLoadingScreen

			GUIDelete($guiLoadingScreen)
			Exit

		Case $guiControlPanel

			GUIDelete($guiControlPanel)

	EndSwitch

EndFunc   ;==>_On_Close

Func _Process_Tree_Close($sPID)

	If IsString($sPID) Then $sPID = ProcessExists($sPID)
	If Not $sPID Then Return SetError(1, 0, 0)

	Return Run(@ComSpec & " /c taskkill /F /PID " & $sPID & " /T", @SystemDir, @SW_HIDE)

EndFunc   ;==>_Process_Tree_Close

Func _Start_Consensus()

	Local $temp = IniRead($pathIniFile, "Paths", "Layouts", "Can't read path 'Layouts' from section 'Paths' in ini-file.") & "%" & $pathInternalDir
	ShellExecute($pathConsensusDir & "\Consensus.jar", $temp, $pathConsensusDir)

EndFunc

Func _Start_Control_Panel()

	$guiControlPanel = GUICreate("Control Panel", 400, 400, 100, 100, $WS_BORDER)
	GUISetBkColor(0xf4f4f4, $guiControlPanel)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")

	GUISetState(@SW_SHOW, $guiControlPanel)
	WinSetOnTop($guiControlPanel, "", 1)

	$btnUnpivot = GUICtrlCreateButton("Unpivot", 0, 0, 150, 100)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnMerge = GUICtrlCreateButton("Merge", 0, 137, 150, 100)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnKillSpaces = GUICtrlCreateButton("Kill spaces", 0, 274, 150, 100)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnConvert = GUICtrlCreateButton("Convert", 246, 137, 150, 100)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnSave = GUICtrlCreateButton("Save as", 246, 0, 150, 100)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnExitControlPanel = GUICtrlCreateButton("Exit", 246, 274, 150, 100)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")


EndFunc

Func _Start_Embedded_Browser()

	$oIE = _IECreateEmbedded()
	$guiEmbeddedBrowser = GUICreate("Embedded Web-Browser", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")
	GUICtrlCreateObj($oIE, 10, 40, (@DesktopWidth - 20), (@DesktopHeight - 50))
	$btnHome = GUICtrlCreateButton("Home", 10, 5, 100, 30)
	GUICtrlSetOnEvent(-1, "_On_Button")
	$btnExitEmbedded = GUICtrlCreateButton("Exit", 1810, 5, 100, 30)
	GUICtrlSetOnEvent(-1, "_On_Button")
	GUISetState(@SW_SHOW, $guiEmbeddedBrowser)
	_IENavigate($oIE, "http://127.0.0.1:8080")
	_IEAction($oIE, "stop")
	_IELinkClickByText($oIE, "My Files")

	;WinWaitActive("Datei zum Hochladen")
	;Sleep(500)
	;Send("^v")

	Return $oIE

EndFunc   ;==>_Start_Embedded_Browser

Func _Start_File_Install()

	$pathTabulaDir = $pathInstallFilesDir & "\Tabula-Win-1.2.0"
	$pathConsensusDir = $pathInstallFilesDir & "\Consensus"
	Local _
			$pathExe7zip = $pathInstallFilesDir & "\7za.exe", _
			$pathMendeleyDir = $pathInstallFilesDir & "\Mendeley", _
			$pathOpenofficeDir = $pathInstallFilesDir & "\OpenOffice", _
			$pathMendeleyZipped = $pathInstallFilesDir & "\Mendeley.7z", _
			$pathTabulaZipped = $pathInstallFilesDir & "\Tabula-Win-1.2.0.7z", _
			$pathOpenofficeZipped = $pathInstallFilesDir & "\OpenOffice.7z", _
			$pathConsensusZipped = $pathInstallFilesDir & "\Consensus.7z"

	#Region ### START Directory in personal documents (default path of installing files) ###

	If Not FileExists($pathInstallFilesDir) Then

		DirCreate($pathInstallFilesDir)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 5)

	If Not FileExists($pathInternalDir) Then

		DirCreate($pathInternalDir)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 10)

	#EndRegion ### START Directory in personal documents (default path of installing files) ###

	#Region ### START Zip installing files (7zip) and set up saving directory ###

	#cs - - - Path of zipped installing files - - -
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\7za.exe"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Mendeley.7z"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Tabula-Win-1.2.0.7z"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\OpenOffice.7z"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\PDFBearbeiten.7z"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Consensus.7z"
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

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Tabula-Win-1.2.0.7z", $pathTabulaZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 25)

	If Not FileExists($pathOpenofficeDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\OpenOffice.7z", $pathOpenofficeZipped)

	EndIf

	If Not FileExists($pathConsensusDir) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Consensus.7z", $pathConsensusZipped)

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

	If Not FileExists($pathConsensusDir) Then

		RunWait($pathExe7zip & ' x ' & $pathConsensusZipped & ' -o' & $pathInstallFilesDir, "", @SW_HIDE)

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

	If FileExists($pathConsensusZipped) Then

		FileDelete($pathConsensusZipped)

	EndIf

	GUICtrlSetData($guiProgressLoadingScreen, 100)
	Sleep(100)

	#EndRegion ### START Delete zipped installing files ###

EndFunc   ;==>_Start_File_Install

Func _Start_GUI_Main()

	#cs - - - Path of pictures - - -
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\ICON_Mendeley.png"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\LOGO_Tabula.png"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\LOGO_UBA.png"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\ICON_Wizard.png"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\ICON_OpenOffice.png"
		"P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\LOGO_Java.png"
	#ce - - - Path of pictures - - -

	Local _
			$pathMendeleyPic = $pathInstallFilesDir & "\ICON_Mendeley.png", _
			$pathTabulaPic = $pathInstallFilesDir & "\LOGO_Tabula.png", _
			$pathUbaPic = $pathInstallFilesDir & "\LOGO_UBA.png", _
			$pathWizardPic = $pathInstallFilesDir & "\ICON_Wizard.png", _
			$pathOpenofficescalcPic = $pathInstallFilesDir & "\ICON_OpenOffice.png", _
			$pathJavaPic = $pathInstallFilesDir & "\LOGO_Java.png"

	If Not FileExists($pathMendeleyPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\ICON_Mendeley.png", $pathMendeleyPic)

	EndIf

	If Not FileExists($pathTabulaPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\LOGO_Tabula.png", $pathTabulaPic)

	EndIf

	If Not FileExists($pathUbaPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\LOGO_UBA.png", $pathUbaPic)

	EndIf

	If Not FileExists($pathWizardPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\ICON_Wizard.png", $pathWizardPic)

	EndIf

	If Not FileExists($pathOpenofficescalcPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\ICON_OpenOffice.png", $pathOpenofficescalcPic)

	EndIf

	If Not FileExists($pathJavaPic) Then

		FileInstall("P:\FG_IV_2.2\Projects\Table_Extractor\External_Software\Pictures\LOGO_Java.png", $pathJavaPic)

	EndIf

	$guiMain = GUICreate("Table Extractor", @DesktopWidth, @DesktopHeight, 0, 0, $WS_POPUP)
	GUISetBkColor(0xf4f4f4, $guiMain)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")

	$lbHeader = GUICtrlCreateLabel("Software for extracting and converting tables from a PDF file", @DesktopWidth / 80, @DesktopHeight / 80, @DesktopWidth / 2, @DesktopHeight / 45)
	GUICtrlSetFont(-1, 16, 600, 0, "MS Sans Serif")
	GUICtrlSetColor(-1, 0x000000)

	GUISetState(@SW_SHOW, $guiMain)

	_GUICtrlPic_Create($pathUbaPic, @DesktopWidth / 1.2, @DesktopHeight / 1201, @DesktopWidth / 6, @DesktopHeight / 7)
	_GUICtrlPic_Create($pathMendeleyPic, @DesktopWidth / 20, @DesktopHeight / 10, @DesktopWidth / 5, @DesktopHeight / 3, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)
	_GUICtrlPic_Create($pathTabulaPic, @DesktopWidth / 3, @DesktopHeight / 11.5, @DesktopWidth / 5, @DesktopHeight / 3, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)
	_GUICtrlPic_Create($pathOpenofficescalcPic, @DesktopWidth / 20, @DesktopHeight / 1.8, @DesktopWidth / 5, @DesktopHeight / 3, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)
	_GUICtrlPic_Create($pathWizardPic, @DesktopWidth / 1.15, @DesktopHeight / 5, @DesktopWidth / 12, @DesktopHeight / 8, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)
	_GUICtrlPic_Create($pathJavaPic, @DesktopWidth / 2.7, @DesktopHeight / 1.75, @DesktopWidth / 8, @DesktopHeight / 3.5, BitOR($SS_CENTERIMAGE, $SS_SUNKEN, $SS_NOTIFY), Default)

	$btnTabula = GUICtrlCreateButton("Tabula", (@DesktopWidth / 3), (@DesktopHeight / 2.3), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnMendeley = GUICtrlCreateButton("Mendeley", (@DesktopWidth / 20), (@DesktopHeight / 2.3), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x0F0000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnOpenoffice = GUICtrlCreateButton("OpenOffice", (@DesktopWidth / 20), (@DesktopHeight / 1.124), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnConsensus = GUICtrlCreateButton("Consensus", (@DesktopWidth / 3), (@DesktopHeight / 1.124), (@DesktopWidth / 5), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnWizard = GUICtrlCreateButton("Assistant", (@DesktopWidth / 1.2), (@DesktopHeight / 6), (@DesktopWidth / 6), (@DesktopHeight / 25))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnExit = GUICtrlCreateButton("Exit", (@DesktopWidth / 1.113), (@DesktopHeight / 1.075), (@DesktopWidth / 10), (@DesktopHeight / 15))
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	While 1
		Sleep(10)
	WEnd

EndFunc   ;==>_Start_GUI_Main

Func _Start_Loading_Screen()

	If (FileFindNextFile(FileFindFirstFile($pathConfigDir & "\*.*")) <> "Table_Extractor.ini") Then

		If Not FileExists($pathConfigDir) Then

			DirCreate($pathConfigDir)
			$pathIniFile = $pathConfigDir & "\Table_Extractor.ini"
			$pathMainDir = FileSelectFolder("Choose the installation directory", "C:\", $WS_POPUP)
			IniWrite($pathIniFile, "Information", "Explanation", "This file contains parameters for the Software 'Table_Extractor'")
			IniWrite($pathIniFile, "Paths", "Main_Dir", $pathMainDir)
			IniWrite($pathIniFile, "Trigger", "Main_Dir", "1")

		Else

			$pathIniFile = $pathConfigDir & "\Table_Extractor.ini"
			$pathMainDir = FileSelectFolder("Choose the installation directory", "C:\", $WS_POPUP)
			IniWrite($pathIniFile, "Information", "Explanation", "This file contains parameters for the Software 'Table_Extractor'")
			IniWrite($pathIniFile, "Paths", "Main_Dir", $pathMainDir)
			IniWrite($pathIniFile, "Trigger", "Main_Dir", "1")

		EndIf

	EndIf

	; initializing interal used variables
	$pathIniFile = $pathConfigDir & "\Table_Extractor.ini"
	$pathMainDir = IniRead($pathIniFile, "Paths", "Main_Dir", "Can't read key 'Main_Dir' from section 'Paths' in " & ($pathConfigDir & "\Table_Extractor.ini"))
	$pathInstallFilesDir = $pathMainDir & "\InstallFiles"
	$pathInternalDir = $pathMainDir & "\Internal"

	; values for altering the progressbar are triggered in '_Start_File_Install()' GUICtrlSetData($guiProgressLoadingScreen, $i)
	$guiLoadingScreen = GUICreate("Starting Table Extractor", 300, 40, 100, 200)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close", $guiLoadingScreen)
	$guiProgressLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
	GUICtrlSetColor($guiProgressLoadingScreen, $COLOR_GREEN) ; not working with Windows XP Style
	GUISetState(@SW_SHOW, $guiLoadingScreen)

	_Start_File_Install()

	GUIDelete($guiProgressLoadingScreen)
	GUIDelete($guiLoadingScreen)

EndFunc   ;==>_Start_Loading_Screen

Func _Start_Mendeley_Create_Backup()

	Run($pathMendeleyExe, "", @SW_SHOW)

	$pathMendeleyBackup = IniRead($pathIniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file.")
	$pathMendeleyBackupArchiveDir = IniRead($pathIniFile, "Paths", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Paths' in ini-file.")

	ClipPut($pathMendeleyBackup)

	$time = _Date_Time_SystemTimeToFileTime(_Date_Time_GetSystemTime())
	$time = _Date_Time_FileTimeToStr($time)
	$time = StringReplace($time, ":", "_")
	$time = StringReplace($time, "/", "_")
	$time = StringReplace($time, " ", "_")

	FileMove($pathMendeleyBackup, $pathMendeleyBackupArchiveDir & "\" & "Archive_" & $time & ".zip", 1)
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
	Sleep(800)
	Send("^v")
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	Send("{ESC}")
	Sleep(1000)
	Send("!{F4}")

EndFunc   ;==>_Start_Mendeley_Create_Backup

Func _Start_Mendeley_with_AutoImport()

	$pathMendeleyExe = $pathInstallFilesDir & "\Mendeley\Mendeley Desktop\MendeleyDesktop.exe"

	Run($pathMendeleyExe, "", @SW_SHOW)

	$pathMendeleyBackup = IniRead($pathIniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file.")
	_ClipBoard_SetData($pathMendeleyBackup)
	WinWaitActive("Mendeley Desktop")
	Sleep(6000)
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

	If (WinWaitActive("Welcome to Mendeley Desktop","", 10) = 0) Then
		Run($pathMendeleyExe, "", @SW_SHOW)
	EndIf

	WinWaitActive("Welcome to Mendeley Desktop")
	Sleep(500)
	Send("{ENTER}")
	Sleep(1000)

	WinWaitActive("Mendeley Desktop")
	Sleep(3000)

	$pathMendeleyPDFdataDir = IniRead($pathIniFile, "Paths", "Mendeley_PDFdata", "Can't read key 'Mendeley_PDFdata' from section 'Paths' in ini-file.")
	_ClipBoard_SetData($pathMendeleyPDFdataDir)

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
	Send("{TAB}")
	Sleep(200)
	Send("{TAB}")
	Sleep(200)
	Send("{TAB}")
	Sleep(200)
	Send("{TAB}")
	Sleep(200)
	Send("{TAB}")
	Sleep(200)
	Send("{RIGHT}")
	Sleep(200)
	Send("{TAB}")
	Sleep(500)
	Send("{SPACE}")
	Sleep(200)
	Send("^v")
	Sleep(200)
	Send("{ENTER}")
	Sleep(500)
	Send("{LEFT}")
	Sleep(500)
	Send("{ENTER}")

	;BlockInput($BI_ENABLE)

EndFunc   ;==>_Start_Mendeley_with_AutoImport

Func _Start_Table_Calculator_with_CSV($pathCSVfile)

	$pathScalcExe = $pathInstallFilesDir & "\OpenOffice\program\scalc.exe"

	Run($pathScalcExe, "", @SW_SHOW)

	WinWaitActive("OpenOffice Calc")
	Sleep(1000)
	Send("{CTRLDOWN}")
	Sleep(500)
	Send("o")
	Send("{CTRLUP}")
	WinWaitActive("Öffnen")
	ClipPut($pathCSVfile)
	Sleep(500)
	Send("^v")
	Sleep(500)
	Send("{ENTER}")
	WinWaitActive("Textimport")
	Send("{ENTER}")

	_Start_Control_Panel()

EndFunc

Func _Start_Tabula_with_file($pathPDFfile)

	Local $arrayPathCSVfile, $newPathCSVfile
	$pathTabulaDir = $pathInstallFilesDir & "\Tabula-Win-1.2.0"

	FileChangeDir($pathTabulaDir)
	$iPID = Run(@ComSpec & " /k tabula.exe", "", @SW_HIDE) ; Execute the Tabula-Win-1.2.0 software (/k means 'keep' (without it does not executed))

	$guiLoadingScreen = GUICreate("Starting server ...", 300, 40, 100, 200)
	$guiProgressLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
	GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close")
	GUICtrlSetColor(-1, 32250) ; not working with Windows XP Style
	GUISetState(@SW_SHOW)

	$iSavePos = 0

	For $i = $iSavePos To 100

		GUICtrlSetData($guiProgressLoadingScreen, $i)

		Sleep(200)

	Next

	GUIDelete($guiLoadingScreen)

	$oIE = _Start_Embedded_Browser()

	sleep(2000)

	ClipPut($pathPDFfile)

	FileChangeDir($pathTabulaDir)

	$oObject = _IEGetObjByName($oIE, "files[]")
	_IEAction($oObject, "click")

	;;;;;; no auto

	$pathCSVfile = StringReplace($pathPDFfile, ".pdf", ".csv")

	Local $arrayPathCSVfile = StringSplit($pathCSVfile, "\", 2)
	Local $newPathCSVfile

	For $i = 0 To (UBound($arrayPathCSVfile) - 1) Step +1

		If ($i) = (UBound($arrayPathCSVfile) - 1) Then

			$arrayPathCSVfile[$i] = StringReplace($arrayPathCSVfile[$i], $arrayPathCSVfile[$i], "tabula-" & $arrayPathCSVfile[$i])
			$arrayPathCSVfile[$i] = StringReplace($arrayPathCSVfile[$i], " ", "_")

		EndIf

	Next

	FileChangeDir($pathInternalDir)

	$newPathCSVfile = _ArrayToString($arrayPathCSVfile, "\")

	FileChangeDir($pathInternalDir)
	;While NOT FileExists($newPathCSVfile)

	;	If WinActive("Speichern unter") = true Then
	;		Sleep(500)
	;		Send("{ENTER}")
	;	Else
	;		sleep(50)
	;	EndIf

	;WEnd

	WinWaitActive("Download beendet")
	WinClose("Download beendet")

	Sleep(1000)

	_Process_Tree_Close($iPID)

	Sleep(1000)

	GUIDelete($guiEmbeddedBrowser)

	Local $orignalNamePath = StringReplace($newPathCSVfile, "tabula-", "");

	FileMove($newPathCSVfile, $orignalNamePath)

	$newPathCSVfile = $orignalNamePath

	Sleep(1000)

	return $newPathCSVfile

EndFunc   ;==>_Start_Tabula_with_file