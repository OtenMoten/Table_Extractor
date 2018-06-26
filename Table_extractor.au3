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
#include <Excel.au3>
#include <FileConstants.au3>
#include <Crypt.au3>

#EndRegion ### START Library section ###

#Region ### START Variables section ###

; mostly only declaration ==> initializing in Func '_Start_Loading_Screen'
Global Const $CALG_SHA_256 = 0x0000800c ; For func "sha256()"
Global $guiLoadingScreen, $guiProgressLoadingScreen, $guiMain, $guiEmbeddedBrowser, $guiControlPanel, _
		$lbHeader, _
		$btnTabula, $btnMendeley, $btnOpenoffice, $btnConsensus, $btnWizard, $btnHome, $btnExitEmbedded, $btnExit, _
		$btnExitControlPanel, $btnUnpivot, $btnKillSpaces, $btnConvert, $btnSave, _
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

		; Select the folder where the Structure should be created
		IniWrite($pathIniFile, "Paths", "Structure", FileSelectFolder("Select a folder where the Structure should be created: ", "C:\", "*.*"))
		DirCreate(IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file while creating the Structure folder."))
		MsgBox($MB_SYSTEMMODAL, "Structure", "The current path of the Structure folder is: " & IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file."))

		; Create the 1) Backup, 2) Data and 3) Layout folder within the Strucute folder
		; 1) Backup
		IniWrite($pathIniFile, "Paths", "Mendeley_Backup", IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file.") & "\Backup\myBack.zip")
		DirCreate(IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file.") & "\Backup")
		MsgBox($MB_SYSTEMMODAL, "Mendeley_Backup", "The current path of the Mendeley-backup is: " & IniRead($pathIniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file."))
		; 1.1) Backup Archive
		IniWrite($pathIniFile, "Paths", "Mendeley_Backup_Archive", IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file.") & "\Backup\Archive")
		DirCreate(IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file.") & "\Backup\Archive")
		MsgBox($MB_SYSTEMMODAL, "Mendeley_Backup_Archive", "The current path of the Mendeley-backup Archive folder is: " & IniRead($pathIniFile, "Paths", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Paths' in ini-file."))
		; 2) Data
		IniWrite($pathIniFile, "Paths", "Mendeley_PDFdata", IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file.") & "\Data")
		DirCreate(IniRead($pathIniFile, "Paths", "Mendeley_PDFdata", "Can't read key 'Mendeley_PDFdata' from section 'Paths' in ini-file."))
		MsgBox($MB_SYSTEMMODAL, "Mendeley_PDFdata", "The current path of the Mendeley-PDF Data folder is: " & IniRead($pathIniFile, "Paths", "Mendeley_PDFdata", "Can't read key 'Mendeley_PDFdata' from section 'Paths' in ini-file."))
		; 3) Layout
		IniWrite($pathIniFile, "Paths", "Layouts", IniRead($pathIniFile, "Paths", "Structure", "Can't read key 'Structure' from section 'Paths' in ini-file.") & "\Layouts")
		DirCreate(IniRead($pathIniFile, "Paths", "Layouts", "Can't read key 'Layouts' from section 'Paths' in ini-file."))
		MsgBox($MB_SYSTEMMODAL, "Layouts", "The current path of the Layouts folder is: " & IniRead($pathIniFile, "Paths", "Layouts", "Can't read key 'Layouts' from section 'Paths' in ini-file."))
		; 3.1) Layout Checkboxes
		DirCreate(IniRead($pathIniFile, "Paths", "Layouts", "Can't read key 'Layouts' from section 'Paths' in ini-file.") & "\checkboxes")
		MsgBox($MB_SYSTEMMODAL, "Layouts", "The current path of the Layouts-checkboxes folder is: " & IniRead($pathIniFile, "Paths", "Layouts", "Can't read key 'Layouts' from section 'Paths' in ini-file.") & "\checkboxes")

		;IniWrite($pathIniFile, "Paths", "Mendeley_Backup_Archive", FileSelectFolder("Select the Mendeley backup archive folder ", "\\gruppende\IV2.2\Int\WRMG\Table_Extractor\"))
		;IniWrite($pathIniFile, "Paths", "Mendeley_PDFdata", FileSelectFolder("Select the PDF data folder", "\\gruppende\IV2.2\Int\WRMG\Table_Extractor\"))
		;IniWrite($pathIniFile, "Paths", "Layouts", FileSelectFolder("Select the layout folder", "\\gruppende\IV2.2\Int\WRMG\Table_Extractor\"))

		IniWrite($pathIniFile, "Trigger", "Mendeley_Backup", "1")
		IniWrite($pathIniFile, "Trigger", "Mendeley_Backup_Archive", "1")
		IniWrite($pathIniFile, "Trigger", "Mendeley_PDFdata", "1")
		IniWrite($pathIniFile, "Trigger", "Layouts", "1")

	EndIf

EndFunc   ;==>_Enter_Ini_Details

Func _Send_PDF_From_Mendeley_To_InternalDir()

	Local $nameOfPDF, $pathToPDF, $myFileListArrayFullPath, $myTimer = 500, $myTrigger = False, $pathMyPDF

	While $myTrigger = False
		If (WinExists("Export Selected Documents") = True) Then
			If (WinActive("Export Selected Documents") = True) Then
				_BlockInputEx(2)

				; Getting the XML (meta data) of the user-selected PDF
				ClipPut($pathInternalDir & "\metadata.xml")
				Sleep($myTimer)
				ClipGet()
				Sleep($myTimer)
				Send("^v")
				Sleep($myTimer)
				Send("{TAB}")
				Sleep($myTimer)
				Send("{DOWN 3}")
				Sleep($myTimer)
				Send("{ENTER 2}")
				Sleep($myTimer)
				$myTrigger = True
			EndIf
		EndIf
	WEnd
	$myTrigger = False

	While $myTrigger = False
		If (WinExists("Mendeley Desktop") = True) Then
			If (WinActive("Mendeley Desktop") = True) Then
				; Getting the XML (meta data) of the user-selected PDF
				ClipPut($pathInternalDir & "\metadata.xml")
				Sleep($myTimer)
				Send("!{F4}")
				Sleep($myTimer)
				$myTrigger = True
			EndIf
		EndIf
	WEnd
	$myTrigger = False

	$myFileListArrayFullPath = _FileListToArray($pathInternalDir & "\" & "metadata.Data\PDF\", "*.*", 1, True)
	$myFileListArrayFileName = _FileListToArray($pathInternalDir & "\" & "metadata.Data\PDF\", "*.*", 1, False)

	$pathToPDF = $myFileListArrayFullPath[1]
	$nameOfPDF = $myFileListArrayFileName[1]

	$pathMyPDF = $pathInternalDir & "\" & $nameOfPDF

	FileMove($pathToPDF, $pathMyPDF)

	While ((_WinAPI_FileInUse($pathToPDF) = 1) Or (_WinAPI_FileInUse($pathMyPDF) = 1))
		ConsoleWrite("in use: " & $pathToPDF & @CRLF)
		ConsoleWrite("in use: " & $pathMyPDF & @CRLF)
	WEnd

	Sleep($myTimer * 2)

	DirRemove($pathInternalDir & "\metadata.Data", 1)

	Sleep($myTimer * 2)

	_BlockInputEx(0)

	Return $pathMyPDF

EndFunc   ;==>_Send_PDF_From_Mendeley_To_InternalDir

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

			_Start_Mendeley_Create_Backup()

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
			Exit
			GUIDelete($guiEmbeddedBrowser)

		Case $btnUnpivot

			;;_BlockInputEx(1)

			WinActivate("Calc")
			Sleep(500)
			Send("{ALT}")
			Sleep(500)
			Send("{LEFT}")
			Sleep(500)
			Send("{LEFT}")
			Sleep(500)
			Send("{LEFT}")
			Sleep(500)
			Send("{DOWN}")
			Sleep(500)
			Send("{DOWN}")
			Sleep(500)
			Send("{DOWN}")
			Sleep(500)
			Send("{DOWN}")
			Sleep(500)
			Send("{ENTER}")
			Sleep(500)
			WinWaitActive("Decroise")
			WinActivate("Decroise")
			Send("{ALTDOWN}")
			Sleep(500)
			Send("v")
			Sleep(500)
			Send("{ALTUP}")
			Sleep(500)

			;_BlockInputEx(0)

		Case $btnKillSpaces
			MsgBox(0, "Test", "Inside of KillSpaces()")

			Local $trigger = False

			While ($trigger = False)

				If (WinActive("OpenOffice") = True) Then
					MsgBox(0, "Test", "Trigger will set true")
					$trigger = True
				EndIf

			WEnd

			Sleep(500)
			Send("+^s")

			$trigger = False
			While ($trigger = False)

				If (WinActive("Speichern unter") = True) Then
					$trigger = True
				EndIf

			WEnd

			Sleep(500)
			Send(StringReplace($pathCSVfromInternal, ".csv", ""))
			Sleep(500)

			Send("{TAB}")
			Sleep(500)
			Send("{DOWN}")
			Sleep(500)
			Send("{UP 15}")
			Sleep(1000)
			Send("{ENTER}")
			Send("{ENTER}")

			Sleep(2000)

			WinKill("OpenOffice")
			WinWaitClose("OpenOffice")

			Local $iSavePathOfODS = StringReplace($pathCSVfromInternal, ".csv", ".ods")
			MsgBox(0, "Test iSaveName", $iSavePathOfODS)
			_KillSpaces_OpenOffice($iSavePathOfODS)

		Case $btnConvert

			Local $trigger = False

			WinActivate("OpenOffice")
			WinActivate("OpenOffice")
			WinActivate("OpenOffice")
			WinActivate("OpenOffice")
			WinWaitActive("OpenOffice")

			Sleep(500)
			Send("+^s")

			While ($trigger = False)

				If (WinActive("Speichern unter") = True) Then
					$trigger = True
				EndIf

			WEnd

			Sleep(500)
			Send(StringReplace($pathCSVfromInternal, ".csv", ""))
			Sleep(500)

			Send("{TAB}")
			Sleep(500)
			Send("{DOWN}")
			Sleep(500)
			Send("{UP 15}")
			Sleep(1000)
			Send("{ENTER}")
			Send("{ENTER}")

			Sleep(2000)

			WinKill("OpenOffice")
			WinWaitClose("OpenOffice")

			Local $iSavePathOfODS = StringReplace($pathCSVfromInternal, ".csv", ".ods")
			_KillSpaces_OpenOffice($iSavePathOfODS)

		Case $btnSave

			;_BlockInputEx(1)

			WinActivate("Calc")
			WinWaitActive("Calc")

			Sleep(500)
			Send("+^s")
			WinWaitActive("Speichern unter")
			Sleep(500)

			Send("{TAB}")
			Sleep(500)
			Send("{DOWN 7}")
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

			;_BlockInputEx(0)

			Local $myPID = ShellExecute($pathConsensusDir & "\Consensus.jar", IniRead($pathIniFile, "Paths", "Layouts", "Can't read path 'Layouts' from section 'Paths' in ini-file.") & "%" & $pathInternalDir, $pathConsensusDir)

			While (Not ProcessExists($myPID) = 0)
				;
			WEnd

			DirRemove($pathInternalDir, 1)
			Sleep(500)
			DirCreate($pathInternalDir)

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

EndFunc   ;==>_Start_Consensus

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

	$btnKillSpaces = GUICtrlCreateButton("KillSpaces() - not ", 0, 274, 150, 100)
	GUICtrlSetFont(-1, 12, 600, 0, "Leelawadee")
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetOnEvent(-1, "_On_Button")

	$btnConvert = GUICtrlCreateButton("KillSpaces()", 246, 137, 150, 100)
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


EndFunc   ;==>_Start_Control_Panel

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

	_BlockInputEx(2)

	Local $time, $hotFileHandler, $pathArchiveBackup, $myTrigger = False, $myDelayTimer = 500, $myCounter = 0

	$pathMendeleyBackup = IniRead($pathIniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file.")
	$pathMendeleyBackupArchiveDir = IniRead($pathIniFile, "Paths", "Mendeley_Backup_Archive", "Can't read key 'Mendeley_Backup_Archive' from section 'Paths' in ini-file.")

	$time = _Date_Time_SystemTimeToFileTime(_Date_Time_GetSystemTime())
	$time = _Date_Time_FileTimeToStr($time)
	$time = StringReplace($time, ":", "_")
	$time = StringReplace($time, "/", "_")
	$time = StringReplace($time, " ", "_")

	$pathArchiveBackup = $pathMendeleyBackupArchiveDir & "\" & "Archive_" & $time & ".zip"

	FileMove($pathMendeleyBackup, $pathArchiveBackup, 1)

	While ((_WinAPI_FileInUse($pathMendeleyBackup) = 1) Or (_WinAPI_FileInUse($pathArchiveBackup) = 1))
		ConsoleWrite("in use: " & $pathMendeleyBackup & @CRLF)
		ConsoleWrite("in use: " & $pathArchiveBackup & @CRLF)
	WEnd

	; Start Mendeley Desktop
	Run($pathMendeleyExe, "", @SW_SHOW)

	; Waiting for the main GUI of Mendeley to start
	While $myTrigger = False
		If (WinExists("Mendeley Desktop") = True) Then
			If (WinActive("Mendeley Desktop") = True) Then
				$myTrigger = True
			Else
				If ($myCounter > 20) Then
					Run($pathMendeleyExe, "", @SW_SHOW)
					$myCounter = 0
				Else
					$myCounter += 1
					Sleep($myDelayTimer)
				EndIf
			EndIf
		EndIf
	WEnd
	$myTrigger = False
	$myCounter = 0

	; Starting Mendeley, waiting for the update-window, kill the update-window and load the backup-window
	While $myTrigger = False
		Sleep($myDelayTimer)
		If (WinExists("Update") = True) Then
			If (WinActive("Update") = True) Then
				ConsoleWrite("1.)")

				Sleep(2000)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)

				; Put in the destination path for the new backup
				ClipPut($pathMendeleyBackup)

				Sleep($myDelayTimer)
				Send("{ALT}")
				Sleep($myDelayTimer)
				Send("{LEFT}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)
				Send("^v")
				Sleep($myDelayTimer * 2)
				Send("{ENTER}")
				Sleep($myDelayTimer * 2)
				Send("{ESC}")
				Sleep($myDelayTimer * 2)
				Send("!{F4}")

				$myTrigger = True
			Else
				If ($myCounter > 20) Then
					Sleep($myDelayTimer)
					Send("{ALT}")
					Sleep($myDelayTimer)
					Send("{LEFT}")
					Sleep($myDelayTimer)
					Send("{DOWN 5}")
					Sleep($myDelayTimer)
					Send("{RIGHT}")
					Sleep($myDelayTimer)
					Send("{ENTER}")
				Else
					$myCounter += 1
					Sleep($myDelayTimer)
				EndIf
			EndIf
		EndIf
	WEnd
	$myTrigger = False
	$myCounter = 0

	_BlockInputEx(0)

EndFunc   ;==>_Start_Mendeley_Create_Backup

Func _Start_Mendeley_with_AutoImport()

	_BlockInputEx(2)

	MsgBox($MB_ICONWARNING, "Auto Import", "Mendeley will start the automatic import process.", 3)

	Local $myTrigger = False, $myDelayTimer = 500, $myCounter = 0

	$pathMendeleyExe = $pathInstallFilesDir & "\Mendeley\Mendeley Desktop\MendeleyDesktop.exe"
	$pathMendeleyBackup = IniRead($pathIniFile, "Paths", "Mendeley_Backup", "Can't read key 'Mendeley_Backup' from section 'Paths' in ini-file.")
	$pathMendeleyPDFdataDir = IniRead($pathIniFile, "Paths", "Mendeley_PDFdata", "Can't read key 'Mendeley_PDFdata' from section 'Paths' in ini-file.")

	Run($pathMendeleyExe, "", @SW_SHOW)

	; Waiting for the main GUI of Mendeley to start
	While $myTrigger = False
		If (WinExists("Mendeley Desktop") = True) Then
			If (WinActive("Mendeley Desktop") = True) Then
				$myTrigger = True
			Else
				If ($myCounter > 20) Then
					Run($pathMendeleyExe, "", @SW_SHOW)
					$myCounter = 0
				Else
					$myCounter += 1
					Sleep($myDelayTimer)
				EndIf
			EndIf
		EndIf
	WEnd
	$myTrigger = False
	$myCounter = 0

	; Starting Mendeley, waiting for the update-window, kill the update-window and load the backup-window
	While $myTrigger = False
		Sleep($myDelayTimer)
		If (WinExists("Update") = True) Then
			If (WinActive("Update") = True) Then

				Sleep($myDelayTimer * 4)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)
				Send("{ALT}")
				Sleep($myDelayTimer)
				Send("{LEFT}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)

				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{RIGHT}")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)
				Send("{UP}")
				Sleep($myDelayTimer)
				Send("{ENTER}")

				$myTrigger = True
			Else
				If ($myCounter > 20) Then
					Sleep($myDelayTimer)
					Send("{ALT}")
					Sleep($myDelayTimer)
					Send("{LEFT}")
					Sleep($myDelayTimer)
					Send("{DOWN 5}")
					Sleep($myDelayTimer)
					Send("{RIGHT}")
					Sleep($myDelayTimer)
					Send("{ENTER}")
				Else
					$myCounter += 1
					Sleep($myDelayTimer)
				EndIf
			EndIf
		EndIf
	WEnd
	$myTrigger = False
	$myCounter = 0

	; In the backup-window, put in the backup-file-path from the .ini-file
	While $myTrigger = False
		Sleep(100)
		If (WinExists("Restore Backup") = True) Then
			If (WinActive("Restore Backup") = True) Then

				Sleep($myDelayTimer)
				_ClipBoard_SetData($pathMendeleyBackup)
				Sleep($myDelayTimer)
				Send("^v")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)

				$myTrigger = True
			EndIf
		EndIf
	WEnd
	$myTrigger = False

	; Click through the backup-options-window and load the backup
	While $myTrigger = False
		Sleep(100)
		If (WinExists("Restore Backup") = True) Then
			If (WinActive("Restore Backup") = True) Then

				Sleep($myDelayTimer)
				Send("{TAB}{SPACE}")
				Sleep($myDelayTimer)
				Send("{TAB}{TAB}{SPACE}")
				Sleep($myDelayTimer)
				Send("{ENTER}")

				$myTrigger = True
			EndIf
		EndIf
	WEnd
	$myTrigger = False

	; Waiting for the reload of Mendeley and click through the start-window of Mendeley
	While $myTrigger = False
		Sleep(100)
		If (WinExists("Welcome to Mendeley Desktop") = True) Then
			If (WinActive("Welcome to Mendeley Desktop") = True) Then

				Sleep($myDelayTimer)
				Send("{ENTER}")

				$myTrigger = True
			EndIf
		EndIf
	WEnd
	$myTrigger = False

	; Waiting for the main GUI of Mendeley to start
	If (WinExists("Mendeley Desktop") = True) Then
		If (WinActive("Mendeley Desktop") = True) Then
			;
		Else
			If ($myCounter > 20) Then
				Run($pathMendeleyExe, "", @SW_SHOW)
				$myCounter = 0
			Else
				$myCounter += 1
				Sleep($myDelayTimer)
			EndIf
		EndIf
	EndIf

	; Starting Mendeley, waiting for the update-window and kill the update-window
	While $myTrigger = False
		If (WinExists("Update") = True) Then
			If (WinActive("Update") = True) Then

				Sleep(2000)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)

				$myTrigger = True
			Else
				If ($myCounter > 20) Then
					Sleep($myDelayTimer)
					Send("{ALT}")
					Sleep($myDelayTimer)
					Send("{LEFT}")
					Sleep($myDelayTimer)
					Send("{DOWN 5}")
					Sleep($myDelayTimer)
					Send("{RIGHT}")
					Sleep($myDelayTimer)
					Send("{ENTER}")
				Else
					$myCounter += 1
					Sleep($myDelayTimer)
				EndIf
			EndIf
		EndIf
	WEnd
	$myTrigger = False
	$myCounter = 0

	; Enter the PDF-data-path within the menu of Mendeley
	While $myTrigger = False
		Sleep(100)
		If (WinExists("Mendeley Desktop") = True) Then
			If (WinActive("Mendeley Desktop") = True) Then

				Sleep($myDelayTimer)
				Send("{ALT}")
				Sleep($myDelayTimer)
				Send("{LEFT}")
				Sleep($myDelayTimer)
				Send("{LEFT}")
				Sleep($myDelayTimer)
				Send("{DOWN}")
				Sleep($myDelayTimer)
				Send("{UP}")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{RIGHT}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				Send("{SPACE}")
				Sleep($myDelayTimer)
				Send("{TAB}")
				Sleep($myDelayTimer)
				_ClipBoard_SetData($pathMendeleyPDFdataDir)
				Sleep($myDelayTimer)
				Send("^v")
				Sleep($myDelayTimer)
				Send("{ENTER}")
				Sleep($myDelayTimer)
				Send("{LEFT}")
				Sleep($myDelayTimer)
				Send("{ENTER}")

				$myTrigger = True
			EndIf
		EndIf
	WEnd

	$myTrigger = False

	_BlockInputEx(0)

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

EndFunc   ;==>_Start_Table_Calculator_with_CSV

Func _Start_Tabula_with_file($pathPDFfile)

	Local $arrayPathCSVfile, $newPathCSVfile, $arrayPathCSVfile, $newPathCSVfile, $orignalNamePath, $myTimer = 500, $myTrigger = False
	$pathTabulaDir = $pathInstallFilesDir & "\Tabula-Win-1.2.0"

	$iPID = Run(@ComSpec & " /k tabula.exe", $pathTabulaDir, @SW_HIDE) ; Execute the Tabula-Win-1.2.0 software (/k means 'keep' (without it does not executed))

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

	Sleep($myTimer * 4)

	ClipPut($pathPDFfile)

	MsgBox(0, "Hint", "Please select the PDF and extract the areas you want.")

	FileChangeDir($pathInternalDir)

	$pathCSVfile = StringReplace($pathPDFfile, ".pdf", ".csv")

	$arrayPathCSVfile = StringSplit($pathCSVfile, "\", 2)

	For $i = 0 To (UBound($arrayPathCSVfile) - 1) Step +1

		If ($i) = (UBound($arrayPathCSVfile) - 1) Then

			$arrayPathCSVfile[$i] = StringReplace($arrayPathCSVfile[$i], $arrayPathCSVfile[$i], "tabula-" & $arrayPathCSVfile[$i])
			$arrayPathCSVfile[$i] = StringReplace($arrayPathCSVfile[$i], " ", "_")

		EndIf

	Next

	FileChangeDir($pathInternalDir)

	$newPathCSVfile = _ArrayToString($arrayPathCSVfile, "\")

	While Not FileExists($newPathCSVfile)

	WEnd

	ClipPut($pathInternalDir)

	; Waiting for the main GUI of Mendeley to start
	While $myTrigger = False
		If (WinExists("Download beendet") = True) Then
			If (WinActive("Download beendet") = True) Then
				While WinExists("Download beendet")
					WinClose("Download beendet")
					$myTrigger = True

					MsgBox(0, "Automatization", "Extracted data will imported into a (freeware) table calculator", 3)
				WEnd
				_BlockInputEx(2)
			EndIf
		EndIf
	WEnd
	$myTrigger = False

	Sleep($myTimer * 2)

	_Process_Tree_Close($iPID)

	Sleep($myTimer * 2)

	GUIDelete($guiEmbeddedBrowser)

	$orignalNamePath = StringReplace($newPathCSVfile, "tabula-", "")

	FileMove($newPathCSVfile, $orignalNamePath)

	While ((_WinAPI_FileInUse($newPathCSVfile) = 1) Or (_WinAPI_FileInUse($orignalNamePath) = 1))
		ConsoleWrite("in use: " & $newPathCSVfile & @CRLF)
		ConsoleWrite("in use: " & $orignalNamePath & @CRLF)
	WEnd

	$newPathCSVfile = $orignalNamePath

	Sleep($myTimer * 2)

	_BlockInputEx(0)

	Return $newPathCSVfile

EndFunc   ;==>_Start_Tabula_with_file

Func _KillSpaces_OpenOffice($pathToTableFile)

	Local $oOpenOffice, $nRows, $nColumns, $trigger
	Local $delay = 500
	Local $triggerCurrent = False, $triggerEqual = False
	Local $sheetListArray
	Local $numberOfLastSheet

	If (Not WinExists("OpenOffice") = True) Then
		Run($pathScalcExe)
		WinWaitActive("OpenOffice")
		Sleep($delay)
		$oOpenOffice = _OOoCalc_BookOpen($pathToTableFile, False, False, "")
		Sleep($delay)
		_OOoCalc_SheetActivate($oOpenOffice, UBound(_OOoCalc_SheetList($oOpenOffice)) - 1)
		Sleep($delay)
		$myArray = _OOoCalc_ReadSheetToArray($oOpenOffice)
		Sleep($delay)

		$nRows = UBound($myArray, 1) - 1
		$nColumns = UBound($myArray, 2) - 1

		Local $triggerCurrentArray[$nRows + 1]
		Local $triggerEqualArray[$nRows + 1]

		ConsoleWrite("nRows: " & $nRows & @CRLF)
		ConsoleWrite("nColumns: " & $nColumns)

		For $row = $nRows To 0 Step -1 ; Loop rows START

			For $column = 0 To $nColumns Step +1 ; Loop columns START

				If ($myArray[$row][$column] = "" Or $myArray[$row][$column] = " ") Then
					$triggerCurrentArray[$row] = True
					ExitLoop
				Else
					$triggerCurrentArray[$row] = False
				EndIf

			Next ; Loop columns END

		Next ; Loop rows END

		For $row = $nRows To 1 Step -1 ; Loop rows START
			$triggerEqualArray[0] = False
			For $column = 0 To $nColumns Step +1 ; Loop columns START

				If ((StringCompare($myArray[$row][$column], $myArray[$row - 1][$column], 1) = 0) And (Not StringCompare($myArray[$row][$column], "", 1) = 0) And (Not StringCompare($myArray[$row - 1][$column], "", 1) = 0)) Then
					$triggerEqualArray[$row] = True
					ExitLoop
				Else
					$triggerEqualArray[$row] = False
				EndIf

			Next ; Loop columns END

		Next ; Loop rows END

		For $row = $nRows To 1 Step -1 ; Loop rows START

			If (($triggerCurrentArray[$row] = True And $triggerEqualArray[$row] = False) And ($myArray[$row - 1][0] == "" Or $myArray[$row - 1][0] == " ")) Then
				ContinueLoop
			Else

				If ($triggerCurrentArray[$row] = False) Then

					ContinueLoop
				Else

					If ($triggerEqualArray[$row] = True) Then
						ContinueLoop
					Else

						For $column = 0 To $nColumns Step +1 ; Loop columns START
							If ($row > 0) Then
								If ($myArray[$row][$column] == "" Or $myArray[$row][$column] == " ") Then
									ContinueLoop
								Else
									$myArray[$row - 1][$column] = $myArray[$row - 1][$column] & " " & $myArray[$row][$column]
									$myArray[$row][$column] = ""
								EndIf
							EndIf

						Next ; Loop columns END

					EndIf
				EndIf
			EndIf

		Next ; Loop rows END

		_Array2DDeleteEmptyRows($myArray)

		Sleep(500)

		_ArrayDisplay($myArray)

		Sleep(500)

		_OOoCalc_SheetAddNew($oOpenOffice, "finalSheet")

		Sleep(500)

		$sheetListArray = _OOoCalc_SheetList($oOpenOffice)

		Sleep(500)

		$numberOfLastSheet = UBound($sheetListArray, 1) - 2

		Sleep($delay)

		_OOoCalc_SheetActivate($oOpenOffice, $numberOfLastSheet)
		Sleep($delay)
		_OOoCalc_WriteFromArray($oOpenOffice, $myArray, "A1", -1, $numberOfLastSheet)

	Else
		$oOpenOffice = _OOoCalc_BookOpen($pathToTableFile, False, False, "")
		Sleep($delay)
		_OOoCalc_SheetActivate($oOpenOffice, UBound(_OOoCalc_SheetList($oOpenOffice)) - 1)
		Sleep($delay)
		$myArray = _OOoCalc_ReadSheetToArray($oOpenOffice)
		Sleep($delay)

		$nRows = UBound($myArray, 1) - 1
		$nColumns = UBound($myArray, 2) - 1

		Local $triggerCurrentArray[$nRows + 1]
		Local $triggerEqualArray[$nRows + 1]

		ConsoleWrite("nRows: " & $nRows & @CRLF)
		ConsoleWrite("nColumns: " & $nColumns)

		For $row = $nRows To 0 Step -1 ; Loop rows START

			For $column = 0 To $nColumns Step +1 ; Loop columns START

				If ($myArray[$row][$column] = "" Or $myArray[$row][$column] = " ") Then
					$triggerCurrentArray[$row] = True
					ExitLoop
				Else
					$triggerCurrentArray[$row] = False
				EndIf

			Next ; Loop columns END

		Next ; Loop rows END

		For $row = $nRows To 1 Step -1 ; Loop rows START
			$triggerEqualArray[0] = False
			For $column = 0 To $nColumns Step +1 ; Loop columns START

				If (($triggerCurrentArray[$row] = True And $triggerEqualArray[$row] = False) And ($myArray[$row - 1][0] == "" Or $myArray[$row - 1][0] == " ")) Then
					$triggerEqualArray[$row] = True
					ExitLoop
				Else
					$triggerEqualArray[$row] = False
				EndIf

			Next ; Loop columns END

		Next ; Loop rows END

		For $row = $nRows To 1 Step -1 ; Loop rows START

			If (($triggerCurrentArray[$row] = True And $triggerEqualArray[$row] = False) And (Not $myArray[$row - 1][0] == $myArray[$row][0])) Then
				ContinueLoop
			Else

				If ($triggerCurrentArray[$row] = False) Then

					ContinueLoop
				Else

					If ($triggerEqualArray[$row] = True) Then
						ContinueLoop
					Else

						For $column = 0 To $nColumns Step +1 ; Loop columns START
							If ($row > 0) Then
								If ($myArray[$row][$column] == "" Or $myArray[$row][$column] == " ") Then
									ContinueLoop
								Else
									$myArray[$row - 1][$column] = $myArray[$row - 1][$column] & " " & $myArray[$row][$column]
									$myArray[$row][$column] = ""
								EndIf
							EndIf

						Next ; Loop columns END

					EndIf
				EndIf
			EndIf

		Next ; Loop rows END

		_Array2DDeleteEmptyRows($myArray)

		Sleep(500)

		_ArrayDisplay($myArray)

		Sleep(500)

		_OOoCalc_SheetAddNew($oOpenOffice, "finalSheet")

		Sleep(500)

		$sheetListArray = _OOoCalc_SheetList($oOpenOffice)

		Sleep(500)

		$numberOfLastSheet = UBound($sheetListArray, 1) - 2

		Sleep($delay)

		_OOoCalc_SheetActivate($oOpenOffice, $numberOfLastSheet)
		Sleep($delay)
		_OOoCalc_WriteFromArray($oOpenOffice, $myArray, "A1", -1, $numberOfLastSheet)

	EndIf

EndFunc   ;==>_KillSpaces_OpenOffice

Func _Array2DDeleteEmptyRows(ByRef $iArray)
	Local $vEmpty = False
	$nColumns = UBound($iArray, 2)
	Local $iArrayOut[1][$nColumns]
	Local $A = 0
	For $row = 0 To (UBound($iArray, 1) - 1) Step +1
		If ($vEmpty = True) Then
			$A += 1
			$vEmpty = False
		EndIf
		For $column = 0 To (UBound($iArray, 2) - 1) Step 1
			If (StringCompare($iArray[$row][$column], "", 0)) <> 0 Then
				ReDim $iArrayOut[$A + 1][$nColumns]
				$iArrayOut[$A][$column] = $iArray[$row][$column]
				$vEmpty = True
			EndIf
		Next
	Next
	$iArray = $iArrayOut
EndFunc   ;==>_Array2DDeleteEmptyRows

Func sha256($message)
	Return _Crypt_HashData($message, $CALG_SHA_256)
EndFunc   ;==>sha256

