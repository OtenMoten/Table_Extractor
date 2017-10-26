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

Global $guiLoadingScreen, $guiProgressLoadingScreen, $i

Opt("WinTitleMatchMode", 2)
Opt("GUIOnEventMode", 1)

;==> Values for altering the progressbar are triggered in '_Start_File_Install()' GUICtrlSetData($progressbarLoadingScreen, $i)
$guiLoadingScreen = GUICreate("Starting Table Extractor", 300, 40, 100, 200)
GUISetOnEvent($GUI_EVENT_CLOSE, "_On_Close", $guiLoadingScreen)
$guiProgressLoadingScreen = GUICtrlCreateProgress(10, 10, 280, 20)
GUICtrlSetColor($guiProgressLoadingScreen, $COLOR_GREEN) ; not working with Windows XP Style
GUISetState(@SW_SHOW, $guiLoadingScreen)

$i = 0
while ($i < 100)
	GUICtrlSetData($guiProgressLoadingScreen, $i)
	$i = $i + 1
	sleep(50)
WEnd

GUIDelete($guiLoadingScreen)

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