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

Opt("WinTitleMatchMode", 2)

Global $OOobject, $sheetToArray, $nRows, $nColumns



Run("C:\Test_Extraktor\InstallFiles\OpenOffice\program\scalc.exe")

If(WinWaitActive("OpenOffice") = true) Then
	Sleep(500)
	$OOobject = _OOoCalc_BookOpen("C:\Users\Ossenbrueck\Desktop\myDoc.ods")
EndIf

$arrayOfSheet =  _OOoCalc_ReadSheetToArray($OOobject)

$nRows = UBound($arrayOfSheet, 1) - 1
$nColumns = UBound($arrayOfSheet, 2) - 1

ConsoleWrite("n Rows: " & $nRows)
ConsoleWrite(@CRLF)
ConsoleWrite("n Columns: " & $nColumns)
ConsoleWrite(@CRLF)

For $rows = $nRows To 0 Step -1
	For $columns = $nColumns To 0 Step -1

	Next
	ConsoleWrite($rows)
	ConsoleWrite(" > ")
Next