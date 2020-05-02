#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GuiComboBox.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiDateTimePicker.au3>
#include <GuiDateTimePicker.au3>
#include <StaticConstants.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <StringConstants.au3>
#include <GuiEdit.au3>


#include <Date.au3>
#Region ### START Koda GUI section ### Form=C:\Users\VM User\Documents\GitHub\TellMyAge\frmMain.kxf
Opt("GUIOnEventMode", 1)
$frmMain = GUICreate("Tell My Age", 612, 451, 365, 125, BitOR($GUI_SS_DEFAULT_GUI,$WS_SIZEBOX,$WS_THICKFRAME))
GUISetIcon("TellMyAge.ico", -1)
GUISetOnEvent($GUI_EVENT_CLOSE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_RESTORE, "SpecialEvents")

$grpDOB = GUICtrlCreateGroup("Date Of Birth", 8, 8, 593, 121)
$g_hDTP = _GUICtrlDTP_Create($frmMain, 24, 72, 190)

GUICtrlCreateGroup("", -99, -99, 1, 1)
$grpAge = GUICtrlCreateGroup("", 8, 136, 593, 249)
$edtAgeDisplay = GUICtrlCreateEdit("", 16, 152, 577, 225)
GUICtrlSetData(-1, StringFormat("You are:\r\n0 Years, \r\n0 Months,\r\nand 0 days old."))
GUICtrlSetFont(-1, 20, 800, 0, "Arial")
GUICtrlSetColor(-1, 0x808080)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$btnCalculate = GUICtrlCreateButton("CALCULATE", 216, 400, 137, 41)
GUICtrlSetOnEvent($btnCalculate, "ShowAge")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
; Just idle around
While 1
	Sleep(10)
WEnd

Func SetDateCombo($Y = @YEAR,$M = @MON)
Local $dS = "|"
Local $nD = _DateDaysInMonth($Y, $M)
For $c = 1 To $nD
	$dS &= "|" & $c
Next
Return $dS
EndFunc

Func ShowAge()
	$g_aDate = _GUICtrlDTP_GetSystemTime($g_hDTP)

	$dob_d = $g_aDate[2]
	$dob_m = $g_aDate[1]
	$dob_y = $g_aDate[0]

	If Number($dob_y) > @YEAR Then
		MsgBox($MB_SYSTEMMODAL+$MB_OK+$MB_ICONWARNING,"Invalid DOB","The selected Year of birth is invalid")
		Return
	EndIf
	$bday_yr = _DateDiff('Y', $dob_y & "/" & $dob_m & "/" & $dob_d & " 00:00:00", _NowCalc())
	$bday_m =  _DateDiff('M', @YEAR & "/" & $dob_m & "/" & $dob_d & " 00:00:00",_NowCalc())
	$bday_d =  _DateDiff('D', @YEAR & "/" & @MON & "/" & $dob_d & " 00:00:00",_NowCalc())
	_GUICtrlEdit_SetText($edtAgeDisplay, "You are:"& @CRLF & $bday_yr & " Years, " & @CRLF & $bday_m & " Months,"& @CRLF & " and " & $bday_d & " days old." )

EndFunc

Func SpecialEvents()
    Select
        Case @GUI_CtrlId = $GUI_EVENT_CLOSE
            If MsgBox($MB_SYSTEMMODAL+$MB_OKCANCEL+$MB_ICONQUESTION+$MB_DEFBUTTON1, "Confirm Exit", "Press OK to exit Application",5)== $IDOK Then
				_GUICtrlDTP_Destroy($g_hDTP)
				GUIDelete()
				Exit
			EndIf

        Case @GUI_CtrlId = $GUI_EVENT_MINIMIZE
        Case @GUI_CtrlId = $GUI_EVENT_RESTORE

    EndSelect
EndFunc   ;==>SpecialEvents
