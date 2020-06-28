#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=TellMyAge.ico
#AutoIt3Wrapper_Outfile=TellMyAge.Exe
#AutoIt3Wrapper_Outfile_x64=TellMyAge-x64.Exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Comment=Get updates at https://github.com/vinuf666/TellMyAge
#AutoIt3Wrapper_Res_Description=Calculates precise age in Years, months days.
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=©2020 Vinu Felix.Released under GPLv3
#AutoIt3Wrapper_Res_Field=Product name|Tell My Age Utility
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs
	TellMyAge
	Version 1.1 beta
	Google style Age calculator based on Date Of Birth
	Author Vinu Felix

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
#ce
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

#Region ### START Koda GUI section ### Form=frmMain.kxf
Opt("GUIOnEventMode", 1)
$frmMain = GUICreate("Tell My Age", 612, 420, 365, 125, BitOR($GUI_SS_DEFAULT_GUI,$WS_SIZEBOX,$WS_THICKFRAME))
GUISetIcon("TellMyAge.ico", -1)
GUISetOnEvent($GUI_EVENT_CLOSE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "SpecialEvents")
GUISetOnEvent($GUI_EVENT_RESTORE, "SpecialEvents")

$grpDOB = GUICtrlCreateGroup("Date Of Birth", 8, 8, 593, 60)
$g_hDTP = _GUICtrlDTP_Create($frmMain, 24, 32, 190)
Global $g_aRange[14] = [False, @YEAR, 1, 1, 21, 45, 32, True, @YEAR,@MON, @MDAY, @HOUR, @MIN, @SEC]
; Set date range
_GUICtrlDTP_SetRange($g_hDTP, $g_aRange) ;date selection range from range is ignored
GUICtrlCreateGroup("", -99, -99, 1, 1)
$grpAge = GUICtrlCreateGroup("", 8, 75, 590, 249)
$edtAgeDisplay = GUICtrlCreateEdit("", 16, 90, 575, 225)
GUICtrlSetData(-1, StringFormat(""))
GUICtrlSetFont(-1, 20, 800, 0, "Arial")
GUICtrlSetColor(-1, 0x808080)
GUICtrlCreateGroup("", -99, -99, 1, 1)
$btnCalculate = GUICtrlCreateButton("CALCULATE", 216, 350, 137, 41)
GUICtrlSetOnEvent($btnCalculate, "ShowAge")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###
; waiting for user action

While 1
	Sleep(10)
WEnd



Func ShowAge()
	$g_aDate = _GUICtrlDTP_GetSystemTime($g_hDTP)

	$dob_d = $g_aDate[2]
	$dob_m = $g_aDate[1]
	$dob_y = $g_aDate[0]
	If( _DateDaysInMonth(@YEAR,@MON) < $dob_d Or @MDAY < $dob_d) Then ;if it falls within a month calculate from last month
		$refDate = @YEAR & "/"& (@MON - 1) & "/" & $dob_d & " 00:00:00"
	Else
		$refDate = @YEAR & "/"& @MON & "/" & $dob_d & " 00:00:00"
	EndIf

	$fullmonths = _DateDiff('M', $dob_y & "/" & $dob_m & "/" & $dob_d & " 00:00:00", $refDate)

	If $fullmonths > 12 Then
    ;The date is more than a year before. Take year into consideration
		$bday_yr = Int($fullmonths/12) ;Ignore fractions
		$bday_m = $fullmonths - ($bday_yr *12)
		$tmpD = _DateAdd('M',$fullmonths,$dob_y & "/" & $dob_m & "/" & $dob_d & " 00:00:00")
		$bday_d = _DateDiff('D',$tmpD,_NowCalcDate())
	ElseIf $fullmonths < 0 Then ;born this month
		$bday_yr = 0
		$bday_m = 0
		$bday_d = _DateDiff('M', $dob_y & "/" & $dob_m & "/" & $dob_d & " 00:00:00", @YEAR & "/"& @MON & "/" & @MDAY & " 00:00:00")
	Else
;within a year dob
		$bday_yr = 0
		$bday_m = $fullmonths
		$tmpD = _DateAdd('M',$fullmonths,$dob_y & "/" & $dob_m & "/" & $dob_d & " 00:00:00")
		$bday_d = _DateDiff('D',$tmpD,_NowCalcDate())

	EndIf


	_GUICtrlEdit_SetText($edtAgeDisplay, "You are:"& @CRLF & $bday_yr & " Years, " & @CRLF & $bday_m & " Months,"& @CRLF & " and " & $bday_d & " days old." )
	Local $bDOW = _DateToDayOfWeek(@YEAR, $dob_m, $dob_d)
	Switch $bDOW
		Case 1
			_GUICtrlEdit_AppendText($edtAgeDisplay, @CRLF & "This year your birthday falls on a Sunday")
		Case 2
			_GUICtrlEdit_AppendText($edtAgeDisplay, @CRLF & "This year your birthday falls on a Monday")
		Case 3
			_GUICtrlEdit_AppendText($edtAgeDisplay, @CRLF & "This year your birthday falls on a Tuesday")
		Case 4
			_GUICtrlEdit_AppendText($edtAgeDisplay, @CRLF & "This year your birthday falls on a Wednesday")
		Case 5
			_GUICtrlEdit_AppendText($edtAgeDisplay, @CRLF & "This year your birthday falls on a Thursday")
		Case 6
			_GUICtrlEdit_AppendText($edtAgeDisplay, @CRLF & "This year your birthday falls on a Friday")
		Case 7
			_GUICtrlEdit_AppendText($edtAgeDisplay, @CRLF & "This year your birthday falls on a Saturday")
	EndSwitch

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
