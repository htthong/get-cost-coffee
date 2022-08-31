#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.16.0
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

; Open a browser to the basic example, get an object reference
; to the DIV element with the ID "line1". Display the innerText
; of this element to the console.



#include <IE.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <Excel.au3>
#include <Date.au3> 		;_NowDate()): today
#include <Inet.au3>

#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <DateTimeConstants.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <GuiListView.au3>
#include <ColorConstants.au3>


Global $url = 'https://giacaphe.com/gia-ca-phe-noi-dia'
Global $aTableData = getTableData($url)
Global $ListArea = _ArrayToString($aTableData, "|", 1, Default, "|", 0,0)



#Region ### START Koda GUI section ### Form=
Global $FormMain = GUICreate("Giá cà phê", 364, 511, -1, -1)
GUISetFont(10, 400, 0, "Segoe UI")
Global $ButtonGetData = GUICtrlCreateButton("Tải dữ liệu", 224, 24, 99, 25)
Global $ButtonStop = GUICtrlCreateButton("Dừng tải", 224, 64, 99, 25)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $LabelFrom   = GUICtrlCreateLabel("Bắt đầu:", 8, 16, 57, 21)
Global $LabelTo     = GUICtrlCreateLabel("Tới:",     8, 48, 57, 21)
Global $LabelFormat = GUICtrlCreateLabel("Tỉnh :",   8, 80, 57, 21)
Global $DateFrom = GUICtrlCreateDate("2022/08/10 21:42:23", 72, 16, 126, 25, BitOR($GUI_SS_DEFAULT_DATE,$DTS_UPDOWN,$DTS_TIMEFORMAT))
Global $DateTo = GUICtrlCreateDate("2022/08/10 21:42:23", 72, 48, 126, 25, BitOR($GUI_SS_DEFAULT_DATE,$DTS_UPDOWN,$DTS_TIMEFORMAT))
Global $TimesType = "yyyy/MM/dd"
GUICtrlSendMsg($DateFrom, $DTM_SETFORMATW, 0, $TimesType)
GUICtrlSendMsg($DateTo, $DTM_SETFORMATW, 0, $TimesType)
Global $ComboFormat = GUICtrlCreateCombo("Chọn khu vực...", 72, 80, 126, 25)
GUICtrlSetData(-1, $ListArea)
Global $ListView1 = GUICtrlCreateListView("Thời gian|TT nhân xô|Giá trung bình|Thay đổi", 0, 112, 362, 398, -1, $LVS_EX_FULLROWSELECT)

; Create Right-click options
Global $hGuicontext = GUICtrlCreateContextMenu()
Global $MenuItem1 = GUICtrlCreateMenuItem("MenuItem1", $hGuicontext)
Global $MenuItem2 = GUICtrlCreateMenuItem("MenuItem2", $hGuicontext)
Global $MenuItem3 = GUICtrlCreateMenuItem("MenuItem3", $hGuicontext)


GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

GUIRegisterMsg($WM_COMMAND, "_WM_COMMAND")
Global $running = True

If _GetIP() = -1 Then
	MsgBox(16  + 262144,"", "Không có kết nối mạng")
	;Exit
EndIf

While 1
	$nMsg = GUIGetMsg()

	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonGetData
			_GUICtrlListView_DeleteAllItems($ListView1)
			Local $type = GUICtrlRead($ComboFormat)
			If $type == "Chọn khu vực..." Then
				MsgBox(16 + 262144, "Lỗi!! Chưa chọn tỉnh", "Vui lòng chọn tỉnh!!")
				ContinueCase
			EndIf

			Local $timeFrom = GUICtrlRead($DateFrom)
			Local $timeEnd = GUICtrlRead($DateTo)
			$timeEnd = _DateDiff('d', $timeEnd, _NowCalcDate()) > 0 ? $timeEnd : _NowCalcDate()
			GUICtrlSetData($DateTo, $timeEnd)
			showRow($timeFrom, $timeEnd)
	EndSwitch
WEnd


;### Format date : dd-MM-yyy ###
Func DateFormat($sDate)
    Local $aDate, $aTime
    _DateTimeSplit ($sDate, $aDate, $aTime)
	Return $aDate[1]& '-' & ((StringLen($aDate[2])=1)?('0' & $aDate[2]):($aDate[2])) & '-' & ((StringLen($aDate[3])=1)?('0' & $aDate[3]):($aDate[3]))
EndFunc


Func getTableData($ref)
	ConsoleWrite($ref & @CRLF)
	Local $oIE = _IECreate($ref, 0, 0)
	Local $oTable = _IETableGetCollection($oIE, 0)
	if $oTable == 0 Then
		return 0
	EndIf
	Local $aTableData = _IETableWriteToArray($oTable, True)
	_IEQuit($oIE)
	ProcessClose("iexplore.exe")
	return $aTableData
EndFunc


Func getCost($ref, $type)
	Local $aTableData = getTableData($ref)
	if $aTableData == 0 Then
		Local $res = ["Không có dữ liệu", "", ""]
		Return $res
	ElseIf $aTableData == -1 Then
		MsgBox(16 + 262114,0, "here")
		Local $res = ["", "", ""]
		return $res
	EndIf
	Local $value = $type
	Local $iIndex = _ArraySearch($aTableData, $value, 0, 0, 0, 1, 1, 0)
	Local $res = [$aTableData[$iIndex][0], $aTableData[$iIndex][4], $aTableData[$iIndex][5]]
	return $res
EndFunc


Func showRow($timeFrom, $timeEnd)
	GUICtrlSetState($ButtonGetData, $GUI_DISABLE)
	GUICtrlSetState($ButtonStop, $GUI_ENABLE)
	Local $color = False
	Do
		$color = Not $color
		Local $date = DateFormat($timeFrom)
		Local $ref = $url & '-ngay-' & $date
		Local $row = getCost($ref, $type)
		GUICtrlCreateListViewItem($timeFrom & '|' & $row[0] & '|' & $row[1] & '|' & $row[2], $ListView1)
		If $color Then
			GUICtrlSetBkColor(-1, "0xeeeeee")
		EndIf
		$timeFrom = _DateAdd('d', 1, $timeFrom)
	Until _DateDiff('d', $timeFrom, $timeEnd) < 0 Or $running = False
	MsgBox(64 +262144, "", "Đã dừng!!")
	$running = True
	GUICtrlSetState($ButtonGetData, $GUI_ENABLE)
	GUICtrlSetState($ButtonStop, $GUI_DISABLE)
EndFunc

Func _WM_COMMAND($hWnd, $Msg, $wParam, $lParam)
    ; The Stop button was pressed so set the flag
    If BitAND($wParam, 0x0000FFFF) = $ButtonStop Then
        $running = False
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc   ;==>_WM_COMMAND