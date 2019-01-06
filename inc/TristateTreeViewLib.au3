#include-once

;**********************************************************
; Constants
;**********************************************************
;~ If Not IsDeclared("WM_NOTIFY")				Then	Global Const $WM_NOTIFY					= 0x004E

;~ If Not IsDeclared("LR_LOADFROMFILE")		Then	Global Const $LR_LOADFROMFILE			= 0x0010
;~ If Not IsDeclared("LR_LOADTRANSPARENT")		Then	Global Const $LR_LOADTRANSPARENT		= 0x0020
;~ If Not IsDeclared("LR_CREATEDIBSECTION")	Then	Global Const $LR_CREATEDIBSECTION		= 0x2000

;~ If Not IsDeclared("CLR_NONE")				Then	Global Const $CLR_NONE					= 0xFFFFFFFF
;~ If Not IsDeclared("IMAGE_BITMAP")			Then	Global Const $IMAGE_BITMAP				= 0
;~ If Not IsDeclared("NM_CLICK")				Then	Global Const $NM_CLICK					= -2
If Not IsDeclared("VK_SPACE")				Then	Global Const $VK_SPACE					= 32
;~ If Not IsDeclared("GWL_STYLE")				Then	Global Const $GWL_STYLE					= -16

;~ If Not IsDeclared("TV_FIRST")				Then	Global Const $TV_FIRST					= 0x1100
;~ If Not IsDeclared("TVM_SETIMAGELIST")		Then	Global Const $TVM_SETIMAGELIST			= $TV_FIRST + 9
;~ If Not IsDeclared("TVM_GETNEXTITEM")		Then	Global Const $TVM_GETNEXTITEM			= $TV_FIRST + 10
If Not IsDeclared("TVM_GETITEM")			Then	Global Const $TVM_GETITEM				= $TV_FIRST + 12
If Not IsDeclared("TVM_SETITEM")			Then	Global Const $TVM_SETITEM				= $TV_FIRST + 13
;~ If Not IsDeclared("TVM_HITTEST")			Then	Global Const $TVM_HITTEST				= $TV_FIRST + 17
;~ If Not IsDeclared("TVSIL_STATE")			Then	Global Const $TVSIL_STATE				= 2
;~ If Not IsDeclared("TVGN_NEXT")				Then	Global Const $TVGN_NEXT					= 0x1
;~ If Not IsDeclared("TVGN_PARENT")			Then	Global Const $TVGN_PARENT				= 0x3
;~ If Not IsDeclared("TVGN_CHILD")				Then	Global Const $TVGN_CHILD				= 0x4
;~ If Not IsDeclared("TVGN_CARET")				Then	Global Const $TVGN_CARET				= 0x9
;~ If Not IsDeclared("TVIF_STATE")				Then	Global Const $TVIF_STATE				= 0x0008
;~ If Not IsDeclared("TVIF_HANDLE")			Then	Global Const $TVIF_HANDLE				= 0x0010
;~ If Not IsDeclared("TVIS_STATEIMAGEMASK")	Then	Global Const $TVIS_STATEIMAGEMASK		= 0xF000
;~ If Not IsDeclared("TVHT_ONITEMSTATEICON")	Then	Global Const $TVHT_ONITEMSTATEICON		= 0x0040
;~ If Not IsDeclared("TVN_FIRST")				Then	Global Const $TVN_FIRST					= -400
;~ If Not IsDeclared("TVN_KEYDOWN")			Then	Global Const $TVN_KEYDOWN				= $TVN_FIRST - 12


;**********************************************************
; Register
;**********************************************************
GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")


;**********************************************************
; Set an item state
;**********************************************************
Func MyCtrlGetItemState($hTV, $nID)
	Local $hWnd = GUICtrlGetHandle($hTV)
	If $hWnd = 0 Then $hWnd = $hTV

	Local $hItem = GUICtrlGetHandle($nID)
	If $hItem = 0 Then $hItem = $nID

	$nState = GetItemState($hWnd, $hItem)

	Switch $nState
		Case 1
			$nState = $GUI_UNCHECKED
		Case 2
			$nState = $GUI_CHECKED
		Case 3
			$nState = $GUI_INDETERMINATE
		Case 4
			$nState = BitOr($GUI_DISABLE, $GUI_UNCHECKED)
		Case 5
			$nState = BitOr($GUI_DISABLE, $GUI_CHECKED)
		Case Else
			Return 0
	EndSwitch

	Return $nState
EndFunc


;**********************************************************
; Get an item state
;**********************************************************
Func MyCtrlSetItemState($hTV, $nID, $nState)
	Local $hWnd = GUICtrlGetHandle($hTV)
	If $hWnd = 0 Then $hWnd = $hTV

	Local $hItem = GUICtrlGetHandle($nID)
	If $hItem = 0 Then $hItem = $nID

	Switch $nState
		Case $GUI_UNCHECKED
			$nState = 1
		Case $GUI_CHECKED
			$nState = 2
		Case $GUI_INDETERMINATE
			$nState = 3
		Case BitOr($GUI_DISABLE, $GUI_UNCHECKED)
			$nState = 4
		Case BitOr($GUI_DISABLE, $GUI_CHECKED)
			$nState = 5
		Case Else
			Return
	EndSwitch

	SetItemState($hWnd, $hItem, $nState)

	CheckChildItems($hWnd, $hItem, $nState)
	CheckParents($hWnd, $hItem, $nState)

EndFunc


;**********************************************************
; MY_WM_NOTIFY
;**********************************************************
Func MY_WM_NOTIFY($hWnd, $nMsg, $wParam, $lParam)
	Local $stNmhdr		= DllStructCreate("dword;int;int", $lParam)
	Local $hWndFrom		= DllStructGetData($stNmhdr, 1)
	Local $nNotifyCode	= DllStructGetData($stNmhdr, 3)
	Local $hItem		= 0

	; Check if its treeview and only NM_CLICK and TVN_KEYDOWN
	If Not BitAnd(GetWindowLong($hWndFrom, $GWL_STYLE), $TVS_CHECKBOXES) Or _
		Not ($nNotifyCode = $NM_CLICK Or $nNotifyCode = $TVN_KEYDOWN) Then Return $GUI_RUNDEFMSG

	If $nNotifyCode = $TVN_KEYDOWN Then
		Local $lpNMTVKEYDOWN = DllStructCreate("dword;int;int;short;uint", $lParam)

		; Check for 'SPACE'-press
		If DllStructGetData($lpNMTVKEYDOWN, 4) <> $VK_SPACE Then Return $GUI_RUNDEFMSG
		$hItem = SendMessage($hWndFrom, $TVM_GETNEXTITEM, $TVGN_CARET, 0)
	Else
		Local $Point = DllStructCreate("int;int")

		GetCursorPos($Point)
		ScreenToClient($hWndFrom, $Point)

		; Check if clicked on state icon
		Local $tvHit = DllStructCreate("int[2];uint;dword")
		DllStructSetData($tvHit, 1, DllStructGetData($Point, 1), 1)
		DllStructSetData($tvHit, 1, DllStructGetData($Point, 2), 2)

		$hItem = SendMessage($hWndFrom, $TVM_HITTEST, 0, DllStructGetPtr($tvHit))

		If Not BitAnd(DllStructGetData($tvHit, 2), $TVHT_ONITEMSTATEICON) Then Return $GUI_RUNDEFMSG
	EndIf

	If $hItem > 0 Then
		Local $nState = GetItemState($hWndFrom, $hItem)

		$bCheckItems = 1

		If $nState = 1 Then
			$nState = 1
		ElseIf $nState = 2 Then
			$nState = 0
		ElseIf $nState = 3 Then
			$nState = 1
		ElseIf $nState > 3 Then
			$nState = $nState - 1
			$bCheckItems = 0
		EndIf

		SetItemState($hWndFrom, $hItem, $nState)

		$nState += 1

		; If item are disabled there is no chance to change it and it's parents/children
		If $bCheckItems Then
			CheckChildItems($hWndFrom, $hItem, $nState)
			CheckParents($hWndFrom, $hItem, $nState)
		EndIf
	EndIf
EndFunc


;**********************************************************
; Helper functions
;**********************************************************
Func CheckChildItems($hWnd, $hItem, $nState)
	Local $hChild = SendMessage($hWnd, $TVM_GETNEXTITEM, $TVGN_CHILD, $hItem)

	While $hChild > 0
		SetItemState($hWnd, $hChild, $nState)
		CheckChildItems($hWnd, $hChild, $nState)

		$hChild = SendMessage($hWnd, $TVM_GETNEXTITEM, $TVGN_NEXT, $hChild)
	WEnd
EndFunc


Func CheckParents($hWnd, $hItem, $nState)
	Local $nTmpState1 = 0, $nTmpState2 = 0
	Local $bDiff = 0
	Local $i = 0

	Local $hParent = SendMessage($hWnd, $TVM_GETNEXTITEM, $TVGN_PARENT, $hItem)

	If $hParent > 0 Then
		Local $hChild = SendMessage($hWnd, $TVM_GETNEXTITEM, $TVGN_CHILD, $hParent)

		If $hChild > 0 Then
			Do
				$i = $i + 1

				If $hChild = $hItem Then
					$nTmpState2 = $nState
				Else
					$nTmpState2 = GetItemState($hWnd, $hChild)
				EndIf

				If $i = 1 Then $nTmpState1 = $nTmpState2

				If $nTmpState1 <> $nTmpState2 Then
					$bDiff = 1
					ExitLoop
				EndIf

				$hChild = SendMessage($hWnd, $TVM_GETNEXTITEM, $TVGN_NEXT, $hChild)
			Until $hChild <= 0

			If $bDiff Then
				SetItemState($hWnd, $hParent, 3)
				$nState = 3
			Else
				SetItemState($hWnd, $hParent, $nState)
			EndIf

		EndIf

		CheckParents($hWnd, $hParent, $nState)
	EndIf
EndFunc


Func SetItemState($hWnd, $hItem, $nState)
	$nState = BitShift($nState, -12)

	Local $tvItem = DllStructCreate("uint;dword;uint;uint;ptr;int;int;int;int;int;int")

	DllStructSetData($tvItem, 1, $TVIF_STATE)
	DllStructSetData($tvItem, 2, $hItem)
	DllStructSetData($tvItem, 3, $nState)
	DllStructSetData($tvItem, 4, $TVIS_STATEIMAGEMASK)

	SendMessage($hWnd, $TVM_SETITEM, 0, DllStructGetPtr($tvItem))
EndFunc


Func GetItemState($hWnd, $hItem)
	Local $tvItem = DllStructCreate("uint;dword;uint;uint;ptr;int;int;int;int;int;int")

	DllStructSetData($tvItem, 1, $TVIF_STATE)
	DllStructSetData($tvItem, 2, $hItem)
	DllStructSetData($tvItem, 4, $TVIS_STATEIMAGEMASK)

	SendMessage($hWnd, $TVM_GETITEM, 0, DllStructGetPtr($tvItem))

	Local $nState = DllStructGetData($tvItem, 3)

	$nState = BitAnd($nState, $TVIS_STATEIMAGEMASK)
	$nState = BitShift($nState, 12)

	Return $nState
EndFunc


Func LoadStateImage($hTreeView, $sFile)
	Local $hWnd = GUICtrlGetHandle($hTreeView)
	If $hWnd = 0 Then $hWnd = $hTreeView

	Local $hImageList = 0

	If @Compiled Then
		Local $hModule = LoadLibrary(@ScriptFullPath)
		$hImageList = ImageList_LoadImage($hModule, "#170", 16, 1, $CLR_NONE, $IMAGE_BITMAP, BitOr($LR_LOADTRANSPARENT, $LR_CREATEDIBSECTION))
	Else
		$hImageList = ImageList_LoadImage(0, $sFile, 16, 1, $CLR_NONE, $IMAGE_BITMAP, BitOr($LR_LOADFROMFILE, $LR_LOADTRANSPARENT, $LR_CREATEDIBSECTION))
	EndIf

	SendMessage($hWnd, $TVM_SETIMAGELIST, $TVSIL_STATE, $hImageList)
	InvalidateRect($hWnd, 0, 1)
EndFunc


;**********************************************************
; Win32-API functions
;**********************************************************
Func SendMessage($hWnd, $Msg, $wParam, $lParam)
	$nResult = DllCall("user32.dll", "int", "SendMessage", _
											"hwnd", $hWnd, _
											"int", $Msg, _
											"int", $wParam, _
											"int", $lParam)
	Return $nResult[0]
EndFunc


Func GetWindowLong($hWnd, $nIndex)
	$nResult = DllCall("user32.dll", "int", "GetWindowLong", "hwnd", $hWnd, "int", $nIndex)
	Return $nResult[0]
EndFunc


Func GetCursorPos($Point)
	DllCall("user32.dll", "int", "GetCursorPos", "ptr", DllStructGetPtr($Point))
EndFunc


Func ScreenToClient($hWnd, $Point)
    DllCall("user32.dll", "int", "ScreenToClient", "hwnd", $hWnd, "ptr", DllStructGetPtr($Point))
EndFunc


Func InvalidateRect($hWnd, $lpRect, $bErase)
	DllCall("user32.dll", "int", "InvalidateRect", _
								"hwnd", $hWnd, _
								"ptr", $lpRect, _
								"int", $bErase)
EndFunc


Func LoadLibrary($sFile)
	Local $hModule = DllCall("kernel32.dll", "hwnd", "LoadLibrary", "str", $sFile)
	Return $hModule[0]
EndFunc


Func ImageList_LoadImage($hInst, $sFile, $cx, $cGrow, $crMask, $uType, $uFlags)
	Local $hImageList = DllCall("comctl32.dll", "hwnd", "ImageList_LoadImage", _
														"hwnd", $hInst, _
														"str", $sFile, _
														"int", $cx, _
														"int", $cGrow, _
														"int", $crMask, _
														"int", $uType, _
														"int", $uFlags)
	Return $hImageList[0]
EndFunc


Func DestroyImageList()
;~ 	DllCall("comctl32.dll", "int", "ImageList_Destroy", "hwnd", $hImageList)
EndFunc