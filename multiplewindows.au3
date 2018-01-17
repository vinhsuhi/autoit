#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>

Global $g_idButton3 = 9999

gui1()

Func gui1()
	Local $hGUI1 = GUICreate("Gui 1", 200, 200, 100, 100)
	Local $idButton1 = GUICtrlCreateButton("Msgbox 1", 10, 10, 80, 30)
	Local $idButton2 = GUICtrlCreateButton("Show Gui 2", 10, 60, 80, 30)
	GUISetState()

	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop
			Case $idButton1
				MsgBox($MB_OK, "MsgBox 1", "Test from Gui 1")
			Case $idButton2
				GUICtrlSetState($idButton2, $GUI_DISABLE)
				gui2()
			Case $g_idButton3
				MsgBox($MB_OK, "MsgBox 2", "Test from Gui 2")
		EndSwitch
	WEnd
EndFunc   ;==>gui1

Func gui2()
	Local $hGUI2 = GUICreate("Gui 2", 200, 200, 350, 350)
	Global $g_idButton3 = GUICtrlCreateButton("MsgBox 2", 10, 10, 80, 30)
	GUISetState()
EndFunc   ;==>gui2