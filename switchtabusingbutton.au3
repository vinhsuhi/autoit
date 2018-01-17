#include <guiconstants.au3>

$main = GUICreate('main', 200, 300)
$button = GUICtrlCreateButton('Switch', 70, 245, 50, 30)
$tab = GUICtrlCreateTab(10, 10, 180, 230)
$tab1 = GUICtrlCreateTabItem('one')
$tab2 = GUICtrlCreateTabItem('two')
$tab3 = GUICtrlCreateTabItem('three')
GUICtrlCreateTabItem('')

GUISetState()

While 1
    $Msg = GUIGetMsg()
    If $Msg = - 3 Then Exit
    If $Msg = $button Then
        GUICtrlSetState($tab2, $GUI_SHOW)
    EndIf

WEnd