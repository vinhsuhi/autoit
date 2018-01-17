#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>

Global $ChonChungCu = 9999, $Cap  =9999, $OK1 = 9999, $OK2 = 9999, $OK3 = 9999, $OK4 = 9999, $Button3CM = 9999, $Button4CM = 9999
Global $CapMoi, $DiaChiDat
Global $recordID, $recordDate ; record
Global $UserName1, $UserID1[2], $IDdate1, $IDplace1, $UserBirth1, $UserSex1, $UserAddress1 ; User1
Global $UserName2, $UserID2[2], $IDdate2, $IDplace2, $UserBirth2, $UserAddress2 ; User2
Global $propertyAddress, $propertyArea, $DienTichSuDungCC, $ProjectName, $InvestorName, $startYear, $nFloors, $vitritheoBangGiaDat ; property Chung cư
Global $contracName, $contractID
Global $local, $phuongName, $quanName
Global $sellerName, $PriceInContract, $deliveryPaperDate, $priceInLiquidation, $totalPrice
Global $landArea, $oldPaperNumber, $oldLandID ; THUA DAT
Global $editLandArea, $editOldLandID, $editOldPaperNumber ; THỬA ĐẤT
Global $editUserName1, $editUserBirth1, $Radio1CM, $cmnd1, $editUserID1, $editIDPlace1, $editIDdate1, $editUserAddress1
Global $editUserName2, $editUserBirth2, $cmnd2, $editUserID2, $editIDPlace2, $editIDDate2, $editUserAddress2
Global $Radio3CM, $originalOfLand

Global $editPropertyAddress, $editDienTichSuDungCC, $editDienTichSanCC

; ###########

; ###########

;############
; NHÀ KHÔNG CHUNG CƯ
;############
Global $editDienTichSan, $editDienTichSuDung, $editDienTichXayDung, $comboNha
Global $dienTichSanKCC, $dienTichSuDungKCC, $DienTichXayDungKCC, $loaiNhaO
Global $lamPhieuChuyen = 9999, $Form1
ShowUpFirst()

Func ShowUpFirst()
	#Region ### START Koda GUI section ### Form=d:\it\autoit\vilis\dia chi dat.kxf
	$DiaChiDat = GUICreate("DiaChiDat", 437, 376, 280, 182)
	$DiaChiGoc = GUICtrlCreateLabel("Dia Chi Goc", 24, 80, 61, 17)
	$editLocal = GUICtrlCreateInput("", 24, 112, 321, 21)
	$labelQuanHuyen = GUICtrlCreateLabel("Quan Huyen", 24, 144, 64, 17)
	$editQuanHuyen = GUICtrlCreateInput("", 24, 176, 321, 21)
	$OK0 = GUICtrlCreateButton("OK", 24, 328, 91, 33)
	$Group1F = GUICtrlCreateGroup("Hinh Thuc", 24, 216, 305, 97)
	$CapMoiF = GUICtrlCreateRadio("Cap Moi", 40, 240, 81, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$CapDoiF = GUICtrlCreateRadio("Cap Doi", 40, 272, 113, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$Label1F = GUICtrlCreateLabel("So Ho So", 24, 40, 50, 17)
	$editRecordId = GUICtrlCreateInput("", 88, 32, 121, 21)
	$Label2F = GUICtrlCreateLabel("Ngay", 224, 40, 29, 17)
	$editRecordDate = GUICtrlCreateInput("", 264, 32, 121, 21)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	while 1
		$aMsg = GUIGetMsg(1)
		Switch $aMsg[1]
			Case $DiaChiDat
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						ExitLoop
					Case $OK0
						GUISetState($OK0, $GUI_DISABLE)
						$recordID = GUICtrlRead($editRecordId)
						$recordDate = GUICtrlRead($editRecordDate)
						$local = GUICtrlRead($editLocal)
						Local $quanHuyen = StringSplit(GUICtrlRead($editQuanHuyen), ",")
						$quanName = StringStripWS($quanHuyen[0], $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
						$huyenName = StringStripWS($quanHuyen[1], $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
						Local $isCapMoi = GUICtrlRead($CapMoiF)
						If $recordID = "" Or $recordDate = "" Or $local = "" Then
							MsgBox(0, "Thiếu thông tin", "Vui lòng nhập hết thông tin")
						ElseIf $isCapMoi = 0 Then
							MsgBox(0, 0, "cấp đổi")
						ElseIf $isCapMoi = 1 Then
							; DO SOMETHING ABOUT INTERACTIVE BETWEEN TWO SOFTWARES
							ShowUpCapMoi()
						EndIf

				EndSwitch
			Case $CapMoi
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($CapMoi)
						GUICtrlSetState($OK0, $GUI_ENABLE)
					Case $OK1
						$UserName1 = GUICtrlRead($editUserName1)
						$UserID1[1] = GUICtrlRead($editUserID1)
						$UserID1[0] = GUICtrlRead($cmnd1)
						$IDdate1 = GUICtrlRead($editIDdate1)
						$UserBirth1 = GUICtrlRead($editUserBirth1)
						$UserAddress1 = GUICtrlRead($editUserAddress1)
						$IDplace1 = GUICtrlRead($editIDPlace1)
						If GUICtrlRead($Radio1CM) = 1 Then
							$UserSex1 = "o" ; GIOI TINH NAM THI LA "o"
						Else
							$UserSex1 = "b"
						EndIf

						$UserName2 = GUICtrlRead($editUserName2)
						$UserID2[1] = GUICtrlRead($editUserID2)
						$UserID2[0] = GUICtrlRead($cmnd2)
						$IDdate2 = GUICtrlRead($editIDdate2)
						$UserBirth2 = GUICtrlRead($editUserBirth2)
						$UserAddress2 = GUICtrlRead($editUserAddress2)
						$IDplace2 = GUICtrlRead($editIDPlace2)

						If $UserName1 = "" Or $UserID1[1] = "" Then
							MsgBox(0, 0, "Tên người sử dụng đất hoặc CMND không được để trống")
						Else
							MsgBox(0, 0, "ahihi")
							; DO SOMETHING INTERACTIVE BETWEEN TWO SOFTWARE
						EndIf
					Case $OK2
						$landArea = GUICtrlRead($editLandArea)
						$oldPaperNumber = GUICtrlRead($editOldPaperNumber)
						$oldLandID = GUICtrlRead($editOldLandID)
						If GUICtrlRead($Radio3CM) = 1 Then
							$originalOfLand = "case1" ; NHÀ NƯỚC GIAO ĐẤT CÓ THU TIỀN SỬ DỤNG ĐẤT
						Else
							$originalOfLand = "case2" ; NHÀ NƯỚC CÔNG NHẬN QUYỀN SỬ DỤNG ĐẤT NHƯ NHÀ NƯỚC GIAO ĐẤT CÓ THU TIỀN SỬ DỤNG ĐẤT
						EndIf

						If $landArea = "" Then
							MsgBox(0, 0, "Bạn phải nhập diện tích đất")
						Else
							MsgBox(0, 0, "ahihi")
							; DO SOMETHING INTERACTIVE BETWEEN TWO SOFTWARE
						EndIf
					Case $OK3
						$dienTichSanKCC = GUICtrlRead($editDienTichSan)
						$dienTichSuDungKCC = GUICtrlRead($editDienTichSuDung)
						$DienTichXayDungKCC = GUICtrlRead($editDienTichXayDung)
						$loaiNhaO = GUICtrlRead($comboNha)
						If $dienTichSanKCC = "" And $dienTichSuDungKCC = "" And $DienTichXayDungKCC = "" Then
							MsgBox(0, 0, "Bạn phải nhập ít nhất một trong ba diện tích")
						ElseIf $landArea = "" Then
							MsgBox(0, 0, "Bạn phải nhập diện tích đất, đây không phải phần cho căn hộ chung cư")
						Else
							MsgBox(0, 0, "ahihi")
							; DO SOMETHING INTERACTIVE BETWEEN TWO SOFTWARE
						EndIf
					Case $OK4
						$propertyAddress = GUICtrlRead($editPropertyAddress)
						$propertyArea = GUICtrlRead($editDienTichSan)
						$DienTichSuDungCC = GUICtrlRead($editDienTichSuDungCC)
						If $propertyArea = "" Or $propertyAddress = "" Then
							MsgBox(0, 0, "ban phai nhap day du dien tich va so can ho")
						Else
							MsgBox(0, 0, "ahihi")
							; DO SOMETHING INTERACTIVE BETWEEN TWO SOFTWARE
						EndIf
					Case $Button4CM
					Case $Cap
					Case $Button3CM
					Case $lamPhieuChuyen
						GUISetState($lamPhieuChuyen, $GUI_DISABLE)
						LamPhieuChuyenChungCu()
				EndSwitch
			Case $Form1
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($Form1)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
				EndSwitch
		EndSwitch
	WEnd
EndFunc



Func ShowUpCapMoi()
	#Region ### START Koda GUI section ### Form=d:\it\autoit\vilis\cap moi.kxf
	$CapMoi = GUICreate("Cap Moi", 916, 660, 219, 78)
	$Tab1CM = GUICtrlCreateTab(8, 8, 953, 641)
	$TabSheet1CM = GUICtrlCreateTabItem("Chu So Huu")
	GUICtrlCreateLabel("Ho Va Ten", 32, 56, 56, 17)
	$editUserName1 = GUICtrlCreateInput("", 96, 56, 201, 21)
	GUICtrlCreateLabel("Nam Sinh", 312, 56, 50, 17)
	$editUserBirth1 = GUICtrlCreateInput("", 376, 56, 97, 21)
	GUICtrlCreateGroup("Group1", 496, 40, 361, 49)
	$Radio1CM = GUICtrlCreateRadio("Nam", 512, 56, 65, 22)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$Radio2CM = GUICtrlCreateRadio("Nu", 600, 56, 113, 22)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$cmnd1 = GUICtrlCreateCombo("CMND", 32, 104, 57, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	$editUserID1 = GUICtrlCreateInput("", 96, 104, 137, 21)
	GUICtrlCreateLabel("Noi Cap", 248, 104, 42, 17)
	$editIDPlace1 = GUICtrlCreateInput("", 312, 104, 121, 21)
	GUICtrlCreateLabel("Ngay Cap", 445, 104)
	$editIDdate1 = GUICtrlCreateInput("", 512, 104, 121, 21)
	GUICtrlCreateLabel("Dia Chi", 32, 152, 38, 17)
	$editUserAddress1 = GUICtrlCreateEdit("", 96, 152, 537, 89)
	GUICtrlCreateLabel("Ho Va Ten", 32, 280, 56, 17)
	$editUserID2 = GUICtrlCreateInput("", 96, 280, 201, 21)
	GUICtrlCreateLabel("Nam Sinh", 312, 280, 50, 17)
	$editUserBirth2 = GUICtrlCreateInput("", 376, 280, 97, 21)
	$cmnd2 = GUICtrlCreateCombo("CMND", 32, 322, 57, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	$editUserID2 = GUICtrlCreateInput("", 96, 322, 137, 21)
	GUICtrlCreateLabel("Noi Cap", 248, 322, 42, 17)
	$editIDPlace2 = GUICtrlCreateInput("", 312, 322, 121, 21)
	GUICtrlCreateLabel("Ngay Cap", 445, 322)
	$editIDDate2 = GUICtrlCreateInput("", 512, 322, 121, 21)
	GUICtrlCreateLabel("Dia Chi", 32, 368, 38, 17)
	$editUserAddress2 = GUICtrlCreateEdit("", 96, 368, 537, 89)
	$inHoOngBaCheck = GUICtrlCreateCheckbox("In ho ong ba", 645, 368)
	$OK1 = GUICtrlCreateButton("OK", 96, 480, 115, 49)

	; TAB2
	$TabSheet2CM = GUICtrlCreateTabItem("Thua Dat")
	$editLandArea = GUICtrlCreateInput("", 104, 88, 121, 21)
	GUICtrlCreateLabel("Dien Tich", 32, 88, 50, 17)
	GUICtrlCreateLabel("So To Cu", 32, 128, 49, 17)
	GUICtrlCreateLabel("So Thua Cu", 248, 136, 61, 17)
	$editOldPaperNumber = GUICtrlCreateInput("", 104, 128, 121, 21)
	$editOldLandID = GUICtrlCreateInput("", 336, 128, 121, 21)
	GUICtrlCreateGroup("Nguon goc su dung dat", 24, 200, 505, 113)
	$Radio3CM = GUICtrlCreateRadio("Nha nuoc giao dat co thu tien su dung dat", 64, 232, 345, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$Radio4CM = GUICtrlCreateRadio("Cong nhan quyen su dung  dat nhu nha nuoc  giao dat co thu tien su dung dat", 64, 264, 425, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$OK2 = GUICtrlCreateButton("OK", 24, 336, 107, 33)
	GUICtrlCreateLabel("Nha Gan Lien Voi Dat O", 24, 408, 119, 17)
	GUICtrlCreateLabel("Dat O", 32, 56, 32, 17)
	GUICtrlCreateLabel("Dien Tich San", 24, 456, 72, 17)
	GUICtrlCreateLabel("Dien Tich Su Dung", 24, 496, 95, 17)
	$editDienTichSan = GUICtrlCreateInput("", 168, 456, 121, 21)
	$editDienTichSuDung = GUICtrlCreateInput("", 168, 496, 121, 21)
	GUICtrlCreateLabel("Dien Tich Xay Dung", 24, 536, 100, 17)
	$editDienTichXayDung = GUICtrlCreateInput("", 168, 536, 121, 21)
	$comboNha = GUICtrlCreateCombo("Nha o rieng le", 168, 400, 145, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	$OK3 = GUICtrlCreateButton("OK", 24, 584, 107, 33)

	; TAB 3
	GUICtrlCreateTabItem("Can Ho Chung Cu")
	$ChonChungCu = GUICtrlCreateButton("Chon Chung Cu", 48, 64, 139, 41)
	GUICtrlCreateLabel("So Can Ho", 56, 152, 56, 17)
	$editPropertyAddress = GUICtrlCreateInput("", 176, 152, 121, 21)
	$editDienTichSuDungCC = GUICtrlCreateInput("", 176, 200, 121, 21)
	GUICtrlCreateLabel("Dien Tich Su Dung", 56, 200, 95, 17)
	GUICtrlCreateLabel("Dien Tich San", 56, 248, 72, 17)
	$editDienTichSanCC = GUICtrlCreateInput("", 176, 248, 121, 21)
	$OK4 = GUICtrlCreateButton("OK", 48, 312, 99, 49)

	; TAB 4
	$TabSheet4CM = GUICtrlCreateTabItem("Cap GCN")
	$Cap = GUICtrlCreateButton("Cap Nhat GCN", 56, 152, 139, 57)
	GUICtrlCreateLabel("So Giay Chung Nhan", 56, 96, 104, 17)
	$editGCNID = GUICtrlCreateInput("", 184, 96, 121, 21)
	$Button3CM = GUICtrlCreateButton("Ve So Do", 224, 152, 139, 57)
	$Button4CM = GUICtrlCreateButton("So tiep theo", 336, 96, 75, 25)
	$lamPhieuChuyen = GUICtrlCreateButton("PC, BC, TT", 400, 152, 139, 57)
	GUICtrlCreateTabItem("")
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ####include <ButtonConstants.au3>
EndFunc


Func LamPhieuChuyenChungCu()
    #Region ### START Koda GUI section ### Form=d:\it\autoit\vilis\tt, pc.kxf
	$Form1 = GUICreate("TT, PC, BC, chungcu", 910, 437, 222, 193)
    $Label1 = GUICtrlCreateLabel("Ten Du An", 48, 32, 56, 17)
	$Input1 = GUICtrlCreateInput("", 144, 32, 337, 21)
    $Label2 = GUICtrlCreateLabel("Nha Dau Tu", 496, 32, 63, 17)
    $Input2 = GUICtrlCreateInput("", 576, 32, 289, 21)
    $Label3 = GUICtrlCreateLabel("Cong ty Ban Nha", 48, 104, 85, 17)
    $Input3 = GUICtrlCreateInput("", 144, 104, 337, 21)
    $Label4 = GUICtrlCreateLabel("So Tang", 496, 104, 45, 17)
    $Input4 = GUICtrlCreateInput("", 576, 104, 177, 21)
    $Label5 = GUICtrlCreateLabel("Nam hoan cong", 48, 176, 78, 17)
    $Input5 = GUICtrlCreateInput("", 200, 176, 121, 21)
    $Label6 = GUICtrlCreateLabel("Gia trong hop dong", 496, 176, 95, 17)
    $Input6 = GUICtrlCreateInput("", 664, 176, 121, 21)
    $Label7 = GUICtrlCreateLabel("Bien ban ban giao ngay", 48, 248, 116, 17)
    $Input7 = GUICtrlCreateInput("", 200, 240, 121, 21)
    $Label8 = GUICtrlCreateLabel("Gia trong bien ban thanh ly", 496, 240, 131, 17)
    $Input8 = GUICtrlCreateInput("", 664, 240, 121, 21)
    $Label9 = GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 320, 130, 17)
    $Input9 = GUICtrlCreateInput("", 200, 312, 121, 21)
    $Button1 = GUICtrlCreateButton("OK", 424, 328, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 328, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
    #EndRegion ### END Koda GUI section ###
EndFunc