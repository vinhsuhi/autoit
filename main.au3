#cs
	Author: Tong Van Vinh
	KSTN - CNTT - k60, Hanoi University of Science and Technology
	email: Vinhbachkhoait@gmail.com
	sdt: 01654095052
#ce
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <TabConstants.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <Word.au3>
;########################
Global $theGCNname = 'Phan mem In giay chung nhan theo luat dat dai 2013; DVHC: Thanh pho Ha Noi - huyen Chuong My; May chu: 10.1.2.3; CSDL: lis; Nguoi dung: Tong Van Truc'
;~ Run("C:\Program Files (x86)\ViLIS\GCN2014\Phoenix.exe")
Global $configurexa = 1
Global $configurehuyen = 1
Global $ChuSoHuu1[10] ;Name, Birth of Date, sex, ID, day of ID, ID location, Address, Family, PrintFamily
Global $ChuSoHuu2[10]
;#########################

; #################### Windows 1
Global $recordID, $recordDate ; record
; $local = Toa nha 21 tang 671 Hoang Hoa Tham
Global $local, $phuongName, $quanName
; #################### Windows 2 tab 1
Global $UserName1, $UserID1[2], $IDdate1, $IDplace1, $UserBirth1, $UserSex1, $UserAddress1, $laNguoiDungTenPhieuChuyen ; User1
Global $UserName2, $UserID2[2], $IDdate2, $IDplace2, $UserBirth2, $UserSex2 = "b", $UserAddress2 ; User2

; #################### Windows 2 tab 3
Global $propertyAddress, $propertyArea, $DienTichSuDungCC


; #################### Wondows 3
Global $ProjectName, $InvestorName, $startYear, $nFloors, $vitritheoBangGiaDat ; property Chung cư
Global $projectArea
Global $contracName, $contractID, $contractDate


Global $sellerName, $PriceInContract, $deliveryPaperDate, $priceInLiquidation, $totalPrice
Global $landArea, $oldPaperNumber, $oldLandID ; THUA DAT

Global $duAn, $soHuuNhaNuoc, $taiDinhCu
;Region
Global $ChonChungCu = 9999, $Cap  =9999, $OK1 = 9999, $OK2 = 9999, $OK3 = 9999, $OK4 = 9999, $Button3CM = 9999, $Button4CM = 9999
Global $CapMoi, $DiaChiDat
Global $editLandArea, $editOldLandID, $editOldPaperNumber ; THỬA ĐẤT
Global $editUserName1, $editUserBirth1, $Radio1CM, $cmnd1, $editUserID1, $editIDPlace1, $editIDdate1, $editUserAddress1
Global $editUserName2, $editUserBirth2, $cmnd2, $editUserID2, $editIDPlace2, $editIDDate2, $editUserAddress2
Global $Radio3CM, $originalOfLand
Global $editPropertyAddress, $editDienTichSuDungCC, $editDienTichSanCC
Global $laNguoiDungTenPhieuChuyenCB
Global $batDauLamPCButton = 9999, $editNFloor, $editSellerName, $editStartYear
Global $editProjectName, $editInves, $editPriceInContract, $editDeliveDate
Global $editPriceInThanhLy, $editTotalPrice
Global $editRecordDate, $editRecordId
Global $muaNhaDuAnRadio, $soHuuNhaNuocRadio, $taiDinhCuRadio
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
	GUICtrlCreateLabel("Quan Huyen", 24, 144, 64, 17)
	$editQuanHuyen = GUICtrlCreateInput("", 24, 176, 321, 21)
	$OK0 = GUICtrlCreateButton("OK", 24, 328, 91, 33)
	GUICtrlCreateGroup("Hinh Thuc", 24, 216, 305, 97)
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

; If the show up windows is DiaChiDat, then
; 	Nhập các thông tin liên quan tới địa chỉ đất vào dao diện
;		Khi nút OK trên dao diện được nhấp, thì sẽ lấy toàn bộ thông tin từ dao diện vào trong các biến nhớ, và nếu radio box được chọn là cấp mới thì hiện dao diện cấp mới lên
; If the windows CapMoi show up, then
;	nhập người sử dụng, khi nhấn OK thì cập nhật tất cả các giá trị, rồi bắt đầu tương tác với phần mềm để điền các giá trị
;	rồi sau chuyển sang phần chung cư, nếu nhấn ok thì các giá trị sẽ được nhập vào phần mềm,
;	rồi sau đó chuyển sang phần cấp giấy chứng nhận.
	while 1
		$aMsg = GUIGetMsg(1)
		Switch $aMsg[1]
			Case $DiaChiDat
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						ExitLoop
					; KHI NHẤN NÚT OK THÌ BẮT ĐẦU GÁN CÁC GIÁ TRỊ CHO CÁC BIẾN, ĐỒNG THỜI BẬT
					; PHẦN MỀM LÊN VÀ THIẾT LẬP ĐƠN VỊ CẤP GIẤY CHỨNG NHẬN
					Case $OK0
						; GÁN CÁC GIÁ TRỊ CHO CÁC BIẾN
						; BAO GỒM: SỐ HỒ SƠ, NGÀY NHẬN HỒ SƠ, ĐỊA CHỈ ĐỊA PHƯƠNG
						; QUẬN HUYỆN, LÀ CẤP MỚI HAY CẤP ĐỔI???
						; NẾU LÀ CẤP MỚI THÌ MỞ GIAO DIỆN CẤP MỚI LÊN

						$recordID = GUICtrlRead($editRecordId)
						$recordDate = GUICtrlRead($editRecordDate)
						$local = GUICtrlRead($editLocal)
						Local $quanHuyen = StringSplit(GUICtrlRead($editQuanHuyen), ",")
						$phuongName = StringStripWS($quanHuyen[1], $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
						$quanName = StringStripWS($quanHuyen[2], $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
						Local $isCapMoi = GUICtrlRead($CapMoiF)
						GUISetState($OK0, $GUI_DISABLE)
						If $recordID = "" Or $recordDate = "" Or $local = "" Then
							MsgBox(0, "Thiếu thông tin", "Vui lòng nhập hết thông tin")
						ElseIf $isCapMoi = 0 Then
							MsgBox(0, 0, "cấp đổi")
						ElseIf $isCapMoi = 1 Then
							; NẾU CHỌN CẤP MỚI THÌ MỞ PHẦN MỀM CẤP MỚI GIẤY CHỨNG NHẬN LÊN, SAU ĐÓ THIẾT LẬP
							; VÀ CUỐI CÙNG LÀ MỞ GIAO DIỆN CẤP MỚI CỦA MÌNH LÊN
;~ 							Run("C:\Program Files (x86)\ViLIS\GCN2014\Phoenix.exe")
;~ 							Configure($quanName, $phuongName)
							ShowUpCapMoi()
						EndIf
				EndSwitch
			Case $CapMoi
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($CapMoi)
						GUICtrlSetState($OK0, $GUI_ENABLE)
					; Neu nhan nut OK o phan chu so huu
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
							$UserSex2 = "o"
						EndIf

						$UserName2 = GUICtrlRead($editUserName2)
						$UserID2[1] = GUICtrlRead($editUserID2)
						$UserID2[0] = GUICtrlRead($cmnd2)
						$IDdate2 = GUICtrlRead($editIDdate2)
						$UserBirth2 = GUICtrlRead($editUserBirth2)
						$UserAddress2 = GUICtrlRead($editUserAddress2)
						$IDplace2 = GUICtrlRead($editIDPlace2)

						If $IDplace1 = "" Then
							$IDplace1 = "Cục CS ĐKQL cư trú và DLQG về Dân Cư"
						EndIf

						If $IDplace2 = "" Then
							$IDplace2 = "Cục CS ĐKQL cư trú và DLQG về Dân Cư"
						EndIf

						If $UserAddress2 = "" Then
							$UserAddress2 = $UserAddress1
						EndIf

						If $UserName1 = "" Or $UserID1[1] = "" Then
							MsgBox(0, 0, "Tên người sử dụng đất hoặc CMND không được để trống")
						Else
							; DO SOMETHING INTERACTIVE BETWEEN TWO SOFTWARE
						EndIf
					; Neu nhan Nut Ok thu nhat o phan Thua Dat
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
					; Neu nhan nut OK thu 2 o phan Thua Dat
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
					; Neu nhan nut OK o phan can ho chung cu
					Case $OK4
						$propertyAddress = GUICtrlRead($editPropertyAddress)
						$propertyArea = GUICtrlRead($editDienTichSanCC)
						$DienTichSuDungCC = GUICtrlRead($editDienTichSuDungCC)
						If $propertyArea = "" Or $propertyAddress = "" Then
							MsgBox(0, 0, "ban phai nhap day du dien tich va so can ho")
						Else
							MsgBox(0, 0, "ahihi")
							; DO SOMETHING INTERACTIVE BETWEEN TWO SOFTWARE
						EndIf
					; don't care
					Case $Button4CM
					Case $Cap
					Case $Button3CM
					; Neu Bam Nut Lam Phieu Chuyen
					Case $lamPhieuChuyen
						$duAn = GUICtrlRead($muaNhaDuAnRadio)
						$soHuuNhaNuoc = GUICtrlRead($soHuuNhaNuocRadio)
						$taiDinhCu = GUICtrlRead($taiDinhCuRadio)
						GUISetState($lamPhieuChuyen, $GUI_DISABLE)
						; Neu la can ho chung cu, thi lam phieu chuyen cho chcc
						If $propertyAddress <> "" And $duAn <> 0 Then
							LamPhieuChuyenChungCu()
						ElseIf $propertyAddress <> "" And $soHuuNhaNuoc <> 0 Then
;~ 							LamPhieuChuyenChungCuSHNN()
						ElseIf $propertyAddress <> "" And $taiDinhCu <> 0 Then
;~ 							LamPhieuChuyenChungCuTDC()
						ElseIf $propertyAddress = "" And $duAn <> Then
;~ 							LamPhieuChuyenNhaLe()
						ElseIf $propertyAddress = "" And $soHuuNhaNuoc <> Then
;~ 							LamPhieuChuyenNhaLeSHNN()
						EndIf
				EndSwitch
			; GET INFO ABOUT CANHOCHUNGCU
			Case $Form1
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($Form1)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
;~ 						$vitritheoBangGiaDat = GUICtrlRead()
						$ProjectName = GUICtrlRead($editProjectName)
						$InvestorName = GUICtrlRead($editInves)
						$nFloors = GUICtrlRead($editNFloor)
						$sellerName = GUICtrlRead($editSellerName)
						If $sellerName = "" Then
							$sellerName = $InvestorName
						EndIf
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$deliveryPaperDate = GUICtrlRead($editDeliveDate)
						$priceInLiquidation = GUICtrlRead($editPriceInThanhLy)
						$totalPrice = GUICtrlRead($editTotalPrice)
						; GOI Ham Cap Moi
						CapMoiNhaDuAn()
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
	GUICtrlCreateGroup("Gioi tinh", 496, 40, 361, 49)
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
	$laNguoiDungTenPhieuChuyenCB = GUICtrlCreateCheckbox("Dung ten PC", 650, 104)
	GUICtrlSetState(-1, $GUI_CHECKED)
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
	GUICtrlCreateGroup("Loai nghia vu", 550, 80, 300, 137)
	$muaNhaDuAnRadio = GUICtrlCreateRadio("Mua nha du an", 575, 112, 113, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$soHuuNhaNuocRadio = GUICtrlCreateRadio("So huu nha nuoc", 575, 144, 113, 17)
	$taiDinhCuRadio = GUICtrlCreateRadio("Tai dinh cu", 575, 176, 113, 17)
	GUICtrlCreateTabItem("")
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ####include <ButtonConstants.au3>
EndFunc

Func LamPhieuChuyenChungCu()
    #Region ### START Koda GUI section ### Form=d:\it\autoit\vilis\tt, pc.kxf
	$Form1 = GUICreate("TT, PC, BC, chungcu", 910, 437, 222, 193)
	; Con thieu: viTriTrongBangGiaDat, projectArea, hosoNgay
	; Bo sung:
    GUICtrlCreateLabel("Ten Du An", 48, 32, 56, 17)
	$editProjectName = GUICtrlCreateInput("", 144, 32, 337, 21)
    GUICtrlCreateLabel("Nha Dau Tu", 496, 32, 63, 17)
    $editInves = GUICtrlCreateInput("", 576, 32, 289, 21)
    GUICtrlCreateLabel("Cong ty Ban Nha", 48, 104, 85, 17)
    $editSellerName = GUICtrlCreateInput("", 144, 104, 337, 21)
    GUICtrlCreateLabel("So Tang", 496, 104, 45, 17)
    $editNFloor = GUICtrlCreateInput("", 576, 104, 177, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 176, 78, 17)
    $editStartYear = GUICtrlCreateInput("", 200, 176, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 176, 95, 17)
    $editPriceInContract = GUICtrlCreateInput("", 664, 176, 121, 21)
    GUICtrlCreateLabel("Bien ban ban giao ngay", 48, 248, 116, 17)
    $editDeliveDate = GUICtrlCreateInput("", 200, 240, 121, 21)
    GUICtrlCreateLabel("Gia trong bien ban thanh ly", 496, 240, 131, 17)
    $editPriceInThanhLy = GUICtrlCreateInput("", 664, 240, 121, 21)
    GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 320, 130, 17)
    $editTotalPrice = GUICtrlCreateInput("", 200, 312, 121, 21)

    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 328, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 328, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
    #EndRegion ### END Koda GUI section ###
EndFunc

; #########
;CHO NAY LA DE LAM THEM 4 CAI GIAO DIEN VE CAP NHA
;##########

; CHỌN GIÁ TRỊ TRONG COMBOBOX, CÁC THAM SỐ ĐẦU VÀO BAO GỒM TÊN CỬA SỔ CHỌN, GIÁ TRỊ CONTROL, GIÁ TRỊ CẦN CHỌN: $thetext, GIÁ TRỊ CUỐI THÌ KHỎI QUAN TÂM
Func ChooseMyBox($name, $control, $thetext, $xaorhuyen)
	$config = WinWaitActive($name)
	ControlClick($config, "", $control)
	Send("{HOME}")
	Sleep(500)
	$count = 0
	$MAX = 70
	For $a = 1 To $MAX
		$count += 1
		$ctext = ControlGetText($name, "", $control)
		If StringInStr($ctext, $thetext) Then
			Send("{ENTER}")
			ExitLoop
		EndIf
		Send("{DOWN}")
	Next
	If $count == $MAX Then
		MsgBox(0, "Loi", "Nhap sai, vui long nhap lai " & $xaorhuyen)
	EndIf
	Return $count
EndFunc

; chọn nơi đăng ký lúc đầu (đăng nhập)
Func Configure($huyen, $xa)
	Local $name0 = 'Đăng nhập hệ thống'
	WinActivate($name0)
	Local $windows0 = WinWaitActive($name0)
	ControlClick($windows0, "", "WindowsForms10.BUTTON.app.0.33c0d9d1")
	WinWaitActive("Cấu hình hệ thống")
	Send("{RIGHT}")
	Send("{RIGHT}")
	Local $name = "Cấu hình hệ thống"
	Local $control1 = 'WindowsForms10.COMBOBOX.app.0.33c0d9d3'
	Local $control2 = 'WindowsForms10.COMBOBOX.app.0.33c0d9d1'
	$count1 = ChooseMyBox($name, $control1, $huyen, "huyen")
	$count2 = ChooseMyBox($name, $control2, $xa, "xa")
	If $count1 < 70 And $count2 < 70 Then
		ControlClick($name, "", "WindowsForms10.BUTTON.app.0.33c0d9d2")
		Local $thongBao = WinWaitActive("Thông báo")
		ControlClick($thongBao, "", 'WindowsForms10.BUTTON.app.0.33c0d9d1')
		ControlClick($name, "", 'WindowsForms10.BUTTON.app.0.33c0d9d1')
		ControlClick($name0, "", 'WindowsForms10.BUTTON.app.0.33c0d9d3')
	EndIf
EndFunc

Func SetAddress($tinh, $huyen, $xa, $diaPhuong)
	Local $winGCN = WinWaitActive($theGCNname)
	ControlClick($winGCN, "", 'WindowsForms10.BUTTON.app.0.33c0d9d13')
	Local $name = 'Nhập địa chỉ chi tiết'
	Local $Control1 = 'WindowsForms10.COMBOBOX.app.0.33c0d9d3'
	Local $Control2 = 'WindowsForms10.COMBOBOX.app.0.33c0d9d2'
	Local $Control3 = 'WindowsForms10.COMBOBOX.app.0.33c0d9d1'
	Local $Control4 = 'WindowsForms10.EDIT.app.0.33c0d9d2'
	ChooseMyBox($name, $Control1, $tinh, "tinh")
	ChooseMyBox($name, $Control2, $huyen, "huyen")
	ChooseMyBox($name, $Control3, $xa, "xa")
	;Sleep(50)
	ControlSetText($name, "", $Control4, $diaPhuong)
	Send("^s")
	Sleep(1000)
	If WinExists("Cập nhật địa chỉ") Then
		Send("{ENTER}")
	EndIf
	Sleep(1000)
	If WinExists("Thông báo") Then
		Send("{ENTER}")
;~ 		Sleep(1000)
	EndIf
EndFunc

Func setUserInfo()
	Local $name = $theGCNname
	Local $winGCN = WinWaitActive($name)
	MouseClick('primary', 232, 43, 1, 0)
	MouseClick('primary', 281, 62, 1, 0)
	Sleep(1000)
	ControlClick($winGCN, "", 'WindowsForms10.BUTTON.app.0.33c0d9d5')
	Local $winGCN2 = WinWaitActive($name)
	ControlSetText($winGCN2, "", "WindowsForms10.EDIT.app.0.33c0d9d11", $UserName1)
	ControlSetText($winGCN2, "", "WindowsForms10.EDIT.app.0.33c0d9d14", $UserBirth1)
	ControlSetText($winGCN2, "", "WindowsForms10.EDIT.app.0.33c0d9d13", $UserID1[1])
	ControlSend($winGCN2, "", "WindowsForms10.EDIT.app.0.33c0d9d9", $IDdate1)
	; NƠI CẤP CMT, NẾU ĐỂ TRỐNG THÌ TỰ ĐỘNG LẤY GIÁ TRỊ GIẤY CMT MỚI
	If $IDplace1 <> "" Then
		ControlSetText($winGCN2, "", "WindowsForms10.EDIT.app.0.33c0d9d12", $IDplace1)
	Else
		ControlSetText($winGCN2, "", "WindowsForms10.EDIT.app.0.33c0d9d12", "Cục CS ĐKQL cư trú và DLQG về Dân cư")
	EndIf
	; VỀ ĐỊA CHỈ CỦA CÁC USER
	Local $lenUserAddress1 = UBound($UserAddress1)
	Local $lenUserAddress2 = UBound($UserAddress2)
	Local $localName1 = ""
	For $a = 0 To $lenUserAddress1
		If $a < $lenUserAddress1-3 Then
			$localName1 = $localName1 & $UserAddress1[$a]
		EndIf
	Next
	; NHẬP ĐỊA CHỈ CỦA User1
	SetAddress($UserAddress1[$lenUserAddress1-1], $UserAddress1[$lenUserAddress1-2], $UserAddress1[$lenUserAddress1-3], $localName1)
	; NẾU TÊN CỦA NGƯỜI THỨ HAI KHÔNG BỎ TRỐNG
	If $UserName2 <> "" Then
		If $UserAddress2 = "" Then
			$UserAddress2 = $UserAddress1
		EndIf
		If $IDplace2 = "" Then
			$IDplace2 = "Cục CS ĐKQL cư trú và DLQG về Dân cư"
		EndIf
		; WRITE INFORMATION TO GCN SOFT WARE
	EndIf
	Send("{F2}")
	Sleep(1000)
	If WinExists("Thông báo") Then
		Send("{ENTER}")
	EndIf
	Sleep(1000)
	If WinExists("Cập nhật địa chỉ") Then
		Send("{ENTER}")
	EndIf
	Send("{ENTER}")
	Send("{ENTER}")
	Send("{ENTER}")
	Sleep(1000)
	ControlClick($name, "", "WindowsForms10.BUTTON.app.0.33c0d9d21")
EndFunc

; PHAN TAO PHIEU CHUYEN

Func CapMoiNhaDuAn()
	; Neu User1 khong la nguoi dung ten phieu chuyen, thi
	; Phai chuyen User1 thanh User2 va nguoc lai
	$laNguoiDungTenPhieuChuyen = GUICtrlRead($laNguoiDungTenPhieuChuyenCB)
	If $laNguoiDungTenPhieuChuyen = 0 Then
		Local $tempName = $UserName1
		$UserName1 = $UserName2
		$UserName2 = $tempName
		Local $tempID = $UserID1
		$UserID1 = $UserID2
		$UserID2 = $tempID
		Local $tempIDDate = $IDdate1
		$IDdate1 = $IDdate2
		$IDdate2 = $tempIDDate
		Local $tempIDPlace = $IDplace1
		$IDplace1 = $IDplace2
		$IDplace2 = $tempIDPlace
		Local $tempBirth = $UserBirth1
		$UserBirth1 = $UserBirth2
		$UserBirth2 = $tempBirth
		Local $tempSex = $UserSex1
		$UserSex1 = $UserSex2
		$UserSex2 = $tempSex
		Local $tempAddress = $UserAddress1
		$UserAddress1 = $UserAddress2
		$UserAddress2 = $tempAddress
	EndIf
	; Neu la nha chung cu thi...
	If $propertyAddress <> "" Then
		; COPY SAMPLE FILE
		Local $destinationFileName = $recordID & " " & StringUpper($UserName1) & ".doc"
		Local $destinationPath = "C:\Users\vinh\Desktop\" & $destinationFileName
		Local $sourcePath = "source"
		FileCopy($sourcePath, $destinationPath)
		; MO FILE DA COPY LEN VA SUA
		Local $oWord = _Word_Create()
		Local $sDocument = $destinationPath
		Local $oDoc = _Word_DocOpen($oWord, $sDocument)
		WinWaitActive("sample [Compatibility Mode] - Word")
		sleep(1000)
		_Word_DocFindReplace($oDoc, "recordID", $recordID)
		_Word_DocFindReplace($oDoc, "recordDate", $recordDate)

		If $UserSex1 = "o" Then
			; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
			If $UserName2 = "" Then
				_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
				_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
			; Truong hop co hai nguoi, va nguoi dung ten la nam
			ElseIf $UserName2 <> "" Then
				_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
				_Word_DocFindReplace($oDoc, "UserName2", $UserName2)
			EndIf
		; Truong hop User1 la nu
		Else
			If $UserName2 = "" Then
				_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
				_Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
				_Word_DocFindReplace($oDoc, "UserName1", "ông " & $UserName1)
				_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
			Else
				_Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
				_Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
				_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
			EndIf
		EndIf

		_Word_DocFindReplace($oDoc, "UserAddress1", $UserAddress1)
        Local $UserIDD1 = $UserID1[0] & "số " & $UserID1[1]
        _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
        _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
        _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)

        ;now about property
        _Word_DocFindReplace($oDoc, "propertyAddress", $propertyAddress)
        ;#O#OO3O3O3O3O3O3O3O3O3 HASN'T INFORMATION ABOUT PROJECT
        _Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)
        _Word_DocFindReplace($oDoc, "projectArea", $projectArea)
        _Word_DocFindReplace($oDoc, "nFloors", $nFloors)

        _Word_DocFindReplace($oDoc, "startYear", $startYear)
        _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
        _Word_DocFindReplace($oDoc, "ContractName", $contracName)
        _Word_DocFindReplace($oDoc, "contractID", $contractID)
        _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
        _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
        _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
        _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
        _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
        _Word_DocFindReplace($oDoc, "quanName", $quanName)
        _Word_DocFindReplace($oDoc, "SellerName", $sellerName)

;~ 		_Word_DocSaveAs($oDoc, "D:\Unity\lol.doc")
;~ 		_Word_Quit($oWord)
	EndIf
EndFunc


