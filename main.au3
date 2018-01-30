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
Global $theGCNname
Global $configurexa = 1
Global $configurehuyen = 1

;#########################

; #################### Windows 1
Global $recordID, $recordDate ; record
; $local = Toa nha 21 tang 671 Hoang Hoa Tham
Global $local, $phuongName, $quanName
; #################### Windows 2 tab 1
Global $UserName1, $UserID1[2], $IDdate1, $IDplace1, $UserBirth1, $UserSex1, $UserAddress1, $laNguoiDungTenPhieuChuyen ; User1
Global $UserName2, $UserID2[2], $IDdate2, $IDplace2, $UserBirth2, $UserSex2 = "b", $UserAddress2 ; User2

; #################### Windows 2 tab 3
Global $propertyAddress, $propertyArea, $DienTichSuDungCC, $CCTTDate, $CCTTID


; #################### Wondows 3
Global $ProjectName, $InvestorName, $startYear, $nFloors, $vitritheoBangGiaDat ; property Chung cư
Global $projectArea
Global $contractName, $contractID, $contractDate, $contractPlace, $pcCCcd, $pcNLcd
Global $flagGCNID = 0, $newGCNID, $flagCap = 0, $kindOfOldGCN

Global $sellerName, $PriceInContract, $deliveryPaperDate, $priceInLiquidation, $totalPrice
Global $landArea, $oldPaperNumber, $oldLandID ; THUA DAT
Global $sellerName1, $sellerID1[2], $sellerAddress1, $sellerSex1, $sellerBirth1, $sellerBirth2
Global $sellerName2, $sellerID2[2], $sellerAddress2, $sellerSex2
Global $duAn, $soHuuNhaNuoc, $taiDinhCu, $congVanID, $congVanDate, $pricePayed
Global $propertyPercent, $duongDoanDuong, $capHang, $isInHoOngBa = 0
Global $oldGCNID, $oldGCNDate, $oldGCNPlace, $isCapDoiC, $isCapDoiM, $isCapMoi, $isChuyenNhuong = 0, $isTangCho = 0
;Region
Global $ChonChungCu = 9999, $Cap  =9999, $OK1 = 9999, $OK2 = 9999, $OK3 = 9999, $OK4 = 9999, $Button3CM = 9999, $Button4CM = 9999, $OKSeller = 9999
Global $CapMoi, $DiaChiDat, $capDoi
Global $editLandArea, $editOldLandID, $editOldPaperNumber ; THỬA ĐẤT
Global $editUserName1, $editUserBirth1, $Radio1CM, $cmnd1, $editUserID1, $editIDPlace1, $editIDdate1, $editUserAddress1
Global $editUserName2, $editUserBirth2, $cmnd2, $editUserID2, $editIDPlace2, $editIDDate2, $editUserAddress2
Global $Radio3CM, $originalOfLand
Global $editPropertyAddress, $editDienTichSuDungCC, $editDienTichSanCC
Global $laNguoiDungTenPhieuChuyenCB
Global $batDauLamPCButton = 9999, $editNFloor, $editSellerName, $editStartYear, $editSellerAddress1, $editSellerName1, $editSellerName2, $editSellerAddress2
Global $editProjectName, $editInves, $editPriceInContract, $editDeliveDate, $editSellerID1, $editSellerID2, $editSellerBirth1, $editSellerBirth2, $radioSellerSex1, $radioSellerSex11
Global $editPriceInThanhLy, $editTotalPrice
Global $editRecordDate, $editRecordId
Global $muaNhaDuAnRadio, $soHuuNhaNuocRadio, $taiDinhCuRadio, $editDuong, $editCongVanID, $editCongVanDate
Global $editPricePayed, $editPropertyPercent, $editCapHang, $editViTriThuaDat, $editCCTTDate, $editCCTTID, $landAddress
Global $editContractDate, $editProjectArea, $editOldGCNID, $editOldGCNDate, $editOldGCNPlace, $radioChuyenNhuong, $radioTangCho
Global $radioChuyenNhuongCC, $radioTangChoCC, $editGCNID, $editContractName, $editContractID, $editContractDate, $editContractPlace, $editKindOfOldGCN, $inHoOngBaCheck, $cmnd1Seller, $cmnd2Seller
;############
; NHÀ KHÔNG CHUNG CƯ
;############
Global $editDienTichSan, $editDienTichSuDung, $editDienTichXayDung, $comboNha
Global $dienTichSanKCC, $dienTichSuDungKCC, $DienTichXayDungKCC, $loaiNhaO
Global $lamPhieuChuyen = 9999, $pcCC, $ShnnCC, $pcNL, $ShnnNL, $TDC
ShowUpFirstGUI()

Func ShowUpFirstGUI()
	#Region ### START Koda GUI section ### Form=d:\it\autoit\vilis\dia chi dat.kxf
	$DiaChiDat = GUICreate("DiaChiDat", 437, 376, 280, 182)
	$DiaChiGoc = GUICtrlCreateLabel("Dia Chi Goc", 24, 80, 61, 17)
	$editLocal = GUICtrlCreateInput("", 24, 112, 321, 21)
	GUICtrlCreateLabel("Quan Huyen", 24, 144, 64, 17)
	$editQuanHuyen = GUICtrlCreateInput("", 24, 176, 321, 21)
	$OK0 = GUICtrlCreateButton("OK", 24, 328, 91, 33)
	GUICtrlCreateGroup("Hinh Thuc", 24, 216, 305, 97)
	$CapMoiF = GUICtrlCreateRadio("Cap Moi", 40, 230, 81, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$CapDoiF = GUICtrlCreateRadio("Cap doi tu GCN do so cap", 40, 255)
	$CapDoiK = GUICtrlCreateRadio("Cap doi tu GCN khong do so cap", 40, 280)
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
						$isCapMoi = GUICtrlRead($CapMoiF)
						$isCapDoiM = GUICtrlRead($CapDoiF)
						$isCapDoiC = GUICtrlRead($CapDoiK)
						GUISetState($OK0, $GUI_DISABLE)
						If $recordID = "" Or $recordDate = "" Then
							MsgBox(0, "Thiếu thông tin", "Vui lòng nhập hết thông tin")
						ElseIf $isCapDoiM = 1 Then
							MsgBox(0, 0, "cấp đổi mới")
						ElseIf $isCapDoiC = 1 Then
							Run("C:\Program Files (x86)\ViLIS\GCN2014\Phoenix.exe")
							Configure($quanName, $phuongName)
							ShowUpCapDoiCuGUI()
						ElseIf $isCapMoi = 1 Then
							; NẾU CHỌN CẤP MỚI THÌ MỞ PHẦN MỀM CẤP MỚI GIẤY CHỨNG NHẬN LÊN, SAU ĐÓ THIẾT LẬP
							; VÀ CUỐI CÙNG LÀ MỞ GIAO DIỆN CẤP MỚI CỦA MÌNH LÊN
							Run("C:\Program Files (x86)\ViLIS\GCN2014\Phoenix.exe")
							Configure($quanName, $phuongName)
							ShowUpCapMoiGUI()
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
						$isInHoOngBa = GUICtrlRead($inHoOngBaCheck)
						If $IDplace1 = "" Then
							$IDplace1 = "Cục CS ĐKQL cư trú và DLQG về Dân Cư"
						EndIf

						If $IDplace2 = "" Then
							$IDplace2 = "Cục CS ĐKQL cư trú và DLQG về Dân Cư"
						EndIf

						If $UserAddress2 = "" Then
							$UserAddress2 = $UserAddress1
						EndIf

						setUserInfo()
						WinActivate("Cap Moi")
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
							nhapThuaDat()
							WinActivate("Cap Moi")
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
							nhapNhaOThongThuong()
							WinActivate("Cap Moi")
						EndIf
					; Neu nhan nut OK o phan can ho chung cu
					Case $OK4
						$propertyAddress = GUICtrlRead($editPropertyAddress)
						$propertyArea = GUICtrlRead($editDienTichSanCC)
						$DienTichSuDungCC = GUICtrlRead($editDienTichSuDungCC)
						If $propertyArea = "" Or $propertyAddress = "" Then
							MsgBox(0, 0, "ban phai nhap day du dien tich va so can ho")
						Else
							setValueCC()
							WinActivate("Cap Moi")
						EndIf
					; don't care
					Case $Button4CM
						$flagGCNID = 1
						MsgBox(0, "Thong Bao", "Bạn đã lấy số tiếp theo")
					Case $Cap
						CapGiay()
						$flagCap = 1
						WinActivate("Cap Moi")
					Case $Button3CM
						If $flagCap = 0 Then
							MsgBox(0, 0, "Phải cập nhật đã")
						Else
							WinActivate($theGCNname)
							WinWaitActive($theGCNname)
							Send("{F3}")
							Sleep(2000)
							Send("{ENTER}")
							Sleep(500)
							Send("{F4}")
							Local $laySoDo = "[NAME:btnEditParcelDiagram]"
							WinWaitActive($theGCNname)
							ControlClick($theGCNname, "", $laySoDo)
						EndIf
					; Neu Bam Nut Lam Phieu Chuyen
					Case $lamPhieuChuyen
						$duAn = GUICtrlRead($muaNhaDuAnRadio)
						$soHuuNhaNuoc = GUICtrlRead($soHuuNhaNuocRadio)
						$taiDinhCu = GUICtrlRead($taiDinhCuRadio)
						GUISetState($lamPhieuChuyen, $GUI_DISABLE)
						; Neu la can ho chung cu, thi lam phieu chuyen cho chcc
						If $propertyAddress <> "" And $duAn <> 0 Then
							LamPhieuChuyenCCGUI()
						ElseIf $propertyAddress <> "" And $soHuuNhaNuoc <> 0 Then
							LamPhieuChuyenChungCuSHNNGUI()
						ElseIf $propertyAddress <> "" And $taiDinhCu <> 0 Then
							LamPhieuChuyenChungCuTDCGUI()
						ElseIf $propertyAddress = "" And $duAn <> 0 Then
							LamPhieuChuyenNhaLeGUI()
						ElseIf $propertyAddress = "" And $soHuuNhaNuoc <> 0 Then
							LamPhieuChuyenNhaLeSHNNGUI()
						EndIf
				EndSwitch
			Case $capDoi
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($capDoi)
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
						MsgBox(0, 0, $UserName2)
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

						setUserInfo()
						WinActivate("Cap Doi")

					Case $OKSeller
						$sellerName1 = GUICtrlRead($editSellerName1)
						$sellerBirth1 = GUICtrlRead($editSellerBirth1)
						$sellerBirth2 = GUICtrlRead($editSellerBirth2)
						$sellerName2 = GUICtrlRead($editSellerName2)
						$sellerID1[1] = GUICtrlRead($editSellerID1)
						$sellerID1[0] = GUICtrlRead($cmnd1Seller)
						$sellerID2[0] = GUICtrlRead($cmnd2Seller)
						$sellerID2[1] = GUICtrlRead($editSellerID2)
						$sellerSex1 = GUICtrlRead($radioSellerSex1)
						$sellerAddress1 = GUICtrlRead($editSellerAddress1)
						$sellerAddress2 = GUICtrlRead($editSellerAddress2)
						If $sellerAddress2 = "" Then
							$sellerAddress2 = $sellerAddress1
						EndIf
						If $sellerSex1 = 1 Then
							$sellerSex2 = 0
						Else
							$sellerSex2 = 1
						EndIf
						MsgBox(0, 0, "Đã lưu thông tin bên bán, tặng cho")
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
						If GUICtrlRead($radioChuyenNhuong) = 1 Then
							$isChuyenNhuong = 1 And $isTangCho = 0
						Else
							$isTangCho = 1 And $isChuyenNhuong = 0
						EndIf
						If $landArea = "" Then
							MsgBox(0, 0, "Bạn phải nhập diện tích đất")
						Else
							nhapThuaDat()
							WinActivate("Cap Doi")
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
							nhapNhaOThongThuong()
							WinActivate("Cap Doi")
						EndIf

					; Neu nhan nut OK o phan can ho chung cu
					Case $OK4
						$propertyAddress = GUICtrlRead($editPropertyAddress)
						$propertyArea = GUICtrlRead($editDienTichSanCC)
						$DienTichSuDungCC = GUICtrlRead($editDienTichSuDungCC)
						$oldGCNID = GUICtrlRead($editOldGCNID)
						$oldGCNDate = GUICtrlRead($editOldGCNDate)
						$oldGCNPlace = GUICtrlRead($editOldGCNPlace)
						If GUICtrlRead($radioChuyenNhuongCC) = 1 Then
							$isChuyenNhuong = 1 And $isTangCho = 0
						Else
							$isTangCho = 1 And $isChuyenNhuong = 0
						EndIf
						If $propertyArea = "" Or $propertyAddress = "" Then
							MsgBox(0, 0, "ban phai nhap day du dien tich va so can ho")
						Else
							setValueCC()
							WinActivate("Cap Doi")
						EndIf
					; don't care
					Case $Button4CM
						$flagGCNID = 1
						MsgBox(0, "Thong Bao", "Bạn đã lấy số tiếp theo")
					Case $Cap
						If $flagGCNID = 0 And GUICtrlRead($editGCNID) <> "" Then
							$newGCNID = GUICtrlRead($editGCNID)
						EndIf
						$oldGCNID = GUICtrlRead($editOldGCNID)
						$oldGCNDate = GUICtrlRead($editOldGCNDate)
						$oldGCNPlace = GUICtrlRead($editOldGCNPlace)
						$contractName = GUICtrlRead($editContractName)
						$contractDate = GUICtrlRead($editContractDate)
						$contractID = GUICtrlRead($editContractID)
						$contractPlace = GUICtrlRead($editContractPlace)
						$kindOfOldGCN = GUICtrlRead($editKindOfOldGCN)
						CapGiay()
						$flagCap = 1
						WinActivate("Cap Doi")
					Case $Button3CM
						If $flagCap = 0 Then
							MsgBox(0, 0, "Phải cập nhật đã")
						Else
							WinActivate($theGCNname)
							WinWaitActive($theGCNname)
							Send("{F3}")
							Sleep(2000)
							Send("{ENTER}")
							Sleep(500)
							Send("{F4}")
							Local $laySoDo = "[NAME:btnEditParcelDiagram]"
							WinWaitActive($theGCNname)
							ControlClick($theGCNname, "", $laySoDo)
						EndIf
					; Neu Bam Nut Lam Phieu Chuyen
					Case $lamPhieuChuyen
						If $propertyAddress = "" Then
							LamPhieuChuyenNLCDGUI()
						Else
							LamPhieuChuyenCCCDGUI()
						EndIf
				EndSwitch
			; GET INFO ABOUT CANHOCHUNGCU
			Case $pcCC
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($pcCC)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
						$duongDoanDuong = GUICtrlRead($editDuong)
						$ProjectName = GUICtrlRead($editProjectName)
						$projectArea = GUICtrlRead($editProjectArea)
						$InvestorName = GUICtrlRead($editInves)
						$nFloors = GUICtrlRead($editNFloor)
						$sellerName = GUICtrlRead($editSellerName)
						If $sellerName = "" Then
							$sellerName = $InvestorName
						ElseIf $InvestorName = "" Then
							$InvestorName = $sellerName
						EndIf
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$deliveryPaperDate = GUICtrlRead($editDeliveDate)
						$priceInLiquidation = GUICtrlRead($editPriceInThanhLy)
						$totalPrice = GUICtrlRead($editTotalPrice)
						$contractDate = GUICtrlRead($editContractDate)
						$contractID = GUICtrlRead($editContractID)
						$contractName = GUICtrlRead($editContractName)
						; GOI Ham Cap Moi
						CapMoiNhaDuAnMSW()
				EndSwitch
			Case $ShnnCC
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($ShnnCC)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
						$ProjectName = GUICtrlRead($editProjectName)
						$congVanDate = GUICtrlRead($editCongVanDate)
						$congVanID = GUICtrlRead($editCongVanID)
						$sellerName = GUICtrlRead($editSellerName)
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$deliveryPaperDate = GUICtrlRead($editDeliveDate)
						$pricePayed = GUICtrlRead($editPricePayed)
						$totalPrice = GUICtrlRead($editTotalPrice)
						$propertyPercent = GUICtrlRead($editPropertyPercent)
						$duongDoanDuong = GUICtrlRead($editDuong)
						$contractDate = GUICtrlRead($editContractDate)
						$projectArea = GUICtrlRead($editProjectArea)
						$nFloors = GUICtrlRead($editNFloor)
						$capHang = GUICtrlRead($editCapHang)
						CapMoiSHNNCCMSW()
				EndSwitch
			Case $pcNL
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($pcNL)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
						$ProjectName = GUICtrlRead($editProjectName)
						$InvestorName = GUICtrlRead($editInves)
						$sellerName = GUICtrlRead($editSellerName)
						If $sellerName = "" Then
							$sellerName = $InvestorName
						EndIf
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$deliveryPaperDate = GUICtrlRead($editDeliveDate)
						$totalPrice = GUICtrlRead($editTotalPrice)
						$duongDoanDuong = GUICtrlRead($editDuong)
						$contractDate = GUICtrlRead($editContractDate)
						$nFloors = GUICtrlRead($editNFloor)
						$priceInLiquidation = GUICtrlRead($editPriceInThanhLy)
						$contractDate = GUICtrlRead($editContractDate)
						$vitritheoBangGiaDat = GUICtrlRead($editViTriThuaDat)
						CapMoiNhaDuAnMSW()
				EndSwitch
			Case $ShnnNL
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($ShnnNL)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
						$ProjectName = GUICtrlRead($editProjectName)
						$congVanDate = GUICtrlRead($editCongVanDate)
						$congVanID = GUICtrlRead($editCongVanID)
						$sellerName = GUICtrlRead($editSellerName)
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$deliveryPaperDate = GUICtrlRead($editDeliveDate)
						$pricePayed = GUICtrlRead($editPricePayed)
						$totalPrice = GUICtrlRead($editTotalPrice)
						$propertyPercent = GUICtrlRead($editPropertyPercent)
						$duongDoanDuong = GUICtrlRead($editDuong)
						$contractDate = GUICtrlRead($editContractDate)
;~ 						$projectArea = GUICtrlRead($editProjectArea)
						$nFloors = GUICtrlRead($editNFloor)
						$capHang = GUICtrlRead($editCapHang)
						CapMoiSHNNNLMSW()
				EndSwitch
			Case $TDC
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($TDC)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
						$ProjectName = GUICtrlRead($editProjectName)
						$sellerName = GUICtrlRead($editSellerName)
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$deliveryPaperDate = GUICtrlRead($editDeliveDate)
						$totalPrice = GUICtrlRead($totalPrice)
						$duongDoanDuong = GUICtrlRead($editDuong)
						$contractDate = GUICtrlRead($editContractDate)
						$nFloors = GUICtrlRead($editNFloor)
						$capHang = GUICtrlRead($editCapHang)
						$vitritheoBangGiaDat = GUICtrlRead($editViTriThuaDat)
						CapMoiTDCMSW()
				EndSwitch
			Case $pcCCcd
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($pcCCcd)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
						$ProjectName = GUICtrlRead($editProjectName)
						$capHang = GUICtrlRead($editCapHang)
						$propertyPercent = GUICtrlRead($editPropertyPercent)
						$nFloors = GUICtrlRead($nFloors)
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$CCTTDate = GUICtrlRead($editCCTTDate)
						$CCTTID = GUICtrlRead($editCCTTID)
						$projectArea = GUICtrlRead($editProjectArea)
						$duongDoanDuong = GUICtrlRead($editDuong)
						CapDoiCCMSW()
				EndSwitch
			Case $pcNLcd
				Switch $aMsg[0]
					Case $GUI_EVENT_CLOSE
						GUIDelete($pcNLcd)
						GUICtrlSetState($lamPhieuChuyen, $GUI_ENABLE)
					Case $batDauLamPCButton
						$ProjectName = GUICtrlRead($editProjectName)
						$capHang = GUICtrlRead($editCapHang)
						$propertyPercent = GUICtrlRead($editPropertyPercent)
						$nFloors = GUICtrlRead($nFloors)
						$startYear = GUICtrlRead($editStartYear)
						$PriceInContract = GUICtrlRead($editPriceInContract)
						$CCTTDate = GUICtrlRead($editCCTTDate)
						$CCTTID = GUICtrlRead($editCCTTID)
						$projectArea = GUICtrlRead($editProjectArea)
						$duongDoanDuong = GUICtrlRead($editDuong)
						$vitritheoBangGiaDat = GUICtrlRead($editViTriThuaDat)
						CapDoiNLMSW()
				EndSwitch

		EndSwitch
	WEnd
EndFunc

Func ShowUpCapDoiCuGUI()
	$CapDoi = GUICreate("Cap Doi", 916, 660, 219, 78)
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
	$editUserName2 = GUICtrlCreateInput("", 96, 280, 201, 21)
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

	; TAB 1.1
	$TabSheetBenBanCM = GUICtrlCreateTabItem("Ben chuyen nhuong, tang cho")
	GUICtrlCreateLabel("Ho Va Ten", 32, 56, 56, 17)
	$editSellerName1 = GUICtrlCreateInput("", 96, 56, 201, 21)
	GUICtrlCreateLabel("Nam Sinh", 312, 56, 50, 17)
	$editSellerBirth1 = GUICtrlCreateInput("", 376, 56, 97, 21)
	GUICtrlCreateGroup("Gioi tinh", 496, 40, 361, 49)
	$radioSellerSex1 = GUICtrlCreateRadio("Nam", 512, 56, 65, 22)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radioSellerSex11 = GUICtrlCreateRadio("Nu", 600, 56, 113, 22)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	$cmnd1Seller = GUICtrlCreateCombo("CMND", 32, 104, 57, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	$editSellerID1 = GUICtrlCreateInput("", 96, 104, 137, 21)

	GUICtrlCreateLabel("Dia Chi", 32, 152, 38, 17)
	$editSellerAddress1 = GUICtrlCreateEdit("", 96, 152, 537, 89)
	GUICtrlCreateLabel("Ho Va Ten", 32, 280, 56, 17)
	$editSellerName2 = GUICtrlCreateInput("", 96, 280, 201, 21)
	GUICtrlCreateLabel("Nam Sinh", 312, 280, 50, 17)
	$editSellerBirth2 = GUICtrlCreateInput("", 376, 280, 97, 21)
	$cmnd2Seller = GUICtrlCreateCombo("CMND", 32, 322, 57, 25, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
	$editSellerID2 = GUICtrlCreateInput("", 96, 322, 137, 21)
	GUICtrlCreateLabel("Dia Chi", 32, 368, 38, 17)
	$editSellerAddress2 = GUICtrlCreateEdit("", 96, 368, 537, 89)
	$OKSeller = GUICtrlCreateButton("OK", 96, 480, 115, 49)

	; TAB2
	$TabSheet2CM = GUICtrlCreateTabItem("Thua Dat")
	$editLandArea = GUICtrlCreateInput("", 104, 88, 121, 21)
	GUICtrlCreateLabel("Dien Tich", 32, 88, 50, 17)
	GUICtrlCreateLabel("So To Cu", 32, 128, 49, 17)
	GUICtrlCreateLabel("So Thua Cu", 248, 136, 61, 17)
	$editOldPaperNumber = GUICtrlCreateInput("", 104, 128, 121, 21)
	$editOldLandID = GUICtrlCreateInput("", 336, 128, 121, 21)
	GUICtrlCreateGroup("Nguon goc su dung dat", 24, 200, 455, 113)
	$Radio3CM = GUICtrlCreateRadio("Nha nuoc giao dat co thu tien su dung dat", 64, 232, 345, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$Radio4CM = GUICtrlCreateRadio("Cong nhan quyen su dung dat nhu nha nuoc giao dat co thu tien su dung dat", 64, 264, 405, 17)
	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUICtrlCreateGroup("Nguon goc su dung dat 2", 500, 200, 405, 113)
	$radioChuyenNhuong = GUICtrlCreateRadio("Nhan chuyen nhuong", 540, 232, 300, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radioTangCho = GUICtrlCreateRadio("Duoc tang cho", 540, 264, 300, 17)
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
	GUICtrlSetData($comboNha, "Nha lien ke|Nha biet thu|Nha vuon", "Nha o rieng le")
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
	GUICtrlCreateGroup("Nguon goc su dung dat 2", 400, 100, 405, 113)
	$radioChuyenNhuongCC = GUICtrlCreateRadio("Nhan chuyen nhuong", 440, 132, 300, 17)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$radioTangChoCC = GUICtrlCreateRadio("Duoc tang cho", 440, 164, 300, 17)
	$OK4 = GUICtrlCreateButton("OK", 48, 312, 99, 49)

	; TAB 4
	$TabSheet4CM = GUICtrlCreateTabItem("Cap GCN")
	GUICtrlCreateLabel("So Giay Chung Nhan", 56, 96, 104, 17)
	$editGCNID = GUICtrlCreateInput("", 194, 96, 121, 21)
	GUICtrlCreateLabel("So Giay Chung Nhan cu", 56, 140)
	$editOldGCNID = GUICtrlCreateInput("", 194, 140, 121, 21)
	GUICtrlCreateLabel("Loai Giay Chung Nhan cu", 56, 185)
	$editKindOfOldGCN = GUICtrlCreateInput("", 194, 185, 300, 21)
	GUICtrlCreateLabel("Ngay Cap GCN cu", 56, 230)
	$editOldGCNDate = GUICtrlCreateInput("", 194, 230, 121, 21)
	GUICtrlCreateLabel("Noi cap", 56, 275)
	$editOldGCNPlace = GUICtrlCreateInput("", 194, 275, 300, 21)
	GUICtrlCreateLabel("Ten Hop Dong", 56, 320)
	$editContractName = GUICtrlCreateInput("", 194, 320, 300, 21)
	GUICtrlCreateLabel("So cong chung", 56, 365)
	$editContractID = GUICtrlCreateInput("", 194, 365, 121, 21)
	GUICtrlCreateLabel("Hop dong ky ngay", 56, 410)
	$editContractDate = GUICtrlCreateInput("", 194, 410, 121, 21)
	GUICtrlCreateLabel("Cong chung tai", 56, 455)
	$editContractPlace = GUICtrlCreateInput("", 194, 455, 300, 21)
	$Button4CM = GUICtrlCreateButton("So tiep theo", 336, 96, 75, 25)
	$Cap = GUICtrlCreateButton("Cap Nhat GCN", 56, 500, 139, 57)
	$Button3CM = GUICtrlCreateButton("Ve So Do", 224, 500, 139, 57)
	$lamPhieuChuyen = GUICtrlCreateButton("PC, BC, TT", 400, 500, 139, 57)
	GUICtrlCreateTabItem("")
	GUISetState(@SW_SHOW)
EndFunc

Func ShowUpCapMoiGUI()
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
	$editUserName2 = GUICtrlCreateInput("", 96, 280, 201, 21)
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
	GUICtrlSetData($comboNha, "Nha lien ke|Nha biet thu|Nha vuon", "Nha o rieng le")
	$OK3 = GUICtrlCreateButton("OK", 24, 584, 107, 33)

	; TAB 3
	GUICtrlCreateTabItem("Can Ho Chung Cu")
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

Func LamPhieuChuyenCCGUI()
    $pcCC = GUICreate("TT, PC, BC, chungcu", 910, 437, 222, 193)
    GUICtrlCreateLabel("Ten Du An", 48, 30, 56, 17)
	$editProjectName = GUICtrlCreateInput("", 144, 30, 337, 21)
    GUICtrlCreateLabel("Nha Dau Tu", 496, 30, 63, 17)
    $editInves = GUICtrlCreateInput("", 576, 30, 289, 21)
    GUICtrlCreateLabel("Cong ty Ban Nha", 48, 80, 85, 17)
    $editSellerName = GUICtrlCreateInput("", 144, 80, 337, 21)
    GUICtrlCreateLabel("So Tang", 496, 80, 45, 17)
    $editNFloor = GUICtrlCreateInput("", 576, 80, 177, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 130, 78, 17)
    $editStartYear = GUICtrlCreateInput("", 209, 130, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 130, 95, 17)
    $editPriceInContract = GUICtrlCreateInput("", 664, 130, 121, 21)
    GUICtrlCreateLabel("Bien ban ban giao ngay", 48, 180, 116, 17)
    $editDeliveDate = GUICtrlCreateInput("", 209, 180, 121, 21)
    GUICtrlCreateLabel("Gia trong bien ban thanh ly", 496, 180, 131, 17)
    $editPriceInThanhLy = GUICtrlCreateInput("", 664, 180, 121, 21)
    GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 230)
    $editTotalPrice = GUICtrlCreateInput("", 209, 230, 121, 21)
	GUICtrlCreateLabel("Dien tich dat SD chung", 48, 280)
	$editProjectArea = GUICtrlCreateInput("", 209, 280, 121, 21)
	GUICtrlCreateLabel("Duong, doan duong, khu vuc", 48, 330)
	$editDuong = GUICtrlCreateInput("", 209, 330, 121, 21)
	GUICtrlCreateLabel("Hop dong ngay", 496, 230)
	$editContractDate = GUICtrlCreateInput("", 664, 230, 121, 21)
	GUICtrlCreateLabel("Ten hop dong", 496, 280)
	$editContractName = GUICtrlCreateInput("", 664, 280, 121, 21)
	GUICtrlCreateLabel("Hop dong so", 496, 330)
	$editContractID = GUICtrlCreateInput("", 664, 330, 121, 21)
    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 350, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 350, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
EndFunc

Func LamPhieuChuyenChungCuSHNNGUI()
	$ShnnCC = GUICreate("TT, PC, BC, So huu nha nuoc", 910, 437, 222, 193)
    GUICtrlCreateLabel("Ten Du An", 48, 30, 56, 17)
	$editProjectName = GUICtrlCreateInput("", 144, 30, 337, 21)
    GUICtrlCreateLabel("Cong van so", 496, 30)
    $editCongVanID = GUICtrlCreateInput("", 596, 30, 289, 21)
    GUICtrlCreateLabel("Cong ty Ban Nha", 48, 80, 85, 17)
    $editSellerName = GUICtrlCreateInput("", 144, 80, 337, 21)
    GUICtrlCreateLabel("Cong van ngay", 496, 80)
    $editCongVanDate = GUICtrlCreateInput("", 596, 80, 177, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 120, 78, 17)
    $editStartYear = GUICtrlCreateInput("", 200, 120, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 120, 95, 17)
    $editPriceInContract = GUICtrlCreateInput("", 664, 120, 121, 21)
    GUICtrlCreateLabel("Bien ban ban giao ngay", 48, 170, 116, 17)
    $editDeliveDate = GUICtrlCreateInput("", 200, 170, 121, 21)
    GUICtrlCreateLabel("So tien da nop vao TK", 496, 170, 131, 17)
    $editPricePayed = GUICtrlCreateInput("", 664, 170, 121, 21)
    GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 220, 130, 17)
    $editTotalPrice = GUICtrlCreateInput("", 200, 220, 121, 21)
	GUICtrlCreateLabel("Ty le chat luong con lai", 496, 220)
	$editPropertyPercent = GUICtrlCreateInput("", 664, 220, 121, 21)
	GUICtrlCreateLabel("Duong, doan duong, khu vuc", 48, 270)
	$editDuong = GUICtrlCreateInput("", 200, 270, 121, 21)
	GUICtrlCreateLabel("Hop dong ngay", 496, 270)
	$editContractDate = GUICtrlCreateInput("", 664, 270, 121, 21)
	GUICtrlCreateLabel("Dien tich dat SD chung", 48, 320)
	$editProjectArea = GUICtrlCreateInput("", 200, 320, 121, 21)
	GUICtrlCreateLabel("So tang", 496, 320)
	$editNFloor = GUICtrlCreateInput("", 664, 320, 121, 21)
	GUICtrlCreateLabel("Cap, hang nha", 48, 370)
	$editCapHang = GUICtrlCreateInput("", 200, 370, 121, 21)
    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 360, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 360, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
EndFunc

Func LamPhieuChuyenNhaLeGUI()
	$pcNL = GUICreate("TT, PC, BC, nha le", 910, 437, 222, 193)
    GUICtrlCreateLabel("Ten Du An", 48, 30, 56, 21)
	$editProjectName = GUICtrlCreateInput("", 144, 30, 337, 21)
    GUICtrlCreateLabel("Nha Dau Tu", 496, 30, 63, 21)
    $editInves = GUICtrlCreateInput("", 576, 30, 289, 21)
    GUICtrlCreateLabel("Cong ty Ban Nha", 48, 80, 85, 21)
    $editSellerName = GUICtrlCreateInput("", 144, 80, 337, 21)
    GUICtrlCreateLabel("So Tang", 496, 80, 45, 21)
    $editNFloor = GUICtrlCreateInput("", 576, 80, 217, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 130, 78, 21)
    $editStartYear = GUICtrlCreateInput("", 209, 130, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 130, 95, 21)
    $editPriceInContract = GUICtrlCreateInput("", 664, 130, 121, 21)
    GUICtrlCreateLabel("Bien ban ban giao ngay", 48, 180, 116, 21)
    $editDeliveDate = GUICtrlCreateInput("", 209, 180, 121, 21)
    GUICtrlCreateLabel("Gia trong bien ban thanh ly", 496, 180)
    $editPriceInThanhLy = GUICtrlCreateInput("", 664, 180, 121, 21)
    GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 230)
    $editTotalPrice = GUICtrlCreateInput("", 209, 230, 121, 21)
	GUICtrlCreateLabel("Duong, doan duong, khu vuc", 48, 330)
	$editDuong = GUICtrlCreateInput("", 209, 330, 121, 21)
	GUICtrlCreateLabel("Hop dong ngay", 496, 230)
	$editContractDate = GUICtrlCreateInput("", 664, 230, 121, 21)
	GUICtrlCreateLabel("Vi tri thua dat", 496, 280)
	$editViTriThuaDat = GUICtrlCreateInput("", 664, 280, 121, 21)
    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 328, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 328, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
EndFunc

Func LamPhieuChuyenNhaLeSHNNGUI()
	$ShnnNL = GUICreate("TT, PC, BC, So huu nha nuoc", 910, 437, 222, 193)
    GUICtrlCreateLabel("Ten Du An", 48, 30, 56, 17)
	$editProjectName = GUICtrlCreateInput("", 144, 30, 337, 21)
    GUICtrlCreateLabel("Cong van so", 496, 30)
    $editCongVanID = GUICtrlCreateInput("", 596, 30, 289, 21)
    GUICtrlCreateLabel("Cong ty Ban Nha", 48, 80, 85, 17)
    $editSellerName = GUICtrlCreateInput("", 144, 80, 337, 21)
    GUICtrlCreateLabel("Cong van ngay", 496, 80)
    $editCongVanDate = GUICtrlCreateInput("", 596, 80, 177, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 120, 78, 17)
    $editStartYear = GUICtrlCreateInput("", 200, 120, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 120, 95, 17)
    $editPriceInContract = GUICtrlCreateInput("", 664, 120, 121, 21)
    GUICtrlCreateLabel("Bien ban ban giao ngay", 48, 170, 116, 17)
    $editDeliveDate = GUICtrlCreateInput("", 200, 170, 121, 21)
    GUICtrlCreateLabel("So tien da nop vao TK", 496, 170, 131, 17)
    $editPricePayed = GUICtrlCreateInput("", 664, 170, 121, 21)
    GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 220, 130, 17)
    $editTotalPrice = GUICtrlCreateInput("", 200, 220, 121, 21)
	GUICtrlCreateLabel("Ty le chat luong con lai", 496, 220)
	$editPropertyPercent = GUICtrlCreateInput("", 664, 220, 121, 21)
	GUICtrlCreateLabel("Duong, doan duong, khu vuc", 48, 270)
	$editDuong = GUICtrlCreateInput("", 200, 270, 121, 21)
	GUICtrlCreateLabel("Hop dong ngay", 496, 270)
	$editContractDate = GUICtrlCreateInput("", 664, 270, 121, 21)
;~ 	GUICtrlCreateLabel("Dien tich dat SD chung", 48, 320)
;~ 	$editProjectArea = GUICtrlCreateInput("", 200, 320, 121, 21)
	GUICtrlCreateLabel("So tang", 496, 320)
	$editNFloor = GUICtrlCreateInput("", 664, 320, 121, 21)
	GUICtrlCreateLabel("Cap, hang nha", 48, 370)
	$editCapHang = GUICtrlCreateInput("", 200, 370, 121, 21)
    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 350, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 350, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
EndFunc

Func LamPhieuChuyenChungCuTDCGUI()
	$TDC = GUICreate("TT, PC, BC, Tai dinh cu", 910, 437, 222, 193)
    GUICtrlCreateLabel("Ten Du An", 48, 30, 56, 17)
	$editProjectName = GUICtrlCreateInput("", 144, 30, 337, 21)
;~     GUICtrlCreateLabel("Cong van so", 496, 30)
;~     $editCongVanID = GUICtrlCreateInput("", 596, 30, 289, 21)
    GUICtrlCreateLabel("Cong ty Ban Nha", 48, 80, 85, 17)
    $editSellerName = GUICtrlCreateInput("", 144, 80, 337, 21)
;~     GUICtrlCreateLabel("Cong van ngay", 496, 80)
;~     $editCongVanDate = GUICtrlCreateInput("", 596, 80, 177, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 120, 78, 17)
    $editStartYear = GUICtrlCreateInput("", 200, 120, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 30, 95, 17)
    $editPriceInContract = GUICtrlCreateInput("", 664, 30, 121, 21)
    GUICtrlCreateLabel("Bien ban ban giao ngay", 48, 170, 116, 17)
    $editDeliveDate = GUICtrlCreateInput("", 200, 170, 121, 21)
;~     GUICtrlCreateLabel("So tien da nop vao TK", 496, 170, 131, 17)
;~     $editPriceInThanhLy = GUICtrlCreateInput("", 664, 170, 121, 21)
    GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 220, 130, 17)
    $editTotalPrice = GUICtrlCreateInput("", 200, 220, 121, 21)
;~ 	GUICtrlCreateLabel("Ty le chat luong con lai", 496, 220)
;~ 	$editPropertyPercent = GUICtrlCreateInput("", 664, 220, 121, 21)
	GUICtrlCreateLabel("Duong, doan duong, khu vuc", 48, 270)
	$editDuong = GUICtrlCreateInput("", 200, 270, 121, 21)
	GUICtrlCreateLabel("Hop dong ngay", 496, 80)
	$editContractDate = GUICtrlCreateInput("", 664, 80, 121, 21)
;~ 	GUICtrlCreateLabel("Dien tich dat SD chung", 48, 320)
;~ 	$editProjectArea = GUICtrlCreateInput("", 200, 320, 121, 21)
	GUICtrlCreateLabel("So tang", 496, 130)
	$editNFloor = GUICtrlCreateInput("", 664, 130, 121, 21)
	GUICtrlCreateLabel("Cap, hang nha", 48, 320)
	$editCapHang = GUICtrlCreateInput("", 200, 320, 121, 21)
	GUICtrlCreateLabel("Vi tri thua dat", 496, 180)
	$editViTriThuaDat = GUICtrlCreateInput("", 664, 180, 121, 21)
    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 350, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 350, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
EndFunc

Func LamPhieuChuyenCCCDGUI()
	$pcCCcd = GUICreate("TT, PC, BC, chung cu", 910, 437, 222, 193)
    GUICtrlCreateLabel("Ten Du An", 48, 30, 56, 17)
	$editProjectName = GUICtrlCreateInput("", 144, 30, 337, 21)
    GUICtrlCreateLabel("Cap hang nha", 496, 30, 63, 17) ;;;;;;;;
    $editCapHang = GUICtrlCreateInput("", 576, 30, 289, 21)
    GUICtrlCreateLabel("Ty le chat luong con lai", 48, 80, 85, 17) ;;;;;;;
    $propertyPercent = GUICtrlCreateInput("", 144, 80, 337, 21)
    GUICtrlCreateLabel("So Tang", 496, 80, 45, 17)
    $editNFloor = GUICtrlCreateInput("", 576, 80, 177, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 130, 78, 17)
    $editStartYear = GUICtrlCreateInput("", 209, 130, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 130, 95, 17)
    $editPriceInContract = GUICtrlCreateInput("", 664, 130, 121, 21)
    GUICtrlCreateLabel("Cung cap TT ngay", 48, 180, 116, 17) ;;;;;;
    $editCCTTDate = GUICtrlCreateInput("", 209, 180, 121, 21)
    GUICtrlCreateLabel("Cung cap TT so", 496, 180, 131, 17)  ;;;;;;
    $editCCTTID = GUICtrlCreateInput("", 664, 180, 121, 21)
	GUICtrlCreateLabel("Dien tich dat SD chung", 48, 280)
	$editProjectArea = GUICtrlCreateInput("", 209, 280, 121, 21)
	GUICtrlCreateLabel("Duong, doan duong, khu vuc", 48, 330)
	$editDuong = GUICtrlCreateInput("", 209, 330, 121, 21)
    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 350, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 350, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
EndFunc

Func LamPhieuChuyenNLCDGUI()
	$pcNLcd = GUICreate("TT, PC, BC, nha le", 910, 437, 222, 193)
    GUICtrlCreateLabel("Ten Du An", 48, 30, 56, 21)
	$editProjectName = GUICtrlCreateInput("", 144, 30, 337, 21)
    GUICtrlCreateLabel("Cung cap TT ngay", 496, 30, 63, 21) ;;;;;;
    $editCCTTDate = GUICtrlCreateInput("", 576, 30, 289, 21)
    GUICtrlCreateLabel("Cung cap TT so", 48, 80, 85, 21) ;;;;;;;
    $editCCTTID = GUICtrlCreateInput("", 144, 80, 337, 21)
    GUICtrlCreateLabel("So Tang", 496, 80, 45, 21)
    $editNFloor = GUICtrlCreateInput("", 576, 80, 217, 21)
    GUICtrlCreateLabel("Nam hoan cong", 48, 130, 78, 21)
    $editStartYear = GUICtrlCreateInput("", 209, 130, 121, 21)
    GUICtrlCreateLabel("Gia trong hop dong", 496, 130, 95, 21)
    $editPriceInContract = GUICtrlCreateInput("", 664, 130, 121, 21)
    GUICtrlCreateLabel("Ty le chat luong con lai", 48, 180, 116, 21) ;;;;;
    $editPropertyPercent = GUICtrlCreateInput("", 209, 180, 121, 21)
    GUICtrlCreateLabel("Cap hang nha", 496, 180) ;;;;;;
    $editCapHang = GUICtrlCreateInput("", 664, 180, 121, 21)
;~     GUICtrlCreateLabel("Tong gia tren cac hoa don", 48, 230) ;;;;;;
;~     $editTotalPrice = GUICtrlCreateInput("", 209, 230, 121, 21)
	GUICtrlCreateLabel("Duong, doan duong, khu vuc", 48, 330)
	$editDuong = GUICtrlCreateInput("", 209, 330, 121, 21)
;~ 	GUICtrlCreateLabel("Hop dong ngay", 496, 230) ;;;;;;;
;~ 	$editContractDate = GUICtrlCreateInput("", 664, 230, 121, 21)
	GUICtrlCreateLabel("Vi tri thua dat", 496, 280)
	$editViTriThuaDat = GUICtrlCreateInput("", 664, 280, 121, 21)
    $batDauLamPCButton = GUICtrlCreateButton("OK", 424, 328, 139, 57)
    $Button2 = GUICtrlCreateButton("Cap Nhat Du An", 672, 328, 139, 57)
    GUICtrlSetData(-1, "Chon Mau")
    GUISetState(@SW_SHOW)
EndFunc
; #########
;CHO NAY LA DE LAM THEM 4 CAI GIAO DIEN VE CAP NHA
;##########

; CHỌN GIÁ TRỊ TRONG COMBOBOX, CÁC THAM SỐ ĐẦU VÀO BAO GỒM TÊN CỬA SỔ CHỌN, GIÁ TRỊ CONTROL, GIÁ TRỊ CẦN CHỌN: $thetext, GIÁ TRỊ CUỐI THÌ KHỎI QUAN TÂM
Func ChooseMyBox($name, $control, $thetext, $xaorhuyen, $flag = 0)
	$config = WinWaitActive($name)
	ControlClick($config, "", $control)
	Send("{HOME}")
	Sleep(500)
	$count = 0
	$MAX = 70
	For $a = 1 To $MAX
		$count += 1
		$ctext = ControlGetText($name, "", $control)
		If StringInStr($ctext, $thetext) Or StringInStr($thetext, $ctext) Then
			Send("{ENTER}")
			If $flag = 1 Then
				$quanName = $ctext
			ElseIf $flag = 2 Then
				$phuongName = $ctext
			EndIf
			ExitLoop
		EndIf
		Send("{DOWN}")
	Next

	If $count == $MAX Then
		$count = 0
		For $a = 1 To $MAX
			$count += 1
			$ctext = ControlGetText($name, "", $control)
			If StringInStr($ctext, $thetext) Or StringInStr($thetext, $ctext) Then
				Send("{ENTER}")
				If $flag = 1 Then
					$quanName = $ctext
				ElseIf $flag = 2 Then
					$phuongName = $ctext
				EndIf
				ExitLoop
			EndIf
			Send("{UP}")
		Next
	EndIf
	Return $count
EndFunc

; chọn nơi đăng ký lúc đầu (đăng nhập)
;OKOKOKOKOKOKOKOKOKOKOK
Func Configure($huyen, $xa)
	Local $name0 = 'Đăng nhập hệ thống'
	WinActivate($name0)
	Local $windows0 = WinWaitActive($name0)
	ControlClick($windows0, "", "[NAME:btnSetup]")
	WinWaitActive("Cấu hình hệ thống")
	Send("{RIGHT}")
	Send("{RIGHT}")
	Local $name = "Cấu hình hệ thống"
	Local $control1 = '[NAME:cboDistrict]'
	Local $control2 = '[NAME:cboCommune]'
	$count1 = ChooseMyBox($name, $control1, $huyen, "huyen", 1)
	$count2 = ChooseMyBox($name, $control2, $xa, "xa", 2)
	If $count1 < 70 And $count2 < 70 Then
		ControlClick($name, "", "[NAME:btnSave]")
		Local $thongBao = WinWaitActive("Thông báo")
		ControlClick($thongBao, "", '[NAME:cmdCancel]')
		ControlClick($name, "", '[NAME:btnExit]')
		ControlClick($name0, "", '[NAME:btnLogin]')
		Sleep(1000)
		$listWins = WinList()
		For $a = 1 To $listWins[0][0]
			If StringInStr($listWins[$a][0], "Phan mem In giay chung nhan theo luat dat dai 2013") Then
				$theGCNname = $listWins[$a][0]
			EndIf
		Next
	EndIf
EndFunc

;OKOKOKOKOKOKOKOK
Func SetAddress($tinh, $huyen, $xa, $diaPhuong, $controlName)
	Local $winGCN = WinWaitActive($theGCNname)
	ControlClick($winGCN, "", $controlName)
	Local $name = 'Nhập địa chỉ chi tiết'
	Local $Control1 = '[NAME:cboProvince]'
	Local $Control2 = '[NAME:cboDistrict]'
	Local $Control3 = '[NAME:cboCommune]'
	Local $Control4 = '[NAME:txtSoNha]'
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
;OKOKOKOKOKOKOKOKO
Func setUserInfo()
	Local $name = $theGCNname
	WinActivate($name)
	WinWaitActive($name)
	Local $winGCN = WinWaitActive($name)
	MouseClick('primary', 232, 43, 1, 0)
	MouseClick('primary', 281, 62, 1, 0)
	Sleep(1000)
	ControlClick($winGCN, "", '[NAME:btnNewReg]')
	Local $winGCN2 = WinWaitActive($name)
	ControlSetText($winGCN2, "", "[NAME:txtFullName]", $UserName1)
	ControlSetText($winGCN2, "", "[NAME:txtBirthDay]", $UserBirth1)
	;;;;;;;;;;;;;;;
	If $UserSex1 = "b" Then
		ControlClick($theGCNname, "", "[NAME:chkSex]")
	EndIf
	;;;;;;;;;;;;;
	MouseClick("primary", 371, 192)
	Send("{TAB}")
	Send("{TAB}")
	Send("{TAB}")
	Send("{TAB}")
	Send($UserID1[0])
	ControlSetText($winGCN2, "", "[NAME:txtIdentifyNo]", $UserID1[1])
	ControlSend($winGCN2, "", "[NAME:txtIdentifyDate]", $IDdate1)
	$IDdate1 = ControlGetText($theGCNname, "", "[NAME:txtIdentifyDate]")
	; NƠI CẤP CMT, NẾU ĐỂ TRỐNG THÌ TỰ ĐỘNG LẤY GIÁ TRỊ GIẤY CMT MỚI
	If $IDplace1 <> "" Then
		ControlSetText($winGCN2, "", "[NAME:txtIdentifyPlace]", $IDplace1)
	Else
		ControlSetText($winGCN2, "", "[NAME:txtIdentifyPlace]", "Cục CS ĐKQL cư trú và DLQG về Dân cư")
	EndIf
	; VỀ ĐỊA CHỈ CỦA CÁC USER
	$UserAddress11 = StringSplit($UserAddress1, ",")
	Local $lenUserAddress1 = UBound($UserAddress11)

	Local $localName1 = ""
	For $a = 1 To $lenUserAddress1
		If $a < $lenUserAddress1-3 Then
			$localName1 = $localName1 & $UserAddress11[$a]
		EndIf
	Next
	; NHẬP ĐỊA CHỈ CỦA User1
	$buttonUserAddress1 = '[NAME:btnAddressOwner]'
	SetAddress($UserAddress11[$lenUserAddress1-1], $UserAddress11[$lenUserAddress1-2], $UserAddress11[$lenUserAddress1-3], $localName1, $buttonUserAddress1)
	; NẾU TÊN CỦA NGƯỜI THỨ HAI KHÔNG BỎ TRỐNG
	If $UserName2 <> "" Then
		ControlClick($theGCNname, "", "[NAME:chkIsHouseHold]")
		If $UserAddress2 = "" Then
			$UserAddress2 = $UserAddress1
		EndIf
		If $IDplace2 = "" Then
			$IDplace2 = "Cục CS ĐKQL cư trú và DLQG về Dân cư"
		EndIf
		Local $UserNamee2 = "[NAME:txtFullName2]"
		Local $UserBirthh2 = "[NAME:txtBirthDay2]"
		Local $UserIDD2 = "[NAME:txtIdentifyNo2]"
		Local $IDdatee2 = "[NAME:txtIdentifyDate2]"
		Local $IDPlacee2 = "[NAME:txtIssuePlace2]"

		$UserAddress22 = StringSplit($UserAddress2, ",")

		Local $lenUserAddress2 = UBound($UserAddress22)
		Local $localName2 = ""
		For $b = 1 To $lenUserAddress2
			If $b < $lenUserAddress2 - 3 Then
				$localName2 = $localName2 & $UserAddress22[$b]
			EndIf
		Next
		Sleep(300)
		ControlSetText($theGCNname, "", $UserNamee2, $UserName2)
		ControlSetText($theGCNname, "", $UserBirthh2, $UserBirth2)
		MouseClick("primary", 295, 318)
		Send("{TAB}")
		Send("{TAB}")
		Send("{TAB}")
		Send($UserID2[0])
		ControlSetText($theGCNname, "", $UserIDD2, $userid2[1])
		ControlSend($theGCNname, "", $IDdatee2, $IDdate2)
		$IDdate2 = ControlGetText($theGCNname, "", $IDdatee2)
		ControlSetText($theGCNname, "", $IDplacee2, $IDplace2)
		$buttonUserAddress2 = "[NAME:btnAddressOwner2]"
		SetAddress($UserAddress22[$lenUserAddress2-1], $UserAddress22[$lenUserAddress2-2], $UserAddress22[$lenUserAddress2-3], $localName2, $buttonUserAddress2)
		If $isInHoOngBa = 1 Then
			ControlClick($theGCNname, "", "[NAME:chkIsPrintHouseHold]")
		EndIf

		ElseIf $isInHoOngBa = 1 Then
		ControlClick($theGCNname, "", "[NAME:chkIsHouseHold]")
		ControlClick($theGCNname, "", "[NAME:chkIsPrintHouseHold]")
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
	ControlClick($name, "", "[NAME:btnAddAllOwner]")
EndFunc
;OKOKOKOKOKOKOKOK
Func nhapThuaDat()
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	MouseClick("primary", 350, 115)
	WinWaitActive($theGCNname)
	; Nhan vao button them thua dat
	ControlClick($theGCNname, "", "[NAME:btnAddParcel]")
	ControlSend($theGCNname, "", "[NAME:txtMapNo]", "999")
	Sleep(1000)
	Local $maxNumber = ControlGetText($theGCNname, "", "[NAME:lblMaxParcelNo]")
	$arrayMax = StringSplit($maxNumber, " ")
	$realMaxNumber = Number($arrayMax[$arrayMax[0]]) + 1
	ControlSend($theGCNname, "", "[NAME:txtParcelNo]", String($realMaxNumber))
	Sleep(300)
	MouseClick("primary", 374, 155)
	Sleep(300)
	MouseClick("primary", 621, 152)
	ControlSetText($theGCNname, "", "[NAME:txtArea]", $landArea)
	ControlSetText($theGCNname, "", "[NAME:txtParcelNoTmp]", $oldLandID)
	ControlSetText($theGCNname, "", "[NAME:txtMapNoTmp]", $oldPaperNumber)
	ControlCommand($theGCNname, "", "[NAME:chkPrintParcelTmp]", "Check")
	ControlSetText($theGCNname, "", "[NAME:txtSoNha]", $local)
	ControlClick($theGCNname, "", "[NAME:btnChoiseOriginParcel]")
	WinWaitActive("Chọn nguồn gốc sử dụng")
	If $originalOfLand = "case1" Then
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{SPACE}")
	Else
		Send("{SPACE}")
	EndIf
	If $isTangCho = 1 Then
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{SPACE}")
	ElseIf $isChuyenNhuong = 1 Then
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{UP}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{DOWN}")
		Send("{SPACE}")
	EndIf
	ControlClick("Chọn nguồn gốc sử dụng", "", "[NAME:btnChoise]")
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnUpdateParcel]")
	Send("{ENTER}")
	Sleep(1000)
	If WinExists("Thông báo") Then
		Send("{ENTER}")
	EndIf
	WinActivate($theGCNname)
	ChooseMyBox($theGCNname, "[NAME:cboPurposeCertificate]", "ODT", "có gì đó sai sai")
	ChooseMyBox($theGCNname, "[NAME:cboPurposePlan]", "ODT", "có gì đó sai sai")
	ChooseMyBox($theGCNname, "[NAME:cboPurposeInventory]", "ODT", "có gì đó sai sai")
	;Cap nhat
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnUpdateRegParcel]")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Send("{ENTER}")
	WinWaitActive($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnAddAllRegParcel]")
EndFunc
;OKOKOKOKOKOKOKOKO
Func nhapNhaOThongThuong()
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	MouseClick("primary", 422, 116)
	WinWaitActive($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnAddBuilding]")
	If $loaiNhaO = "nha o rieng le" Then
		ChooseMyBox($theGCNname, "[NAME:cboLoaiTaiSan]", "riêng lẻ", "st wrong")
	ElseIf $loaiNhaO = "Nha lien ke" Then
		ChooseMyBox($theGCNname, "[NAME:cboLoaiTaiSan]", "liền kề", "st wrong")
	ElseIf $loaiNhaO = "Nha biet thu" Then
		ChooseMyBox($theGCNname, "[NAME:cboLoaiTaiSan]", "biệt thự", "st wrong")
	ElseIf $loaiNhaO = "Nha vuon" Then
		ChooseMyBox($theGCNname, "[NAME:cboLoaiTaiSan]", "vườn", "st wrong")
	EndIf
	ControlSetText($theGCNname, "", "[NAME:txtFloorArea]", $dienTichSanKCC)
	ControlSetText($theGCNname, "", "[NAME:txtUsingArea]", $dienTichSuDungKCC)
	ControlSetText($theGCNname, "", "[NAME:txtBuildingArea]", $DienTichXayDungKCC)
	ControlClick($theGCNname, "", "[NAME:btnUpdateBuilding]")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Send("{ENTER}")
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnAddAllBuilding]")
EndFunc
;OKOKOKOKOK
Func setValueCC()
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	MouseClick('primary', 432, 116)
	WinWaitActive($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnAddBuilding]")
	; chon chung cu
	MsgBox(0, "Thong bao", "Hay nhan ok khi ban da chon xong chung cu")
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	$cbkindOfHouse = "[NAME:cboLoaiTaiSan]"
	ChooseMyBox($theGCNname, $cbkindOfHouse, "căn hộ chung cư", "AHIHI")
	Local $soCanHo = "[NAME:txtBuildingNo]"
	ControlSetText($theGCNname, "", $soCanHo, $propertyAddress)
	Local $dienTichSan = "[NAME:txtFloorArea]"
	Local $Ssudung = "[NAME:txtUsingArea]"
	Local $xayDung = "[NAME:txtBuildingArea]"
	ControlSetText($theGCNname, "", $dienTichSan, $propertyArea)
	ControlSetText($theGCNname, "", $Ssudung, $DienTichSuDungCC)
	; an nut cap nhat
	ControlClick($theGCNname, "", "[NAME:btnUpdateBuilding]")
	Sleep(500)
	Send("{ENTER}")
	Sleep(500)
	Send("{ENTER}")
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnAddAllBuilding]")
EndFunc
;OKOKOKOKOKOKOKOK
Func CapGiay()
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	MouseClick('primary', 53, 114)
	WinWaitActive($theGCNname)
	Local $maHoSo = "[NAME:txtMaDon]"
	ControlSetText($theGCNname, "", $maHoSo, $recordID)
	Local $ngayHoSo = "[NAME:txtRegDate]"
	ControlSend($theGCNname, "", $ngayHoSo, $recordDate)
	$recordDate = ControlGetText($theGCNname, "", $ngayHoSo)
	Send("{F3}")
	WinWaitActive("Thông báo")
	Send("{ENTER}")
	WinActivate($theGCNname)
	WinWaitActive($theGCNname)
	Sleep(1000)
	Local $themGiayMoi = "[NAME:btnAddNewCertificate]"
	ControlClick($theGCNname, "", $themGiayMoi)
	MouseClick("primary", 489, 201)
	MouseClick("primary", 940, 201)
	Local $soCap = "[NAME:chkAdministrativeUnit2]"
	ControlCommand($theGCNname, "", $soCap, "Check", "")
	$LaySeries = "[NAME:btnMaxOriginDocumentNo]"
	ControlClick($theGCNname, "", $LaySeries)
	Local $soTiepTheo = "[NAME:cmdQuanLySoHieuGCN]"
	If $flagGCNID = 1 Then
		ControlClick($theGCNname, "", $soTiepTheo)
	EndIf
	; NẾU LÀ TẶNG CHO HOẶC CHUYỂN NHƯỢNG THÌ SẼ CÓ GHI CHÚ
	If $isCapDoiC = 1 Or $isCapDoiM = 1 Then
		ControlSend($theGCNname, "", "[NAME:txtDescriptionPage2]", "- Số tờ, số thửa sẽ được điều chỉnh khi có bản đồ địa chính chính quy")
		Local $uuu
		If $UserSex1 = "o" Then
			; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
			If $UserName2 = "" Then
				$uuu = "ông " & $UserName1
			; Truong hop co hai nguoi, va nguoi dung ten la nam
			ElseIf $UserName2 <> "" Then
				$uuu = "ông " & $UserName1 & " và bà " & $UserName2
			EndIf
		; Truong hop User1 la nu
		Else
			If $UserName2 = "" Then
				$uuu = "bà " & $UserName1
			Else
				$uuu = "bà " & $UserName1 & " ông " & $UserName2
			EndIf
		EndIf
		$cntc = "nhận chuyển nhượng"
		If $isTangCho = 1 Then
			$cntc = "được tặng cho"
		EndIf
		ControlSend($theGCNname, "", "[NAME:txtDescriptionPage2]", "- Giấy chứng nhận này được cấp đổi từ " & $kindOfOldGCN & " số " & $oldGCNID & " được " & $oldGCNPlace & " cấp ngày " & $oldGCNDate & " do " & $uuu & " " & $cntc & " theo " & $contractName & " số công chứng " & $contractID & " ký ngày " & $contractDate & " tại " & $contractPlace)
	EndIf

	; KẾT THÚC GHI CHÚ
	Local $capNhat = "[NAME:btnUpdateCertificate]"
	ControlClick($theGCNname, "", $capNhat)
	WinWaitActive("Thông báo")
	Send("{ENTER}")
	WinActivate($theGCNname)
	ControlClick($theGCNname, "", "[NAME:btnViewCertificate]")
EndFunc
;PHAN TAO PHIEU CHUYEN
Func CapMoiNhaDuAnMSW()
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
		Local $destinationPath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\" & $destinationFileName
		Local $sourcePath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\sampleChungCu.doc"
		FileCopy($sourcePath, $destinationPath)
		; MO FILE DA COPY LEN VA SUA
		Local $oWord = _Word_Create()
		Local $sDocument = $destinationPath
		Local $oDoc = _Word_DocOpen($oWord, $sDocument)
		WinWaitActive($recordID & " " & StringUpper($UserName1) & " [Compatibility Mode] - Microsoft Word")
		sleep(1000)
		_Word_DocFindReplace($oDoc, "recordID", $recordID)
		_Word_DocFindReplace($oDoc, "recordDate", $recordDate)

		If $UserSex1 = "o" Then
			; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
			If $UserName2 = "" Then
				_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
				_Word_DocFindReplace($oDoc, "bà UserName1", "ông " & $UserName1)
				_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
			; Truong hop co hai nguoi, va nguoi dung ten la nam
			ElseIf $UserName2 <> "" Then
				_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
				_Word_DocFindReplace($oDoc, "bà UserName1", "ông " & $UserName1)
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
        Local $UserIDD1 = $UserID1[0] & " " & "số " & $UserID1[1]
        _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
        _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
        _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)

        ;now about property
        _Word_DocFindReplace($oDoc, "propertyAddress", $propertyAddress)
        ;#O#OO3O3O3O3O3O3O3O3O3 HASN'T INFORMATION ABOUT PROJECT
		_Word_DocFindReplace($oDoc, "PropertyArea", $propertyArea)

        _Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)
        _Word_DocFindReplace($oDoc, "projectArea", $projectArea)
        _Word_DocFindReplace($oDoc, "nFloors", $nFloors)
		_Word_DocFindReplace($oDoc, "vitriTheoBangGiaDat", $duongDoanDuong)
        _Word_DocFindReplace($oDoc, "startYear", $startYear)
        _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
        _Word_DocFindReplace($oDoc, "ContractName", $contractName)
        _Word_DocFindReplace($oDoc, "contractID", $contractID)
        _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
        _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
        _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
        _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
        _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
        _Word_DocFindReplace($oDoc, "quanName", $quanName)
        _Word_DocFindReplace($oDoc, "SellerName", $sellerName)
		_Word_DocFindReplace($oDoc, "priceInLiquidation", $priceInLiquidation)

;~ 		_Word_DocSaveAs($oDoc, "D:\Unity\lol.doc")
;~ 		_Word_Quit($oWord)
	ElseIf $propertyAddress = "" Then
		; COPY SAMPLE FILE
		Local $destinationFileName = $recordID & " " & StringUpper($UserName1) & ".doc"
		Local $destinationPath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\" & $destinationFileName
		Local $sourcePath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\sampleNhaDat.doc"
		FileCopy($sourcePath, $destinationPath)
		; MO FILE DA COPY LEN VA SUA
		Local $oWord = _Word_Create()
		Local $sDocument = $destinationPath
		Local $oDoc = _Word_DocOpen($oWord, $sDocument)
		WinWaitActive($recordID & " " & StringUpper($UserName1) & " [Compatibility Mode] - Microsoft Word")
		sleep(1000)
		_Word_DocFindReplace($oDoc, "recordID", $recordID)
		_Word_DocFindReplace($oDoc, "recordDate", $recordDate)

		If $UserSex1 = "o" Then
            ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
            If $UserName2 = "" Then

                Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
                _Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
                _Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
                _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            ; Truong hop co hai nguoi, va nguoi dung ten la nam
            ElseIf $UserName2 <> "" Then
                _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
                _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
                If $UserAddress1 <> $UserAddress2 Then
                    Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
                    _Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
                Else
                    Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
                    _Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
                EndIf
                _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
                _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
            EndIf
        ; Truong hop User1 la nu
        Else
            If $UserName2 = "" Then
                Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
                _Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
                _Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
                _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
                _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            Else
                If $UserAddress1 <> $UserAddress2 Then
                    Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
                    _Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
                Else
                    Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
                    _Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
                EndIf
                _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
                _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
                _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            EndIf
        EndIf

        _Word_DocFindReplace($oDoc, "UserAddress1", $UserAddress1)
        Local $UserIDD1 = $UserID1[0] & "số " & $UserID1[1]
        If $UserName2 <> "" Then
            $UserIDD2 = $UserID2[0] & "số " & $UserID2[1]
        EndIf
        _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
        _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
        _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)
        _Word_DocFindReplace($oDoc, "UserID2", $UserIDD2)
        _Word_DocFindReplace($oDoc, "IDdate2", $IDdate2)
        _Word_DocFindReplace($oDoc, "IDplace2", $IDplace2)
        _Word_DocFindReplace($oDoc, "UserBirth1", $UserBirth1)
        _Word_DocFindReplace($oDoc, "UserBirth2", $UserBirth2)

        ;now about property
		Local $landAddress = $local & " " & $phuongName & " " & $quanName & " thành phố Hà Nội"
        _Word_DocFindReplace($oDoc, "landAddress", $landAddress)
        ;#O#OO3O3O3O3O3O3O3O3O3 HASN'T INFORMATION ABOUT PROJECT
		_Word_DocFindReplace($oDoc, "PropertyArea", $propertyArea)

        _Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)
        _Word_DocFindReplace($oDoc, "landArea", $landArea)
        _Word_DocFindReplace($oDoc, "nFloors", $nFloors)
		_Word_DocFindReplace($oDoc, "viTriThuaDat", $vitritheoBangGiaDat)
		_Word_DocFindReplace($oDoc, "vitriTheoBangGiaDat", $duongDoanDuong)
		Local $nguonGoc
		If $originalOfLand = "case1" Then
			$nguonGoc = "Nhà nước giao đất có thu tiền sử dụng đất"
		Else
			$nguonGoc = "Nhà nước công nhận quyền sử dụng đất như nhà nước giao đất có thu tiền sử dụng đất"
		EndIf
		_Word_DocFindReplace($oDoc, "NguonGoc", $nguonGoc)
		_Word_DocFindReplace($oDoc, "loaiNhaO", $loaiNhaO)
		_Word_DocFindReplace($oDoc, "dienTichXayDung", $DienTichXayDungKCC)
		_Word_DocFindReplace($oDoc, "capHangNha", $capHang)
		_Word_DocFindReplace($oDoc, "propertyPercent", $propertyPercent)
        _Word_DocFindReplace($oDoc, "startYear", $startYear)
        _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
        _Word_DocFindReplace($oDoc, "ContractName", $contractName)
        _Word_DocFindReplace($oDoc, "contractID", $contractID)
        _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
        _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
        _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
		_Word_DocFindReplace($oDoc, "moneyPayed", $pricePayed)
        _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
        _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
        _Word_DocFindReplace($oDoc, "quanName", $quanName)
        _Word_DocFindReplace($oDoc, "SellerName", $sellerName)

;~ 		_Word_DocSaveAs($oDoc, "D:\Unity\lol.doc")
;~ 		_Word_Quit($oWord)
	EndIf

EndFunc

Func CapMoiSHNNCCMSW()
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


	Local $destinationFileName = $recordID & " " & StringUpper($UserName1) & ".doc"
    Local $destinationPath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\01.SỞ HỮU NHÀ NƯỚC\" & $destinationFileName
    Local $sourcePath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\01.SỞ HỮU NHÀ NƯỚC\sampleSHNNChungCu.doc"
    FileCopy($sourcePath, $destinationPath)
    ; MO FILE DA COPY LEN VA SUA
    Local $oWord = _Word_Create()
    Local $sDocument = $destinationPath
    Local $oDoc = _Word_DocOpen($oWord, $sDocument)
    WinWaitActive($recordID & " " & StringUpper($UserName1) & " [Compatibility Mode] - Microsoft Word")
    sleep(1000)
    _Word_DocFindReplace($oDoc, "recordID", $recordID)
    _Word_DocFindReplace($oDoc, "recordDate", $recordDate)

    If $UserSex1 = "o" Then
        ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
        If $UserName2 = "" Then
			Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        ; Truong hop co hai nguoi, va nguoi dung ten la nam
        ElseIf $UserName2 <> "" Then
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
			_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
        EndIf
    ; Truong hop User1 la nu
    Else
        If $UserName2 = "" Then
			Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        Else
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        EndIf
    EndIf

    _Word_DocFindReplace($oDoc, "UserAddress1", $UserAddress1)
    Local $UserIDD1 = $UserID1[0] & "số " & $UserID1[1]
	If $UserName2 <> "" Then
		$UserIDD2 = $UserID2[0] & "số " & $UserID2[1]
	EndIf
    _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
    _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
    _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)
	_Word_DocFindReplace($oDoc, "UserID2", $UserIDD2)
    _Word_DocFindReplace($oDoc, "IDdate2", $IDdate2)
    _Word_DocFindReplace($oDoc, "IDplace2", $IDplace2)
	_Word_DocFindReplace($oDoc, "UserBirth1", $UserBirth1)
	_Word_DocFindReplace($oDoc, "UserBirth2", $UserBirth2)

    _Word_DocFindReplace($oDoc, "propertyAddress", $propertyAddress)
    _Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)
    _Word_DocFindReplace($oDoc, "projectArea", $projectArea)
    _Word_DocFindReplace($oDoc, "nFloors", $nFloors)

    _Word_DocFindReplace($oDoc, "startYear", $startYear)
    _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
    _Word_DocFindReplace($oDoc, "ContractName", $contractName)
    _Word_DocFindReplace($oDoc, "contractID", $contractID)
    _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
    _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
    _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
    _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
    _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
    _Word_DocFindReplace($oDoc, "quanName", $quanName)
    _Word_DocFindReplace($oDoc, "SellerName", $sellerName)
	_Word_DocFindReplace($oDoc, "PropertyArea", $propertyArea)
	_Word_DocFindReplace($oDoc, "vitritheoBangGiaDat", $duongDoanDuong)
	_Word_DocFindReplace($oDoc, "capHangNha", $capHang)
	_Word_DocFindReplace($oDoc, "propertyPercent", $propertyPercent)
	_Word_DocFindReplace($oDoc, "moneyPayed", $pricePayed)
	_Word_DocFindReplace($oDoc, "congVanDate", $congVanDate)
	_Word_DocFindReplace($oDoc, "congVanID", $congVanID)
EndFunc

Func CapMoiSHNNNLMSW()
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


	Local $destinationFileName = $recordID & " " & StringUpper($UserName1) & ".doc"
    Local $destinationPath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\01.SỞ HỮU NHÀ NƯỚC\" & $destinationFileName
    Local $sourcePath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\01.SỞ HỮU NHÀ NƯỚC\sampleSHNNNhaDat.doc"
    FileCopy($sourcePath, $destinationPath)
    ; MO FILE DA COPY LEN VA SUA
    Local $oWord = _Word_Create()
    Local $sDocument = $destinationPath
    Local $oDoc = _Word_DocOpen($oWord, $sDocument)
    WinWaitActive($recordID & " " & StringUpper($UserName1) & " [Compatibility Mode] - Microsoft Word")
    sleep(1000)
    _Word_DocFindReplace($oDoc, "recordID", $recordID)
    _Word_DocFindReplace($oDoc, "recordDate", $recordDate)

    If $UserSex1 = "o" Then
        ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
        If $UserName2 = "" Then

			Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        ; Truong hop co hai nguoi, va nguoi dung ten la nam
        ElseIf $UserName2 <> "" Then
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
			_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
        EndIf
    ; Truong hop User1 la nu
    Else
        If $UserName2 = "" Then
			Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        Else
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        EndIf
    EndIf

    _Word_DocFindReplace($oDoc, "UserAddress1", $UserAddress1)
    Local $UserIDD1 = $UserID1[0] & "số " & $UserID1[1]
	If $UserName2 <> "" Then
		$UserIDD2 = $UserID2[0] & "số " & $UserID2[1]
	EndIf
    _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
    _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
    _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)
	_Word_DocFindReplace($oDoc, "UserID2", $UserIDD2)
    _Word_DocFindReplace($oDoc, "IDdate2", $IDdate2)
    _Word_DocFindReplace($oDoc, "IDplace2", $IDplace2)
	_Word_DocFindReplace($oDoc, "UserBirth1", $UserBirth1)
	_Word_DocFindReplace($oDoc, "UserBirth2", $UserBirth2)
	Local $landAddress = $local & " " & $phuongName & " " & $quanName & " thành phố Hà Nội"
    _Word_DocFindReplace($oDoc, "landAddress", $landAddress)
    _Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)
    _Word_DocFindReplace($oDoc, "landArea", $landArea)
    _Word_DocFindReplace($oDoc, "nFloors", $nFloors)
	Local $nguonGoc
	If $originalOfLand = "case1" Then
		$nguonGoc = "Nhà nước giao đất có thu tiền sử dụng đất"
	Else
		$nguonGoc = "Nhà nước công nhận quyền sử dụng đất như nhà nước giao đất có thu tiền sử dụng đất"
	EndIf
	_Word_DocFindReplace($oDoc, "loaiNhaO", $loaiNhaO)
	_Word_DocFindReplace($oDoc, "NguonGoc", $nguonGoc)
    _Word_DocFindReplace($oDoc, "startYear", $startYear)
    _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
    _Word_DocFindReplace($oDoc, "ContractName", $contractName)
    _Word_DocFindReplace($oDoc, "contractID", $contractID)
    _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
    _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
    _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
    _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
    _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
    _Word_DocFindReplace($oDoc, "quanName", $quanName)
    _Word_DocFindReplace($oDoc, "SellerName", $sellerName)
	_Word_DocFindReplace($oDoc, "PropertyArea", $propertyArea)
	_Word_DocFindReplace($oDoc, "vitritheoBangGiaDat", $duongDoanDuong)
	_Word_DocFindReplace($oDoc, "capHangNha", $capHang)
	_Word_DocFindReplace($oDoc, "propertyPercent", $propertyPercent)
	_Word_DocFindReplace($oDoc, "moneyPayed", $pricePayed)
	_Word_DocFindReplace($oDoc, "congVanDate", $congVanDate)
	_Word_DocFindReplace($oDoc, "congVanID", $congVanID)

;~      _Word_DocSaveAs($oDoc, "D:\Unity\lol.doc")
;~      _Word_Quit($oWord)
EndFunc

Func CapMoiTDCMSW()
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


	Local $destinationFileName = $recordID & " " & StringUpper($UserName1) & ".doc"
    Local $destinationPath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\_TÁI ĐỊNH CƯ" & $destinationFileName
    Local $sourcePath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\01. BC CẤP MỚI\_TÁI ĐỊNH CƯ\sampleTDCChungCu.doc"
    FileCopy($sourcePath, $destinationPath)
    ; MO FILE DA COPY LEN VA SUA
    Local $oWord = _Word_Create()
    Local $sDocument = $destinationPath
    Local $oDoc = _Word_DocOpen($oWord, $sDocument)
    WinWaitActive($recordID & " " & StringUpper($UserName1) & " [Compatibility Mode] - Microsoft Word")
    sleep(1000)
    _Word_DocFindReplace($oDoc, "recordID", $recordID)
    _Word_DocFindReplace($oDoc, "recordDate", $recordDate)

    If $UserSex1 = "o" Then
        ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
        If $UserName2 = "" Then

			Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        ; Truong hop co hai nguoi, va nguoi dung ten la nam
        ElseIf $UserName2 <> "" Then
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
			_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
        EndIf
    ; Truong hop User1 la nu
    Else
        If $UserName2 = "" Then
			Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        Else
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        EndIf
    EndIf

    _Word_DocFindReplace($oDoc, "UserAddress1", $UserAddress1)
    Local $UserIDD1 = $UserID1[0] & "số " & $UserID1[1]
	If $UserName2 <> "" Then
		$UserIDD2 = $UserID2[0] & "số " & $UserID2[1]
	EndIf
    _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
    _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
    _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)
	_Word_DocFindReplace($oDoc, "UserID2", $UserIDD2)
    _Word_DocFindReplace($oDoc, "IDdate2", $IDdate2)
    _Word_DocFindReplace($oDoc, "IDplace2", $IDplace2)
	_Word_DocFindReplace($oDoc, "UserBirth1", $UserBirth1)
	_Word_DocFindReplace($oDoc, "UserBirth2", $UserBirth2)

    _Word_DocFindReplace($oDoc, "propertyAddress", $propertyAddress)
    _Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)
    _Word_DocFindReplace($oDoc, "projectArea", $projectArea)
    _Word_DocFindReplace($oDoc, "nFloors", $nFloors)

    _Word_DocFindReplace($oDoc, "startYear", $startYear)
    _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
    _Word_DocFindReplace($oDoc, "ContractName", $contractName)
    _Word_DocFindReplace($oDoc, "contractID", $contractID)
    _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
    _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
    _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
    _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
    _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
    _Word_DocFindReplace($oDoc, "quanName", $quanName)
    _Word_DocFindReplace($oDoc, "SellerName", $sellerName)
	_Word_DocFindReplace($oDoc, "PropertyArea", $propertyArea)
	_Word_DocFindReplace($oDoc, "vitritheoBangGiaDat", $duongDoanDuong)
	_Word_DocFindReplace($oDoc, "viTriThuaDat", $vitritheoBangGiaDat)
	_Word_DocFindReplace($oDoc, "capHangNha", $capHang)
	_Word_DocFindReplace($oDoc, "propertyPercent", $propertyPercent)
	_Word_DocFindReplace($oDoc, "moneyPayed", $pricePayed)
	_Word_DocFindReplace($oDoc, "congVanDate", $congVanDate)
	_Word_DocFindReplace($oDoc, "congVanID", $congVanID)

;~      _Word_DocSaveAs($oDoc, "D:\Unity\lol.doc")
;~      _Word_Quit($oWord)
EndFunc

Func CapDoiNLMSW()
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


	Local $destinationFileName = $recordID & " " & StringUpper($UserName1) & ".doc"
    Local $destinationPath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\02. BC BIẾN ĐỘNG\" & $destinationFileName
    Local $sourcePath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\02. BC BIẾN ĐỘNG\sampleLeBD.doc"
    FileCopy($sourcePath, $destinationPath)
    ; MO FILE DA COPY LEN VA SUA
    Local $oWord = _Word_Create()
    Local $sDocument = $destinationPath
    Local $oDoc = _Word_DocOpen($oWord, $sDocument)
    WinWaitActive($recordID & " " & StringUpper($UserName1) & " [Compatibility Mode] - Microsoft Word")
    sleep(1000)
    _Word_DocFindReplace($oDoc, "recordID", $recordID)
    _Word_DocFindReplace($oDoc, "recordDate", $recordDate)
	$landAddress = $local & " " & $phuongName & " " & $quanName & " thành phố Hà Nội"
	If  $sellerSex1 = 1 Then
        ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
        If $sellerName2 = "" Then
			Local $sellerInfo = "Ông SellerNamen1, sinh năm SellerBirth1, SellerID1, hộ khẩu thường trú tại SellerAddress1"
			_Word_DocFindReplace($oDoc, "SellerInfo", $sellerInfo)
			_Word_DocFindReplace($oDoc, "ông SellerNamen1 và bà SellerNamen2", "ông " & $sellerName1)
            _Word_DocFindReplace($oDoc, "SellerNamen1", $sellerName1)
        ; Truong hop co hai nguoi, va nguoi dung ten la nam
        ElseIf $sellerName2 <> "" Then
            _Word_DocFindReplace($oDoc, "SellerNamen1", $sellerName1)
            _Word_DocFindReplace($oDoc, "SellerNamen2", $sellerName2)
			If $sellerAddress1 <> $sellerAddress2 Then
				Local $sellerInfo = "Ông SellerNamen1, sinh năm SellerBirth1, SellerID1, hộ khẩu thường trú tại SellerAddress1 và vợ là bà SellerNamen2, sinh năm SellerBirth2, SellerID2, hộ khẩu thường trú tại SellerAddress2"
				_Word_DocFindReplace($oDoc, "SellerInfo", $sellerInfo)
			Else
				Local $sellerInfo = "Ông SellerNamen1, sinh năm SellerBirth1, SellerID1 và vợ là bà SellerNamen2, sinh năm SellerBirth2, SellerID2, cả hai cùng đăng ký hộ khẩu thường trú tại SellerAddress1"
				_Word_DocFindReplace($oDoc, "SellerInfo", $sellerInfo)
			EndIf
			_Word_DocFindReplace($oDoc, "SellerNamen1", $sellerName1)
            _Word_DocFindReplace($oDoc, "SellerNamen2", $sellerName2)
        EndIf
    ; Truong hop User1 la nu
    Else
        If $UserName2 = "" Then
			Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        Else
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        EndIf
    EndIf

	If $UserSex1 = "o" Then
        ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
        If $UserName2 = "" Then
			Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        ; Truong hop co hai nguoi, va nguoi dung ten la nam
        ElseIf $UserName2 <> "" Then
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
			_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
        EndIf
    ; Truong hop User1 la nu
    Else
        If $UserName2 = "" Then
			Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        Else
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        EndIf
    EndIf

    _Word_DocFindReplace($oDoc, "UserAddress1", $UserAddress1)
	_Word_DocFindReplace($oDoc, "UserAddress2", $UserAddress2)


    Local $UserIDD1 = $UserID1[0] & " số " & $UserID1[1]
	Local $UserIDD2
	If $UserName2 <> "" Then
		$UserIDD2 = $UserID2[0] & " số " & $UserID2[1]
	EndIf

	Local $sellerIDD1 = $sellerID1[0] & " số " & $sellerID1[1]
	Local $sellerIDD2
	If $UserName2 <> "" Then
		$sellerIDD2 = $sellerID2[0] & " số " & $sellerID2[1]
	EndIf
    _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
    _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
    _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)
	_Word_DocFindReplace($oDoc, "UserID2", $UserIDD2)
    _Word_DocFindReplace($oDoc, "IDdate2", $IDdate2)
    _Word_DocFindReplace($oDoc, "IDplace2", $IDplace2)
	_Word_DocFindReplace($oDoc, "UserBirth1", $UserBirth1)
	_Word_DocFindReplace($oDoc, "UserBirth2", $UserBirth2)

	_Word_DocFindReplace($oDoc, "SellerID1", $sellerID1)
	_Word_DocFindReplace($oDoc, "SellerID2", $sellerID2)
	_Word_DocFindReplace($oDoc, "SellerBirth1", $sellerBirth1)
	_Word_DocFindReplace($oDoc, "SellerBirth2", $sellerBirth2)

    _Word_DocFindReplace($oDoc, "startYear", $startYear)
    _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
    _Word_DocFindReplace($oDoc, "ContractName", $contractName)
    _Word_DocFindReplace($oDoc, "contractID", $contractID)
    _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
    _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
    _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
    _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
    _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
    _Word_DocFindReplace($oDoc, "quanName", $quanName)
	_Word_DocFindReplace($oDoc, "PropertyArea", $dienTichSanKCC)
	_Word_DocFindReplace($oDoc, "vitritheoBangGiaDat", $duongDoanDuong)
	_Word_DocFindReplace($oDoc, "capHangNha", $capHang)
	_Word_DocFindReplace($oDoc, "propertyPercent", $propertyPercent)
	_Word_DocFindReplace($oDoc, "moneyPayed", $pricePayed)
	_Word_DocFindReplace($oDoc, "congVanDate", $congVanDate)
	_Word_DocFindReplace($oDoc, "congVanID", $congVanID)

	_Word_DocFindReplace($oDoc, "landArea", $landArea)
	Local $loaiNghiaVuTaiChinh
	Local $loaiNghiaVuTaiChinhRutGon
	If $isChuyenNhuong = 1 Then
		$loaiNghiaVuTaiChinh = "Nhận chuyển nhượng"
		$loaiNghiaVuTaiChinhRutGon = "chuyển nhượng"
	Else
		$loaiNghiaVuTaiChinh = "Được tặng cho"
		$loaiNghiaVuTaiChinhRutGon = "tặng cho"
	EndIf
	_Word_DocFindReplace($oDoc, "loaiNghiaVuTaiChinhRutGon", $loaiNghiaVuTaiChinhRutGon)
	_Word_DocFindReplace($oDoc, "loaiNghiaVuTaiChinh", $loaiNghiaVuTaiChinh)
	_Word_DocFindReplace($oDoc, "vanBanCungCapThongTinDay", $CCTTDate)
	_Word_DocFindReplace($oDoc, "vanBanCungCapThongTinID", $CCTTID)
	_Word_DocFindReplace($oDoc, "viTriThuaDat", $vitritheoBangGiaDat)
	_Word_DocFindReplace($oDoc, "landID", $oldLandID)
	_Word_DocFindReplace($oDoc, "paperNumber", $oldPaperNumber)
	_Word_DocFindReplace($oDoc, "LandAddress", $landAddress)
	_Word_DocFindReplace($oDoc, "dienTichXayDung", $DienTichXayDungKCC)
	_Word_DocFindReplace($oDoc, "contractPlace", $contractPlace)
	_Word_DocFindReplace($oDoc, "kindOfOldGCN", $kindOfOldGCN)
	_Word_DocFindReplace($oDoc, "oldGCNID", $oldGCNID)
	_Word_DocFindReplace($oDoc, "oldGCNDate", $oldGCNDate)
	_Word_DocFindReplace($oDoc, "dvCapOldGCN", $oldGCNPlace)
	_Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)

	_Word_DocFindReplace($oDoc, "nFloors", $nFloors)
	_Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
;~ 	_Word_DocFindReplace($oDoc, "vanBanCungCapThongTinDay", $vitritheoBangGiaDat)
EndFunc

Func CapDoiCCMSW()
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


	Local $destinationFileName = $recordID & " " & StringUpper($UserName1) & ".doc"
    Local $destinationPath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\02. BC BIẾN ĐỘNG\" & $destinationFileName
    Local $sourcePath = "\\10.1.2.3\thu muc dung chung toan co quan\PHÒNG ĐK&CGCN\HIỀN\02. BC BIẾN ĐỘNG\sampleLeBD.doc"
    FileCopy($sourcePath, $destinationPath)
    ; MO FILE DA COPY LEN VA SUA
    Local $oWord = _Word_Create()
    Local $sDocument = $destinationPath
    Local $oDoc = _Word_DocOpen($oWord, $sDocument)
    WinWaitActive($recordID & " " & StringUpper($UserName1) & " [Compatibility Mode] - Microsoft Word")
    sleep(1000)
    _Word_DocFindReplace($oDoc, "recordID", $recordID)
    _Word_DocFindReplace($oDoc, "recordDate", $recordDate)

	If  $sellerSex1 = 1 Then
        ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
        If $sellerName2 = "" Then
			Local $sellerInfo = "Ông SellerNamen1, sinh năm SellerBirth1, SellerID1, hộ khẩu thường trú tại SellerAddress1"
			_Word_DocFindReplace($oDoc, "SellerInfo", $sellerInfo)
			_Word_DocFindReplace($oDoc, "ông SellerNamen1 và bà SellerNamen2", "ông " & $sellerName1)
            _Word_DocFindReplace($oDoc, "SellerNamen1", $sellerName1)
        ; Truong hop co hai nguoi, va nguoi dung ten la nam
        ElseIf $sellerName2 <> "" Then
            _Word_DocFindReplace($oDoc, "SellerNamen1", $sellerName1)
            _Word_DocFindReplace($oDoc, "SellerNamen2", $sellerName2)
			If $sellerAddress1 <> $sellerAddress2 Then
				Local $sellerInfo = "Ông SellerNamen1, sinh năm SellerBirth1, SellerID1, hộ khẩu thường trú tại SellerAddress1 và vợ là bà SellerNamen2, sinh năm SellerBirth2, SellerID2, hộ khẩu thường trú tại SellerAddress2"
				_Word_DocFindReplace($oDoc, "SellerInfo", $sellerInfo)
			Else
				Local $sellerInfo = "Ông SellerNamen1, sinh năm SellerBirth1, SellerID1 và vợ là bà SellerNamen2, sinh năm SellerBirth2, SellerID2, cả hai cùng đăng ký hộ khẩu thường trú tại SellerAddress1"
				_Word_DocFindReplace($oDoc, "SellerInfo", $sellerInfo)
			EndIf
			_Word_DocFindReplace($oDoc, "SellerNamen1", $sellerName1)
            _Word_DocFindReplace($oDoc, "SellerNamen2", $sellerName2)
        EndIf
    ; Truong hop User1 la nu
    Else
        If $UserName2 = "" Then
			Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        Else
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        EndIf
    EndIf

	If $UserSex1 = "o" Then
        ; Truong hop chi co mot nguoi dung ten, va nguoi do la nam
        If $UserName2 = "" Then
			Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "ông " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        ; Truong hop co hai nguoi, va nguoi dung ten la nam
        ElseIf $UserName2 <> "" Then
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
			_Word_DocFindReplace($oDoc, "UserName1", $UserName1)
            _Word_DocFindReplace($oDoc, "UserName2", $UserName2)
        EndIf
    ; Truong hop User1 la nu
    Else
        If $UserName2 = "" Then
			Local $buyerInfo = "Bà UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1"
			_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			_Word_DocFindReplace($oDoc, "ông UserName1 và bà UserName2", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        Else
			If $UserAddress1 <> $UserAddress2 Then
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1, hộ khẩu thường trú tại UserAddress1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, hộ khẩu thường trú tại UserAddress2"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			Else
				Local $buyerInfo = "Ông UserName1, sinh năm UserBirth1, UserID1 và vợ là bà UserName2, sinh năm UserBirth2, UserID2, cả hai cùng đăng ký hộ khẩu thường trú tại UserAddress1"
				_Word_DocFindReplace($oDoc, "BuyerInfo", $buyerInfo)
			EndIf
            _Word_DocFindReplace($oDoc, "ông UserName1", "bà " & $UserName1)
            _Word_DocFindReplace($oDoc, "bà UserName2", "ông " & $UserName2)
            _Word_DocFindReplace($oDoc, "UserName1", $UserName1)
        EndIf
    EndIf

    _Word_DocFindReplace($oDoc, "UserAddress1", $UserAddress1)
	_Word_DocFindReplace($oDoc, "UserAddress2", $UserAddress2)


    Local $UserIDD1 = $UserID1[0] & " số " & $UserID1[1]
	Local $UserIDD2
	If $UserName2 <> "" Then
		$UserIDD2 = $UserID2[0] & " số " & $UserID2[1]
	EndIf

	Local $sellerIDD1 = $sellerID1[0] & " số " & $sellerID1[1]
	Local $sellerIDD2
	If $UserName2 <> "" Then
		$sellerIDD2 = $sellerID2[0] & " số " & $sellerID2[1]
	EndIf
    _Word_DocFindReplace($oDoc, "UserID1", $UserIDD1)
    _Word_DocFindReplace($oDoc, "IDdate1", $IDdate1)
    _Word_DocFindReplace($oDoc, "IDplace1", $IDplace1)
	_Word_DocFindReplace($oDoc, "UserID2", $UserIDD2)
    _Word_DocFindReplace($oDoc, "IDdate2", $IDdate2)
    _Word_DocFindReplace($oDoc, "IDplace2", $IDplace2)
	_Word_DocFindReplace($oDoc, "UserBirth1", $UserBirth1)
	_Word_DocFindReplace($oDoc, "UserBirth2", $UserBirth2)

	_Word_DocFindReplace($oDoc, "SellerID1", $sellerID1)
	_Word_DocFindReplace($oDoc, "SellerID2", $sellerID2)
	_Word_DocFindReplace($oDoc, "SellerBirth1", $sellerBirth1)
	_Word_DocFindReplace($oDoc, "SellerBirth2", $sellerBirth2)

    _Word_DocFindReplace($oDoc, "startYear", $startYear)
    _Word_DocFindReplace($oDoc, "InvestorName", $InvestorName)
    _Word_DocFindReplace($oDoc, "ContractName", $contractName)
    _Word_DocFindReplace($oDoc, "contractID", $contractID)
    _Word_DocFindReplace($oDoc, "contractDate", $contractDate)
    _Word_DocFindReplace($oDoc, "priceInContract", $PriceInContract)
    _Word_DocFindReplace($oDoc, "deliveryPaperDate", $deliveryPaperDate)
    _Word_DocFindReplace($oDoc, "totalPrice", $totalPrice)
    _Word_DocFindReplace($oDoc, "phuongName", $phuongName)
    _Word_DocFindReplace($oDoc, "quanName", $quanName)

	_Word_DocFindReplace($oDoc, "PropertyArea", $propertyArea)
	_Word_DocFindReplace($oDoc, "vitritheoBangGiaDat", $duongDoanDuong)
	_Word_DocFindReplace($oDoc, "capHangNha", $capHang)
	_Word_DocFindReplace($oDoc, "propertyPercent", $propertyPercent)
	_Word_DocFindReplace($oDoc, "moneyPayed", $pricePayed)
	_Word_DocFindReplace($oDoc, "congVanDate", $congVanDate)
	_Word_DocFindReplace($oDoc, "congVanID", $congVanID)

	Local $loaiNghiaVuTaiChinh
	Local $loaiNghiaVuTaiChinhRutGon
	If $isChuyenNhuong = 1 Then
		$loaiNghiaVuTaiChinh = "Nhận chuyển nhượng"
		$loaiNghiaVuTaiChinhRutGon = "chuyển nhượng"
	Else
		$loaiNghiaVuTaiChinh = "Được tặng cho"
		$loaiNghiaVuTaiChinhRutGon = "tặng cho"
	EndIf
	Local $landAddress = $local & " " & $phuongName & " " & $quanName & " thành phố Hà Nội"
	_Word_DocFindReplace($oDoc, "loaiNghiaVuTaiChinhRutGon", $loaiNghiaVuTaiChinhRutGon)
	_Word_DocFindReplace($oDoc, "loaiNghiaVuTaiChinh", $loaiNghiaVuTaiChinh)
	_Word_DocFindReplace($oDoc, "vanBanCungCapThongTinDay", $CCTTDate)
	_Word_DocFindReplace($oDoc, "vanBanCungCapThongTinID", $CCTTID)
	_Word_DocFindReplace($oDoc, "viTriThuaDat", $vitritheoBangGiaDat)
	_Word_DocFindReplace($oDoc, "paperNumber", $oldPaperNumber)
	_Word_DocFindReplace($oDoc, "dienTichXayDung", $DienTichXayDungKCC)
	_Word_DocFindReplace($oDoc, "contractPlace", $contractPlace)
	_Word_DocFindReplace($oDoc, "kindOfOldGCN", $kindOfOldGCN)
	_Word_DocFindReplace($oDoc, "oldGCNID", $oldGCNID)
	_Word_DocFindReplace($oDoc, "oldGCNDate", $oldGCNDate)
	_Word_DocFindReplace($oDoc, "dvCapOldGCN", $oldGCNPlace)
	_Word_DocFindReplace($oDoc, "propertyAddress", $propertyAddress)
	_Word_DocFindReplace($oDoc, "ProjectName", $ProjectName)
	_Word_DocFindReplace($oDoc, "projectArea", $projectArea)
	_Word_DocFindReplace($oDoc, "nFloors", $nFloors)
	_Word_DocFindReplace($oDoc, "priceInLiquidation", $priceInLiquidation)
EndFunc