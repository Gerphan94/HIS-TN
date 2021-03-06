#NoEnv
#NoTrayIcon
#SingleInstance Ignore
FileEncoding, UTF-8

;AHK SETUP
GuiWidth := "w330"
Tabwidth := "w310"
AHK_tittle := "Tiếp nhận"
Global g_Resolution := A_ScreenWidth . "x" . A_ScreenHeight
;Gắn các giá trị theo tỷ lệ màn hình sử dụng
;Màn hình tỷ lệ 720p
If ( g_Resolution != "1366x768") AND ( g_Resolution != "1920x1080" )
{
	MsgBox, 48, % "Oop!",% "Lỗi độ phần giải màn hình`n1366x768 hoặc 1920x1080`nHiện tại: " . g_Resolution
	ExitApp
}
;FIle
exe := A_ScriptDir + "/Reception.exe"
Global path_titleFolder := A_ScriptDir . "\data\Title"
Global path_ModuleTitle := A_ScriptDir . "\data\ModuleTitle"
Global path_Config := A_ScriptDir . "\config.ini"
global path_PosConfig := A_ScriptDir . "\PosConfig.ini"
Global path_Company := A_ScriptDir . "\data\Company"
Global path_addressName := A_ScriptDir . "\data\Address"
global path_DepartFolde := A_ScriptDir . "\data\Depart"
IfNotExist, % path_ModuleTitle
{
	MsgBox, 48, % "Oop!",% "File" . path_ModuleTitle . " không tồn tại"
	ExitApp
}
arr_phongkham4 := ["Phòng cấp cứu"]
Global arr_doituong1 := ["Thu phí","BHYT"]
Global arr_doituong2 := ["Thu phí","BHYT", "Thẻ tạm"]
global arr_addcode := ["HCTDTB","DIBHTD","DILKXU","DILKPB","DILKBV","DILKXL","DILKBS","DILKBV"]

Global arr_addressName := fn_readSimpleFile(path_addressName)
Global arr_Company := fn_GetArrayOfCompany()
global arr_ModuleTitle := initArrayTitle()
; global arr_HISType := HIStypeList()
Global arr_DepartGroup := fn_GetlistFile(path_DepartFolde)
Global arr_PKList := {}
;KHỞI TẠO BIẾN
;Các biến Position 
Global HIS_PK, HIS_PLV, HIS_MaBN, HIS_MuaSo, HIS_BNUT, HIS_BHYT5nam, HIS_SoCT, HIS_Ver, HIS_Email, HIS_CMND, HIS_Nguoithan
Global xHIS_PK, xHIS_PLV, xHIS_MaBN, xHIS_MuaSo, xHIS_BNUT, xHIS_BHYT5nam, xHIS_SoCT, xHIS_Ver, xHIS_Email, xHIS_CMND, xHIS_Nguoithan
Global yHIS_PK, yHIS_PLV, yHIS_MaBN, yHIS_MuaSo, yHIS_BNUT, yHIS_BHYT5nam, yHIS_SoCT, yHIS_Ver, yHIS_Email, yHIS_CMND, yHIS_Nguoithan
;Các biến file config
Global iPOS, Phongkhamid, Doituongid, SoBNcallQMS, g_DepartGID, g_Email, g_PKNhi
Global gTITLE_TN , gTITLE_CC , gTITLE_TNVP , EXEFILE
Global g_MaBV , SoBNcallQMS , g_HISver , g_HISType

global g_nWait := 2
;Khởi tạo data
initDATA()
initPOS()
initTitle()
TN_thuong := {"Phòng Hành chính - Quầy tiếp nhận 1":1,"Phòng Hành chính - Quầy tiếp nhận 2":1}
TN_dichvu := {"Phòng Hành chính - Quầy tiếp nhận VIP 1":1}
TN_capcuu := {"Khoa Cấp cứu - Phòng cấp cứu":1}
TN_khoasan := {"Khoa Sản phụ - Tiếp nhận Sản":1}
only_female := {"PK Khoa sản":1, "PK Phụ khoa":1, "PK Thai":1}
depart_diff := {"PK Nội tổng quát 1":1, "PK Nội tổng quát 2":1,"PK Hô hấp":1,"PK Nội tiết":1,"PK Nội tim mạch":1}

;GUI
if FileExist(A_ScriptDir + "\Reception.exe")
	Menu, Tray, icon, % A_ScriptDir . "\Reception.exe"
Menu, FileMenu, Add,
Menu, FileMenu, Add, % "Logfile", gotologfile
Menu, FileMenu, Add, % "Mở TN Thường", OpenTNThuong
Menu, FileMenu, Add,
Menu, FileMenu, Add, % "Reload (F12)", Reload
Menu, FileMenu, Add,
Menu, FileMenu, Add, % "Exit", MenuExit

Menu, OptionMenu, Add, % "Xem Log", Log
Menu, OptionMenu, Add, % "Danh mục", Catelogy
Menu, OptionMenu, Add,
Menu, OptionMenu, Add, % "Config", Config
Menu, InfoMenu, Add, Thông tin, Info 
Menu, MyMenuBar, Add, &File, :FileMenu
Menu, MyMenuBar, Add, &Option, :OptionMenu
Menu, MyMenuBar, Add, &Info, :InfoMenu

Gui, Menu, MyMenuBar

Gui, 1:Font, s11 BOLD c000080 
Gui, 1:Add, Tab3, x10 %Tabwidth% vmyTAB ,% "Tiếp nhận|."
Gui, 1:Tab, 1
Gui, 1:Font, s10 cNavy, Segoe UI
Gui, 1:Font, s10 normal c000000
Gui, 1:Add, Text, x20 y60 w120 h23 +0x200, % "Loại TN:"
Gui, 1:Add, Text, x200 y60 h23 +0x200, % "Số BN:"
Gui, 1:Add, Text, x20 y90 w120 h23 +0x200, % "Bệnh nhân:"
Gui, 1:Add, Text, x20 y120 w120 h23 +0x200, % "Phòng khám:"
Gui, 1:Add, Text, x20 y150 w120 h23 +0x200, % "Đối tượng:"
Gui, 1:Add, DropDownList, x100 y60 w80 vddlLoaiTN gddlLoaiTN, % "Thường||Cấp cứu"
Gui, 1:Add, DropDownList, x255 y60 w50 vddlSoBN,
Gui, 1:Add, DropDownList, x100 y90 w80 vddlBenhnhan gddlBenhnhan choose1, % "||Trẻ em"
Gui, 1:Add, DropDownList, x100 y120 w210 vddlPhongkham gddlPhongkham choose1, % ConvertARRtoString(arr_PKList)
Gui, 1:Add, DropDownList, x100 y150 w90 vddlDoituong gddlDoituong choose2, % ConvertARRtoString(arr_doituong1)
Gui, 1:Add, DropDownList, x200 y150 w110 vddlTuyen, % "Đúng tuyến||Thông tuyến|Chuyển tuyến"
Gui, 1:Font, s10
Gui, 1:Add, Text, x100 y190 w195 h2 w190 +0x10
Gui, 1:Add, Checkbox, x100 y200 h23 vcbBHYT5nam, % "BHYT 5 năm"
Gui, 1:Add, CheckBox, x100 y225 h23 vcbMuaSKB, % "Mua sổ KB"
Gui, 1:Add, CheckBox, x220 y225 h23 vcbBNUuTien, % "Ưu tiên"
Gui, 1:Add, CheckBox, x100 y250 w70 h23 vcbDST, % "DST"
Gui, 1:Add, CheckBox, x220 y250 w150 h23 vTNcddv disabled, % "Chỉ định DV"
Gui, 1:Add, Checkbox, x100 y275 vcbThutien, % "Thu tiền"
Gui, 1:Tab
Gui, 1:Tab, 2
Gui, Font, s9
Gui, 1:Add, Checkbox, x30 y60 vcbFull gcbFull, % "Full"
Gui, 1:Add, Checkbox, x30 y80 vcbESC, % "Email, SĐT, Công ty"
Gui, 1:Add, Checkbox, x30 y100 vcbCMND, % "CMND"
Gui, 1:Add, Checkbox, x30 y120 vcbNguoithan, % "Người thân"
Gui, 1:Font
Gui, 1:Tab

Gui, 1:Add, Groupbox, x10 y308 h55 w310
Gui, 1:Add, Button, x250 y325 w60 h32 vbtnRunTN gbtnRunTN, % "Bắt đầu"
Gui, 1:Font, s8
Gui, 1:Add, StatusBar,, 
SB_SetParts(90,65,50)
SB_SetText(" " . g_HISver, 1)
SB_SetText(A_ScreenWidth . "x" . A_ScreenHeight, 2)
SB_SetText(a_hour . ":" . a_min . ":" . a_sec, 3)
SB_SetText(fn_dayVietNam(), 4)

Gui, 1:Font
Gui, 1:Show, %GUIwidth% , %AHK_tittle%
Gui, 1:Submit, Nohide

initSOBNTN(SoBNcallQMS)
SetTimer, RefreshTime, 1000
;Gán Control của GUI
ctrl_BHYT5Nam := "Button1"
ctrl_MuaSKB := "Button2"
ctrl_BNUT := "Button3"
ctrl_DST := "Button4"
ctrl_CDDV := "Button5"
ctrl_Thutien := "Button6"
ctrl_Full := "Button7"
ctrl_ESC := "Button8"
ctrl_CMND := "Button9"
ctrl_Nguoithan := "Button10"
;Gui 2 - Thông tin
Gui, 2:Font, s11 bold cNavy
Gui, 2:Add, Text, x20 y20, % "Thông tin"
Gui, 2:Font
Gui, 2:Add, Edit, x20 y40 w300 h400 v2_edtINFO +ReadOnly,

Gui, 3:Default
Gui, 3:Font, s14 cNavy BOLD, Segoe UI
Gui, 3:Add, Text, x20 y20, % "CẤU HÌNH"
Gui, 3:Font
Gui, 3:Font, s10 cNavy BOLD
Gui, 3:Add, Tab3, x20 y60 w550 h270 vmyTab, % "Chung|Pos"
Gui, 3:Font
Gui, 3:Tab, 1
Gui, 3:Font, s8 cBlack, Segoe UI
Gui, 3:Add, Text, x40 y100 h24 +0x200, % "Số BN/Gọi:"
Gui, 3:Add, DropDownList, x110 y100 w70 v3_ddlSoBN, % "1|2|3|4|5|6|7"
Gui, 3:Add, Text, x340 y100 h24 +0x200 , % "POS:"
Gui, 3:Add, DropDownList, x380 y100 w70 v3_ddlPOS, % "Default|Custom"
Gui, 3:Add, Text, x40 y130 h24 +0x200, % "PK cho TE:"
Gui, 3:Add, DropDownList, x110 y130 w150 v3_ddlPKNhi, % ConvertARRtoString(arr_PKList)
Gui, 3:Add, Text, x40 y160 h24 +0x200, % "Đuôi Email:"
Gui, 3:Add, Edit, x110 y160 w150 +0x200 v3_edtEmail,
Gui, 3:Font
Gui, 3:Tab
Gui, 3:Tab, 2
;Line 1
Gui, 3:Add, Text, x40 y100 w90 h24 +0x200 v3_lb1, % "Phòng làm việc:"
Gui, 3:Add, Text, x140 y100 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x170 y100 w40 v3_edtxPLV Disabled Right,
Gui, 3:Add, Text, x220 y100 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x250 y100 w40 v3_edtyPLV Disabled Right,
Gui, 3:Add, Text, x330 y100 w60 h24 +0x200, % "Mã BN:"
Gui, 3:Add, Text, x400 y100 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x430 y100 w40 v3_edtxMaBN Disabled Right, 
Gui, 3:Add, Text, x480 y100 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x510 y100 w40 v3_edtyMaBN Disabled Right, 
;Line 2
Gui, 3:Add, Text, x40 y125 w90 h24 +0x200, % "Số CT:"
Gui, 3:Add, Text, x140 y125 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x170 y125 w40 v3_edtxSoCT Disabled Right, 
Gui, 3:Add, Text, x220 y125 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x250 y125 w40 v3_edtySoCT Disabled Right, 
Gui, 3:Add, Text, x330 y125 w60 h24 +0x200, % "BN Ưu tiên:"
Gui, 3:Add, Text, x400 y125 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x430 y125 w40 v3_edtxBNUT Disabled Right, 
Gui, 3:Add, Text, x480 y125 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x510 y125 w40 v3_edtyBNUT Disabled Right, 
;Line 3
Gui, 3:Add, Text, x40 y150 w90 h24 +0x200, % "BHYT 5 năm:"
Gui, 3:Add, Text, x140 y150 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x170 y150 w40 v3_edtxBHYT5nam Disabled Right,
Gui, 3:Add, Text, x220 y150 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x250 y150 w40 v3_edtyBHYT5nam Disabled Right, 
Gui, 3:Add, Text, x330 y150 w60 h24 +0x200, % "Mua sổ KB:"
Gui, 3:Add, Text, x400 y150 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x430 y150 w40 v3_edtxMuaSOKB Disabled Right, 
Gui, 3:Add, Text, x480 y150 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x510 y150 w40 v3_edtyMuaSOKB Disabled Right, 
;Line 4
Gui, 3:Add, Text, x40 y175 w90 h24 +0x200, % "Phòng khám:"
Gui, 3:Add, Text, x140 y175 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x170 y175 w40 v3_edtxPK Disabled Right,
Gui, 3:Add, Text, x220 y175 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x250 y175 w40 v3_edtyPK Disabled Right,
Gui, 3:Add, Text, x330 y175 w60 h24 +0x200 , % "Version:"
Gui, 3:Add, Text, x400 y175 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x430 y175 w40 v3_edtxVersion Disabled Right,  
Gui, 3:Add, Text, x480 y175 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x510 y175 w40 v3_edtyVersion Disabled Right,
;Line 5
Gui, 3:Add, Text, x40 y200 w90 h24 +0x200, % "Email:"
Gui, 3:Add, Text, x140 y200 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x170 y200 w40 v3_edtxEmail Disabled Right,
Gui, 3:Add, Text, x220 y200 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x250 y200 w40 v3_edtyEmail Disabled Right,
Gui, 3:Add, Text, x330 y200 w60 h24 +0x200 , % "CMND:"
Gui, 3:Add, Text, x400 y200 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x430 y200 w40 v3_edtxCMND Disabled Right,
Gui, 3:Add, Text, x480 y200 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x510 y200 w40 v3_edtyCMND Disabled Right,
;Line 6
Gui, 3:Add, Text, x40 y225 w90 h24 +0x200, % "Người thân:"
Gui, 3:Add, Text, x140 y225 w20 h24 +0x200 cBlue, % "X ="
Gui, 3:Add, Edit, x170 y225 w40 v3_edtxNguoithan Disabled Right,
Gui, 3:Add, Text, x220 y225 w20 h24 +0x200 cBlue, % "Y ="
Gui, 3:Add, Edit, x250 y225 w40 v3_edtyNguoithan Disabled Right, 
Gui, 3:Add, Groupbox, x40 y260 w510 h60
Gui, 3:Add, Checkbox, x330 y280 h21 v3_cbfindPos g3_cbfindPos, % "PosFinder"
Gui, 3:Add, DropDownList, x400 y280 w120 v3_ddlField, % "|Phòng làm việc|Mã BN|Số CT|BN Ưu tiên|BHYT 5 năm|Mua sổ KB|Phòng khám|Version|Email|CMND|Người thân"
Gui, 3:Tab
Gui, 3:Add, Button, x420 y340 w70 h32 g3_btnSave, % "Lưu"
Gui, 3:Add, Button, x500 y340 w70 h32 g3_btnClose, % "Đóng"
SetTimer, Update, 250
SetTimer, Update, Off
Gui, 3:Add, StatusBar,,
SB_SetParts(250,20,250)

;GUI 4 - Xem log
Gui, 4:Default ;Không có dòng này sẽ không format được listview
Gui, 4:Font, S16 cNavy BOLD
GUi, 4:Add, Text, x20 y20 w150 , % "LOG SỰ KIỆN"
Gui, 4:Font
Gui, 4:Add, Text, x20 y60 w30 h23 +0x200, % "Ngày:"
Gui, 4:Add, DateTime, x60 y60 w100 v4_date, dd/MM/yyyy
Gui, 4:Add, Button, x170 y60 w70 g4_btnView, % "Xem"
Gui, 4:Add, ListView, x20 y90 w660 h300 +Grid vMYLIST +AltSubmit, % "STT|Thời gian|Mã TN|Mã BN|Họ tên|Phòng khám|TG gọi số|TG lưu BN|TG lưu TN"
	LV_ModifyCol(1, "40 Center")
	LV_ModifyCol(2, "120 Center")
	LV_ModifyCol(3, "100 Center")
	LV_ModifyCol(4, "70 Center")
	LV_ModifyCol(5, "140 Right")
	LV_ModifyCol(6, "100 Right")
	LV_ModifyCol(7, "80 Right")
	LV_ModifyCol(8, "80 Right")
	LV_ModifyCol(9, "80 Right")
Gui, 4:Add, StatusBar

Gui, 5:Default
Gui, 5:Font, s16 cMaroon Bold
Gui, 5:Add, Text, x20 y20 , % "Danh mục"
Gui, 5:Font

Gui, 5:Font, s10 bold cGray
Gui, 5:Add,Tab3, x20 y50 w380 h260 v5_tab, % "Màn hình|Phòng"
Gui, 5:Font
Gui, 5:Font, s9
Gui, 5:Tab, 1
Gui, 5:Add, DropDownList, x110 y90 w100 v5_ddlHisType g5_ddlHisType, % ConvertARRtoString(fn_GetlistFile(path_titleFolder))
Gui, 5:Add, Text, x40 y120 w60 h21 +0x200, % "TN thường:"
Gui, 5:ADd, Edit, x110 y120 w270 h21 v5_edtTNT +ReadOnly,
Gui, 5:Add, Text, x40 y150 w60 h24 +0x200, % "TN CC:"
Gui, 5:ADd, Edit, x110 y150 w270 h21 v5_edtTNCC +ReadOnly, 
Gui, 5:Add, Text, x40 y180 w60 h24 +0x200, % "TN VP:"
Gui, 5:ADd, Edit, x110 y180 w270 h21 v5_edtTNVP +ReadOnly,

Gui, 5:Tab
Gui, 5:Tab, 2
Gui, 5:Add, Text, x40 y90 h24 +0x200, % "Nhóm:"
Gui, 5:Add, DropDownList, x90 y90 w100 v5_ddlNhom g5_ddlNhom, % ConvertARRtoString(arr_DepartGroup)
Gui, 5:Add, ListView, x200 y90 w180 h200 v5LISTVIEW -Hdr +Grid +AltSubmit, % "column name"
	LV_ModifyCol(1, "175 Left")
Gui, 5:Tab
Gui, 5:Add, Button, x250 y315 w70 h32 g5_btnSave, % "Lưu"
Gui, 5:Add, Button, x330 y315 w70 h32 g5_btnCLose, % "Đóng"
Gui, 5:Add, StatusBar

SB_SetText("ReadOnly")
Return

MenuExit:
GuiEscape:
GuiClose:
    ExitApp
	Return
F12::
	MsgBox, 48, Warning!, Dừng Script
Reload:
	Reload
	Return
; Khởi tạo Data từ file
initDATA()
{
	IniRead, iPOS, % path_Config, section1, iPOS
	IniRead, SoBNcallQMS, % path_Config, section1, SoBN
	IniRead, g_MaBV, % path_Config, section1, MaBV
	IniRead, g_HISver, % path_Config, section1, version
	IniRead, g_HISType, % path_Config, section1, HISType
	IniRead, g_DepartGID, % path_Config, section1, DepartID
	IniRead, g_Email, % path_Config, section1, Email
	IniRead, g_PKNhi, % path_Config, section1, PKNhi
	arr_PKList := fn_readSimpleFile(path_DepartFolde . "\" . arr_DepartGroup[g_DepartGID])
	Return
}
initTitle()
{
	filename_path := A_ScriptDir . "\data\Title\" . fn_GetlistFile(path_titleFolder)[g_HISType]
	gTITLE_TN := fn_GetTitleModule(filename_path).TN
	gTITLE_CC := fn_GetTitleModule(filename_path).CC
	gTITLE_TNVP := fn_GetTitleModule(filename_path).VP
	Return
}
initPOS()
{
	IniRead, HIS_PK, % path_PosConfig , Custom, HIS_PK
	IniRead, HIS_PLV, % path_PosConfig , Custom, HIS_PLV
	IniRead, HIS_MaBN, % path_PosConfig , Custom, HIS_MaBN
	IniRead, HIS_MuaSo, % path_PosConfig , Custom, HIS_MuaSo
	IniRead, HIS_BNUT, % path_PosConfig , Custom, HIS_BNUT
	IniRead, HIS_BHYT5nam, % path_PosConfig , Custom, HIS_BHYT5nam
	IniRead, HIS_SoCT, % path_PosConfig , Custom, HIS_SoCT
	IniRead, HIS_Ver, % path_PosConfig , Custom, HIS_Ver
	IniRead, HIS_Email, % path_PosConfig , Custom, HIS_Email
	IniRead, HIS_CMND, % path_PosConfig , Custom, HIS_CMND
	IniRead, HIS_Nguoithan, % path_PosConfig , Custom, HIS_Nguoithan
	;Nếu sử dụng Custom
	If ( iPOS = 2 )
	{
		xHIS_PLV := fn_getXY(HIS_PLV).xPOS , yHIS_PLV := fn_getXY(HIS_PLV).yPOS
		xHIS_MaBN := fn_getXY(HIS_MaBN).xPOS , yHIS_MaBN := fn_getXY(HIS_MaBN).yPOS
		xHIS_BHYT5nam := fn_getXY(HIS_BHYT5nam).xPOS , yHIS_BHYT5nam := fn_getXY(HIS_BHYT5nam).yPOS
 		xHIS_SoCT := fn_getXY(HIS_SoCT).xPOS , yHIS_SoCT := fn_getXY(HIS_SoCT).yPOS
		xHIS_MuaSo := fn_getXY(HIS_Muaso).xPOS , yHIS_MuaSo := fn_getXY(HIS_Muaso).yPOS
		xHIS_BNUT := fn_getXY(HIS_BNUT).xPOS , yHIS_BNUT := fn_getXY(HIS_BNUT).yPOS
		xHIS_PK := fn_getXY(HIS_PK).xPOS , yHIS_PK := fn_getXY(HIS_PK).yPOS
		xHIS_Ver := fn_getXY(HIS_Ver).xPOS , yHIS_Ver := fn_getXY(HIS_Ver).yPOS
		xHIS_Email := fn_getXY(HIS_Email).xPOS , yHIS_Email := fn_getXY(HIS_Email).yPOS
		xHIS_CMND := fn_getXY(HIS_CMND).xPOS , yHIS_CMND := fn_getXY(HIS_CMND).yPOS
		xHIS_Nguoithan := fn_getXY(HIS_Nguoithan).xPOS , yHIS_Nguoithan := fn_getXY(HIS_Nguoithan).yPOS
	}
	Else
	{
		If (g_Resolution = "1366x768") {
			xHIS_PK := 120 , yHIS_PK := 429
			xHIS_PLV := 1020 , yHIS_PLV := 36
			xHIS_Ver := 1275 , yHIS_Ver := 722
			xHIS_MaBN := 124 , yHIS_MaBN := 120
			xHIS_MuaSo := 940 , yHIS_MuaSo := 430
			xHIS_BNUT := 1064 , yHIS_BNUT := 431
			xHIS_cb5nam := 721 , yHIS_cb5nam := 320
			xHIS_SoCT := 124 , yHIS_SoCT := 348
			xHIS_Email := 130 , yHIS_Email := 277
			xHIS_CMND := 130 , yHIS_CMND := 367
			xHIS_Nguoithan := 130 , yHIS_Nguoithan := 397
		}
		Else if (g_Resolution = "1920x1080")
		{
			xHIS_PK := 120 , yHIS_PK := 429
			xHIS_PLV := 1574 , yHIS_PLV := 36
			xHIS_Ver := 1829 , yHIS_Ver := 1034
			xHIS_MaBN := 124 , yHIS_MaBN := 120
			xHIS_MuaSo := 1494 , yHIS_MuaSo := 430
			xHIS_BNUT := 1618 , yHIS_BNUT := 431
			xHIS_cb5nam := 721 , yHIS_cb5nam := 320
			xHIS_SoCT := 120 , yHIS_SoCT := 345
		}
	}
	Return
}
initArrayTitle()
{
	AR := {}
	Loop,
	{
		row := A_Index
		FileReadLine, OutputVar, %path_ModuleTitle%, % row
		If ErrorLevel
			Break
		tmpAR := []
		Loop, Parse, OutputVar, CSV
		{
			if (A_Index = 1)
				str1 := A_LoopField
			if (A_Index = 2)
				str2 := A_LoopField
			Else
				str3 := A_LoopField
		}
		AR.Push({1:str1, 2:str2, 3:str3})
	}
	Return, AR
}
gotologfile:
	Run, %A_ScriptDir%\log
	Return
;Mở tiếp nhận thường
OpenTNThuong:
	Msgbox,, % "Thông báo", % "Chức năng đang xây dựng"
	Return

Info:
	info =
	(LTrim Comments
	----------------------------------------------------
	Version: 20220225.1
	
	Làm được gì?
		- Tự động thao tác hành động tiếp nhận bệnh nhân vào Phòng khám
	-----------------------------------------------------
	Các thông số trên phần mềm:
		- Version sẽ hiển thị Version hiện tại của phần mềm HIS
		- Hiển thị Độ phân giải màn hình đang chạy
		- Ngày, giờ hệ thống
	)
	GuiControl, 2:, 2_edtINFO, % info
	Gui, 2:Show, w340 h500, % "Thông tin"
	Return

;Vào Danh mục phòng
Catelogy:
	Gui,5:Default
	Gui, 5:Show, w420 h375, % "Danh mục"
	;Xử lý TAB1
	GuiCOntrol, 5:Choose, 5_ddlHisType, % g_HISType
	ControlGetText, var, ComboBox1 , % "Danh mục"
	GuiControl, 5:, 5_edtTNT, % gTITLE_TN
	GuiControl, 5:, 5_edtTNCC, % gTITLE_CC
	GuiControl, 5:, 5_edtTNVP, % gTITLE_TNVP
	;Xử lý TAB2
	GuiControl, 5:Choose, 5_ddlNhom, % g_DepartGID
	Gui, 5:Submit, NoHide
	Gui, 5:ListView, 5LISTVIEW
	addIntoLV(5_ddlNhom)
	Return

	5_btnCLose:
	5GuiClose:
	5GuiEscape:
		Gui, 5:Cancel
		Gui, 1:Default
		Return

	addIntoLV(filename)
	{
		path := path_DepartFolde . "\" . filename
		tmpAR := fn_readSimpleFile(path)
		LV_Delete()
		Loop, % tmpAR.Length()
		{
			iSring := tmpAR[A_index]
			LV_Add("", iSring)
		}
		Return		
	}
	5_ddlHisType:
		ControlGetText, var, ComboBox1 , % "Danh mục"
		GuiControl, 5:, 5_edtTNT, % fn_GetTitleModule(A_ScriptDir . "\data\Title\" . var).TN
		GuiControl, 5:, 5_edtTNCC, % fn_GetTitleModule(A_ScriptDir . "\data\Title\" . var).CC
		GuiControl, 5:, 5_edtTNVP, % fn_GetTitleModule(A_ScriptDir . "\data\Title\" . var).VP
		Return

	5_ddlNhom:
		Gui, 5:Submit, NoHide
		addIntoLV(5_ddlNhom)
		Return
	
	5_btnSave:
		Gui, 5:Submit, NoHide
		If ( 5_tab = "Màn hình")
		{
			ControlGetText, var, ComboBox1 , % "Danh mục"
			idHISType := fn_indexInString(var, ConvertARRtoString(fn_GetlistFile(path_titleFolder)), "|")
			if (idHISType != g_HISType)
			{
				Iniwrite, % idHISType, % path_Config , section1, HISType
				IniRead, g_HISType, % path_Config , section1, HISType
				initTitle()
				SB_SetText("Lưu thành công")
			}
			Else
				SB_SetText("Không có giá trị thay đổi")
		}
		Else
		{
			SB_SetText("Chức năng Lưu phòng đang xây dựng")
		}
		Return

	


;GUI3 ------------------------------------------------------------------------------------
;Vào chức năng gui 3
Config:
	Gui 3:Default
	GuiControl, 3:choose, 3_ddlSoBN, % SoBNcallQMS
	GuiControl, 3:choose, 3_ddlHISType, % g_HISType
	GuiControl, 3:Choose, 3_ddlPOS, % iPOS
	GuiControl, 3:, 3_edtEmail, % g_Email
	
	Loop, % arr_PKList.Length()
	{
		i := A_Index
		tmpString := arr_PKList[i]
		If ( Lowercase(g_PKNhi) = Lowercase(bodau(tmpString)) )
			GuiControl, 3:choose, 3_ddlPKNhi, % i
	}
	Gui, 3:+ToolWindow
	Gui, 3:Show, NoActivate w590 h400, % "Cấu hình"
	ControlGetText, OutputVar, ComboBox4, % "Cấu hình"
	Gui, 1:+Disabled
	Gui, 3:Default
	WinSet, AlwaysOnTop ,On, % "Cấu hình"


	Guicontrol, 3:Choose, 3_ddlField, 1
	GuiControl, 3:, 3_edtxPLV, % fn_getXY(HIS_PLV).xPOS
	GuiControl, 3:, 3_edtyPLV, % fn_getXY(HIS_PLV).yPOS
	GuiControl, 3:, 3_edtxMaBN, % fn_getXY(HIS_MaBN).xPOS
	GuiControl, 3:, 3_edtyMaBN, % fn_getXY(HIS_MaBN).yPOS
	GuiControl, 3:, 3_edtxBHYT5nam, % fn_getXY(HIS_BHYT5nam).xPOS
	GuiControl, 3:, 3_edtyBHYT5nam, % fn_getXY(HIS_BHYT5nam).yPOS
	GuiControl, 3:, 3_edtxSoCT, % fn_getXY(HIS_SoCT).xPOS
	GuiControl, 3:, 3_edtySoCT, % fn_getXY(HIS_SoCT).yPOS
	GuiControl, 3:, 3_edtxMuaSoKB, % fn_getXY(HIS_Muaso).xPOS
	GuiControl, 3:, 3_edtyMuaSoKB, % fn_getXY(HIS_Muaso).yPOS
	GuiControl, 3:, 3_edtxBNUT, % fn_getXY(HIS_BNUT).xPOS
	GuiControl, 3:, 3_edtyBNUT, % fn_getXY(HIS_BNUT).yPOS
	GuiControl, 3:, 3_edtxPK, % fn_getXY(HIS_PK).xPOS
	GuiControl, 3:, 3_edtyPK, % fn_getXY(HIS_PK).yPOS
	GuiControl, 3:, 3_edtxVersion, % fn_getXY(HIS_Ver).xPOS
	GuiControl, 3:, 3_edtyVersion, % fn_getXY(HIS_Ver).yPOS
	GuiControl, 3:, 3_edtxEmail, % fn_getXY(HIS_Email).xPOS
	GuiControl, 3:, 3_edtyEmail, % fn_getXY(HIS_Email).yPOS
	GuiControl, 3:, 3_edtxCMND, % fn_getXY(HIS_CMND).xPOS
	GuiControl, 3:, 3_edtyCMND, % fn_getXY(HIS_CMND).yPOS
	GuiControl, 3:, 3_edtxNguoithan, % fn_getXY(HIS_Nguoithan).xPOS
	GuiControl, 3:, 3_edtyNguoithan, % fn_getXY(HIS_Nguoithan).yPOS
	Return

	3_btnClose:
	3Guiclose:
	3GuiEscape:
		Gui, 3:Cancel
		Control, Uncheck,, Button2
		SetTimer, Update, Off
		Gui, 1:-Disabled
		Gui, 1:Default
		Gui, 1:Show
		Return

	3_cbfindPos:
		Gui, 3:Submit, NoHide
		If (3_cbfindPos) {
			SetTimer, Update, On
			SB_SetText("Giữ Ctrl để Pause",1)
		}	
		Else {
			SetTimer, Update, Off
			SB_SetText("",1)
		}
		Return

	Update:
		Gui, 3:Default
		CoordMode, Mouse, Relative
		MouseGetPos, , , msWin, msCtrl
		WinGetTitle, WinTitle, ahk_id %msWin%
		If (WinTitle = "Cấu hình") Or (WinTitle = "")
			Return
		ControlGetText, ctrlName, %msCtrl%, % WinTitle
		SB_SetText(WinTitle, 1)
		SB_SetText(ctrlName, 3)
		ControlGetPos, ctrlX, ctrlY, , , %msCtrl%, % WinTitle
		ControlGetText, var, ComboBox4, % "Cấu hình"
		Switch, var
		{
			Case "Phòng làm việc":
				GuiControl, 3:, 3_edtxPLV, % ctrlX
				GuiControl, 3:, 3_edtyPLV, % ctrlY
			Case "Mã BN":
				GuiControl, 3:, 3_edtxMaBN, % ctrlX
				GuiControl, 3:, 3_edtyMabn, % ctrlY
			Case "Số CT":
				GuiControl, 3:, 3_edtxSoCT, % ctrlX
				GuiControl, 3:, 3_edtySoCT, % ctrlY
			Case "BN Ưu tiên":
				GuiControl, 3:, 3_edtxBNUT, % ctrlX
				GuiControl, 3:, 3_edtyBNUT, % ctrlY
			Case "BHYT 5 năm":
				GuiControl, 3:, 3_edtxBHYT5Nam, % ctrlX
				GuiControl, 3:, 3_edtyBHYT5Nam, % ctrlY
			Case "Mua sổ KB":
				GuiControl, 3:, 3_edtxMuaSoKB, % ctrlX
				GuiControl, 3:, 3_edtyMuaSoKB, % ctrlY
			Case "Phòng khám":
				GuiControl, 3:, 3_edtxPK, % ctrlX
				GuiControl, 3:, 3_edtyPK, % ctrlY
			Case "Version":
				GuiControl, 3:, 3_edtxVersion, % ctrlX
				GuiControl, 3:, 3_edtyVersion, % ctrlY
			Case "Email":
				GuiControl, 3:, 3_edtxEmail, % ctrlX
				GuiControl, 3:, 3_edtyEmail, % ctrlY
			Case "CMND":
				GuiControl, 3:, 3_edtxCMND, % ctrlX
				GuiControl, 3:, 3_edtyCMND, % ctrlY
			Case "Người thân":
				GuiControl, 3:, 3_edtxNguoithan, % ctrlX
				GuiControl, 3:, 3_edtyNguoithan, % ctrlY
		}
		Return

		~*Ctrl::
			If (3_cbfindPos)
				SetTimer, Update, Off
			Return

		~*Ctrl up::
			If (3_cbfindPos)
				SetTimer, Update, On
			return

	3_btnSave:
		Gui, 3:Submit, NoHide
		If (myTab = "Chung")
		{
			new_SoBNcallQMS := 3_ddlSoBN
			new_HISType := fn_indexInArray(3_ddlHISType, arr_HISType)
			new_Email := Trim(Bodau(3_edtEmail))
			GuiControl, 3:, 3_edtEmail, % new_Email
			If ( 3_ddlPOS = "Default")
				tempPOS := 1
			Else
				tempPOS := 2
			If (new_Email = "")
			{
				SB_SetText("Email không được để trống")
				ControlFocus, Edit1, % "Cấu hình"
				Return
			}
			IniWrite, % new_Email, %path_Config%, section1, Email
			IniWrite, % new_SoBNcallQMS, %path_Config%, section1, SoBN
			IniWrite, % Bodau(3_ddlPKNhi), %path_Config%, section1, PKNhi
			IniWrite, % tempPOS, %path_Config%, section1, iPOS
			SB_SetText("Lưu thành công")
			; ;Load lại config
			initDATA()
			initSOBNTN(new_SoBNcallQMS)
		}
		Else if (myTab = "Pos")
		{
			newPhongLV := 3_edtxPLV . "," . 3_edtyPLV
			newMaBN := 3_edtxMaBN . "," . 3_edtyMaBN
			newSoCT := 3_edtxSoCT . "," . 3_edtySoCT
			newBNUT := 3_edtxBNUT . "," . 3_edtyBNUT
			newBHYT5Nam := 3_edtxBHYT5Nam . "," . 3_edtyBHYT5Nam
			newMuaSoKB := 3_edtxMuaSoKB . "," . 3_edtyMuaSoKB
			newPK := 3_edtxPK . "," . 3_edtyPK
			newVersion := 3_edtxVersion . "," . 3_edtyVersion
			newEmail := 3_edtxEmail . "," . 3_edtyEmail
			newCMND := 3_edtxCMND . "," . 3_edtyCMND
			newNguoithan := 3_edtxNguoithan . "," . 3_edtyNguoithan

			;Lưu vào file Config
			IniWrite, % newPK, % path_PosConfig , Custom, HIS_PK
			IniWrite, % newPhongLV, % path_PosConfig , Custom, HIS_PLV
			IniWrite, % newMaBN, % path_PosConfig , Custom, HIS_MaBN
			IniWrite, % newMuaSoKB, % path_PosConfig , Custom, HIS_MuaSo
			IniWrite, % newBNUT, % path_PosConfig , Custom, HIS_BNUT
			IniWrite, % newBHYT5Nam, % path_PosConfig , Custom, HIS_BHYT5nam
			IniWrite, % newSoCT, % path_PosConfig , Custom, HIS_SoCT
			IniWrite, % newVersion, % path_PosConfig , Custom, HIS_Ver
			IniWrite, % newEmail, % path_PosConfig , Custom, HIS_Email
			IniWrite, % newCMND, % path_PosConfig , Custom, HIS_CMND
			IniWrite, % newNguoithan, % path_PosConfig , Custom, HIS_Nguoithan
			SB_SetText("Lưu thành công")
		}
		Else
			SB_SetText("Có lỗi khi lưu")
		Return
	

;Vào chức năng xem log
;/////////////////////////////////////////////////////////////////////////
LOG:
	Gui, 4:+ToolWindow
	Gui, 4:Show, w700 h450, % "LOG"
	
	Gui, 1:+Disabled
	Gui, 4:Default
	WinSet, AlwaysOnTop ,On, % "LOG"
	Gui, 4:Submit, NoHide
	Readlog(4_date)
	Return

	4Guiclose:
	4GuiEscape:
		;WinSet, Style, ^0x20000, % "LOG"
		Gui, 4:Cancel
		Gui, 1:-Disabled
		Gui, 1:Default
		Gui, 1:Show
		Return

	Readlog(idate)
	{
		Gui, 4:ListView, MYLIST
		LV_Delete()
		FormatTime, D1, % idate, yyyyMM
		FormatTime, D2, % idate, dd
		path := A_ScriptDir . "\log\" . D1 . "\" . D2 . ".txt"
		IfNotExist, % path
			Return
		Loop
		{
			i := A_Index
			FileReadLine, OutputVar, % path, % i
			If ErrorLevel
				Break
			Loop, Parse, % OutputVar, CSV
			{
				Switch % A_Index
				{
					Case 1:
						Col1 := A_LoopField
					Case 2:
						Col2 := A_LoopField
					Case 3:
						Col3 := A_LoopField
					Case 4:
						Col4 := A_LoopField
					Case 5:
						Col5 := A_LoopField
					Case 6:
						Col6 := fn_ThousandSeperate(A_LoopField)
					Case 7:
						Col7 := fn_ThousandSeperate(A_LoopField)
					Case 8:
						Col8 := fn_ThousandSeperate(A_LoopField)
				}
			}
			LV_Add("",i, Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8)
		}
		Return
	}

	4_btnView:
		Gui, 4:Submit, Nohide
		Readlog(4_date)
		Return
;///////////////////////////////////////////////////////////////////////////

;;
;KẾT THÚC XỬ LÝ GUI
;CLOCKTIME
RefreshTime:
	iTime := a_hour . ":" . a_min . ":" . a_sec
    SB_SetText(iTime, 3)
    Return
;Hiển thị Thứ,ngày...tháng... Năm
fn_dayVietNam()
{
	result := ""
	Switch % A_DDDD
	{
		Case "Sunday":
			result := "Chủ nhật"
		Case "Monday":
			result := "Thứ hai"
		Case "Tuesday":
			result := "Thứ ba"
		Case "Wednesday":
			result := "Thứ tư"
		Case "Thursday":
			result := "Thứ năm"
		Case "Friday":
			result := "Thứ sáu"
		Case "Saturday":
			result := "Thứ bảy"
	}
	result := result . ", " . A_DD . "/" . A_MM . "/" . A_YYYY
	Return, result
}

;Xử lý kho chọn Dropdownlist bệnh nhân
ddlBenhnhan:
	Gui, 1:Submit, NoHide
	If (ddlBenhnhan = "Trẻ em")
	{
		Control, Check,, % ctrl_Nguoithan
		Control, UnCheck,, % ctrl_Full
		Control, UnCheck,, % ctrl_CMND
		Control, UnCheck,, % ctrl_ESC
		GuiControl, Disable ,cbFull
		GuiControl, Disable ,% ctrl_CMND
		GuiControl, Disable ,% ctrl_ESC
		GuiControl, Disable ,% ctrl_Nguoithan
		Loop, % arr_PKList.Length()
		{
			i := A_Index
			tmpString := arr_PKList[i]
			If (Lowercase(bodau(tmpString)) = g_PKNhi)
				GuiControl, Choose, ddlPhongKham, % i
		}
	}
	Else
	{
		Control, UnCheck,, % ctrl_Nguoithan
		GuiControl, Enable ,cbFull
		GuiControl, Enable ,% ctrl_CMND
		GuiControl, Enable ,% ctrl_ESC
		GuiControl, Enable ,% ctrl_Nguoithan
		GuiControl, 1:Choose, ddlPhongKham, 1
	}
		
	Return
;Xử lý kho chọn Dropdownlist Đối tượng
ddlDoituong:
	Gui, 1:Submit, NoHide
	If (ddlDoituong = "Thu phí")
	{
		Control, Uncheck, ,% ctrl_Dungtuyen
		Control, Uncheck,, % ctrl_BHYT5Nam
		Control, Uncheck, , % ctrl_GiayCT
		GuiControl, disable, % ctrl_BHYT5Nam
		GuiControl, enable, % ctrl_Thutien
		Guicontrol, disable, ddlTuyen
		Control, check,, % ctrl_Thutien
	}
	Else if (ddlDoituong = "Thẻ tạm")
	{
		Control, Uncheck,, % cbGiayCT
		Control, Uncheck,, % cbBHYT5nam
		Guicontrol, disable, ddlTuyen
	}
	Else 
	{
		Guicontrol, Enable, ddlTuyen
		Control, Uncheck, ,% ctrl_Thutien
		GuiControl, enable, % ctrl_Dungtuyen
		GuiControl, enable, % ctrl_BHYT5Nam
		GuiControl, enable, % ctrl_GiayCT
		GuiControl, enable, % ctrl_GiayCT
	}
	return

ddlLoaiTN:
	Gui, 1:Submit, NoHide
	Switch ddlLoaiTN
	{
		Case "Thường":
			GuiControl, , ddlDoituong, % "|" . ConvertARRtoString(arr_doituong1)
			GuiControl, , ddlPhongkham, % "|" . ConvertARRtoString(arr_phongkham1)
			GuiControl, Disable, TNcddv
			GuiControl, Enable, Thutien
		Case "Cấp cứu":
			GuiControl, , ddlPhongkham, % "|Phòng cấp cứu"
			
	}
	GuiControl, Choose, ddlPhongkham, 1
	GuiControl, choose, ddlDoituong, 2
	Return

ddlPhongkham:
	Gui, 1:Submit, NoHide
	If (ddlDoituong="BHYT") {
		If (ddlPhongkham="Phòng Cấp cứu") 
			GuiControl, Disable, cbThutien			
		Else {
			If (depart_diff[ddlPhongkham])
				GuiControl, Enable, cbThutien
			Else {
				Control, Uncheck,, %cbThutien%
				GuiControl, Disable, cbThutien
			}
		}
	}
	Else {
		GuiControl, Enable, cbThutien
			If (ddlPhongkham="Phòng khám chung") {
				GuiControl, Enable, TNcddv
			}
			Else {
				Control, Uncheck,, Button9
				GuiControl, Disable, TNcddv
			}
	}
	Return

cbFull:
	Gui, 1:Submit, NoHide
	If (cbFull) {
		Control, Check,, % ctrl_ESC
		Control, Check,, % ctrl_CMND
		Control, Check,, % ctrl_Nguoithan
	}
	Else {
		Control, UnCheck,, % ctrl_ESC
		Control, UnCheck,, % ctrl_CMND
		Control, UnCheck,, % ctrl_Nguoithan
	}
	Return


initSOBNTN(n)
{
	ds = 1
	Loop, 10
	{
		ds .= "|"
		ds .= A_Index*n
	}
	Guicontrol,1: , ddlSoBN, |%ds%
	GuiControl, 1:choose, ddlSoBN, 2
	Return
}
;RUN/////////////////////
btnRunTN:
	Gui, 1:Submit, NoHide
	TimeGetPatient=NONE
	;Kiểm tra Màn hình tiếp nhận đã mở hay chưa?
	If (ddlLoaiTN = "Cấp cứu") {
		IfWinNotExist,	% gTITLE_CC
		{
			MsgBox,48, % "Oop!",% "Chưa mở `n" . gTITLE_CC
			Return
		}
		title_tiepnhan := gTITLE_CC
	}
	Else if (ddlLoaiTN = "Thường") {
		IfWinNotExist, % gTITLE_TN
		{
			MsgBox, 48, % "Oop!",% "Chưa mở `n" . gTITLE_TN
			Return
		}
		title_tiepnhan := gTITLE_TN
	}
	Sleep 300
	;Kiểm tra quầy tiếp nhận
	WinActivate, %title_tiepnhan%
	Sleep 100
	;Khởi tạo các control trong form thông tin BN
	ClassNN_Email := ""
	HISctrl_CNMD := ""
	ClassNN_Nguoithan := ""
	;Lấy COntrol của Tên phòng và mã BN
	iCheck := 0
	WinGet, ilist, ControlList, %title_tiepnhan%
	HISctrl_Thutien := ""
	Loop, Parse, ilist, `n`r
	{
		ControlGetText, OutputVar, %A_LoopField%, %title_tiepnhan%
		ControlGetPos, x, y, , , %A_LoopField%, %title_tiepnhan%
		If (OutputVar = "Gọi") {
			HISctrl_Goiso := A_LoopField
			iCheck++
			Continue
		}
		If (OutputVar = "Lưu") {
			HISctrl_Luu := A_LoopField
			iCheck++
			Continue
		}	
		If (OutputVar = "Nhập mới") {
			HISctrl_Nhapmoi := A_LoopField
			iCheck++
			Continue
		}
		If (OutputVar = "Nhập sinh hiệu") {
			HISctrl_Nhapsinhhieu := A_LoopField
			iCheck++
			Continue
		}
		If (OutputVar = "Thu tiền") {
			HISctrl_Thutien := A_LoopField
			iCheck++
			Continue
		}
		If (x = xHIS_PLV and y = yHIS_PLV) {
			ClassNN_tenphong := A_LoopField
			iCheck++
			Continue
		}
		If (x = xHIS_MaBN and y = yHIS_MaBN) {
			ClassNN_MaBN := A_LoopField
			iCheck++
			Continue
		}
		If ( x = xHIS_Ver ) AND ( y = yHIS_Ver ) {
			ClassNN_version := A_LoopField
			iCheck++
			Continue
		}
		If ( x = xHIS_PK ) AND ( y = yHIS_PK ) {
			ClassNN_PK := A_LoopField
			iCheck++
			Continue
		}
		If ( x = xHIS_MuaSo ) AND ( y = yHIS_MuaSo ) {
			ClassNN_MuaSKB := A_LoopField
			iCheck++
			Continue
		}
		If ( x = xHIS_BNUT ) AND ( y = yHIS_BNUT ) {
			ClassNN_BNUT := A_LoopField
			iCheck++
			Continue
		}
		If ( x = xHIS_cb5nam ) AND ( y = yHIS_cb5nam ) {
			ClassNN_cb5nam := A_LoopField
			iCheck++
			Continue
		}
		If ( x = xHIS_SoCT ) AND ( y = yHIS_SoCT ) {
			ClassNN_SoCT := A_LoopField
			iCheck++
			Continue
		}
	}
	;Kiểm tả version HIS và ver trên TOOL
	;Nếu giống nhau thì bỏ qua
	;Nếu khác nhau thì lưu ver HIS hiện tại lại
	ControlGetText, version, %ClassNN_version%, %title_tiepnhan%
	If (HisVer != version) {
		IniWrite, % version, %path_Config%, section1, version
		SB_SetText(version, 1)
	}
	ControlGetText, tenphong, %ClassNN_tenphong%, %title_tiepnhan%
	If (tenphong = "") {
		Msgbox, 16, % "Lỗi", % "Kiểm tra lại tên phòng làm việc.`nKhông lấy được tên phòng"
		Return
	}
	Switch ddlLoaiTN
	{
		Case "Thường":
			If (NOT TN_thuong[tenphong]) {
				Msgbox, 16, % "Oop!", % "Cần vào phòng tiếp nhận thường, thử lại!"
				WinActivate, % AHK_tittle
				Return
			}
		Case "Cấp cứu":
			If (NOT TN_capcuu[tenphong]) {
				Msgbox, 16, % "Oop!", % "Cần vào phòng tiếp nhận CC, thử lại!"
				WinActivate, % AHK_tittle
				Return
			}
		Case "Khoa sản":
			If (NOT TN_khoasan[tenphong]) {
				Msgbox, 16, % "Oop!", % "Cần vào phòng tiếp nhận khoa sản, thử lại!"
				WinActivate, %AHK_tittle%
				Return
			}
	}
	;Kiểm tra có button Thu tiền khi sử dụng chức năng thu tiền hay không
	If ( cbThutien ) AND ( HISctrl_Thutien = "" ) {
		Msgbox, 16, % "Lỗi", % "User không có quyền thu tiền, `nthử lại"
		WinActivate, % AHK_tittle
		Return
	}
	;Bắt đầu:
	;Thời gian bắt đầu chạy Script
	Script_start := A_TickCount
	;Chọn phòng khám, (tiếp nhận cấp cứu sẽ không chọn)
	If (title_tiepnhan = gTITLE_TN) {
		Sleep 300
		ControlClick, %ClassNN_PK%, %title_tiepnhan%
		ddl_choose(ddlPhongkham, 300)
		sleep 200
	}
	icount = 0
	SoBN := ddlSoBN
	Loop, %SoBN%
	{
		Clipboard := ""
		;Check cứ SoBNcallQMS BN sẽ click gọi số một lần
		If (ddlLoaiTN = "Thường") Or (ddlLoaiTN = "Dịch vụ") OR (ddlLoaiTN = "Khoa sản") {
			Starttime := A_TickCount
			If (mod(icount, SoBNcallQMS)=0) {
				WinActivate, %title_tiepnhan%
				ControlClick, %HISctrl_Goiso%, %title_tiepnhan%
				Sleep 200
				Loop {
					ControlGet, var, Enabled,, %HISctrl_Goiso%, %title_tiepnhan%
					If (var=1)
						Break
				}
			Endtime := A_TickCount
			time_goiso := (Endtime-Starttime-200)
			}
		}
		Sleep 300
		ControlClick, %HISctrl_Nhapmoi%, %title_tiepnhan%
		Sleep 300
		;Nhâp và tiếp nhậN BN
		FormatTime, randhoten, , HHmmss
		Send % "N " . randhoten
		;Sleep 500
		Send {tab}
		Partient_Form := "THÔNG TIN BỆNH NHÂN"
		waitform(Partient_Form)
		If (only_female[Phongkham]) ;Kiểm tra nếu phòng khám thuộc PK chỉ dành cho nữ thì gán giới tính = Nữ(iGT=1)
			iGT := 1
		Else
			Random, iGT, 1,2 ;Nếu là 1: Nam, 2:Nữ
		Partient_name := fn_CreatePersonName(iGT)
		full_name := Partient_name.hovaten ;Tạo HỌ và TÊN
		edit_choose(full_name, 300)
		;Tạo ngày sinh
		If (ddlBenhnhan = "Trẻ em") {
			If (ddlDoituong = "Thẻ tạm")
				bday := RandomDate(GETyearNyearAGO(1) ,GETyearNyearAGO(0) ,"ddMMyyyy") ;Tạo ngày sinh <1T
			Else
				bday := RandomDate(GETyearNyearAGO(6) ,GETyearNyearAGO(1) ,"ddMMyyyy") ;Tạo ngày sinh <6T
		}
		Else if (phongkham="PK Khoa sản" Or phongkham="PK Thai")
			bday := RandomDate(GETyearNyearAGO(40) ,GETyearNyearAGO(24),"ddMMyyyy") ;Tạo ngày sinh từ 24T -> 40T
		Else
			bday := RandomDate(GETyearNyearAGO(90) ,GETyearNyearAGO(6),"ddMMyyyy") ;Tạo ngày sinh từ 6T -> 100T
		edit_choose(bday, 300)
		;CHỌN GIỚI TÍNH
		If ( iGT = 1)
			edit_choose("Nam", 200)
		Else
			edit_choose("Nu", 200)
		;HÀM TÍNH TUỔI
		yearold := fn_tinhtuoi(bday)
		Job := ""
		BHYT_header := ""
		If (yearold <= 6) {
			BHYT_header := "TE1"
			Job = "Trẻ <6 tuổi đi học"	
		}
		Else If (yearold <= 18) {
			BHYT_header := "HS4"
			Job = "Sinh viên, học sinh"
		}
		Else If (yearold <= 23) {
			BHYT_header := "SV4"
			Job = "Sinh viên, học sinh"	
		}
		Else if (yearold >= 60) {
			Job = "Hưu và >60 tuổi"
		}
		If (Job <> "") {
			ControlClick, x487 y124, % Partient_Form
			edit_choose(Job, 300)
		}
		Sleep 300
		;Random tên địa chỉ
		edit_choose(fn_randChoice(arr_addressName), 300)
		;RANDOM CODE ĐỊA CHỊ
		edit_choose(fn_randChoice(arr_addcode),300)
		IfWinActive, % "Thông báo"
			Send {Enter}
		If ( cbESC != 0 ) OR ( cbEmail != 0 ) OR ( cbNguoithan != 0 )
		{
			If (ClassNN_Email = "")
			{
				WinGet, ilist, ControlList, % Partient_Form
				Loop, Parse, ilist, `n`r
				{
					ControlGetPos, x, y, , , %A_LoopField%, % Partient_Form
					If (x = xHIS_Email and y = yHIS_Email)
						ClassNN_Email := A_LoopField
					If (x = xHIS_CMND and y = yHIS_CMND)
						ClassNN_CMND := A_LoopField
					If (x = xHIS_Nguoithan and y = yHIS_Nguoithan)
						ClassNN_Nguoithan := A_LoopField
				}
			}
		}
		;Nhập thông tin Email, SĐT, công ty
		If cbESC
		{
			If ( ClassNN_Email = "")
			{
				Msgbox, 16, % "Oop!", % "Có lỗi"
				Return
			}
			ControlFocus, % ClassNN_Email, % Partient_Form
			Sleep, 300
			Email := fn_CreateEmail(Partient_name.Ho, Partient_name.Lot, Partient_name.Ten, bday)
			edit_choose(Email, 200)
			Sdt := "09" . RandomNumRange(8)
			edit_choose(Sdt, 200)
			CPN := fn_GetArrayOfCompany()
			random, r, 1, % CPN.Length()
			edit_choose(CPN[r][1], 200)
			edit_choose(CPN[r][2], 200)
			edit_choose(CPN[r][3], 200)
		}
		;Nhập thông tin CMND
		If cbCMND {
			If (yearold >= 18) {
				ControlFocus, % ClassNN_CMND, % Partient_Form
				CMND := RandomNumRange(9)
				CMND_date := RandomDate(GETyearNyearAGO(16) ,GETyearNyearAGO(18) ,"ddMMyyyy")
				CMND_location := "CA Tỉnh"
				edit_choose(CMND, 200)
				edit_choose(CMND_date, 200)
				edit_choose(CMND_location, 200)
			}
		}
		;Nhập thông tin người thân
		If cbNguoithan
		{
			ControlFocus, % ClassNN_Nguoithan, % Partient_Form
			Sleep, 100
			relativePeople := fn_CreatePersonName(2).hovaten
			edit_choose(relativePeople, 200)
			phonenumber := "09" . RandomNumRange(8)
			edit_choose(phonenumber, 200)
			edit_choose("Mẹ", 200)
			add_relativePeople := "Địa chỉ của " . relativePeople
			edit_choose(add_relativePeople, 200)
		}
		Send ^{s}
		Startime := A_TickCount
		WinWaitActive, % "Thông báo"
		Endtime := A_TickCount
		Time_saveInfoPatient := Endtime - Startime	;Tính thời gian lưu thông tin BN
		Sleep 200
		Send {Enter}
		Sleep 200
		WinWaitActive, % title_tiepnhan
		MaBN := ""
		Loop, 10
		{
			Sleep, 100
			ControlGetText, MaBN, %ClassNN_MaBN%, %title_tiepnhan%
			If (MaBN != "")
				Break
		}
		If (ddlDoituong != "Thu phí") 
		{
			Sleep, 300
			Send {F2}
			WinWaitActive, % "THÔNG TIN THẺ BHYT" ,, 2
			If ErrorLevel
			{
				Msgbox, 48, % "oop", % "Có lỗi khi vào màn hình THÔNG TIN THẺ BHYT"
				Return
			}
			Sleep 200
			WinGet, ilist, ControlList, % "THÔNG TIN THẺ BHYT"
			
			If (ddlDoituong = "Thẻ tạm") {
				Loop, Parse, ilist, `n`r
				{
					ControlGetText, OutputVar, %A_LoopField%, % "THÔNG TIN THẺ BHYT"
					If ( OutputVar = "Thẻ tạm" ) {
						btnTHETAM := A_LoopField
						ControlClick, % btnTHETAM, % "THÔNG TIN THẺ BHYT"
						Break
					}
				}
			}	
			Else 
			{
				If ( ddlDoituong = "BHYT 80%" )
					MH := 80
				Else if ( ddlDoituong = "BHYT 95%" )
					MH := 95
				Else if ( ddlDoituong = "BHYT 100%")
					MH := 100
				Else
					MH := ""
				BHYT := fn_CreateBHYTCode(BHYT_header, ddlTuyen, MH)
				bhyt_code := BHYT.bhytcode
				Bhyt_hoscode := BHYT.hoscode		
				Bhyt_from := BHYT.fromdate
				Bhyt_to := BHYT.todate
				edit_choose(bhyt_code,300)
				if (ddlTuyen = "Đúng tuyến")
					Send {tab}
				else
					ddl_choose(bhyt_hoscode,300)
				edit_choose(bhyt_from,300)
				edit_choose(bhyt_to,300)
				Send ^{s}
				WinWaitActive, % "Thông báo",, 2
				If ErrorLevel
				{
					Send {Enter}
					Sleep 300
					Send ^{q}
				}
				Sleep 300
				Send {enter}
				Sleep 300
				Send {enter}
				WinWaitActive, % "KIỂM TRA THÔNG TUYẾN"
			}
			;}
			Sleep 300
			Send ^{q}
			Waitform(title_tiepnhan)
			sleep 300
		}
		;Nhập thông tin thẻ 5 năm 6 tháng
		If ( cbBHYT5Nam ) {
			ControlClick, % ClassNN_cb5nam, % title_tiepnhan
			Send {tab}
			rd := RandomNum(1,24)
			date5nam := A_Now
			date5nam += -%rd%, Days
			FormatTime, date5nam, %date5nam%, ddMMyyyy
			edit_choose(date5nam, 200)
		}
		;Nhập giấy chuyển tuyến
		if ( ddlTuyen = "Chuyển tuyến")
		{
			ControlFocus, % ClassNN_SoCT, % title_tiepnhan
			edit_choose("CT" . RandomNumRange(5), 200)
			FormatTime, NgayCT,, ddMMyyyy
			edit_choose(NgayCT, 200)
			ddl_choose(Bhyt_hoscode, 300)
			Send {tab}{tab}{tab}
			ddl_choose("I10", 300)
		}
		;MUA SỔ KHÁM BỆNH
		If ( cbMuaSKB ) {
			ControlClick, %ClassNN_MuaSKB%, % title_tiepnhan
			Sleep, 200
		}
		;TÍCH BN ƯU TIÊN
		If ( cbBNUuTien ) {
			ControlClick, %ClassNN_BNUT%, % title_tiepnhan
			Sleep, 200
		}
		Sleep 500
		Send ^{s}
		Starttime := A_TickCount
		WinWaitActive, % "Thông báo"
		fYESctrl := ""
		;Kiểm tra xem có thông báo xác nhậnh Yes/No hay không
		WinGet, iList, ControlList, % "Thông báo"
		Loop, Parse, iList, `r`n
		{
			ControlGetText, OutputVar, %A_LoopField%, % "Thông báo"
			if ( OutputVar = "&Yes") {
				fYESctrl := A_LoopField
				Break
			}
		}
		If (fYESctrl != "") {
			Send {Y}
			fOKctrl := ""
			Loop
			{
				WinGet, iList, ControlList, % "Thông báo"
				Loop, Parse, iList, `r`n
				{
					ControlGetText, OutputVar, %A_LoopField%, % "Thông báo"
					if ( OutputVar = "&OK") {
						fOKctrl := A_LoopField
						Break
					}
				}
				If (fOKctrl != "") {
					Send {Enter}
					Break
				}
			}
		}
		Else
		{
			WinActivate, % "Thông báo"
			Sleep, 200
			Send {enter}
		}
		Sleep, 200
		Loop {
			Controlget, var, Enabled,, %HISctrl_Luu%, %title_tiepnhan%
			If (var=1)
				Break
		}
		Endtime := A_TickCount
		TimeSaveReception := (Endtime-Starttime)
		iMATN := Clipboard
		WinActivate, %title_tiepnhan%
		
		icount++
		FormatTime, iNow,, dd/MM/yyyy HH:mm:ss
		;Ghi log tiếp nhận
		tiepnhan_log := iNow
						. "," . iMATN
						. "," . MaBN
						. "," . full_name
						. "," . ddlPhongkham
						. "," . time_goiso
						. "," . Time_saveInfoPatient
						. "," . TimeSaveReception
		;tiepnhan_log := StrReplace(tiepnhan_log, "|",";")
		create_log(tiepnhan_log)
		;Thu tiền
		If ( cbThutien )
		{
			ControlClick, % HISctrl_Thutien, % title_tiepnhan
			WinWaitActive, % gTITLE_TNVP
			Loop,
			{
				WinGet, ilist, ControlList, % gTITLE_TNVP
				VPbtnThutien := ""
				Loop, Parse, ilist, `n`r
				{
					ControlGetText, OutputVar, %A_LoopField%, % gTITLE_TNVP
					If (OutputVar = "Thu tiền") {
						VPbtnThutien := A_LoopField
						Break
					}
				}
				If (VPbtnThutien != "")
					Break
			}
			Loop 
			{
				Controlget, var, Enabled,, %VPbtnThutien%, %gTITLE_TNVP%
				If (var=1)
					Break
			}
			Sleep 300
			ControlClick, % VPbtnThutien, % gTITLE_TNVP
			WinWaitActive, % "THU TIỀN"
			Send ^{s}
			WinWaitActive, % "Thông báo"
			Send {Enter}
			Sleep 200
			Send ^{t}
			WinWaitActive, % title_tiepnhan
		}

		;Nhập dấu sinh tồn
		If cbDST
		{
			ControlClick, % HISctrl_Nhapsinhhieu, % title_tiepnhan
			WinWaitActive, % "SINH HIỆU"
			edit_choose(RandomNum(155,170), 200)
			edit_choose(RandomNum(50,70), 200)
			edit_choose(RandomNum(90,130), 200)
			edit_choose(RandomNum(60,90), 200)
			edit_choose(RandomNum(70,150), 200)
			edit_choose(RandomNum(35.0,37.0), 200)
			edit_choose(RandomNum(95,100), 200)
			edit_choose(RandomNum(15,30), 200)
			Send {Enter}
			WinWaitActive, % "Thông báo"
			Send {Enter}
		}
		;END LOOP
	}
	Script_end := A_TickCount
	TotalTime := ConvertMilisec(Script_end-Script_start)
	TotalTime_H := TotalTime.hour
	TotalTime_M := TotalTime.min
	TotalTime_S := TotalTime.sec
	MsgTime := ""
	If (TotalTime_H <> 0) {
		MsgTime .= TotalTime_H . " giờ "
	}	
	If (TotalTime_M <> 0)
		MsgTime .= TotalTime_M . " phút "
	MsgTime .= TotalTime_S . " giây "
	MsgBox, 64, % "WoW", % "Chạy Script thành công!`nSố BN: " . SoBN . "`nThời gian: " . MsgTime
	WinActivate, % AHK_tittle
	Return

;FUNCTION
;BASIC FUNCTION
;Random Số
RandomNum(f_num, t_num)
{
	Random, r, %f_num%, %t_num%
	Return % r
}
;Random dãy số
RandomNumRange(len, i = 48, x = 57)  ; length, lowest and highest Asc value
{
	Loop, % len
	{
		Random, r, i, x
		s .= Chr(r)
	}
	Return, s
}
;Phân tách số hàng ngàn
fn_ThousandSeperate(k)
{
	If (k < 1000)
		Return k . " ms"
	Else
		Return, Format("{:d}", (k - mod(k,1000))/1000) . "," . mod(k,1000) . " ms"
}
;ConvertARRtoString
ConvertARRtoString(arr)
{
	i := arr.Length()
	Loop, % i
	{
		_str .= arr[A_Index]
		If (A_Index <> i)
			_str .= "|"
	}
	Return, _str
}
;Kiểm tra giá trị tồn tại trong mảng
fn_isInArray(string, array)
{
	L := array.Length()
	Loop, % L
	{
		i := A_index
		If (array[i] = string)
			Return, True
	}
	Return, False
}
;Trả về index của giá trị trong mảng
fn_indexInArray(string, array)
{
	L := array.Length()
	Loop, % L
	{
		i := A_index
		If (array[i] = string)
			Return, i
	}
	Return, 1
}
;Trả về index của giá trị trong string
fn_indexInString(needle, string, seperate)
{
	Loop, Parse, string, % "`" . seperate
	{
		If (needle = A_LoopField)
			Return, A_Index
	}
	Return, 0

}
;Hàm get tất cả file trong folder
fn_GetlistFile(path)
	{
		tmpAR := []
		path := path . "\*.*"
		tmpString := ""
		Loop, % path
		{
			tmpAR.push(A_LoopFileName)
		}
		Return, tmpAR
	}

;Specical Function
;Tách x,y từ dữ liệu (XY=123,36 =>x=123 y=36)
fn_getXY(XY)
{
	Loop, parse, XY, `,
	{
		If (A_index = 1)
			Str1 := A_LoopField
		Else if (A_index = 2)
			Str2 := A_LoopField
		Else
			Continue
	}
	Return {xPOS:Str1, yPOS:Str2}
}

;Sử dụng cho các TH chọn dữ liệu dạng DropDownList
ddl_choose(mydata, ctime)      
{
	Send {Text}%mydata%
    Sleep %ctime%
    Send {down}
    Sleep %ctime%
    Send {enter}
    Sleep %ctime%
    Send {tab}
    sleep %ctime%
}
;Sử dụng cho các TH nhập sử liệu vào ô Edit
edit_choose(mydata, ctime)				
{
	If (mydata <> 0)
		Send {Text}%mydata%
	Sleep %ctime%
	Send {tab}
	Sleep %ctime%
}
;Hàm wait form
Waitform(form_name) {
	WinWaitActive, % form_name,, % g_nWait
    If ErrorLevel
    {
        Msgbox, 16, % "Oop!", % "Không vào được form " . form_name
		Return
        exit
    }
}
;RANDOM DATE FROM-TO
RandomDate(startDate,endDate,Format) 
{
	startDate := RegExReplace(startDate,"/"), max := endDate :=	RegExReplace(endDate,"/")
	max -= startDate, days
	Random, days, 1, %max%
	startDate += days, days
	FormatTime, newDate, %startDate%, %Format%
	return	newDate
}
GETyearNyearAGO(N) 
{
	n_day := A_DD
	n_month := A_Mon
	n_year := A_Year-N
	idate := n_year . "/" . n_month . "/" . n_day
	Return % idate
}
ConvertMilisec(mil)
{
    If (mil<1000)
        sec = 0
    Else
        sec := SubStr(mil, 1, Strlen(mil)-3)
    min := Floor(sec/60)
    sec := sec-(min*60)
    hour := Floor(min/60)
    min := min-(hour*60)
	If (hour<10)
		hour := "0" . hour
	If (min<10)
		min := "0" . min
	If (sec<10)
		sec := "0" . sec
    Return, {hour:hour, min:min, sec:sec, Mil:mil}
}
;Tạo thông tin HỌ va TÊN
fn_CreatePersonName(Gend)
{
	;Data
	firstname_arr := ["Nguyễn","Bùi","Nguyễn","Châu","Đặng","Nguyễn","Đinh","Đỗ","Đoàn","Dương","Hà","Hồ","Hứa","Huỳnh","Lê","Lý","Mạc","Mai","Ngô","Nguyễn","Phạm","Phan","Quách","Tăng","Thạch","Thái","Tô","Tôn","Trần","Triệu","Trịnh","Trương","Võ","Vương"]
	midname_male_arr := ["Ngọc", "Minh", "Bảo","Văn","Gia","Hoàng","Thiên","Khánh","Thái","Tuấn"]
	midname_female_arr := ["Hồng","Thị","Thị","Thị","Thị","Thị","Thị"]
	lastname_male_arr := ["An","Anh","Bảo","Bình","Biên","Công","Chung","Cường","Danh","Du","Duy","Dũng","Dương","Đường","Đạt","Được","Đăng","Định","Đức","Hoài","Hoàng","Hải","Hùng","Huy","Hậu","Huấn","Hưng","Kiên","Linh","Lương","Lăng","Ly","Long","Mạnh","Minh","Mẫn","Nam","Năm","Nghị","Phúc","Phước","Phong","Phi","Quyền","Quãng","Tư","Tứ","Tuấn","Tùng","Tấn","Tiến","Toàn","Thịnh","Thông","Thương","Tài","Thắng","Thanh","Vũ","Vy","Văn"]
	lastname_female_arr := ["Ánh","Bống","Chi","Chung","Châu","Dung","Dương","Duyên","Hằng","Hoài","Hương","Hai","Hạnh","Hồng","Hoa","Kiều","Linh","Lan","Ly","Liễu","Mai","Mơ","Nhung","Nhi","Như","Nga","Ngân","Phượng","Phương","Quyên ","Tình","Tư","Thương","Thảo","Thơ","Thơm","Thi","Thúy","Thủy","Thanh","Uyên","Yến","Vân Anh","Bảo Anh","Kiều Anh","Ngọc Anh","Trâm Anh","Trà Long","Trà My","Gia Mỹ","Kiều Tiên","Thúy Kiều ","Thúy Vân","Vân Kiều"]
	;--------------------------------------
    F_name := fn_randChoice(firstname_arr)
    If (Gend = 1) {
        M_name := fn_randChoice(midname_male_arr)
        L_Name := fn_randChoice(lastname_male_arr)
    }
    Else {
        M_name := fn_randChoice(midname_female_arr)
        L_Name := fn_randChoice(lastname_female_arr)
    }
    Full_name := F_name . " " . M_name . " " . L_Name
    Return, {HovaTen:Full_name, Ho:F_name, Lot:M_name, Ten:L_Name}
}
; Trả về 1 giá trị ngẫu nhiên trong mảng
fn_randChoice(arr)
{
	arr_len := arr.Length()
	Random, r, 1, % arr_len
	return, arr[r]
}
;Hàm tính tuổi theo ngày sinh
fn_tinhtuoi(birthday)
{
	byear := SubStr(birthday, 5, 4)
	year_now := A_Year
	yearold := year_now-byear
	Return yearold
}
;Hàm tạo thẻ BHYT
fn_CreateBHYTCode(BH_header, tuyen, MH)
{
	ar_BHYT := ["DN4","HX4","CH4","NN4","TK4","HC4","XK4","TB4","NO4","CT2","XB4","TN4","CS4","XN4","MS4","HD4","TQ4","TA4","TY4","HG4","LS4","PV4","GB4","GD4","HT3","TC3","CN3","CC1","CK2","CB2","KC2","HN2","DT2","DK2","XD2","BT2","TS2","QN5","CA5","CY5"]
	code := ["79011","75011","79012","79013","79014","79015","79016"]
	If (tuyen = "Đúng tuyến")
	{
		maBV := g_MaBV
		maTinh := SubStr(maBV, 1, 2)
	}
	Else
	{
		maBV := fn_randChoice(code)
		maTinh := SubStr(maBV, 1, 2)
	}
	If (BH_header = "")
		bhyt := fn_randChoice(ar_BHYT) . maTinh . RandomNumRange(10)
	Else
		bhyt := BH_header . maTinh . RandomNumRange(10)
	FromDate := RandomDate(GETyearNyearAGO(1) ,GETyearNyearAGO(0) ,"ddMMyyyy")
	ToDate := SubStr(FromDate, 1, 4) . SubStr(FromDate, 5, 4)+1
    Return, {bhytcode:bhyt, hoscode:maBV, fromdate:Fromdate, todate:todate}
}
;Hàmm tạo địa chỉ
fn_CreateAddress()
{
	xpath = % A_ScriptDir . "\data\Address"
	IfNotExist, % xpath
	{
		MsgBox, 16, % "Lỗi Script", % "File không tồn tại"
		Return
	}
    Loop, read, %xpath%
    {
    	max = % A_Index
    }
    random, i, 1, %max% 
    FileReadLine, iline, %xpath%, %i%
    If ErrorLevel
    	MsgBox, 16, % "Lỗi Script", % "Không tìm được file"
	Random, iNum, 10, 1500
	Random, r1, 1, 2
	If (r1 = 1)
	{
		Random, R2, 1, 20
		iNum := iNum . "/" . R2
	}
    address := "Số " . iNum ", " . iline
    return address
}
;Hàm đọc SIMPLE file (File có dữ liệu trên mỗi dòng)
fn_readSimpleFile(path)
{
	tmpAR := []
	Loop,
	{
		i := A_Index
		FileReadLine, OutputVar, % path, % i
		If ErrorLevel
			Break
		tmpAR.Push(OutputVar)
	}
	Return, tmpAR
}
;Gán tiêu đề các màn hình làm việc
fn_GetTitleModule(file)
{
	Loop,
	{
		i := A_Index
		FileReadLine, var, % file, % i
		If ErrorLevel
			Break
		if ( i = 1)
			str1 := var
		else if ( i = 2)
			str2 := var
		else if ( i =3 )
			str3 := var
	}
	Return {TN:str1, CC:str2, VP:str3}
}
;Hàm tạo Email theo Họ và tên
fn_CreateEmail(Ho,Lot,Ten,bday)
{
	email := ""
	StringLower, Ho, Ho
	StringLower, Lot, Lot
	email := Bodau(Ten) . SubStr(Bodau(Ho), 1, 1) . SubStr(Bodau(Lot), 1, 1) . SubStr(bday, -1) . g_Email
	Return, email
}
;Lấy thông tin Cty từ file chuyển thanh array
fn_GetArrayOfCompany()
{
	tmpARR := []
	Loop,
	{
		i := A_Index
		FileReadLine, OutputVar, %path_Company%, % i
		If ErrorLevel
			Break
		Loop, Parse, OutputVar, CSV
		{
			If (A_Index = 1)
				name := A_LoopField
			If (A_Index = 2)
				MST := A_LoopField
			If (A_Index = 3)
				addr := A_LoopField
		}
		tmpARR.Push({1:name,2:MST,3:addr})
	}
	Return, tmpARR
}
;Hàm Lowercase
Lowercase(string)
{
	StringLower, OutputVar, string
	Return, % OutputVar
}
;Hàm bỏ dấu
Bodau(myString)
{    
    N_String := ""    
    Loop, % StrLen(myString)    
    {        
        temp := SubStr(myString, A_Index, 1)        
        Switch Asc(temp)       
        {            
            Case 273:                
                N_String .= "d"            
            Case 272:                
                N_String .= "D"            
            Case 224, 225, 226, 227, 259, 7841, 7843, 7845, 7847, 7849, 7851, 7853, 7855, 7857, 7859, 7861, 7863:                
                N_String .= "a"            
            Case 192, 193, 194, 195, 258, 7840, 7842, 7844, 7846, 7848, 7850, 7852, 7854, 7856, 7858, 7860, 7862:                
                N_String .= "A"            
            Case 232, 233, 234, 7865, 7867, 7869, 7871, 7873, 7875, 7877, 7879:                
                N_String .= "e"            
            Case 200, 201, 202, 7864, 7866, 7868, 7870, 7872, 7874, 7876, 7878:                
                N_String .= "E"            
            Case 236, 237, 297, 7881, 7883:                
                N_String .= "i"            
            Case 204, 205, 296, 7880, 7882:                
                N_String .= "I"            
            Case 242, 243, 244, 245, 417, 7885, 7887, 7889, 7891, 7893, 7895, 7897, 7899, 7901, 7903, 7905, 7907:                
                N_String .= "o"            
            Case 210, 211, 212, 213, 416, 7884, 7886, 7888, 7890, 7892, 7894, 7896, 7898, 7900, 7902, 7904, 7906:                
                N_String .= "O"            
            Case 249, 250, 361, 432, 7909, 7911, 7913, 7915, 7917, 7919, 7921:                
                N_String .= "u"            
            Case 217, 218, 360, 431, 7908, 7910, 7912, 7914, 7916, 7918, 7920:                
                N_String .= "U"            
            Case 253, 7923, 7925, 7927, 7929:                
                N_String .= "y"            
            Case 221, 7922, 7924, 7926, 7928:                
                N_String .= "Y"
            Case 32:
                N_String .= " "         
            Default:                
                N_String .= temp        
        }    
    }    
    Return, N_String
}
;Tạo file log
create_log(string)
{
	string := string . "`n"
	FormatTime, folder,, yyyyMM
	FormatTime, filename, , dd
	Filepath := A_ScriptDir . "\log\" . folder 
	FileCreateDir, %Filepath%
	Filepath := A_ScriptDir . "\log\" . folder . "\" . filename . ".txt"
	FileAppend, %string%, %Filepath%
}