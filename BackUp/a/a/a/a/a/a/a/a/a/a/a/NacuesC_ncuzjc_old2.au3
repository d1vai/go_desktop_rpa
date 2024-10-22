#region ;**** 参数创建于 ACNWrapper_GUI ****
#AutoIt3Wrapper_icon=apple.ico
#AutoIt3Wrapper_outfile=NacuesC_ncuzjc.exe
#AutoIt3Wrapper_Run_Tidy=y
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/cs=0
#endregion ;**** 参数创建于 ACNWrapper_GUI ****
#include <GUIConstants.au3>
#include <Constants.au3>
#include <GuiToolbar.au3>
#include <GUIStatusBar.au3>
#include <File.au3>
#include <Array.au3>
#include <Misc.au3>
#include <excel.au3>
#include <WinAPIEx.au3>
#include <GuiMenu.au3>

Opt("WinTitleMatchMode", 4)
Opt("WinWaitDelay", 10) ;10 milliseconds
$oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")




Global $Cilp_last_username
Global $Cilp_last_password
Global $Cilp_last_sf
Global $Cilp_last_college = ""
Global $exe_dir, $exe_name, $exe_path
Global $a_kszt
Global $Cilp_tdd_file_last = ""
Global $Cilp_output_dir_last = ""
Global $G_ss_py_dm[34][3] = [["北京市", "Beijing", "11"],["天津市", "Tianjin", "12"],["河北省", "Hebei", "13"],["山西省", "Shanxi14", "14"],["内蒙古自治区", "Neimenggu", "15"],["辽宁省", "Liaoning", "21"],["吉林省", "Jilin", "22"],["黑龙江省", "Heilongjiang", "23"],["上海市", "Shanghai", "31"],["江苏省", "Jiangsu", "32"],["浙江省", "Zhejiang", "33"],["安徽省", "Anhui", "34"],["福建省", "Fujian", "35"],["江西省", "Jiangxi", "36"],["山东省", "Shandong", "37"],["河南省", "Henan", "41"],["湖北省", "Hubei", "42"],["湖南省", "Hunan", "43"],["广东省", "Guangdong", "44"],["广西壮族自治区", "Guangxi", "45"],["海南省", "Hainan", "46"],["重庆市", "Chongqing", "50"],["四川省", "Sichuan", "51"],["贵州省", "Guizhou", "52"],["云南省", "Yunnan", "53"],["西藏自治区", "Xizang", "54"],["陕西省", "Shanxi61", "61"],["甘肃省", "Gansu", "62"],["青海省", "Qinghai", "63"],["宁夏回族自治区", "Ningxia", "64"],["新疆维吾尔自治区", "Xinjiang", "65"],["台湾省", "Taiwan", "71"],["香港特别行政区", "Hongkong", "72"],["澳门特别行政区", "Macao", "73"]]

Global Const $lnk_name = "院校子系统基本版 2012.lnk"
Global Const $data_dir = @MyDocumentsDir & "\.NacuesCStorage2012\"
Global Const $data_university_mdb_filename = "NacuesCUniv.mdb"
Global Const $data_university_share_mdb_filename = "NacuesCUnivShare.mdb"

Global Const $mdb_adoProvider = 'Microsoft.Jet.OLEDB.4.0; '


;~ kszt
;~ 0未定
;~ 1录取
;~ 2退档
;~ 3已退档
;~ 4录取中
;~ 5已录取

$a_kszt = StringSplit("未定|录取|退档|已退档|录取中|已录取", "|", 3)
;~ _ArrayDisplay($a_kszt)

$self_Process_list = ProcessList(@ScriptName)
If $self_Process_list[0][0] > 1 Then
	MsgBox(0, "", "程序运行中...", 0.5)
	Exit
EndIf



;HotKeySet("{ScrollLock}", "doNacuesC")
;HotKeySet("^{ScrollLock}", "doNacuesC_config")
HotKeySet("#`", "runExe")
HotKeySet("#1", "doNacuesC")
HotKeySet("#2", "doNacuesC")
HotKeySet("#3", "doNacuesC")
HotKeySet("#4", "doNacuesC")
HotKeySet("#5", "doNacuesC")
HotKeySet("#6", "doNacuesC")
HotKeySet("#7", "doNacuesC")
HotKeySet("#8", "doNacuesC")
HotKeySet("#9", "doNacuesC")
HotKeySet("#0", "doNacuesC")
HotKeySet("#q", "dianjiEx_exe")
HotKeySet("#w", "Output_Tdd")



;;;; Body of program would go here ;;;;
;While 1
;	Sleep(100)
;WEnd

Opt("TrayMenuMode", 3) ; Default tray menu items (Script Paused/Exit) will not be shown.
$qg_str = StringSplit("迅速|强大|专业|有用|偷懒|方便|快捷|聪明|自动|牛！", "|")
$doitem = TrayCreateMenu("登录操作[&D](" & $qg_str[Random(1, UBound($qg_str, 1) - 1, 1)] & ")")
$doitem1 = TrayCreateItem("1071800 登录[&1]" & @TAB & "Win+1", $doitem)
$doitem2 = TrayCreateItem("10718aa 登录[&2]" & @TAB & "Win+2", $doitem)
$doitem3 = TrayCreateItem("10718ab 登录[&3]" & @TAB & "Win+3", $doitem)
$doitem4 = TrayCreateItem("10718ac 登录[&4]" & @TAB & "Win+4", $doitem)
$doitem5 = TrayCreateItem("10718ad 登录[&5]" & @TAB & "Win+5", $doitem)
$doitem6 = TrayCreateItem("10718ae 登录[&6]" & @TAB & "Win+6", $doitem)
$doitem7 = TrayCreateItem("10718af 登录[&7]" & @TAB & "Win+7", $doitem)
$doitem8 = TrayCreateItem("10718ag 登录[&8]" & @TAB & "Win+8", $doitem)
$doitem9 = TrayCreateItem("10718ah 登录[&9]" & @TAB & "Win+9", $doitem)
$doitem0 = TrayCreateItem("备用 登录[&0]" & @TAB & "Win+0", $doitem)
TrayCreateItem("")
#CS
	生成最新投档单 Win+W
	生成历史投档单
	重新生成第1次投档单
	重新生成第2次投档单
	重新生成第3次投档单
	重新生成第4次投档单
	...第n次(输入)...
	
	当前投档单目录
	投档单根目录
	过滤考生[慎用]
	取消考生隐藏功能
	只显示最投档考生
	只显示第1次投档考生
	只显示第2次投档考生
	...第n次(输入).
	转换全部投档单
#CE
$tdd_item = TrayCreateMenu("投档单操作[&T](" & $qg_str[Random(1, UBound($qg_str, 1) - 1, 1)] & ")")
$tdd_tdd_last = TrayCreateItem("生成最新投档单 [&1]" & @TAB & "Win+W", $tdd_item)
$tdd_tdd_update = TrayCreateItem("重新输入控制线生成", $tdd_item)
$tdd_tdd = TrayCreateMenu("生成历史投档单[&2]", $tdd_item)
TrayCreateItem("", $tdd_item)
$tdd_dir = TrayCreateItem("投档单根目录[&3]", $tdd_item)
$tdd_dir_last = TrayCreateItem("最后的投档单目录[&4]", $tdd_item)
TrayCreateItem("", $tdd_item)
$out_put = TrayCreateItem("备份到自动目录[&5]", $tdd_item)
$out_put_dir = TrayCreateItem("打开最后备份目录[&6]", $tdd_item)
TrayCreateItem("", $tdd_item)
$tdd_filter = TrayCreateMenu("过滤考生(慎用)[&7]", $tdd_item)
$tdd_convert = TrayCreateItem("转换全部投档单[&8]", $tdd_item)


$tdd_tdd1 = TrayCreateItem("重新生成第1次投档单", $tdd_tdd)
$tdd_tdd2 = TrayCreateItem("重新生成第2次投档单", $tdd_tdd)
$tdd_tdd3 = TrayCreateItem("重新生成第3次投档单", $tdd_tdd)
$tdd_tdd4 = TrayCreateItem("重新生成第4次投档单", $tdd_tdd)
$tdd_tdd5 = TrayCreateItem("重新生成第5次投档单", $tdd_tdd)
$tdd_tdd6 = TrayCreateItem("重新生成第6次投档单", $tdd_tdd)
$tdd_tdd7 = TrayCreateItem("重新生成第7次投档单", $tdd_tdd)
$tdd_tdd8 = TrayCreateItem("重新生成第8次投档单", $tdd_tdd)
$tdd_tdd9 = TrayCreateItem("重新生成第9次投档单", $tdd_tdd)
TrayCreateItem("", $tdd_tdd)
$tdd_tddn = TrayCreateItem("重新生成第n次投档单...", $tdd_tdd)

TrayCreateItem("", $tdd_filter)
$tdd_filter_1 = TrayCreateItem("显示全部考生", $tdd_filter, -1, 1)
;~ TrayItemSetState(-1, $TRAY_CHECKED)
TrayItemSetState($tdd_filter_1, $TRAY_CHECKED)
$tdd_filter0 = TrayCreateItem("只显示最新投档考生", $tdd_filter, -1, 1)
$tdd_filter1 = TrayCreateItem("只显示第1次投档考生", $tdd_filter, -1, 1)
$tdd_filter2 = TrayCreateItem("只显示第2次投档考生", $tdd_filter, -1, 1)
$tdd_filter3 = TrayCreateItem("只显示第3次投档考生", $tdd_filter, -1, 1)
$tdd_filter4 = TrayCreateItem("只显示第4次投档考生", $tdd_filter, -1, 1)
$tdd_filter5 = TrayCreateItem("只显示第5次投档考生", $tdd_filter, -1, 1)
$tdd_filter6 = TrayCreateItem("只显示第6次投档考生", $tdd_filter, -1, 1)
$tdd_filter7 = TrayCreateItem("只显示第7次投档考生", $tdd_filter, -1, 1)
$tdd_filter8 = TrayCreateItem("只显示第8次投档考生", $tdd_filter, -1, 1)
$tdd_filter9 = TrayCreateItem("只显示第9次投档考生", $tdd_filter, -1, 1)
$tdd_filtern = TrayCreateItem("只显示第n次投档考生...", $tdd_filter, -1, 1)
TrayCreateItem("", $tdd_filter)



TrayCreateItem("")
$runNacuesC = TrayCreateItem("运行院校子系统&Z Win+`")
$rerunNacuesC = TrayCreateItem("重启院校子系统&R")
TrayCreateItem("")
$dianjiEx = TrayCreateItem("下载考生信息&D Win+Q")
TrayCreateItem("")
$copy_last_name = TrayCreateItem("复制用户名[&U]")
$copy_last_password = TrayCreateItem("复制密码[&P]")
TrayCreateItem("")
$aboutitem = TrayCreateItem("关于[&A](睿智)")
TrayCreateItem("")
$exititem = TrayCreateItem("退出[&X](无语)")

TraySetState()
TraySetToolTip("陕西师范大学院校子系统")
;TraySetClick(16)
;TraySetClick(64)


If Not FileExists(@ScriptDir & "\NacuesC_ncuzjc.dbf") Then
	MsgBox(0, "", "数据表NacuesC_ncuzjc.dbf不存在" & @CR & "请检查：" & @ScriptDir & "\NacuesC_ncuzjc.dbf是否存在")
	Exit
EndIf

;~ MsgBox(0, "", @DesktopDir & "\院校子系统基本版 2012.lnk" & @CRLF & FileExists(@DesktopDir & "\院校子系统基本版 2012.lnk"))
;~ 得到院校子系统所在目录
;~ @DesktopCommonDir
;~ MsgBox(0, "", @DesktopDir)

If FileExists(@DesktopDir & "\" & $lnk_name) Or FileExists(@DesktopCommonDir & "\" & $lnk_name) Then
	Local $t = FileGetShortcut(_Iif(FileExists(@DesktopDir & "\" & $lnk_name), @DesktopDir, @DesktopCommonDir) & "\" & $lnk_name)
	Local $szDrive, $szDir, $szFName, $szExt
	Local $TestPath = _PathSplit($t[0], $szDrive, $szDir, $szFName, $szExt)
;~ 	得到院校子系统所在目录
	$exe_dir = $szDrive & $szDir ;$t[1] & $t[2]
	$exe_name = $szFName & $szExt
	$exe_path = $t[0]
;~ 	MsgBox(0, "", $exe_name)
Else
	MsgBox(0, "", "桌面上不存在 " & $lnk_name & " 的快捷方式" & @CRLF & "插件的大部分功能可能无法正常使用。")
;~ 	Return
	Exit
EndIf

aboutme(2)
runExe()

FileCreateShortcut(@ScriptFullPath, @DesktopDir & "\陕西师范大学院校子系统.lnk")

While 1
	$msg = TrayGetMsg()
	Select
		Case $msg = 0
			ContinueLoop
		Case $msg = $aboutitem
			aboutme()
		Case $msg = $doitem1
			doNacuesCx("1071800")
		Case $msg = $doitem2
			doNacuesCx("10718aa")
		Case $msg = $doitem3
			doNacuesCx("10718ab")
		Case $msg = $doitem4
			doNacuesCx("10718ac")
		Case $msg = $doitem5
			doNacuesCx("10718ad")
		Case $msg = $doitem6
			doNacuesCx("10718ae")
		Case $msg = $doitem7
			doNacuesCx("10718af")
		Case $msg = $doitem8
			doNacuesCx("10718ag")
		Case $msg = $doitem9
			doNacuesCx("10718ah")
		Case $msg = $doitem0
			doNacuesCx("0")

		Case $msg = $tdd_tdd_last
			Output_Tdd()
		Case $msg = $tdd_tdd_update
			Output_Tdd(0, True)
		Case $msg = $tdd_tdd1
			Output_Tdd(1)
		Case $msg = $tdd_tdd2
			Output_Tdd(2)
		Case $msg = $tdd_tdd3
			Output_Tdd(3)
		Case $msg = $tdd_tdd4
			Output_Tdd(4)
		Case $msg = $tdd_tdd5
			Output_Tdd(5)
		Case $msg = $tdd_tdd6
			Output_Tdd(6)
		Case $msg = $tdd_tdd7
			Output_Tdd(7)
		Case $msg = $tdd_tdd8
			Output_Tdd(8)
		Case $msg = $tdd_tdd9
			Output_Tdd(9)
		Case $msg = $tdd_tddn
			$i = Number(InputBox("请输入", "需要导出哪次投档单:", "10", "", 100, 50, (@DesktopWidth - 100) / 2, (@DesktopHeight - 50) / 2, 30))
			If $i > 0 Then Output_Tdd($i)
		Case $msg = $tdd_filter_1
			HideKs(-1)
		Case $msg = $tdd_filter0
			HideKs()
		Case $msg = $tdd_filter1
			HideKs(1)
		Case $msg = $tdd_filter2
			HideKs(2)
		Case $msg = $tdd_filter3
			HideKs(3)
		Case $msg = $tdd_filter4
			HideKs(4)
		Case $msg = $tdd_filter5
			HideKs(5)
		Case $msg = $tdd_filter6
			HideKs(6)
		Case $msg = $tdd_filter7
			HideKs(7)
		Case $msg = $tdd_filter8
			HideKs(8)
		Case $msg = $tdd_filter9
			HideKs(9)
		Case $msg = $tdd_filtern
			$i = Number(InputBox("请输入", "只显示哪次投档单考生:", "10", "", 100, 50, (@DesktopWidth - 100) / 2, (@DesktopHeight - 50) / 2, 30))
			If $i > 0 Then HideKs($i)
		Case $msg = $tdd_dir
;~ 				!explorer /e, "&xy_top"
			If Not FileExists(@ScriptDir & "\tdd投档单") Then DirCreate(@ScriptDir & "\tdd投档单")
			Run(@ComSpec & " /c start explorer /e,/select, """ & @ScriptDir & "\tdd投档单""", "", @SW_HIDE)
;~ 				Run(@SystemDir&"\explorer.exe /select, """&@ScriptDir&"\tdd投档单""")
;~ 				Sleep(300)
;~ 				Send("{TAB}")
		Case $msg = $tdd_dir_last
;~ 				!explorer /e, "&xy_top"
			If $Cilp_tdd_file_last <> "" And FileExists($Cilp_tdd_file_last) Then Run(@ComSpec & " /c start explorer /e,/select, """ & $Cilp_tdd_file_last & """", "", @SW_HIDE)
;~ 				Run(@SystemDir&"\explorer.exe /select, """&@ScriptDir&"\tdd投档单""")
;~ 				Sleep(300)
;~ 				Send("{TAB}")
		Case $msg = $out_put
			output_data()
		Case $msg = $out_put_dir
			If $Cilp_output_dir_last <> "" And FileExists($Cilp_output_dir_last) Then Run(@ComSpec & " /c start explorer /e, """ & $Cilp_output_dir_last & """", "", @SW_HIDE)
		Case $msg = $tdd_convert
			converDb()
		Case $msg = $runNacuesC
			runExe()
		Case $msg = $rerunNacuesC
			If MsgBox(0x101, "重启进程警告", "真的要重启吗?可能造成当前操作没有保存...", 3) = 1 Then
				;MsgBox(0,"","呵呵,重启")
				ProcessClose($exe_name)
				ProcessWaitClose($exe_name)
				runExe()
			EndIf
		Case $msg = $dianjiEx
			dianjiEx_exe()
		Case $msg = $copy_last_name
			copyToClip($Cilp_last_username)
		Case $msg = $copy_last_password
			copyToClip($Cilp_last_password)
		Case $msg = $exititem
			ExitLoop
	EndSelect
WEnd


Func copyToClip($txt)
	If $txt <> "" Then
		ClipPut($txt)
		MsgBox(0, "", "已经将 " & $txt & " 复制到粘贴板", 0.5)
	EndIf
EndFunc   ;==>copyToClip

Func runExe()
	If ProcessExists($exe_name) Then
		;MsgBox(0,"","院校子系统基本版 已经运行",0.8)
		If Not WinActivate("登录系统", "选择省、自治区或直辖市") Then
			WinActivate("全国普通高校招生网上录取 - 院校子系统")
			
			
			
		Else
;~ 	自动填写省份
			If $Cilp_last_sf <> "" And ControlGetText("登录系统", "选择省、自治区或直辖市", "TComboBox1") = "" Then
				ControlCommand("登录系统", "选择省、自治区或直辖市", "TComboBox1", "SelectString", $Cilp_last_sf)
			EndIf
		EndIf
	Else
		If FileExists($exe_path) Then
			;FileGetShortcut
;~ 			Run(@ComSpec & " /c start " & FileGetShortName(@DesktopDir & "\院校子系统基本版.lnk"), "", @SW_HIDE)

			ShellExecute($exe_path)
			
			
			;提交关闭启动信息窗口。。
			If WinWait("启动院校子系统", "", 5) Then
;~ 				WinSetState("[CLASS:TSplashForm; INSTANCE:0; title:启动院校子系统]","",@SW_HIDE)
				WinSetState("启动院校子系统", "", @SW_HIDE)
				WinActivate("登录系统", "选择省、自治区或直辖市")
			EndIf
			
			If WinWaitActive("登录系统", "选择省、自治区或直辖市", 5) Then
;~ 	自动填写省份
				If $Cilp_last_sf <> "" And ControlGetText("登录系统", "选择省、自治区或直辖市", "TComboBox1") = "" Then
					ControlCommand("登录系统", "选择省、自治区或直辖市", "TComboBox1", "SelectString", $Cilp_last_sf)
				EndIf
			EndIf
		Else
			MsgBox(0, "", "不存在院校子系统" & $exe_path & " 程序", 0.3)
			Return
		EndIf
	EndIf
EndFunc   ;==>runExe


;;;;;;;;
Func aboutme($o_time = 0)
	#cs
		GUICreate("招生院校子系统插件 - 陕西师范大学招办")
		;~ 		GUICtrlCreateLabel("招生院校子系统插件-陕西师范大学招办", "1.Win+1：输入 1071800 的登录信息。" & 	@CRLF & _
		"2.Win+2：输入 科院 的登录信息。" & @CRLF & _
		"3.Win+3：输入 共青 的登录信息。" & @CRLF & _
		"4.Win+4：输入 抚医 的登录信息。" & @CRLF & _
		"5.Win+5：输入 软院 的登录信息。" & @CRLF & _
		"5.Win+5：输入 软院 的登录信息。" & @CRLF & _
		@CRLF & _
		@TAB &@TAB&@TAB &"QQ:70197918",30,30)
		
		Sleep(200)
		Do
		Sleep(200)
		Until 1=WinActivate("招生院校子系统插件 - 陕西师范大学招办")
	#ce
	MsgBox(0, "招生院校子系统插件-陕西师范大学招办", "1.Win+1：输入 1071800 的登录信息。" & @CRLF & _
			"2.Win+2：输入 107180aa 的登录信息。" & @CRLF & _
			"(Win键在Ctrl、Alt键之间,`键在Tab键上方)" & @CRLF & _
			@CRLF & _
			@TAB & @TAB & @TAB & "QQ:70197918", $o_time)
EndFunc   ;==>aboutme


Func doNacuesC()
	Select
		Case @HotKeyPressed = "#2"
			$user_class = "10718aa"
		Case @HotKeyPressed = "#3"
			$user_class = "10718ab"
		Case @HotKeyPressed = "#4"
			$user_class = "10718ac"
		Case @HotKeyPressed = "#5"
			$user_class = "10718ad"
		Case @HotKeyPressed = "#6"
			$user_class = "10718ae"
		Case @HotKeyPressed = "#7"
			$user_class = "10718af"
		Case @HotKeyPressed = "#8"
			$user_class = "10718ag"
		Case @HotKeyPressed = "#9"
			$user_class = "10718ah"
		Case @HotKeyPressed = "#0"
			$user_class = "0"
		Case Else ;@HotKeyPressed = "#1"
			$user_class = "1071800"
	EndSelect
	doNacuesC_exe($user_class)
EndFunc   ;==>doNacuesC
Func doNacuesCx($user_class)
	WinActivate("登录系统", "选择省、自治区或直辖市")
	doNacuesC_exe($user_class)
EndFunc   ;==>doNacuesCx

Func doNacuesC_exe($user_class)
	Local $sf, $ip, $port
	Local $conn, $rs, $sql_where, $sql
	;If WinExists("登录系统", "选择省、自治区或直辖市") Then
	;	WinActivate("登录系统", "选择省、自治区或直辖市")
	;如果当前活动窗口不是院校子系统,则退出
	If Not WinActive("登录系统", "选择省、自治区或直辖市") Then Return

	If $Cilp_last_sf <> "" And ControlGetText("登录系统", "选择省、自治区或直辖市", "TComboBox1") = "" Then
		ControlCommand("登录系统", "选择省、自治区或直辖市", "TComboBox1", "SelectString", $Cilp_last_sf)
	EndIf

	;$whd=WinGetHandle("登录系统","选择省、自治区或直辖市")
	$sf = StringRegExpReplace(ControlGetText("登录系统", "选择省、自治区或直辖市", "TComboBox1"), "(^\s*)|(\s*$)", "")
	$ip = ControlGetText("登录系统", "选择省、自治区或直辖市", "TEdit4")
	$port = ControlGetText("登录系统", "选择省、自治区或直辖市", "TEdit3")
	;		MsgBox(4096, "", @ScriptDir & $sf & "," & $ip & ":" & $port)
	If $sf = "" Then
		MsgBox(0, "错误提示", "请先选择好省份", 0.7)
		WinActivate("登录系统", "选择省、自治区或直辖市")
		Return
	EndIf


	If Not FileExists(@ScriptDir & "\NacuesC_ncuzjc.dbf") Then
		MsgBox(0, "", "数据表NacuesC_ncuzjc.dbf不存在" & @CR & "请检查：" & @ScriptDir & "\NacuesC_ncuzjc.dbf是否存在")
		Return
	EndIf

	$conn = ObjCreate("ADODB.Connection")
	$rs = ObjCreate("ADODB.Recordset")
	$sql_where = "sf='" & StringReplace($sf, "'", "") & "' AND (user_class='" & $user_class & "' OR user_class=' ' ) AND is_default='1'"
	$sql = "SELECT * from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\NacuesC_ncuzjc.dbf"), ".*\\", "") & "] WHERE " & $sql_where & " ORDER BY user_class desc"
	;$conn.Open("Driver=Microsoft Visual FoxPro Driver;SourceType=DBF;SourceDB=" & @ScriptDir)
	$conn.Open("Provider=Microsoft.Jet.OLEDB.4.0;dBase IV;HDR=NO;IMEX=2;DATABASE=" & @ScriptDir)
	$rs.ActiveConnection = $conn
	$rs.Open($sql);,1,1
	If @error Then
		MsgBox(0, "", "无法打开数据表NacuesC_ncuzjc.dbf" & @CRLF & "可能需要安装BDE驱动。")
		Return
	EndIf
	Dim $A_data[1][8]
	Dim $A_data_i = -1
	If Not ($rs.eof Or $rs.bof) Then
		$A_data_i = 0
		$A_data[$A_data_i][0] = StringReplace($rs.Fields("sf").Value, " ", "")
		$A_data[$A_data_i][1] = StringReplace($rs.Fields("user_class").Value, " ", "")
		$A_data[$A_data_i][2] = StringReplace($rs.Fields("ip").Value, " ", "")
		$A_data[$A_data_i][3] = StringReplace($rs.Fields("port").Value, " ", "")
		$A_data[$A_data_i][4] = StringReplace($rs.Fields("user_name").Value, " ", "")
		$A_data[$A_data_i][5] = StringReplace($rs.Fields("password").Value, " ", "")
		$A_data[$A_data_i][6] = StringReplace($rs.Fields("is_default").Value, " ", "")
		$A_data[$A_data_i][7] = StringReplace($rs.Fields("bz").Value, " ", "")
		$rs.movenext
	EndIf
	While Not ($rs.eof Or $rs.bof)
		$A_data_i = UBound($A_data)
		ReDim $A_data[$A_data_i + 1][8]
		$A_data[$A_data_i][0] = StringReplace($rs.Fields("sf").Value, " ", "")
		$A_data[$A_data_i][1] = StringReplace($rs.Fields("user_class").Value, " ", "")
		$A_data[$A_data_i][2] = StringReplace($rs.Fields("ip").Value, " ", "")
		$A_data[$A_data_i][3] = StringReplace($rs.Fields("port").Value, " ", "")
		$A_data[$A_data_i][4] = StringReplace($rs.Fields("user_name").Value, " ", "")
		$A_data[$A_data_i][5] = StringReplace($rs.Fields("password").Value, " ", "")
		$A_data[$A_data_i][6] = StringReplace($rs.Fields("is_default").Value, " ", "")
		$A_data[$A_data_i][7] = StringReplace($rs.Fields("bz").Value, " ", "")
		$rs.movenext
	WEnd
	$rs.Close
	$conn.Close
	If $A_data_i = -1 Then
		;MsgBox(0, "", "没有找到 " & $user_class & " 在 " & $sf & " 的用户名!", 0.3)
		TrayTip("", "没有找到 " & $user_class & " 在 " & $sf & " 的用户名!", 0.3, 17)
		Return
	EndIf
	$A_data_i = 0
	#cs
		IPV6正则表达式
		^([\da-fA-F]{1,4}:){6}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^::([\da-fA-F]{1,4}:){0,4}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:):([\da-fA-F]{1,4}:){0,3}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){2}:([\da-fA-F]{1,4}:){0,2}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){3}:([\da-fA-F]{1,4}:){0,1}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){4}:((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){7}[\da-fA-F]{1,4}$|^:((:[\da-fA-F]{1,4}){1,6}|:)$|^[\da-fA-F]{1,4}:((:[\da-fA-F]{1,4}){1,5}|:)$|^([\da-fA-F]{1,4}:){2}((:[\da-fA-F]{1,4}){1,4}|:)$|^([\da-fA-F]{1,4}:){3}((:[\da-fA-F]{1,4}){1,3}|:)$|^([\da-fA-F]{1,4}:){4}((:[\da-fA-F]{1,4}){1,2}|:)$|^([\da-fA-F]{1,4}:){5}:([\da-fA-F]{1,4})?$|^([\da-fA-F]{1,4}:){6}:$
	#ce
	;IP或端口有问题
	$ip_regExp = "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$"
	$ip_regExpV6 = "^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "6}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){"
	$ip_regExpV6 += "3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^::([\da-fA-F]{1,4}:){0,4}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){"
	$ip_regExpV6 += "3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:):([\da-fA-F]{1,4}:){0,3}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){"
	$ip_regExpV6 += "3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "2}:([\da-fA-F]{1,4}:){0,2}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){"
	$ip_regExpV6 += "3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "3}:([\da-fA-F]{1,4}:){0,1}((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){"
	$ip_regExpV6 += "3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "4}:((25[0-5]|2[0-4]\d|[01]?\d\d?)\.){"
	$ip_regExpV6 += "3}(25[0-5]|2[0-4]\d|[01]?\d\d?)$|^([\da-fA-F]{1,4}:){7"
	$ip_regExpV6 += "}[\da-fA-F]{1,4}$|^:((:[\da-fA-F]{1,4}){1,6}|:)$|^[\da-fA-F]{1,4}:((:[\da-fA-F]{1,4}){1,5}|:)$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "2}((:[\da-fA-F]{1,4}){1,4}|:)$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "3}((:[\da-fA-F]{1,4}){1,3}|:)$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "4}((:[\da-fA-F]{1,4}){1,2}|:)$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "5}:([\da-fA-F]{1,4})?$|^([\da-fA-F]{1,4}:){"
	$ip_regExpV6 += "6}:$)"
	;校对IP或PV6、端口
	If StringRegExp($A_data[$A_data_i][2], _Iif(StringInStr($A_data[$A_data_i][2], ":") > 0, $ip_regExpV6, $ip_regExp), 0) == 0 Or StringRegExp($A_data[$A_data_i][3], "^\d{1,7}$", 0) == 0 Then
		ClipPut("browse for " & $sql_where)
		;MsgBox(0, "", $sf & "的IP有误。" & @CRLF & $A_data[$A_data_i][2] & ":" & $A_data[$A_data_i][3] & "已经将VF修正此问题的语句复制到粘贴板。", 0.8)
		;TrayTip( "", $sf & "的IP有误。" & @CRLF & $A_data[$A_data_i][2] & ":" & $A_data[$A_data_i][3] & "已经将VF修正此问题的语句复制到粘贴板。",1,17)
		;Return
		;没必要停止执行，只要不修改IP和端口就行了
		If MsgBox(1, "登录服务器信息有误，确认是否继续", $sf & "的IP有误。" & @CRLF & $A_data[$A_data_i][2] & ":" & $A_data[$A_data_i][3] & "已经将VF修正此问题的语句复制到粘贴板。此时粘贴以查看。" & @CRLF & "点 确认 不修改登录服务器信息，只填充用户名和密码。") == 2 Then Return
	Else
		;如果通过检验，修改登录的IP和端口
		If Not ($ip == $A_data[$A_data_i][2] And $port == $A_data[$A_data_i][3]) Then
			ControlClick("登录系统", "选择省、自治区或直辖市", "TButton3")
			WinWait("修改服务器地址", "")
			ControlSetText("修改服务器地址", "", "TEdit4", $A_data[$A_data_i][2])
			ControlSetText("修改服务器地址", "", "TEdit3", $A_data[$A_data_i][3])
			ControlClick("修改服务器地址", "", "TButton3")
			WinActivate("登录系统", "选择省、自治区或直辖市")
		EndIf
	EndIf

	ControlSetText("登录系统", "选择省、自治区或直辖市", "TEdit2", $A_data[$A_data_i][4])
	ControlSetText("登录系统", "选择省、自治区或直辖市", "TEdit1", $A_data[$A_data_i][5])
	$Cilp_last_sf = StringReplace($sf, "'", "")
	$Cilp_last_username = $A_data[$A_data_i][4]
	$Cilp_last_password = $A_data[$A_data_i][5]
	$Cilp_last_college = $user_class
	ClipPut($Cilp_last_password)
	TrayTip("", $user_class & " 的登录信息已经输入.." & @CR & "密码也已经复制到粘贴板", 1, 17)



	Local $r = 0
	$sf_py = ""
	For $r = 0 To UBound($G_ss_py_dm, 1) - 1
		If $G_ss_py_dm[$r][0] = $Cilp_last_sf Then
			$sf_py = $G_ss_py_dm[$r][1]
			ExitLoop
		EndIf
	Next

	
	#CS ;~ ConsoleWrite(@CRLF)
		;~ ConsoleWrite($db_dir&@CRLF)
		
		Local $file_name_t, $db_file_exists = True
		;~ 检查文件是否存在
		For $file_name_t In StringSplit("t_tdd.db|t_jhk.db", "|", 3)
		If Not FileExists($db_dir & "\" & $file_name_t) Then $db_file_exists = False
		Next
	#CE
	Local $db_mdb = $data_dir & $sf_py & "\" & $Cilp_last_username & "\" & $data_university_mdb_filename
	If FileExists($db_mdb) Then
		$conn = _OpenConnToMdb($db_mdb)
		If IsObj($conn) Then
;~ 		$conn = ObjCreate("ADODB.Connection")
;~ 		$rs = ObjCreate("ADODB.Recordset")
			;$sql_where = "sf='" & StringReplace($sf, "'", "") & "' AND (user_class='" & $user_class & "' OR user_class=' ' ) AND is_default='1'"
			;$sql = "SELECT * from NacuesC_ncuzjc WHERE " & $sql_where & " ORDER BY user_class desc"
			;$conn.Open("Driver=Microsoft Visual FoxPro Driver;SourceType=DBF;SourceDB=" & @ScriptDir)
;~ 		$conn.Open("DRIVER={Driver do Microsoft Paradox (*.db )};DriverId=26;dbq=" & $db_dir)

;~ 显示所有考生
;~ 	--全部移回0-9
;~ 	update t_tdd set jhxz=chr(asc(jhxz)+15) where (asc(jhxz) between 33 and 42) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
;~ 	--全部移回A-G
;~ 	update t_tdd set jhxz=chr(asc(jhxz)+ 7) where (asc(jhxz) between 58 and 64) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
			$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)+15) where (asc(jhxz) between 33 and 42) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")
			$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)+ 7) where (asc(jhxz) between 58 and 64) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")
			$conn.Close()
		EndIf
	EndIf
	TrayItemSetState($tdd_filter_1, $TRAY_CHECKED) ;切换显示全部勾
	TrayItemSetState($tdd_filter0, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter1, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter2, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter3, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter4, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter5, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter6, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter7, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter8, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filter9, $TRAY_UNCHECKED)
	TrayItemSetState($tdd_filtern, $TRAY_UNCHECKED)

EndFunc   ;==>doNacuesC_exe

Func dianjiEx_exe()
	#cs
		0前一考生
		1后一考生
		2
		3固定页面
		4
		5调整专业
		6退档原因
		7
		8打印
		9
		10结束浏览
		11
		12帮助
		13
		0 10372930170009 - 杨永志
		1
		2报名信息
		3成绩与志愿
		4体检信息
		5指纹
		6体检表
		7附加表
		8
	#ce
;~ If Not WinExists("全国普通高校招生网上录取 - 院校子系统") Then Return
	Sleep(100)
	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return

	;	WinActivate("全国普通高校招生网上录取 - 院校子系统")
;~ 	Send("!{Tab}{Tab}")
;~ 	Send("{ESC}")
;~ 	MsgBox(0,"",WinGetState("全国普通高校招生网上录取 - 院校子系统"))

	closeOtherWin()
	If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return
;~ 	Sleep(20)
	Send("^{home}")
	Send("^i")
	Sleep(50)
	If Not WinWaitActive("[Class:TDNForm]", "", 5) Then Return
;~ [TITLE:考生信息;CLASS:TDNForm;]
	Local $hWnd_ksxx = WinGetHandle("[LAST]")
	Local $hWnd_Tray_Toolbar = ControlGetHandle($hWnd_ksxx, '', 'TToolBar1')
	Local $hWnd_Tray_Toolbar2 = ControlGetHandle($hWnd_ksxx, '', 'TToolBar2')
	Local $iCount = _GUICtrlToolbar_ButtonCount($hWnd_Tray_Toolbar)
	Local $iCount2 = _GUICtrlToolbar_ButtonCount($hWnd_Tray_Toolbar2)
;~ 	Local $Direction_is_down = False
	Local $i_bnt = 0 ;指示点击 上一个考生 还是下一个考生
	Local $Toolbar2_curr = 2
	Local $Toolbar2_curr_last = 2
;~ Sleep(100)

	;ConsoleWrite($iCount&@CRLF)
	;ConsoleWrite(_GUICtrlToolbar_IsButtonEnabled($hWnd_Tray_Toolbar,0)&@CRLF)
	;固定页面
	;_GUICtrlToolbar_CheckButton($hWnd_Tray_Toolbar,_GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, 3),True)
	;MsgBox(0,"",_GUICtrlToolbar_IsButtonChecked($hWnd_Tray_Toolbar,_GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, 3)))
	If Not _GUICtrlToolbar_IsButtonChecked($hWnd_Tray_Toolbar, _GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, 3)) Then
		_GUICtrlToolbar_ClickButton($hWnd_Tray_Toolbar, _GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, 3), "left", False)
	EndIf
	Sleep(30)
	;点击 第三个按钮:报名信息
	_GUICtrlToolbar_ClickButton($hWnd_Tray_Toolbar2, _GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, 2), "left", False)
	Sleep(30)
	;_GUICtrlToolbar_CheckButton($hWnd_Tray_Toolbar,_GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, 3),True)
	;_GUICtrlToolbar_CheckButton($hWnd_Tray_Toolbar2,_GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar2, 2),True)
;~ 	Local $iStart = TimerInit()
;~ 用BitAND(WinGetState("[TITLE:考生信息;CLASS:TDNForm;]"),2) 代替 WinExists("[TITLE:考生信息;CLASS:TDNForm;]").因为TDNForm可能只是不可见..
;~ 1 = 窗口存在
;~  2 = 窗口可见
;~  4 = 窗口可用(未被禁用)
;~  8 = 窗口为激活状态
;~  16 = 窗口为最小化状态
;~  32 = 窗口为最大化状态
	While BitAND(WinGetState($hWnd_ksxx), 2) And _GUICtrlToolbar_IsButtonEnabled($hWnd_Tray_Toolbar, 0)
		While BitAND(WinGetState($hWnd_ksxx), 2) And (Not WinActive($hWnd_ksxx))
			WinWaitActive($hWnd_ksxx, "", 1)
		WEnd
		If Not BitAND(WinGetState($hWnd_ksxx), 2) Then Return
		_GUICtrlToolbar_ClickButton($hWnd_Tray_Toolbar, _GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, 0), "left", True)
	WEnd
	For $i_2 = 2 To $iCount2 - 2

;~ 		因为 体检信息 是与tdd一起发数据过来的,所以可以跳过 体检信息 栏(非"体检表")
		If $i_2 == 4 Then ContinueLoop
		
		While BitAND(WinGetState($hWnd_ksxx), 2) And (Not WinActive($hWnd_ksxx))
			;如果如果等待超过1秒钟,看是否有提示信息页面出错,有则关闭提示
			If WinWaitActive($hWnd_ksxx, "", 1) = 0 And WinExists("[TITLE:考生信息;CLASS:#32770;]", "生成考生信息页面出错:") Then WinClose("[TITLE:考生信息;CLASS:#32770;]", "生成考生信息页面出错:")
		WEnd
		$Toolbar2_curr = $i_2
		;_GUICtrlToolbar_CheckButton($hWnd_Tray_Toolbar2,_GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar2,$Toolbar2_curr),True)
		_GUICtrlToolbar_ClickButton($hWnd_Tray_Toolbar2, _GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, $Toolbar2_curr), "left", False)
		Sleep(80)
;~ 	如果点击按钮没有导致改变,说明是非可点按钮,进入下一个循环体
		If $i_2 == 2 Or (Not _GUICtrlToolbar_IsButtonChecked($hWnd_Tray_Toolbar2, _GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar2, $Toolbar2_curr_last))) Then
			$Toolbar2_curr_last = $Toolbar2_curr
		Else
			ContinueLoop
		EndIf
		

;~ 		$Direction_is_down = (Not $Direction_is_down)
		$i_bnt = Mod($i_bnt + 1, 2);切换点击"前一考生" 还是"后一考生"按钮.
		While BitAND(WinGetState($hWnd_ksxx), 2) And _GUICtrlToolbar_IsButtonEnabled($hWnd_Tray_Toolbar, $i_bnt)
;~ 		在"成绩与志愿"中,如果"下载全部志愿"按钮可用,点击下载
			If $i_2 == 3 And ControlCommand($hWnd_ksxx, "", "TButton2", "IsVisible", "") == 1 Then
				While BitAND(WinGetState($hWnd_ksxx), 2) And (Not WinActive($hWnd_ksxx))
					;如果如果等待超过1秒钟,看是否有提示信息页面出错,有则关闭提示
					If WinWaitActive($hWnd_ksxx, "", 1) = 0 And WinExists("[TITLE:考生信息;CLASS:#32770;]", "生成考生信息页面出错:") Then WinClose("[TITLE:考生信息;CLASS:#32770;]", "生成考生信息页面出错:")
				WEnd
				If Not BitAND(WinGetState($hWnd_ksxx), 2) Then Return
				ControlClick($hWnd_ksxx, "", "TButton2")
			EndIf

			While BitAND(WinGetState($hWnd_ksxx), 2) And (Not WinActive($hWnd_ksxx))
				;如果如果等待超过1秒钟,看是否有提示信息页面出错,有则关闭提示
				If WinWaitActive($hWnd_ksxx, "", 1) = 0 And WinExists("[TITLE:考生信息;CLASS:#32770;]", "生成考生信息页面出错:") Then WinClose("[TITLE:考生信息;CLASS:#32770;]", "生成考生信息页面出错:")
			WEnd
			If Not BitAND(WinGetState($hWnd_ksxx), 2) Then Return
			_GUICtrlToolbar_ClickButton($hWnd_Tray_Toolbar, _GUICtrlToolbar_IndexToCommand($hWnd_Tray_Toolbar, $i_bnt), "left", True)
		WEnd
	Next
EndFunc   ;==>dianjiEx_exe


Func Output_Tdd($tdd_xh = 0, $update_czx = False)
	Local $c_yx_mc, $c_sf_mc, $c_pc_dm, $c_pc_mc, $c_kl_dm, $c_kl_mc, $c_jhxz_dm, $c_jhxz_mc, $c_zt, $c_tddw_dm, $c_tddw_mc, $c_zy_dm, $c_zy_mc, $c_ks_td, $c_ks_not_zy
	Local $t, $i
	Local $sql_tdd_where = ""
	Local $conn, $conn_share, $conn_dbf, $conn_exe
	Local $rs, $rs_share, $rs_dbf, $rs_exe
	Local $db_dir
;~ 如果没有运行院校子系统过程退出
	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return
	If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return

	#CS 	If Not FileExists($exe_dir & "\T_SSZT.DB") Then
		MsgBox(0, "错误", "数据表T_SSZT.DB不在" & @CRLF & "请检查：" & $exe_dir & "\T_SSZT.DB是否存在")
		;~ 	Return
		Return ;Exit
		EndIf
	#CE

;~ Dim $tdd_xh
	If Not IsDeclared("tdd_xh") Then Assign("tdd_xh", 0)
	If Not IsNumber($tdd_xh) Then $tdd_xh = 0

	If Not IsDeclared("update_czx") Then Assign("update_czx", False)
	If Not IsBool($update_czx) Then $update_czx = False
;~ 获得子系统筛选工具条上的值
	$statusbarhwnd = ControlGetHandle('[Class:TNCMainForm]', '', '[CLASS:TToolBar; INSTANCE:2]')
;~ Local $panelcount = _GUICtrlToolbar_ButtonCount($statusbarhwnd)
;~ For $panel = 0 to $panelcount-1
;~ 	$statusbartext &= $panel & ": " & _GUICtrlToolbar_GetString($statusbarhwnd, $panel) & @CRLF
;~ 	$statusbartext &= $panel & ": " & _GUICtrlToolbar_GetButtonText($statusbarhwnd, $panel) & @CRLF
;~ Next
	#cs
		0: YX_SSBtn
		0: 陕西师范大学▲ - 江西省考生
		1: PCBtn
		1: 0 提前本科
		2: KL_JHXZBtn
		2: 1 文史 - 2 国防生
		3: ZTBtn
		3: [阅档中]
		4: TDDWBtn
		4: 所有的投档单位  |||| 1 江西考生
		5: ZYBtn
		5: 所有的专业 |||| 江西考生的所有专业 |||| 24 会计学⊙ ||||预退档考生 |||| 未确定专业考生
		
		
		有一张表：t_tddzdsx 投档单字段属性 说明了t_tdd中所使用的字段及含义
		其中给出了bh是投档单编号。（可能最重要的内容。。。）
		
	#ce

	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 0), " - ", 3)
	If @error Then Return

	$c_yx_mc = $t[0]
	$c_sf_mc = StringReplace($t[1], "录取进程", "")
;~ MsgBox(0, "",$c_sf_mc)
	If $c_sf_mc = "" Or $Cilp_last_sf <> $c_sf_mc Then
		MsgBox(0, "错误", "插件所存储的登录信息与院校子系统不一致" & @CRLF & "请使用插件自动填写登录信息再使用此功能." & @CRLF & "插件所存储的省份：" & $Cilp_last_sf & "" & @CRLF & "院校子系统登录省份为：" & $c_sf_mc)
		Return ;Exit
		;### Tidy Error -> "endif" is closing previous "func" on line 718
	EndIf


	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 1), " ", 3)
	$c_pc_dm = $t[0]
	$c_pc_mc = $t[1]
;~ MsgBox(0, "",$c_pc_mc)
	$sql_tdd_where &= _Iif($sql_tdd_where == "", "", " and") & " t_tdd.pcdm='" & $c_pc_dm & "'"

;~ 获得科类代码及计划性质
	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 2), " - ", 3)
	$t = StringSplit($t[0], " ", 3)
	$c_kl_dm = $t[0]
	$c_kl_mc = $t[1]
;~ MsgBox(0, "",$c_kl_mc)
	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 2), " - ", 3)
	$t = StringSplit($t[1], " ", 3)
	$c_jhxz_dm = $t[0]
	$c_jhxz_mc = $t[1]
;~ MsgBox(0, "",$c_jhxz_mc)
	$sql_tdd_where &= _Iif($sql_tdd_where == "", "", " and") & " t_tdd.kldm='" & $c_kl_dm & "'"
	$sql_tdd_where &= _Iif($sql_tdd_where == "", "", " and") & " t_tdd.jhxz='" & $c_jhxz_dm & "'"

;~ 获得投档单位筛选信息
	If _GUICtrlToolbar_GetButtonText($statusbarhwnd, 4) <> "所有的投档单位" Then
		$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 4), " ", 3)
		$c_tddw_dm = $t[0]
		$c_tddw_mc = $t[1]
;~ MsgBox(0, "",$c_pc_mc)
		$sql_tdd_where &= _Iif($sql_tdd_where == "", "", " and") & " t_tdd.tddw='" & $c_tddw_dm & "'"
	EndIf



;~ 获得专业筛选信息
	$t = _GUICtrlToolbar_GetButtonText($statusbarhwnd, 5)
	Select
		Case $t = "所有的专业" Or $t = ($c_tddw_mc + "的所有专业")
		Case $t = "预退档考生"
			$c_ks_td = True
			$sql_tdd_where &= _Iif($sql_tdd_where == "", "", " and") & " t_tdd.kszt='2'"
;~ 		$sql_tdd_where &=_Iif($sql_tdd_where=="",""," and") &" t_tdd.tdyydm is null"
		Case $t = "未确定专业考生"
			$c_ks_not_zy = True
			$sql_tdd_where &= _Iif($sql_tdd_where == "", "", " and") & " t_tdd.kszt='0'"
;~ 		$sql_tdd_where &=_Iif($sql_tdd_where=="",""," and") &" t_tdd.tdyydm is null"
		Case Else
			$t = StringSplit($t, " ", 3)
			$c_zy_dm = $t[0]
			$c_zy_mc = $t[1]
;~ 		$sql_tdd_where &=_Iif($sql_tdd_where=="",""," and") &" t_tdd.kszt='0'"
			$sql_tdd_where &= _Iif($sql_tdd_where == "", "", " and") & " t_tdd.lqzy='" & $c_zy_dm & "'"
	EndSelect

	Dim $ks_count
	$ks_count = Number(StringReplace(_GUICtrlStatusBar_GetText(ControlGetHandle('[Class:TNCMainForm]', '', 'TStatusBar1'), 1), "考生数", ""))
;~ ConsoleWrite($ks_count)

	#CS 	$conn_exe = ObjCreate("ADODB.Connection")
		$rs_exe = ObjCreate("ADODB.Recordset")
		$conn_exe.Open("DRIVER={Driver do Microsoft Paradox (*.db )};DriverId=26;dbq=" & $exe_dir)
		;select sspy from t_sszt where ssmc='江西省'
		$rs_exe = $conn_exe.execute("select sspy,ssdm from t_sszt where ssmc='" & $Cilp_last_sf & "'")
		If @error Then
		MsgBox(0, "", "无法打开t_sszt.db在:" & $exe_dir)
		;Return
		Return ;Exit
		EndIf
		Local $sf_py = "", $sf_dm = ""
		If Not ($rs_exe.eof Or $rs_exe.bof) Then
		;~ 	ConsoleWrite($RS_exe.Fields(0).value)
		$sf_py = $rs_exe.Fields(0).value
		$sf_dm = $rs_exe.Fields(1).value
		Else
		MsgBox(0, "", "省份不存在:" & $Cilp_last_sf)
		Return
		EndIf
		$db_dir = $exe_dir & $sf_py & "\" & $Cilp_last_username
		;~ ConsoleWrite(@CRLF)
		;~ ConsoleWrite($db_dir&@CRLF)
	#CE
	Local $r = 0
	Local $sf_py = "", $sf_dm = ""
	For $r = 0 To UBound($G_ss_py_dm, 1) - 1
		If $G_ss_py_dm[$r][0] = $Cilp_last_sf Then
			$sf_py = $G_ss_py_dm[$r][1]
			$sf_dm = $G_ss_py_dm[$r][2]
			ExitLoop
		EndIf
	Next



	#CS ;~ 检查文件是否存在
		For $file_name_t In StringSplit("t_tdd.db|t_jhk.db|TD_XBDM.DB", "|", 3)
		If Not FileExists($db_dir & "\" & $file_name_t) Then
		MsgBox(0, "", "数据表" & $file_name_t & "不存在" & @CR & "请检查：" & $db_dir & "\" & $file_name_t & "是否存在")
		;~ 		Return
		Return ;Exit
		EndIf
		Next
		
		
		$conn = ObjCreate("ADODB.Connection")
		$rs = ObjCreate("ADODB.Recordset")
		;$sql_where = "sf='" & StringReplace($sf, "'", "") & "' AND (user_class='" & $user_class & "' OR user_class=' ' ) AND is_default='1'"
		;$sql = "SELECT * from NacuesC_ncuzjc WHERE " & $sql_where & " ORDER BY user_class desc"
		;$conn.Open("Driver=Microsoft Visual FoxPro Driver;SourceType=DBF;SourceDB=" & @ScriptDir)
		$conn.Open("DRIVER={Driver do Microsoft Paradox (*.db )};DriverId=26;dbq=" & $db_dir)
	#CE

	Local $db_mdb = $data_dir & $sf_py & "\" & $Cilp_last_username & "\" & $data_university_mdb_filename
	Local $db_mdb_share = $data_dir & $sf_py & "\" & $Cilp_last_username & "\" & $data_university_share_mdb_filename
	If Not FileExists($db_mdb) Then
		MsgBox(0, "", "数据库" & $db_mdb & "不存在" & @CR & "请检查：" & $db_mdb & "是否存在")
		Return
	EndIf
	
	If Not FileExists($db_mdb_share) Then
		MsgBox(0, "", "数据库" & $db_mdb_share & "不存在" & @CR & "请检查：" & $db_mdb_share & "是否存在")
		Return
	EndIf
	
	$conn = _OpenConnToMdb($db_mdb)
	If Not IsObj($conn) Then
		MsgBox(0, "", "连接数据库" & $db_mdb & "失败。")
		Return
	EndIf

	$conn_share = _OpenConnToMdb($db_mdb_share)
	If Not IsObj($conn_share) Then
		MsgBox(0, "", "连接数据库" & $db_mdb_share & "失败。")
		Return
	EndIf
	

;~ 检查文件是否存在(控制线、投档单考生、投档单考生信息数据表,模板)
;~ $files_list=StringSplit("Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")
	For $file_name_t In StringSplit("czx.dbf|tdd_ks_info.dbf|tdd_ks.dbf|模板.xls", "|", 3)
		If Not FileExists(@ScriptDir & "\" & $file_name_t) Then
			MsgBox(0, "", "数据表" & $file_name_t & "不存在" & @CR & "请检查：" & @ScriptDir & "\" & $file_name_t & "是否存在")
;~ 		Return
			Return ;Exit
		EndIf
	Next


	Local $c_czx = 0, $cj_zd = "tdcj", $cj_zd_mc = ""

	$conn_dbf = ObjCreate("ADODB.Connection")
;~ $rs_dbf = ObjCreate("ADODB.Recordset")
	;$conn.Open("Driver=Microsoft Visual FoxPro Driver;SourceType=DBF;SourceDB=" & @ScriptDir)
	$conn_dbf.Open("Provider=Microsoft.Jet.OLEDB.4.0;dBase IV;HDR=NO;IMEX=2;DATABASE=" & @ScriptDir)
;~ $conn_dbf.execute("delete from czx")
;~ $rs_dbf=$conn_dbf.execute("select * from czx where sf='' and pcdm='' and kldm='' and jhxz=''")
	$rs_dbf = $conn_dbf.execute("select czx,tdcj_zd from czx where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "'")
	If Not ($rs_dbf.eof Or $rs_dbf.bof) Then
		$c_czx = $rs_dbf.fields(0).value
		$cj_zd = StringReplace($rs_dbf.fields(1).value, " ", "")
		If Not $update_czx Then
			If $cj_zd <> "tdcj" And $cj_zd <> "" Then
				$rs = $conn.execute("select zdmc from t_tddzdsx where zddh='" & StringUpper($cj_zd) & "'")
				If Not ($rs.eof Or $rs.bof) Then
					$cj_zd_mc = $rs.Fields(0).Value
				Else
;~ 		在属性表中没有找到成绩字段，重置为空
					$cj_zd = ""
				EndIf
			Else
				$cj_zd = ""
			EndIf
		Else
			$c_czx = InputBox("请重新输入控制线", $c_sf_mc & " " & $c_pc_dm & $c_pc_mc & @CRLF & $c_kl_dm & $c_kl_mc & " - " & $c_jhxz_dm & $c_jhxz_mc & @CRLF & "的控制线是:" & @CRLF & "(如控制线不是指投档成绩," & @CRLF & "请加一个逗号,例:“521,”):", $c_czx)
			If @error Then
				Return ;Exit
			EndIf
			If StringInStr($c_czx, ",") > 0 Then
				$cj_zd = InputBox("请输入控制线所指成绩字段名", StringReplace(StringReplace("tdcj,投档成绩(默认)|yxdrcj,院校导入成绩|tzcj,特征成绩|cj,辅助投档分|gkcjx01,代号为01的高考成绩|更多请使用[转换全部投档单]获得字段名", "|", @CRLF), ",", @TAB), $cj_zd)
				If @error Then
					$cj_zd = "tdcj"
					SetError(0)
				EndIf
			Else
				$cj_zd = ""
			EndIf
			$c_czx = Number(StringReplace($c_czx, ",", ""))
			If $c_czx <= 0 Then
				MsgBox(0, "错误", "" & $c_sf_mc & @CRLF & $c_pc_dm & $c_pc_mc & @CRLF & $c_kl_dm & $c_kl_mc & @CRLF & $c_jhxz_dm & $c_jhxz_mc & @CRLF & "的控制线是:" & $c_czx)
				Return ;Exit
			EndIf
			$cj_zd = StringReplace(StringLower($cj_zd), " ", "")
			If $cj_zd <> "tdcj" And $cj_zd <> "" Then
				$rs = $conn.execute("select zdmc from t_tddzdsx where zddh='" & StringUpper($cj_zd) & "'")
				If Not ($rs.eof Or $rs.bof) Then
					$cj_zd_mc = $rs.Fields(0).Value
				Else
;~ 		在属性表中没有找到成绩字段，重置为空
					$cj_zd = ""
				EndIf
			Else
				$cj_zd = ""
			EndIf

			;" & $c_czx & ",'" & $cj_zd & "'
			$conn_dbf.execute("update czx set czx=" & $c_czx & ",tdcj_zd='" & $cj_zd & "' where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "'")

		EndIf
	Else
		$rs_dbf = $conn_dbf.execute("select czx from czx where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "'")
		If Not ($rs_dbf.eof Or $rs_dbf.bof) Then $c_czx = $rs_dbf.fields(0).value
		$c_czx = InputBox("请输入控制线", $c_sf_mc & " " & $c_pc_dm & $c_pc_mc & @CRLF & $c_kl_dm & $c_kl_mc & " - " & $c_jhxz_dm & $c_jhxz_mc & @CRLF & "的控制线是:" & @CRLF & "(如控制线不是指投档成绩," & @CRLF & "请加一个逗号,例:“521,”):", $c_czx)
		If @error Then
			Return ;Exit
		EndIf
		If StringInStr($c_czx, ",") > 0 Then
			$cj_zd = InputBox("请输入控制线所指成绩字段名", StringReplace(StringReplace("tdcj,投档成绩(默认)|yxdrcj,院校导入成绩|tzcj,特征成绩|cj,辅助投档分|gkcjx01,代号为01的高考成绩|更多请使用[转换全部投档单]获得字段名", "|", @CRLF), ",", @TAB), $cj_zd)
			If @error Then
				$cj_zd = "tdcj"
				SetError(0)
			EndIf
		EndIf
		$c_czx = Number(StringReplace($c_czx, ",", ""))
		If $c_czx <= 0 Then
			MsgBox(0, "错误", "" & $c_sf_mc & @CRLF & $c_pc_dm & $c_pc_mc & @CRLF & $c_kl_dm & $c_kl_mc & @CRLF & $c_jhxz_dm & $c_jhxz_mc & @CRLF & "的控制线是:" & $c_czx)
			Return ;Exit
		EndIf
		$cj_zd = StringReplace(StringLower($cj_zd), " ", "")
		If $cj_zd <> "tdcj" And $cj_zd <> "" Then
			$rs = $conn.execute("select zdmc from t_tddzdsx where zddh='" & StringUpper($cj_zd) & "'")
			If Not ($rs.eof Or $rs.bof) Then
				$cj_zd_mc = $rs.Fields(0).Value
			Else
;~ 		在属性表中没有找到成绩字段，重置为空
				$cj_zd = ""
			EndIf
		Else
			$cj_zd = ""
		EndIf


;~ 	ConsoleWrite(     "insert into czx (sf,pcdm,kldm,jhxz,czx) values ('"&$c_sf_mc&"','"&$c_pc_dm&"','"&$c_kl_dm&"','"&$c_jhxz_dm&"',"&$c_czx&")"&@CRLF)
		$conn_dbf.execute("insert into czx (sf,pcdm,kldm,jhxz,czx,tdcj_zd) values ('" & $c_sf_mc & "','" & $c_pc_dm & "','" & $c_kl_dm & "','" & $c_jhxz_dm & "'," & $c_czx & ",'" & $cj_zd & "')")
	EndIf


	#CS ;~ MsgBox(0,"",$cj_zd&@TAB&$cj_zd_mc)
		
		If Not FileExists($db_dir & "\t_tdd.db") Then
		MsgBox(0, "错误", "数据表t_tdd.db不在" & @CRLF & "请检查：" & $db_dir & "\t_tdd.db是否存在")
		Return ;Exit
		EndIf
		$rs = $conn.execute("select yxmc from t_yxzt")
		If @error Then
		MsgBox(0, "", "无法打开t_yxzt.db在:" & $db_dir)
		;Return
		Return ;Exit
		EndIf
		If Not ($rs.eof Or $rs.bof) Then
		If $c_yx_mc <> $rs.Fields(0).value Then
		MsgBox(0, "", "院校名称不对，请使用插件登录子系统再用此功能:" & $rs.Fields(0).value)
		If AscW($rs.Fields(0).value) < 256 Then
		MsgBox(0, "解决方案：", "导出投档单时,提示院校名称是乱码的解决方案 " & @CRLF & _
		"打开windows【控制面板】－【BDE  Administrator】，鼠标点左边窗口上方的【Configuration】: " & @CRLF & _
		"1.在窗口中依次点开【Drivers】－【Native】－【Paradox】，然后在右边的参数列表窗口中，修改【LangDriver】的值为Paradox China 936。 " & @CRLF & _
		"2.在窗口中依次点开【System】－【INIT】然后在右边的参数列表窗口中，修改【LangDriver】的值为Paradox China 936。 " & @CRLF & _
		"3.最后点击‘Object’菜单下的‘Apply’保存设置。 " & @CRLF & _
		"这样创建出来的db表的Table Lanague都为Paradox China 936了,即使用中文编码。 " & @CRLF & _
		"退出院校子系统,在院校子系统程序目录,下面删除所有的省份文件夹(或重命名,如:jiangxi改为jiangxixxx),重新联机登录下载数据.")
		EndIf
		Return ;Exit
		EndIf
		EndIf
	#CE
	#CS ;~ ConsoleWrite($sql_tdd_where&@CRLF)
		$rs = $conn.execute("SELECT count(1) FROM T_tdd" & _Iif($sql_tdd_where == "", "", " where " & $sql_tdd_where))
		If @error Then
		MsgBox(0, "", "无法打T_tdd.db在:" & $db_dir)
		;Return
		Return ;Exit
		EndIf
		If Not ($rs.eof Or $rs.bof) Then
		;~ 		MsgBox(0, "", "院校名称不对，请使用插件登录子系统再用此功能:"&$RS.Fields(0).value)
		;~ ConsoleWrite($RS.Fields(0).value&@CRLF)
		
		EndIf
	#CE


;~ $RS=$conn.execute("select * from t_tdd order by pcdm,kldm,tdcj desc,ksh")
;~ If @error Then
;~ 	MsgBox(0, "", "无法打开t_ttd.db在:"&$db_dir)

;~ 	;Return
;~ 	Exit
;~ EndIf

;~ 显示所有考生
;~ 	--全部移回0-9
;~ 	update t_tdd set jhxz=chr(asc(jhxz)+15) where (asc(jhxz) between 33 and 42) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
;~ 	--全部移回A-G
;~ 	update t_tdd set jhxz=chr(asc(jhxz)+ 7) where (asc(jhxz) between 58 and 64) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
	$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)+15) where (asc(jhxz) between 33 and 42) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")
	$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)+ 7) where (asc(jhxz) between 58 and 64) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")

	If IsDeclared("tdd_filter_1") Then ;如果定义了过滤器菜单项目,重置它.
		Dim $tdd_filter_1, $tdd_filter0
		Dim $tdd_filter1
		Dim $tdd_filter2
		Dim $tdd_filter3
		Dim $tdd_filter4
		Dim $tdd_filter5
		Dim $tdd_filter6
		Dim $tdd_filter7
		Dim $tdd_filter8
		Dim $tdd_filter9
		Dim $tdd_filtern
		TrayItemSetState($tdd_filter_1, $TRAY_CHECKED) ;切换显示全部勾
		TrayItemSetState($tdd_filter0, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter1, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter2, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter3, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter4, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter5, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter6, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter7, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter8, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filter9, $TRAY_UNCHECKED)
		TrayItemSetState($tdd_filtern, $TRAY_UNCHECKED)
	EndIf


	Local $c_pc_kl_jhxz_ksh_last = "'#0000000000000'" ;已经生成考生号列表
	Local $c_pc_kl_jhxz_ksh = "'#0000000000000'" ;本投档单考生号列表
	Local $o_tdd_xh = 0 ;本次生成投档单序号
	Local $o_tdd_is_old = False
;~ @YEAR
	Local $date = @YEAR & "-" & @MON & "-" & @MDAY
	Local $sj = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
	Local $sj_str = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC

	Local $tdd_ks_dm = 1, $tdd_cnt = 0




;~ 得到当前批次科类性质的已经投档次数
	$rs_dbf = $conn_dbf.execute("select count(1) from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "'")
	If Not ($rs_dbf.eof Or $rs_dbf.bof) Then
		$tdd_cnt = $rs_dbf.Fields(0).Value
	Else
		Return
	EndIf
	;如果要取得的投档单序号小于等投档单数目
	If $tdd_xh > 0 And $tdd_xh <= $tdd_cnt Then
		$o_tdd_is_old = True
		$o_tdd_xh = $tdd_xh
;~ 得到当前批次科类性质的已经投档次数
		$rs_dbf = $conn_dbf.execute("select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' order by tdd_ks_dm")
		$i = 0
		While Not ($rs_dbf.eof Or $rs_dbf.bof)
			$i += 1
			If $i = $tdd_xh Then
				$tdd_ks_dm = $rs_dbf.Fields(0).Value
				ExitLoop
			EndIf
			$rs_dbf.movenext
		WEnd
		;本次考生号列表
		$c_pc_kl_jhxz_ksh = "'#0000000000000'"
		$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm =" & $tdd_ks_dm & "")
		While Not ($rs_dbf.eof Or $rs_dbf.bof)
			$c_pc_kl_jhxz_ksh &= ",'" & $rs_dbf.Fields(0).Value & "'"
			$rs_dbf.movenext
		WEnd
		;本次以前考生号列表
		$c_pc_kl_jhxz_ksh_last = "'#0000000000000'"
		$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm <" & $tdd_ks_dm & " And tdd_ks_dm in(select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "')")
		While Not ($rs_dbf.eof Or $rs_dbf.bof)
			$c_pc_kl_jhxz_ksh_last &= ",'" & $rs_dbf.Fields(0).Value & "'"
			$rs_dbf.movenext
		WEnd
	Else
		;获得已经生成过投档单的考生列表
		$c_pc_kl_jhxz_ksh_last = "'#0000000000000'"
		$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm in(select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "')")
		While Not ($rs_dbf.eof Or $rs_dbf.bof)
			$c_pc_kl_jhxz_ksh_last &= ",'" & $rs_dbf.Fields(0).Value & "'"
			$rs_dbf.movenext
		WEnd
;~ 查询是否有新考生
		$rs = $conn.execute("SELECT ksh FROM T_tdd where pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh not in(" & $c_pc_kl_jhxz_ksh_last & ")")
;~ 	有新投档考生
		If Not ($rs.eof Or $rs.bof) Then
			$o_tdd_xh = $tdd_cnt + 1
			$o_tdd_is_old = False

			#CS 		;~ 如果是最新的投档单.$tdd_ks_dm=max+1
				$rs_dbf=$conn_dbf.execute("select max(tdd_ks_dm) as max_dm from tdd_ks_info")
				If Not ($rs_dbf.eof Or $rs_dbf.bof) Then
				$tdd_ks_dm=Number($rs_dbf.Fields(0).Value)+1
				Else
				$tdd_ks_dm=1
				EndIf
			#CE

			Local $oRec = ObjCreate("ADODB.Recordset")
			If IsObj($oRec) = 0 Then Return SetError(2)
			With $oRec
				;$conn_dbf.execute("insert INTO tdd_ks_info (tdd_ks_dm,sf,pcdm,kldm,jhxz,sj) values ("&$tdd_ks_dm&",'"&$c_sf_mc&"','"&$c_pc_dm&"','"&$c_kl_dm&"','"&$c_jhxz_dm&"','"&$sj&"')")

				.Open("SELECT * FROM [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where tdd_ks_dm in(select max(tdd_ks_dm) as max_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "])", $conn_dbf, 0, 2)
				If Not ($oRec.eof Or $oRec.bof) Then
					$tdd_ks_dm = Number($oRec.Fields(0).Value) + 1
				Else
					$tdd_ks_dm = 1
				EndIf
				.AddNew
				.Fields.Item("tdd_ks_dm") = $tdd_ks_dm
				.Fields.Item("sf") = $c_sf_mc
				.Fields.Item("pcdm") = $c_pc_dm
				.Fields.Item("kldm") = $c_kl_dm
				.Fields.Item("jhxz") = $c_jhxz_dm
				.Fields.Item("sj") = $sj
				.Update
				.close

				.Open("SELECT tdd_ks_dm,ksh FROM tdd_ks where 1=2", $conn_dbf, 3, 3)

				;本次考生号列表
				$c_pc_kl_jhxz_ksh = "'#0000000000000'"
				While Not ($rs.eof Or $rs.bof)
					.AddNew
					.Fields.Item(0) = $tdd_ks_dm
					$c_pc_kl_jhxz_ksh &= ",'" & $rs.Fields(0).Value & "'"
					.Fields.Item(1) = $rs.Fields(0).Value
					$rs.movenext
				WEnd
				.Update
				.Close
			EndWith
		Else
			$o_tdd_xh = $tdd_cnt
			$o_tdd_is_old = True ;设置为使用最后一次的投档单
			$rs_dbf = $conn_dbf.execute("select max(tdd_ks_dm) as max_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "'")
			If Not ($rs_dbf.eof Or $rs_dbf.bof) Then $tdd_ks_dm = Number($rs_dbf.Fields(0).Value)
			;本次考生号列表
			$c_pc_kl_jhxz_ksh = "'#0000000000000'"
			$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm =" & $tdd_ks_dm & "")
			While Not ($rs_dbf.eof Or $rs_dbf.bof)
				$c_pc_kl_jhxz_ksh &= ",'" & $rs_dbf.Fields(0).Value & "'"
				$rs_dbf.movenext
			WEnd
			;本次以前考生号列表
			$c_pc_kl_jhxz_ksh_last = "'#0000000000000'"
			$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm <" & $tdd_ks_dm & " And tdd_ks_dm in(select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "')")
			While Not ($rs_dbf.eof Or $rs_dbf.bof)
				$c_pc_kl_jhxz_ksh_last &= ",'" & $rs_dbf.Fields(0).Value & "'"
				$rs_dbf.movenext
			WEnd
		EndIf

	EndIf
	If $c_pc_kl_jhxz_ksh = "'#0000000000000'" Then
		MsgBox(0, "错误", "本次投档单人数为0,无法生成.请试试生成历史投档单")
		Return ;Exit
	EndIf

;~ ConsoleWrite("第"&$o_tdd_xh&"次投档"&@CRLF)
	TrayTip("", "投档单生成后台进行中...请稍候!", 25, 1)

	Local $jhs
	Local $max_tdcj
	Local $min_tdcj
	Local $ks_cnt
	Local $lq_cnt
	Local $td_cnt
	Local $ylq_cnt
	Local $qe
	Local $sm = ""


;~ '计划数
;~ select pcdm,kldm,jhxz,sum(jhzxs) as jhs from t_jhk group by pcdm,kldm,jhxz
	$rs = $conn.execute("select sum(jhzxs) as jhs from t_jhk where pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' group by pcdm,kldm,jhxz")
	If Not ($rs.eof Or $rs.bof) Then
		$jhs = $rs.Fields(0).Value
	Else
		MsgBox(0, "错误", "计划数获取失败:")
		Return ;Exit
	EndIf


;~ '本次最高数、最低分、人数
;~ select pcdm,kldm,jhxz,max(tdcj) as max_tdcj,min(tdcj) as min_tdcj,count(1) as rs from t_tdd group by pcdm,kldm,jhxz
	If $cj_zd_mc <> "" Then
		$rs = $conn.execute("select max(" & $cj_zd & ") as max_tdcj,min(" & $cj_zd & ") as min_tdcj,count(1) as rs from t_tdd where pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") group by pcdm,kldm,jhxz")
	Else
		$rs = $conn.execute("select max(tdcj) as max_tdcj,min(tdcj) as min_tdcj,count(1) as rs from t_tdd where pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") group by pcdm,kldm,jhxz")
	EndIf

	If Not ($rs.eof Or $rs.bof) Then
		$max_tdcj = $rs.Fields(0).Value
		$min_tdcj = $rs.Fields(1).Value
		$ks_cnt = $rs.Fields(2).Value
	Else
		MsgBox(0, "错误", "本次最高分、最低分、人数获取失败:" & @CRLF & "可能是本次投档人数为0,或隐藏了考生." & @CRLF & "也可以试试“生成历史投档单”里面的第１次投档单")
		Return ;Exit
	EndIf

;~ '本次投档单,拟录取人数
;~ select pcdm,kldm,jhxz,count(1) as lq_cnt from t_tdd where (kszt='1' or kszt>='4') group by pcdm,kldm,jhxz
	$rs = $conn.execute("select count(1) as lq_cnt from t_tdd where (kszt='1' or kszt>='4') and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") group by pcdm,kldm,jhxz")
	$lq_cnt = 0
	If Not ($rs.eof Or $rs.bof) Then $lq_cnt = $rs.Fields(0).Value

;~ '本次投档单,拟退档人数
;~ select pcdm,kldm,jhxz,count(1) as td_cnt from t_tdd where kszt in('2','3') group by pcdm,kldm,jhxz
	$rs = $conn.execute("select count(1) as td_cnt from t_tdd where kszt in('2','3') and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") group by pcdm,kldm,jhxz")
	$td_cnt = 0
	If Not ($rs.eof Or $rs.bof) Then $td_cnt = $rs.Fields(0).Value


;~ '已录取人数,非本次投档单
;~ select pcdm,kldm,jhxz,count(1) as ylq_cnt from t_tdd where (kszt='1' or kszt>='4') group by pcdm,kldm,jhxz
	$rs = $conn.execute("select count(1) as ylq_cnt from t_tdd where (kszt='1' or kszt>='4') and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh_last & ") group by pcdm,kldm,jhxz")
	$ylq_cnt = 0
	If Not ($rs.eof Or $rs.bof) Then $ylq_cnt = $rs.Fields(0).Value


	$qe = $jhs - $ylq_cnt - $lq_cnt
	$sm = "第" & $o_tdd_xh & "次投档" & _Iif($o_tdd_xh > 1, ",原已录取" & $ylq_cnt & "人", "") & _Iif($cj_zd_mc <> "", ",以 " & $cj_zd_mc & " 为依据", "")



;~ 复制模板Excel文件
;~ DirCreate("sf\pc\kl_jhxz\"
	Local $output_dir, $output_file
	$output_dir = @ScriptDir & "\tdd投档单\" & $sf_dm & $c_sf_mc & "\" & $c_pc_dm & $c_pc_mc & "\" & $c_kl_dm & $c_kl_mc & "_" & $c_jhxz_dm & $c_jhxz_mc & "\"
	$output_file = $output_dir & $c_sf_mc & "" & $c_pc_mc & "" & $c_kl_mc & "" & $c_jhxz_mc & "第" & $o_tdd_xh & "次投档单" & $sj_str & ".xls"
	FileCopy(@ScriptDir & "\模板.xls", $output_dir, 8)
	FileMove($output_dir & "模板.xls", $output_file)
	$Cilp_tdd_file_last = $output_file
	If @error Then
		MsgBox(0, "错误", "文件复制失败:" & $output_file)
		;Return
		Return ;Exit
	EndIf

	$oExcel = _ExcelBookOpen($output_file)
	$oExcel.Application.DisplayAlerts = False
	$oExcel.Application.ScreenUpdating = False


	$oExcel.ActiveSheet.cells.replace("{year}", @YEAR)
	$oExcel.ActiveSheet.cells.replace("{sj}", Number(@MON) & "月" & Number(@MDAY) & "日")
	$oExcel.ActiveSheet.cells.replace("{sf}", $c_sf_mc)
	$oExcel.ActiveSheet.cells.replace("{pc}", $c_pc_mc)
	$oExcel.ActiveSheet.cells.replace("{jhxz}", $c_jhxz_mc)
	$oExcel.ActiveSheet.cells.replace("{kl}", $c_kl_mc)
	$oExcel.ActiveSheet.cells.replace("{czx}", $c_czx)
;~ 	MsgBox(0, "", $Cilp_last_college)
	If StringLen(StringStripWS($Cilp_last_college, 3)) > 1 And $Cilp_last_college <> "本部" Then
		$oExcel.ActiveSheet.cells.replace("{zk}", _Iif(StringInStr($c_pc_dm, "专") > 0, "(" & StringStripWS($Cilp_last_college, 3) & "、专科)", "(" & StringStripWS($Cilp_last_college, 3) & ")"))
	Else
		$oExcel.ActiveSheet.cells.replace("{zk}", _Iif(StringInStr($c_pc_dm, "专") > 0, "(专科)", ""))
	EndIf



	$oExcel.ActiveSheet.cells.replace("{jhs}", $jhs)
	$oExcel.ActiveSheet.cells.replace("{max_tdcj}", $max_tdcj)
	$oExcel.ActiveSheet.cells.replace("{min_tdcj}", $min_tdcj)
	$oExcel.ActiveSheet.cells.replace("{ks_cnt}", $ks_cnt)
	$oExcel.ActiveSheet.cells.replace("{lq_cnt}", $lq_cnt)
	$oExcel.ActiveSheet.cells.replace("{td_cnt}", $td_cnt)
	$oExcel.ActiveSheet.cells.replace("{qe}", _Iif($qe = 0, "计划完成", _Iif($qe > 0, "缺", "超录") & Abs($qe) & "人"))
	$oExcel.ActiveSheet.cells.replace("{sm}", $sm)

	$oExcel.ActiveSheet.cells.replace("{*}", "") ;删除没有被替换的大括号


;~ 生成投档单
	_ExcelSheetActivate($oExcel, "投档单")
	If @error Then
		MsgBox(0, "错误", "模板.xls中没有 投档单 工作表")
		$oExcel.Visible = True
		Return ;Exit
	EndIf

;~ select ksh,xm,xbmc,tdcj,tdzy,zydh1,jhk1.zymc as zymc1,zydh2,jhk2.zymc as zymc2,zydh3,jhk3.zymc as zymc3,zydh4,jhk4.zymc as zymc4,zydh5,jhk5.zymc as zymc5,zydh6,jhk6.zymc as zymc6,zyzytj,lxdh,zsyj,lqzy,jhk.zymc as lqzymc,kszt,tdyymc from (((((((((t_tdd left join td_xbdm on td_xbdm.xbdm=t_tdd.xbdm) left join t_jhk as jhk1 on jhk1.zydh=t_tdd.zydh1 and jhk1.pcdm=t_tdd.pcdm and jhk1.kldm=t_tdd.kldm and jhk1.jhxz=t_tdd.jhxz) left join t_jhk as jhk2 on jhk2.zydh=t_tdd.zydh2 and jhk2.pcdm=t_tdd.pcdm and jhk2.kldm=t_tdd.kldm and jhk2.jhxz=t_tdd.jhxz) left join t_jhk as jhk3 on jhk3.zydh=t_tdd.zydh3 and jhk3.pcdm=t_tdd.pcdm and jhk3.kldm=t_tdd.kldm and jhk3.jhxz=t_tdd.jhxz) left join t_jhk as jhk4 on jhk4.zydh=t_tdd.zydh4 and jhk4.pcdm=t_tdd.pcdm and jhk4.kldm=t_tdd.kldm and jhk4.jhxz=t_tdd.jhxz) left join t_jhk as jhk5 on jhk5.zydh=t_tdd.zydh5 and jhk5.pcdm=t_tdd.pcdm and jhk5.kldm=t_tdd.kldm and jhk5.jhxz=t_tdd.jhxz) left join t_jhk as jhk6 on jhk6.zydh=t_tdd.zydh6 and jhk6.pcdm=t_tdd.pcdm and jhk6.kldm=t_tdd.kldm and jhk6.jhxz=t_tdd.jhxz) left join t_jhk as jhk on jhk.zydh=t_tdd.lqzy and jhk.pcdm=t_tdd.pcdm and jhk.kldm=t_tdd.kldm and jhk.jhxz=t_tdd.jhxz) left join TD_TDYYDM ON t_tdd.TDYYDM=TD_TDYYDM.TDYYDM)  where ksh='10369552112178'
;~ ConsoleWrite("select ksh,xm,xbmc,tdcj,tdzy,zydh1,jhk1.zymc as zymc1,zydh2,jhk2.zymc as zymc2,zydh3,jhk3.zymc as zymc3,zydh4,jhk4.zymc as zymc4,zydh5,jhk5.zymc as zymc5,zydh6,jhk6.zymc as zymc6,zyzytj,lxdh,zsyj,lqzy,jhk.zymc as lqzymc,kszt,tdyymc from (((((((((t_tdd left join td_xbdm on td_xbdm.xbdm=t_tdd.xbdm) left join t_jhk as jhk1 on jhk1.zydh=t_tdd.zydh1 and jhk1.pcdm=t_tdd.pcdm and jhk1.kldm=t_tdd.kldm and jhk1.jhxz=t_tdd.jhxz) left join t_jhk as jhk2 on jhk2.zydh=t_tdd.zydh2 and jhk2.pcdm=t_tdd.pcdm and jhk2.kldm=t_tdd.kldm and jhk2.jhxz=t_tdd.jhxz) left join t_jhk as jhk3 on jhk3.zydh=t_tdd.zydh3 and jhk3.pcdm=t_tdd.pcdm and jhk3.kldm=t_tdd.kldm and jhk3.jhxz=t_tdd.jhxz) left join t_jhk as jhk4 on jhk4.zydh=t_tdd.zydh4 and jhk4.pcdm=t_tdd.pcdm and jhk4.kldm=t_tdd.kldm and jhk4.jhxz=t_tdd.jhxz) left join t_jhk as jhk5 on jhk5.zydh=t_tdd.zydh5 and jhk5.pcdm=t_tdd.pcdm and jhk5.kldm=t_tdd.kldm and jhk5.jhxz=t_tdd.jhxz) left join t_jhk as jhk6 on jhk6.zydh=t_tdd.zydh6 and jhk6.pcdm=t_tdd.pcdm and jhk6.kldm=t_tdd.kldm and jhk6.jhxz=t_tdd.jhxz) left join t_jhk as jhk on jhk.zydh=t_tdd.lqzy and jhk.pcdm=t_tdd.pcdm and jhk.kldm=t_tdd.kldm and jhk.jhxz=t_tdd.jhxz) left join TD_TDYYDM ON t_tdd.TDYYDM=TD_TDYYDM.TDYYDM)  where t_tdd.pcdm='"&$c_pc_dm&"' and t_tdd.kldm='"&$c_kl_dm&"' and t_tdd.jhxz='"&$c_jhxz_dm&"' and ksh not in("&$c_pc_kl_jhxz_ksh_last&")")
;~ ConsoleWrite("select ksh,xm,xbmc,tdcj"&_Iif($cj_zd_mc<>"",",t_tdd."&$cj_zd,"")&",tdzy,jhk1.zydh as zydh1,jhk1.zymc as zymc1,jhk2.zydh as zydh2,jhk2.zymc as zymc2,jhk3.zydh as zydh3,jhk3.zymc as zymc3,jhk4.zydh as zydh4,jhk4.zymc as zymc4,jhk5.zydh as zydh5,jhk5.zymc as zymc5,jhk6.zydh as zydh6,jhk6.zymc as zymc6,zyzytj,lxdh,zsyj,lqzy,jhk.zymc as lqzymc,kszt,tdyymc from (((((((((t_tdd left join td_xbdm on td_xbdm.xbdm=t_tdd.xbdm) left join t_jhk as jhk1 on jhk1.zydh=t_tdd.zydh1 and jhk1.pcdm=t_tdd.pcdm and jhk1.kldm=t_tdd.kldm and jhk1.jhxz=t_tdd.jhxz) left join t_jhk as jhk2 on jhk2.zydh=t_tdd.zydh2 and jhk2.pcdm=t_tdd.pcdm and jhk2.kldm=t_tdd.kldm and jhk2.jhxz=t_tdd.jhxz) left join t_jhk as jhk3 on jhk3.zydh=t_tdd.zydh3 and jhk3.pcdm=t_tdd.pcdm and jhk3.kldm=t_tdd.kldm and jhk3.jhxz=t_tdd.jhxz) left join t_jhk as jhk4 on jhk4.zydh=t_tdd.zydh4 and jhk4.pcdm=t_tdd.pcdm and jhk4.kldm=t_tdd.kldm and jhk4.jhxz=t_tdd.jhxz) left join t_jhk as jhk5 on jhk5.zydh=t_tdd.zydh5 and jhk5.pcdm=t_tdd.pcdm and jhk5.kldm=t_tdd.kldm and jhk5.jhxz=t_tdd.jhxz) left join t_jhk as jhk6 on jhk6.zydh=t_tdd.zydh6 and jhk6.pcdm=t_tdd.pcdm and jhk6.kldm=t_tdd.kldm and jhk6.jhxz=t_tdd.jhxz) left join t_jhk as jhk on jhk.zydh=t_tdd.lqzy and jhk.pcdm=t_tdd.pcdm and jhk.kldm=t_tdd.kldm and jhk.jhxz=t_tdd.jhxz) left join TD_TDYYDM ON t_tdd.TDYYDM=TD_TDYYDM.TDYYDM)  where t_tdd.pcdm='"&$c_pc_dm&"' and t_tdd.kldm='"&$c_kl_dm&"' and t_tdd.jhxz='"&$c_jhxz_dm&"' and ksh not in("&$c_pc_kl_jhxz_ksh_last&") order by "&_Iif($cj_zd_mc<>"",""&$cj_zd&" desc,","")&"tdcj desc,ksh"&@CRLF)
;~ 	$rs = $conn.execute("select ksh,xm,xbmc,tdcj" & _Iif($cj_zd_mc <> "", ",t_tdd." & $cj_zd, "") & ",tdzy,cstr(jhk1.zydh) as zydh1,jhk1.zymc as zymc1,cstr(jhk2.zydh) as zydh2,jhk2.zymc as zymc2,cstr(jhk3.zydh) as zydh3,jhk3.zymc as zymc3,cstr(jhk4.zydh) as zydh4,jhk4.zymc as zymc4,cstr(jhk5.zydh) as zydh5,jhk5.zymc as zymc5,cstr(jhk6.zydh) as zydh6,jhk6.zymc as zymc6,zyzytj,lxdh,zsyj,cstr(t_tdd.lqzy) as lqzy,jhk.zymc as lqzymc,kszt,cstr(t_tdd.tdyydm) as tdyydm,tdyymc,t_tdd.bz from (((((((((t_tdd left join td_xbdm on td_xbdm.xbdm=t_tdd.xbdm) left join t_jhk as jhk1 on jhk1.zydh=t_tdd.zydh1 and jhk1.pcdm=t_tdd.pcdm and jhk1.kldm=t_tdd.kldm and jhk1.jhxz=t_tdd.jhxz and jhk1.tddw=t_tdd.tddw) left join t_jhk as jhk2 on jhk2.zydh=t_tdd.zydh2 and jhk2.pcdm=t_tdd.pcdm and jhk2.kldm=t_tdd.kldm and jhk2.jhxz=t_tdd.jhxz and jhk2.tddw=t_tdd.tddw) left join t_jhk as jhk3 on jhk3.zydh=t_tdd.zydh3 and jhk3.pcdm=t_tdd.pcdm and jhk3.kldm=t_tdd.kldm and jhk3.jhxz=t_tdd.jhxz and jhk3.tddw=t_tdd.tddw) left join t_jhk as jhk4 on jhk4.zydh=t_tdd.zydh4 and jhk4.pcdm=t_tdd.pcdm and jhk4.kldm=t_tdd.kldm and jhk4.jhxz=t_tdd.jhxz and jhk4.tddw=t_tdd.tddw) left join t_jhk as jhk5 on jhk5.zydh=t_tdd.zydh5 and jhk5.pcdm=t_tdd.pcdm and jhk5.kldm=t_tdd.kldm and jhk5.jhxz=t_tdd.jhxz and jhk5.tddw=t_tdd.tddw) left join t_jhk as jhk6 on jhk6.zydh=t_tdd.zydh6 and jhk6.pcdm=t_tdd.pcdm and jhk6.kldm=t_tdd.kldm and jhk6.jhxz=t_tdd.jhxz and jhk6.tddw=t_tdd.tddw) left join t_jhk as jhk on jhk.zydh=t_tdd.lqzy and jhk.pcdm=t_tdd.pcdm and jhk.kldm=t_tdd.kldm and jhk.jhxz=t_tdd.jhxz and jhk.tddw=t_tdd.tddw) left join TD_TDYYDM ON t_tdd.TDYYDM=TD_TDYYDM.TDYYDM)  where t_tdd.pcdm='" & $c_pc_dm & "' and t_tdd.kldm='" & $c_kl_dm & "' and t_tdd.jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") order by " & _Iif($cj_zd_mc <> "", "" & $cj_zd & " desc,", "") & "tdcj desc,ksh")
	;$rs = $conn.execute("select ksh,xm,xbmc,tdcj" & _Iif($cj_zd_mc <> "", ",t_tdd." & $cj_zd, "") & ",tdzy,jhk1.zydh as zydh1,jhk1.zymc as zymc1,cstr(jhk2.zydh) as zydh2,jhk2.zymc as zymc2,cstr(jhk3.zydh) as zydh3,jhk3.zymc as zymc3,cstr(jhk4.zydh) as zydh4,jhk4.zymc as zymc4,cstr(jhk5.zydh) as zydh5,jhk5.zymc as zymc5,cstr(jhk6.zydh) as zydh6,jhk6.zymc as zymc6,zyzytj,lxdh,zsyj,cstr(t_tdd.lqzy) as lqzy,jhk.zymc as lqzymc,kszt,cstr(t_tdd.tdyydm) as tdyydm,tdyymc,t_tdd.bz from (((((((((t_tdd left join td_xbdm on td_xbdm.xbdm=t_tdd.xbdm) left join t_jhk as jhk1 on jhk1.zydh=t_tdd.zydh1 and jhk1.pcdm=t_tdd.pcdm and jhk1.kldm=t_tdd.kldm and jhk1.jhxz=t_tdd.jhxz and jhk1.tddw=t_tdd.tddw) left join t_jhk as jhk2 on jhk2.zydh=t_tdd.zydh2 and jhk2.pcdm=t_tdd.pcdm and jhk2.kldm=t_tdd.kldm and jhk2.jhxz=t_tdd.jhxz and jhk2.tddw=t_tdd.tddw) left join t_jhk as jhk3 on jhk3.zydh=t_tdd.zydh3 and jhk3.pcdm=t_tdd.pcdm and jhk3.kldm=t_tdd.kldm and jhk3.jhxz=t_tdd.jhxz and jhk3.tddw=t_tdd.tddw) left join t_jhk as jhk4 on jhk4.zydh=t_tdd.zydh4 and jhk4.pcdm=t_tdd.pcdm and jhk4.kldm=t_tdd.kldm and jhk4.jhxz=t_tdd.jhxz and jhk4.tddw=t_tdd.tddw) left join t_jhk as jhk5 on jhk5.zydh=t_tdd.zydh5 and jhk5.pcdm=t_tdd.pcdm and jhk5.kldm=t_tdd.kldm and jhk5.jhxz=t_tdd.jhxz and jhk5.tddw=t_tdd.tddw) left join t_jhk as jhk6 on jhk6.zydh=t_tdd.zydh6 and jhk6.pcdm=t_tdd.pcdm and jhk6.kldm=t_tdd.kldm and jhk6.jhxz=t_tdd.jhxz and jhk6.tddw=t_tdd.tddw) left join t_jhk as jhk on jhk.zydh=t_tdd.lqzy and jhk.pcdm=t_tdd.pcdm and jhk.kldm=t_tdd.kldm and jhk.jhxz=t_tdd.jhxz and jhk.tddw=t_tdd.tddw) left join TD_TDYYDM ON t_tdd.TDYYDM=TD_TDYYDM.TDYYDM)  where t_tdd.pcdm='" & $c_pc_dm & "' and t_tdd.kldm='" & $c_kl_dm & "' and t_tdd.jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") order by " & _Iif($cj_zd_mc <> "", "" & $cj_zd & " desc,", "") & "tdcj desc,ksh")
	$rs = $conn.execute("select ksh,xm,xbmc,tdcj" & _Iif($cj_zd_mc <> "", ",t_tdd." & $cj_zd, "") & ",tdzy,jhk1.zydh as zydh1,jhk1.zymc as zymc1,jhk2.zydh as zydh2,jhk2.zymc as zymc2,jhk3.zydh as zydh3,jhk3.zymc as zymc3,jhk4.zydh as zydh4,jhk4.zymc as zymc4,jhk5.zydh as zydh5,jhk5.zymc as zymc5,jhk6.zydh as zydh6,jhk6.zymc as zymc6,zyzytj,lxdh,zsyj,lqzy,jhk.zymc as lqzymc,kszt,t_tdd.tdyydm,tdyymc,t_tdd.bz from (((((((((t_tdd left join td_xbdm on td_xbdm.xbdm=t_tdd.xbdm) left join t_jhk as jhk1 on jhk1.zydh=t_tdd.zydh1 and jhk1.pcdm=t_tdd.pcdm and jhk1.kldm=t_tdd.kldm and jhk1.jhxz=t_tdd.jhxz and jhk1.tddw=t_tdd.tddw) left join t_jhk as jhk2 on jhk2.zydh=t_tdd.zydh2 and jhk2.pcdm=t_tdd.pcdm and jhk2.kldm=t_tdd.kldm and jhk2.jhxz=t_tdd.jhxz and jhk2.tddw=t_tdd.tddw) left join t_jhk as jhk3 on jhk3.zydh=t_tdd.zydh3 and jhk3.pcdm=t_tdd.pcdm and jhk3.kldm=t_tdd.kldm and jhk3.jhxz=t_tdd.jhxz and jhk3.tddw=t_tdd.tddw) left join t_jhk as jhk4 on jhk4.zydh=t_tdd.zydh4 and jhk4.pcdm=t_tdd.pcdm and jhk4.kldm=t_tdd.kldm and jhk4.jhxz=t_tdd.jhxz and jhk4.tddw=t_tdd.tddw) left join t_jhk as jhk5 on jhk5.zydh=t_tdd.zydh5 and jhk5.pcdm=t_tdd.pcdm and jhk5.kldm=t_tdd.kldm and jhk5.jhxz=t_tdd.jhxz and jhk5.tddw=t_tdd.tddw) left join t_jhk as jhk6 on jhk6.zydh=t_tdd.zydh6 and jhk6.pcdm=t_tdd.pcdm and jhk6.kldm=t_tdd.kldm and jhk6.jhxz=t_tdd.jhxz and jhk6.tddw=t_tdd.tddw) left join t_jhk as jhk on jhk.zydh=t_tdd.lqzy and jhk.pcdm=t_tdd.pcdm and jhk.kldm=t_tdd.kldm and jhk.jhxz=t_tdd.jhxz and jhk.tddw=t_tdd.tddw) left join TD_TDYYDM ON t_tdd.TDYYDM=TD_TDYYDM.TDYYDM)  where t_tdd.pcdm='" & $c_pc_dm & "' and t_tdd.kldm='" & $c_kl_dm & "' and t_tdd.jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") order by " & _Iif($cj_zd_mc <> "", "" & $cj_zd & " desc,", "") & "tdcj desc,ksh")
;~  ConsoleWrite("select ksh,xm,xbmc,tdcj" & _Iif($cj_zd_mc <> "", ",t_tdd." & $cj_zd, "") & ",tdzy,cstr(jhk1.zydh) as zydh1,jhk1.zymc as zymc1,cstr(jhk2.zydh) as zydh2,jhk2.zymc as zymc2,cstr(jhk3.zydh) as zydh3,jhk3.zymc as zymc3,cstr(jhk4.zydh) as zydh4,jhk4.zymc as zymc4,cstr(jhk5.zydh) as zydh5,jhk5.zymc as zymc5,cstr(jhk6.zydh) as zydh6,jhk6.zymc as zymc6,zyzytj,lxdh,zsyj,cstr(t_tdd.lqzy) as lqzy,jhk.zymc as lqzymc,kszt,cstr(t_tdd.tdyydm) as tdyydm,tdyymc,t_tdd.bz from (((((((((t_tdd left join td_xbdm on td_xbdm.xbdm=t_tdd.xbdm) left join t_jhk as jhk1 on jhk1.zydh=t_tdd.zydh1 and jhk1.pcdm=t_tdd.pcdm and jhk1.kldm=t_tdd.kldm and jhk1.jhxz=t_tdd.jhxz and jhk1.tddw=t_tdd.tddw) left join t_jhk as jhk2 on jhk2.zydh=t_tdd.zydh2 and jhk2.pcdm=t_tdd.pcdm and jhk2.kldm=t_tdd.kldm and jhk2.jhxz=t_tdd.jhxz and jhk2.tddw=t_tdd.tddw) left join t_jhk as jhk3 on jhk3.zydh=t_tdd.zydh3 and jhk3.pcdm=t_tdd.pcdm and jhk3.kldm=t_tdd.kldm and jhk3.jhxz=t_tdd.jhxz and jhk3.tddw=t_tdd.tddw) left join t_jhk as jhk4 on jhk4.zydh=t_tdd.zydh4 and jhk4.pcdm=t_tdd.pcdm and jhk4.kldm=t_tdd.kldm and jhk4.jhxz=t_tdd.jhxz and jhk4.tddw=t_tdd.tddw) left join t_jhk as jhk5 on jhk5.zydh=t_tdd.zydh5 and jhk5.pcdm=t_tdd.pcdm and jhk5.kldm=t_tdd.kldm and jhk5.jhxz=t_tdd.jhxz and jhk5.tddw=t_tdd.tddw) left join t_jhk as jhk6 on jhk6.zydh=t_tdd.zydh6 and jhk6.pcdm=t_tdd.pcdm and jhk6.kldm=t_tdd.kldm and jhk6.jhxz=t_tdd.jhxz and jhk6.tddw=t_tdd.tddw) left join t_jhk as jhk on jhk.zydh=t_tdd.lqzy and jhk.pcdm=t_tdd.pcdm and jhk.kldm=t_tdd.kldm and jhk.jhxz=t_tdd.jhxz and jhk.tddw=t_tdd.tddw) left join TD_TDYYDM ON t_tdd.TDYYDM=TD_TDYYDM.TDYYDM)  where t_tdd.pcdm='" & $c_pc_dm & "' and t_tdd.kldm='" & $c_kl_dm & "' and t_tdd.jhxz='" & $c_jhxz_dm & "' and ksh in(" & $c_pc_kl_jhxz_ksh & ") order by " & _Iif($cj_zd_mc <> "", "" & $cj_zd & " desc,", "") & "tdcj desc,ksh")

	Local $r_i = 0, $c_i = 0, $ii, $jj, $tdd_ks_sqls = "", $lqzy_xh = 0
	If $cj_zd_mc <> "" Then
		$oExcel.ActiveSheet.columns(5).insert
		$oExcel.ActiveSheet.columns(5).ColumnWidth = 4
		$oExcel.ActiveSheet.cells(2, 5).value = "专业成绩"
	EndIf
	If @error Then
		MsgBox(0, "错误", "投档单查询错误，有可能是数据库被锁引起，请重启院校子系统。")
		Return
	EndIf

	While Not ($rs.eof Or $rs.bof)
		$r_i += 1
;~ 	$tdd_ks_sqls&="INSERT INTO tdd_ks (tdd_ks_dm,ksh) VALUES ("&$tdd_ks_dm&",'"&$rs.Fields("ksh").Value&"')"&@CRLF
		;$tdd_ks_sqls&="INSERT INTO tdd_ks SELECT distinct tdd_ks_dm,'"&$rs.Fields("ksh").Value&"' as ksh FROM tdd_ks_info where tdd_ks_dm="&$tdd_ks_dm&" and NOT exists(select 1 FROM tdd_ks WHERE ksh='"&$rs.Fields("ksh").Value&"' and tdd_ks_dm="&$tdd_ks_dm&")"&@CRLF
		$oExcel.ActiveSheet.rows($r_i + 3).insert
;~ 	$oExcel.ActiveSheet.cells($r_i+3,1).value=$r_i
		$c_i = 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).Formula = "=row()-2"
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = StringStripWS($rs.Fields("ksh").Value, 2)
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = StringStripWS($rs.Fields("xm").Value, 2)
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = StringStripWS($rs.Fields("xbmc").Value, 2)
		If $cj_zd_mc <> "" Then
			$c_i += 1
			$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = Round($rs.Fields($cj_zd).Value, 3)
		EndIf
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = Round($rs.Fields("tdcj").Value, 3)
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = $rs.Fields("tdzy").Value
;~ 	最长49个字符
;~ 	BinaryLen(StringToBinary("a你好"));长度为5
		For $ii = 3 To 25
			$t = ""
			For $jj = 1 To 6
				$t &= StringLeft(binToStr($rs.Fields("zydh" & $jj).Value) & $rs.Fields("zymc" & $jj).Value, $ii)
				$t = _Iif(StringRight($t, 1) == ",", $t, $t & ",")
			Next
			If $t == "" Then ExitLoop
			If BinaryLen(StringToBinary($t)) >= 50 Then ExitLoop
		Next

		$t = ""
		For $jj = 6 To 2 Step -1
			$t = StringLeft(binToStr($rs.Fields("zydh" & $jj).Value) & $rs.Fields("zymc" & $jj).Value, $ii - 1) & $t
			$t = _Iif($t == "" Or StringLeft($t, 1) == ",", $t, "," & $t)
		Next
		$t = StringLeft(binToStr($rs.Fields("zydh1").Value) & $rs.Fields("zymc1").Value, 1 + (49 - BinaryLen(StringToBinary($t))) / 2) & _Iif($t <> ",", $t, "")
		$t = _Iif(StringLeft($t, 1) == ",", StringMid($t, 2), $t)

		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = $t
;~ 	是否服从调剂
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = _Iif($rs.Fields("zyzytj").Value == "1", "是", "否")
;~ 	StringRegExpReplace($rs.Fields("lxdh").Value,"\s+$","")
;~ 	$oExcel.ActiveSheet.cells($r_i+3,9).value=StringRegExpReplace($rs.Fields("lxdh").Value&_Iif(StringInStr($rs.Fields("lxdh").Value,"增加一个考生联系电话")>0,","&StringMid($rs.Fields("lxdh").Value,43),""),"\s","")
		$t = StringReplace($rs.Fields("zsyj").Value, " ", "")
		$t = _Iif(StringInStr($t, "增加一个考生联系电话") > 0, StringMid($t, 43), "")
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = StringReplace($rs.Fields("lxdh").Value & _Iif($t <> "", "," & $t, ""), " ", "")
;~ 	录取专业
		$lqzy_xh = 0
		If StringStripWS(binToStr($rs.Fields("lqzy").Value), 3) <> "?" Then
			For $jj = 1 To 6
				If binToStr($rs.Fields("zydh" & $jj).Value) == binToStr($rs.Fields("lqzy").Value) Then
					$lqzy_xh = $jj
					ExitLoop
				EndIf
			Next
		EndIf
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = _Iif(StringStripWS(binToStr($rs.Fields("lqzy").Value), 3) <> "?", StringLeft($lqzy_xh & StringStripWS($rs.Fields("lqzymc").Value, 2), 8), "")
;~ 	考生状态
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = _Iif(Number($rs.Fields("kszt").Value) >= 0 And Number($rs.Fields("kszt").Value) < UBound($a_kszt), $a_kszt[Number($rs.Fields("kszt").Value)], $rs.Fields("kszt").Value & "未知")
		$c_i += 1
		$oExcel.ActiveSheet.cells($r_i + 3, $c_i).value = StringReplace(StringLeft($rs.Fields("bz").Value, 8) & binToStr($rs.Fields("tdyydm").Value) & StringLeft($rs.Fields("tdyymc").Value, 8), " ", "")
		$rs.movenext
	WEnd


	$oExcel.ActiveSheet.rows($r_i + 1 + 3).delete
	$oExcel.ActiveSheet.rows(3).delete
;~ 2010年陕西师范大学山东艺术文本科提前批投档单
	$oExcel.ActiveSheet.cells(1, 1).value = @YEAR & "年陕西师范大学" & $c_sf_mc & $c_pc_mc & $c_kl_mc & $c_jhxz_mc & "投档单"

;~ If IsObj($oExcel.ActiveSheet.PageSetup) And IsString("$oExcel.ActiveSheet.PageSetup.RightFooter") Then
;~ 	$oExcel.ActiveSheet.PageSetup.RightFooter = "导出时间:"&$sj
;~ EndIf
	$oExcel.ActiveSheet.PageSetup.RightFooter = "导出时间:" & $sj


	_ExcelBookSave($oExcel)
	$oExcel.Application.DisplayAlerts = True
	$oExcel.Application.ScreenUpdating = True
	$oExcel.Visible = True

	#cs
		;~ 如果是最新的投档单.插入投档单考生信息
		If $o_tdd_is_old =False Then
		;~ 如果是最新的投档单.$tdd_ks_dm=max+1
		$rs_dbf=$conn_dbf.execute("select max(tdd_ks_dm) as max_dm from tdd_ks_info")
		If Not ($rs_dbf.eof Or $rs_dbf.bof) Then $tdd_ks_dm=Number($rs_dbf.Fields(0).Value)+1
		;记录投档单信息
		;~ 	ConsoleWrite("insert INTO tdd_ks_info (tdd_ks_dm,sf,pcdm,kldm,jhxz,sj) values ("&$tdd_ks_dm&",'"&$c_sf_mc&"','"&$c_pc_dm&"','"&$c_kl_dm&"','"&$c_jhxz_dm&"',{^"&$sj&"})")
		$conn_dbf.execute("insert INTO tdd_ks_info (tdd_ks_dm,sf,pcdm,kldm,jhxz,sj) values ("&$tdd_ks_dm&",'"&$c_sf_mc&"','"&$c_pc_dm&"','"&$c_kl_dm&"','"&$c_jhxz_dm&"','"&$sj&"')")
		
		$tdd_ks_sqls=StringSplit($tdd_ks_sqls,@CRLF,3)
		For $sql In $tdd_ks_sqls
		If StringLen($sql)>1 Then
		;~ 	ConsoleWrite($sql&@crlf)
		$conn_dbf.execute($sql)
		EndIf
		Next
		EndIf
	#ce

	If IsObj($conn) Then $conn.close
	If IsObj($conn_share) Then $conn_share.close
	If IsObj($conn_exe) Then $conn_exe.close
	If IsObj($conn_dbf) Then $conn_dbf.close
;~ 取消气泡
	TrayTip("取消泡泡提示", "", 1)
	MsgBox(0, "提示", "投档单生成成功！")
EndFunc   ;==>Output_Tdd

;~ 隐藏院校子系统上显示的考生,
;~ 参数$tdd_xh 为-1的显示所有考生
;~ 	0的时候,只显示最新投档单
;~ 	n的时候,只显示第n次投档单
Func HideKs($tdd_xh = 0)
	Local $c_yx_mc, $c_sf_mc, $c_pc_dm, $c_pc_mc, $c_kl_dm, $c_kl_mc, $c_jhxz_dm, $c_jhxz_mc
	Local $t, $i
	Local $conn, $conn_share, $conn_dbf, $conn_exe
	Local $rs, $rs_share, $rs_dbf, $rs_exe
	Local $db_dir

;~ 如果没有运行院校子系统过程退出
;~ If Not WinActivate("全国普通高校招生网上录取 - 院校子系统")Then	Return
;~ If Not FileExists($exe_dir & "\T_SSZT.DB") Then
;~ 	MsgBox(0, "错误", "数据表T_SSZT.DB不在" & @CRLF & "请检查：" & $exe_dir & "\T_SSZT.DB是否存在")
;~ 	Return
;~ EndIf

	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return

	;	WinActivate("全国普通高校招生网上录取 - 院校子系统")
;~ 	Send("!{Tab}{Tab}")
;~ 	Send("{ESC}")
;~ 	MsgBox(0,"",WinGetState("全国普通高校招生网上录取 - 院校子系统"))

;~ 枚举进程的窗体
	Local $NacuesC_wins = _WinAPI_EnumProcessWindows(ProcessExists($exe_name))
	For $i = 1 To $NacuesC_wins[0][0]
		If $NacuesC_wins[$i][1] <> "TNCMainForm" And $NacuesC_wins[$i][1] <> "TApplication" Then WinClose($NacuesC_wins[$i][0])
	Next
;~ _ArrayDisplay($NacuesC_wins, '_WinAPI_EnumProcessWindows')
	If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return

;~ 获得子系统筛选工具条上的值
	Local $statusbarhwnd = ControlGetHandle('[Class:TNCMainForm]', '', '[CLASS:TToolBar; INSTANCE:2]')
	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 0), " - ", 3)
	If @error Then Return

	$c_yx_mc = $t[0]
	$c_sf_mc = StringReplace($t[1], "录取进程", "")
;~ MsgBox(0, "",$c_sf_mc)
	If $c_sf_mc = "" Or $Cilp_last_sf <> $c_sf_mc Then
		MsgBox(0, "错误", "插件所存储的登录信息与院校子系统不一致" & @CRLF & "请使用插件自动填写登录信息再使用此功能." & @CRLF & "插件所存储的省份：" & $Cilp_last_sf & "")
;~ 	Return
		Return ;Exit
	EndIf


	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 1), " ", 3)
	$c_pc_dm = $t[0]
	$c_pc_mc = $t[1]
;~ MsgBox(0, "",$c_pc_mc)

;~ 获得科类代码及计划性质
	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 2), " - ", 3)
	$t = StringSplit($t[0], " ", 3)
	$c_kl_dm = $t[0]
	$c_kl_mc = $t[1]
;~ MsgBox(0, "",$c_kl_mc)
	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 2), " - ", 3)
	$t = StringSplit($t[1], " ", 3)
	$c_jhxz_dm = $t[0]
	$c_jhxz_mc = $t[1]
;~ MsgBox(0, "",$c_jhxz_mc)

	Local $r = 0
	Local $sf_py = "", $sf_dm = ""
	For $r = 0 To UBound($G_ss_py_dm, 1) - 1
		If $G_ss_py_dm[$r][0] = $Cilp_last_sf Then
			$sf_py = $G_ss_py_dm[$r][1]
			$sf_dm = $G_ss_py_dm[$r][2]
			ExitLoop
		EndIf
	Next



	#CS ;~ 检查文件是否存在
		For $file_name_t In StringSplit("t_tdd.db|t_jhk.db|TD_XBDM.DB", "|", 3)
		If Not FileExists($db_dir & "\" & $file_name_t) Then
		MsgBox(0, "", "数据表" & $file_name_t & "不存在" & @CR & "请检查：" & $db_dir & "\" & $file_name_t & "是否存在")
		;~ 		Return
		Return ;Exit
		EndIf
		Next
		
		
		$conn = ObjCreate("ADODB.Connection")
		$rs = ObjCreate("ADODB.Recordset")
		;$sql_where = "sf='" & StringReplace($sf, "'", "") & "' AND (user_class='" & $user_class & "' OR user_class=' ' ) AND is_default='1'"
		;$sql = "SELECT * from NacuesC_ncuzjc WHERE " & $sql_where & " ORDER BY user_class desc"
		;$conn.Open("Driver=Microsoft Visual FoxPro Driver;SourceType=DBF;SourceDB=" & @ScriptDir)
		$conn.Open("DRIVER={Driver do Microsoft Paradox (*.db )};DriverId=26;dbq=" & $db_dir)
	#CE

	Local $db_mdb = $data_dir & $sf_py & "\" & $Cilp_last_username & "\" & $data_university_mdb_filename
	Local $db_mdb_share = $data_dir & $sf_py & "\" & $Cilp_last_username & "\" & $data_university_share_mdb_filename
	If Not FileExists($db_mdb) Then
		MsgBox(0, "", "数据库" & $db_mdb & "不存在" & @CR & "请检查：" & $db_mdb & "是否存在")
		Return
	EndIf
	
	If Not FileExists($db_mdb_share) Then
		MsgBox(0, "", "数据库" & $db_mdb_share & "不存在" & @CR & "请检查：" & $db_mdb_share & "是否存在")
		Return
	EndIf
	
	$conn = _OpenConnToMdb($db_mdb)
	If Not IsObj($conn) Then
		MsgBox(0, "", "连接数据库" & $db_mdb & "失败。")
		Return
	EndIf

	$conn_share = _OpenConnToMdb($db_mdb_share)
	If Not IsObj($conn_share) Then
		MsgBox(0, "", "连接数据库" & $db_mdb_share & "失败。")
		Return
	EndIf

;~ 检查文件是否存在(控制线、投档单考生、投档单考生信息数据表,模板)
;~ $files_list=StringSplit("Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")
	For $file_name_t In StringSplit("czx.dbf|tdd_ks_info.dbf|tdd_ks.dbf|模板.xls", "|", 3)
		If Not FileExists(@ScriptDir & "\" & $file_name_t) Then
			MsgBox(0, "", "数据表" & $file_name_t & "不存在" & @CR & "请检查：" & @ScriptDir & "\" & $file_name_t & "是否存在")
;~ 		Return
			Return ;Exit
		EndIf
	Next

	$conn_dbf = ObjCreate("ADODB.Connection")
;~ $rs_dbf = ObjCreate("ADODB.Recordset")
	;$conn.Open("Driver=Microsoft Visual FoxPro Driver;SourceType=DBF;SourceDB=" & @ScriptDir)
	$conn_dbf.Open("Provider=Microsoft.Jet.OLEDB.4.0;dBase IV;HDR=NO;IMEX=2;DATABASE=" & @ScriptDir)
;~ $conn_dbf.execute("delete from czx")
;~ $rs_dbf=$conn_dbf.execute("select * from czx where sf='' and pcdm='' and kldm='' and jhxz=''")


;~ 	--全部移回0-9
;~ 	update t_tdd set jhxz=chr(asc(jhxz)+15) where (asc(jhxz) between 33 and 42) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
;~ 	--全部移回A-G
;~ 	update t_tdd set jhxz=chr(asc(jhxz)+ 7) where (asc(jhxz) between 58 and 64) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
	$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)+15) where (asc(jhxz) between 33 and 42) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")
	$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)+ 7) where (asc(jhxz) between 58 and 64) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")

;~ 负数的时候,显示所有考生,清除计划性质代码的加密
	If $tdd_xh < 0 Then

	Else
		Local $c_pc_kl_jhxz_ksh_last = "'#0000000000000'" ;已经生成考生号列表
		Local $c_pc_kl_jhxz_ksh = "'#0000000000000'" ;本投档单考生号列表
		Local $o_tdd_xh = 0 ;本次生成投档单序号
		Local $o_tdd_is_old = False
;~ @YEAR
		Local $date = @YEAR & "-" & @MON & "-" & @MDAY
		Local $sj = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
		Local $sj_str = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC

		Local $tdd_ks_dm = 1, $tdd_cnt = 0




;~ 得到当前批次科类性质的已经投档次数
		$rs_dbf = $conn_dbf.execute("select count(1) from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "'")
		If Not ($rs_dbf.eof Or $rs_dbf.bof) Then
			$tdd_cnt = $rs_dbf.Fields(0).Value
		Else
			Return
		EndIf
		;如果要取得的投档单序号小于等投档单数目
		If $tdd_xh > 0 And $tdd_xh <= $tdd_cnt Then
			$o_tdd_is_old = True
			$o_tdd_xh = $tdd_xh
;~ 得到当前批次科类性质的已经投档次数
			$rs_dbf = $conn_dbf.execute("select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' order by tdd_ks_dm")
			$i = 0
			While Not ($rs_dbf.eof Or $rs_dbf.bof)
				$i += 1
				If $i = $tdd_xh Then
					$tdd_ks_dm = $rs_dbf.Fields(0).Value
					ExitLoop
				EndIf
				$rs_dbf.movenext
			WEnd
			;本次考生号列表
			$c_pc_kl_jhxz_ksh = "'#0000000000000'"
			$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm =" & $tdd_ks_dm & "")
			While Not ($rs_dbf.eof Or $rs_dbf.bof)
				$c_pc_kl_jhxz_ksh &= ",'" & $rs_dbf.Fields(0).Value & "'"
				$rs_dbf.movenext
			WEnd
			;本次以前考生号列表
			$c_pc_kl_jhxz_ksh_last = "'#0000000000000'"
			$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm <" & $tdd_ks_dm & " And tdd_ks_dm in(select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "')")
			While Not ($rs_dbf.eof Or $rs_dbf.bof)
				$c_pc_kl_jhxz_ksh_last &= ",'" & $rs_dbf.Fields(0).Value & "'"
				$rs_dbf.movenext
			WEnd
		Else
			;获得已经生成过投档单的考生列表
			$c_pc_kl_jhxz_ksh_last = "'#0000000000000'"
			$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm in(select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "')")
			While Not ($rs_dbf.eof Or $rs_dbf.bof)
				$c_pc_kl_jhxz_ksh_last &= ",'" & $rs_dbf.Fields(0).Value & "'"
				$rs_dbf.movenext
			WEnd
;~ 查询是否有新考生
			$rs = $conn.execute("SELECT ksh FROM T_tdd where pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh not in(" & $c_pc_kl_jhxz_ksh_last & ")")
;~ 	有新投档考生
			If Not ($rs.eof Or $rs.bof) Then
				$o_tdd_xh = $tdd_cnt + 1
				$o_tdd_is_old = False

				#CS 		;~ 如果是最新的投档单.$tdd_ks_dm=max+1
					$rs_dbf=$conn_dbf.execute("select max(tdd_ks_dm) as max_dm from tdd_ks_info")
					If Not ($rs_dbf.eof Or $rs_dbf.bof) Then
					$tdd_ks_dm=Number($rs_dbf.Fields(0).Value)+1
					Else
					$tdd_ks_dm=1
					EndIf
				#CE

				Local $oRec = ObjCreate("ADODB.Recordset")
				If IsObj($oRec) = 0 Then Return SetError(2)
				With $oRec
					;$conn_dbf.execute("insert INTO tdd_ks_info (tdd_ks_dm,sf,pcdm,kldm,jhxz,sj) values ("&$tdd_ks_dm&",'"&$c_sf_mc&"','"&$c_pc_dm&"','"&$c_kl_dm&"','"&$c_jhxz_dm&"','"&$sj&"')")

					.Open("SELECT * FROM [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where tdd_ks_dm in(select max(tdd_ks_dm) as max_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "])", $conn_dbf, 0, 2)
					If Not ($oRec.eof Or $oRec.bof) Then
						$tdd_ks_dm = Number($oRec.Fields(0).Value) + 1
					Else
						$tdd_ks_dm = 1
					EndIf
					.AddNew
					.Fields.Item("tdd_ks_dm") = $tdd_ks_dm
					.Fields.Item("sf") = $c_sf_mc
					.Fields.Item("pcdm") = $c_pc_dm
					.Fields.Item("kldm") = $c_kl_dm
					.Fields.Item("jhxz") = $c_jhxz_dm
					.Fields.Item("sj") = $sj
					.Update
					.close

					.Open("SELECT tdd_ks_dm,ksh FROM tdd_ks where 1=2", $conn_dbf, 3, 3)

					;本次考生号列表
					$c_pc_kl_jhxz_ksh = "'#0000000000000'"
					While Not ($rs.eof Or $rs.bof)
						.AddNew
						.Fields.Item(0) = $tdd_ks_dm
						$c_pc_kl_jhxz_ksh &= ",'" & $rs.Fields(0).Value & "'"
						.Fields.Item(1) = $rs.Fields(0).Value
						$rs.movenext
					WEnd
					.Update
					.Close
				EndWith
			Else
				$o_tdd_xh = $tdd_cnt
				$o_tdd_is_old = True ;设置为使用最后一次的投档单
				$rs_dbf = $conn_dbf.execute("select max(tdd_ks_dm) as max_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "'")
				If Not ($rs_dbf.eof Or $rs_dbf.bof) Then $tdd_ks_dm = Number($rs_dbf.Fields(0).Value)
				;本次考生号列表
				$c_pc_kl_jhxz_ksh = "'#0000000000000'"
				$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm =" & $tdd_ks_dm & "")
				While Not ($rs_dbf.eof Or $rs_dbf.bof)
					$c_pc_kl_jhxz_ksh &= ",'" & $rs_dbf.Fields(0).Value & "'"
					$rs_dbf.movenext
				WEnd
				;本次以前考生号列表
				$c_pc_kl_jhxz_ksh_last = "'#0000000000000'"
				$rs_dbf = $conn_dbf.execute("select ksh from tdd_ks where tdd_ks_dm <" & $tdd_ks_dm & " And tdd_ks_dm in(select tdd_ks_dm from [" & StringRegExpReplace(FileGetShortName(@ScriptDir & "\tdd_ks_info.dbf"), ".*\\", "") & "] where sf='" & $c_sf_mc & "' and pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "')")
				While Not ($rs_dbf.eof Or $rs_dbf.bof)
					$c_pc_kl_jhxz_ksh_last &= ",'" & $rs_dbf.Fields(0).Value & "'"
					$rs_dbf.movenext
				WEnd
			EndIf

		EndIf

;~ If $c_pc_kl_jhxz_ksh="'#0000000000000'" Then
;~ 	MsgBox(0, "错误", "本次投档单人数为0.")
;~ 	Return ;Exit
;~ EndIf

;~ --移走0-9,最好限制pc,kl,jhxz
;~ update t_tdd set jhxz=chr(asc(jhxz)-15) where ksh not in ('11370103170001','11370203170002') and (asc(jhxz) between 48 and 57) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
;~ --移走A-G,最好限制pc,kl,jhxz
;~ update t_tdd set jhxz=chr(asc(jhxz)- 7) where ksh not in ('11370103170001','11370203170002') and (asc(jhxz) between 65 and 71) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))
;~ ConsoleWrite($c_pc_kl_jhxz_ksh)
;~ ConsoleWrite(@CRLF)

		$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)-15) where pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh not in(" & $c_pc_kl_jhxz_ksh & ") and (asc(jhxz) between 48 and 57) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")
		$conn.execute("update t_tdd set jhxz=chr(asc(jhxz)- 7) where pcdm='" & $c_pc_dm & "' and kldm='" & $c_kl_dm & "' and jhxz='" & $c_jhxz_dm & "' and ksh not in(" & $c_pc_kl_jhxz_ksh & ") and (asc(jhxz) between 65 and 71) and not exists(select * from t_jhk where (asc(jhxz) between 33 and 42) or (asc(jhxz) between 58 and 64))")


	EndIf;tdd_xh<0

	#CS
		;新版刷新投档单
		;通过修改显示的列项目来重新载入数据,先添加一显示列,再取消该显示列.打开两次"选项"对话框
		If IsObj($conn) Then $conn.close
		If IsObj($conn_share) Then $conn_share.close
		If IsObj($conn_exe) Then $conn_exe.close
		If IsObj($conn_dbf) Then $conn_dbf.close
		
		closeOtherWin()
		;For $i = 1 To 2
		If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return
		Send("!vo")
		If Not WinWaitActive("[TITLE:选项; CLASS:TOptionsForm;]", "", 3) Then Return ;打开"选项"对话框
		;WinSetState("[TITLE:选项; CLASS:TOptionsForm;]", "",@SW_HIDE)
		ControlSend("[TITLE:选项; CLASS:TOptionsForm; ]", "", "TCoolListDBExt2", "{end}") ;发送选择末尾项.
		ControlClick("[TITLE:选项; CLASS:TOptionsForm; ]", "", "TBitBtn4") ;点击向右图标,添加显示列
		ControlClick("[TITLE:选项; CLASS:TOptionsForm; ]", "", "TButton3") ;点击"确认"按钮;列表会刷新.
		
		If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return
		Send("!vo")
		If Not WinWaitActive("[TITLE:选项; CLASS:TOptionsForm;]", "", 3) Then Return ;打开"选项"对话框
		;WinSetState("[TITLE:选项; CLASS:TOptionsForm;]", "",@SW_HIDE)
		ControlSend("[TITLE:选项; CLASS:TOptionsForm; ]", "", "TCoolListDBExt1", "{end}") ;发送选择末尾项.
		ControlClick("[TITLE:选项; CLASS:TOptionsForm; ]", "", "TBitBtn3") ;点击向左图标,减少显示列
		ControlClick("[TITLE:选项; CLASS:TOptionsForm; ]", "", "TButton3") ;点击"确认"按钮;列表会刷新.
		;Next
	#CE


	
;~ 刷新投档单
	Local $xh = 0
	Local $is_refresh_pc = False
	
	$rs = $conn.execute("select count(1) from (select  kldm,jhxz  from t_jhk where pcdm='" & $c_pc_dm & "' group by kldm,jhxz) a")
	If ($rs.eof Or $rs.bof) Then Return
	;如果有2个以上的科类、性质,刷新科类就行了....
	If $rs.Fields(0).Value >= 2 Then
		$rs = $conn.execute("select count(1) from (select  kldm+'_'+jhxz as kldmjhxz  from t_jhk where pcdm='" & $c_pc_dm & "' group by kldm,jhxz) a where kldmjhxz<='" & $c_kl_dm & "_" & $c_jhxz_dm & "'")
	Else
		$is_refresh_pc = True
		$rs = $conn.execute("select count(1) from t_jhk where pcdm<='" & $c_pc_dm & "'")
	EndIf
	If ($rs.eof Or $rs.bof) Then Return
	$xh = $rs.Fields(0).Value
	
	If IsObj($conn) Then $conn.close
	If IsObj($conn_share) Then $conn_share.close
	If IsObj($conn_exe) Then $conn_exe.close
	If IsObj($conn_dbf) Then $conn_dbf.close
	
	;MsgBox(0, "", $xh)
	
;~    如果不是可见,退出
	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return
;~    激活前置失败,退出
	If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return
	WinSetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]", "", @SW_MAXIMIZE)
	;Sleep(300)
	
	
	
	
	While BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) And (Not WinActive("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"))
		WinWaitActive("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]", "", 1)
	WEnd
	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return
	
	Local $s_t
	
	Send("^{end}")
	$s_t = StringRegExpReplace(_GUICtrlStatusBar_GetText(ControlGetHandle('[Class:TNCMainForm]', '', 'TStatusBar1'), 2), "\D", "")
	ConsoleWrite($s_t & @CRLF)
	
	Send("!f")
	If Not $is_refresh_pc Then Send("{down}")
	Send("{enter}");打开选择菜单
	Sleep(30)
	If $xh = 1 Then Send("{down}");如果当前是第一个项目,多发送一个向下
	Send("{down}{enter}");

	;ConsoleWrite(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 2)&@crlf);
	;ConsoleWrite(@CRLF)
	;ConsoleWrite(WinGetText("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]", "") & @CRLF)
	
	For $i = 1 To 20
		Sleep(50)
		;考生数为0
		$s_t = StringRegExpReplace(_GUICtrlStatusBar_GetText(ControlGetHandle('[Class:TNCMainForm]', '', 'TStatusBar1'), 1), "\D", "")
		If $s_t = "0" Then ExitLoop
		
		;或 第1个考生
		$s_t = StringRegExpReplace(_GUICtrlStatusBar_GetText(ControlGetHandle('[Class:TNCMainForm]', '', 'TStatusBar1'), 2), "\D", "")
		ConsoleWrite($s_t & @CRLF)
		If $s_t = "1" Then ExitLoop
	Next
	
	;Sleep(1000)
	
	While BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) And (Not WinActive("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"))
		WinWaitActive("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]", "", 1)
	WEnd
	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return
	
	Send("!f")
	If Not $is_refresh_pc Then Send("{down}")
	Send("{enter}");打开选择菜单
	Sleep(30)
	For $i = 1 To $xh
		Send("{down}");
	Next
	Send("{enter}");



;~ $rs=$conn.execute("select pcdm,kldm,jhxz from t_jhk group by pcdm,kldm,jhxz where pcdm='"&$c_pc_dm&"'")


EndFunc   ;==>HideKs





Func output_data()

	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return
	;	WinActivate("全国普通高校招生网上录取 - 院校子系统")
;~ 	Send("!{Tab}{Tab}")
;~ 	Send("{ESC}")
;~ 	MsgBox(0,"",WinGetState("全国普通高校招生网上录取 - 院校子系统"))

;~ 枚举进程的窗体
	Local $NacuesC_wins = _WinAPI_EnumProcessWindows(ProcessExists($exe_name))
	For $i = 1 To $NacuesC_wins[0][0]
		If $NacuesC_wins[$i][1] <> "TNCMainForm" And $NacuesC_wins[$i][1] <> "TApplication" Then WinClose($NacuesC_wins[$i][0])
	Next
;~ _ArrayDisplay($NacuesC_wins, '_WinAPI_EnumProcessWindows')


;~ Exit
	;Send("{esc}")
;~ Sleep(200)
	If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return
;~ Send("^{home}")

;~ 获得子系统筛选工具条上的值
	Local $statusbarhwnd = ControlGetHandle('[Class:TNCMainForm]', '', '[CLASS:TToolBar; INSTANCE:2]')
	Local $t = ""
	$t = StringSplit(_GUICtrlToolbar_GetButtonText($statusbarhwnd, 0), " - ", 3)
	If @error Then Return

	Local $c_sf_mc = StringReplace($t[1], "录取进程", "")
	$t = StringSplit($t[0], " ", 3)
	Local $c_yx_mc = $t[1]

;~ MsgBox(0, "",$c_sf_mc)
	If $c_sf_mc = "" Or $Cilp_last_sf <> $c_sf_mc Then
		MsgBox(0, "错误", "插件所存储的登录信息与院校子系统不一致" & @CRLF & "请使用插件自动填写登录信息再使用此功能." & @CRLF & "插件所存储的省份：" & $Cilp_last_sf & "")
;~ 	Return
		Return ;Exit
	EndIf



	Local $sj = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC
	Local $sj_str = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC


	Local $r = 0
	Local $sf_dm = ""
	For $r = 0 To UBound($G_ss_py_dm, 1) - 1
		If $G_ss_py_dm[$r][0] = $c_sf_mc Then
			$sf_dm = $G_ss_py_dm[$r][2]
			ExitLoop
		EndIf
	Next

	Local $output_dir = @ScriptDir & "\data备份\" & $sf_dm & "" & $c_sf_mc & "\" & $Cilp_last_username & $c_yx_mc & "\" & $Cilp_last_username & $c_yx_mc & $sj_str & "\"
	If Not FileExists($output_dir) Then DirCreate($output_dir)

	If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return
	Send("!fs")
	If Not WinWaitActive("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", 2) Then Return

	;点击全选
	ControlClick("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TButton2")
	;设置导出目录
;~ 	ControlSetText("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TEdit1", $output_dir&"tdd.dbf")
	
	;最多尝试5次填写路径
	For $i = 1 To 5
		WinClose("[TITLE:另存为;CLASS:#32770;]")
;~ 	点“。。”选择文件位置， [CLASS:TButton; INSTANCE:1]
		ControlClick("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TButton1")
		
		If Not WinWaitActive("[TITLE:另存为;CLASS:#32770;]", "", 3) Then Return
		;Send("{ENTER}")
		Sleep($i * 20)
		;ControlSetText("[TITLE:另存为;CLASS:#32770;]", "", "Edit1", $output_dir&"tdd.dbf")
		;MsgBox(0,"","设置好了吧"&@CRLF&$output_dir&"tdd.dbf")
		ControlSetText("[TITLE:另存为;CLASS:#32770;]", "", "Edit1", $output_dir & "tdd.dbf")
		;	MsgBox(0,"","设置好了吧"&@CRLF&$output_dir&"tdd.dbf")
		ControlClick("[TITLE:另存为;CLASS:#32770;]", "", "Button2")
		If Not WinWaitActive("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", 3) Then Return
		If ControlGetText("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TEdit1") == ($output_dir & "tdd.dbf") Then ExitLoop
	Next

	If Not ControlGetText("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TEdit1") == ($output_dir & "tdd.dbf") Then Return
	;ControlSetText("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TEdit1", $output_dir & "tdd.dbf")

	;设置可用
;~ 	ControlEnable("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TButton6")
	;点击全选
	ControlClick("[TITLE:导出投档单;CLASS:TDumpToFileForm;]", "", "TButton6")


	
	$Cilp_output_dir_last = $output_dir
	If Not WinWaitActive("[TITLE:导出投档单;CLASS:#32770;]", "", 30) Then Return
	WinClose("[TITLE:导出投档单;CLASS:#32770;]")
	If Not @error Then
		If MsgBox(4, "数据已经备份,是否打开", "数据已经备份在:" & @CRLF & $output_dir & @CRLF & "点 {是} 打开备份的目录", 8) == 6 Then Run(@ComSpec & " /c start explorer /e,, """ & $output_dir & """", "", @SW_HIDE)
	EndIf
EndFunc   ;==>output_data



Func converDb()
	Local $db_dir, $conn, $rs, $oExcel
	Local $conn_exe, $rs_exe

;~ 如果没有运行院校子系统过程退出
	If Not BitAND(WinGetState("[TITLE:全国普通高校招生网上录取 - 院校子系统;CLASS:TNCMainForm;]"), 2) Then Return
	If Not WinActivate("全国普通高校招生网上录取 - 院校子系统") Then Return

	#CS 	If Not FileExists($exe_dir & "\T_SSZT.DB") Then
		MsgBox(0, "错误", "数据表T_SSZT.DB不在" & @CRLF & "请检查：" & $exe_dir & "\T_SSZT.DB是否存在")
		;~ 	Return
		Return ;Exit
		EndIf
	#CE
	If $Cilp_last_sf = "" Then Return
	Local $r = 0
	Local $sf_py = "", $sf_dm = ""
	For $r = 0 To UBound($G_ss_py_dm, 1) - 1
		If $G_ss_py_dm[$r][0] = $Cilp_last_sf Then
			$sf_py = $G_ss_py_dm[$r][1]
			$sf_dm = $G_ss_py_dm[$r][2]
			ExitLoop
		EndIf
	Next


	Local $db_mdb = $data_dir & $sf_py & "\" & $Cilp_last_username & "\" & $data_university_mdb_filename
	Local $db_mdb_share = $data_dir & $sf_py & "\" & $Cilp_last_username & "\" & $data_university_share_mdb_filename
	If Not FileExists($db_mdb) Then
		MsgBox(0, "", "数据库" & $db_mdb & "不存在" & @CR & "请检查：" & $db_mdb & "是否存在")
		Return
	EndIf
	
	If Not FileExists($db_mdb_share) Then
		MsgBox(0, "", "数据库" & $db_mdb_share & "不存在" & @CR & "请检查：" & $db_mdb_share & "是否存在")
		Return
	EndIf

	$oExcel = _ExcelBookNew(0)
	;$oExcel.ActiveSheet.Range("A1").Value="'00123"
	#CS
		With $oExcel.ActiveSheet.QueryTables.Add(StringSplit( _
		"OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Password="""";User ID=Admin;Data Source=" & $db_mdb & ";Mode=Share Deny Write;Extende" _
		& "|" & _
		"d Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=82;" _
		& "|" & _
		"Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Databas" _
		& "|" & _
		"e Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=Fal" _
		& "|" & "se;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False", "|", 3), _
		$oExcel.Range("A1"))
	#CE
	#CS "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & $db_mdb & ";" _
			, _
			"Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Datab" _
			, _
			"ase Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";J" _
			, _
			"et OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Co" _
			, "mpact Without Replica Repair=False;Jet OLEDB:SFP=False")
		.CommandType = xlCmdSql
	#CE
	
	With $oExcel.ActiveSheet.QueryTables.Add(StringSplit( _
			"OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & $db_mdb & ";" _
			 & "|" & _
			"Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Datab" _
			 & "|" & _
			"ase Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";J" _
			 & "|" & _
			"et OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Co" _
			 & "|" & "mpact Without Replica Repair=False;Jet OLEDB:SFP=False", "|", 3), _
			$oExcel.Range("A1"))
;~         .CommandType = xlCmdTable=3
;~         .CommandType = xlCmdDefault=4
;~         .CommandType = xlCmdSql=2
		.CommandType = 4
		.CommandText = "select * from t_tdd where 1=1 order by pcdm,kldm,jhxz,tdcj desc,ksh"
;~         .Name = "T_JHK.DB"
		.FieldNames = True
		.RowNumbers = False
		.FillAdjacentFormulas = False
		.PreserveFormatting = True
		.RefreshOnFileOpen = False
		.BackgroundQuery = True
;~         .RefreshStyle = $oExcel.xlInsertDeleteCells
		.RefreshStyle = 1
		.SavePassword = False
		.SaveData = True
		.AdjustColumnWidth = True
		.RefreshPeriod = 0
		.PreserveColumnInfo = True
;~         .SourceDataFile = "E:\NacuesC\Jiangxi\1040300\T_JHK.DB"
		.Refresh(False)
	EndWith

;~ 	If FileExists($db_dir & "\T_TDDZDSX.DB") Then
	If $oExcel.ActiveWorkbook.Sheets.Count < 2 Then
		$oExcel.ActiveWorkbook.Sheets.Add
	Else
		$oExcel.ActiveWorkbook.Sheets(2).Select()
	EndIf
;~ $oExcel.ActiveWorkbook.Sheets.Add
	With $oExcel.ActiveSheet.QueryTables.Add(StringSplit( _
			"OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & $db_mdb & ";" _
			 & "|" & _
			"Mode=Read;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Datab" _
			 & "|" & _
			"ase Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";J" _
			 & "|" & _
			"et OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Co" _
			 & "|" & "mpact Without Replica Repair=False;Jet OLEDB:SFP=False", "|", 3), _
			$oExcel.Range("A1"))
;~         .CommandType = xlCmdTable=3
;~         .CommandType = xlCmdDefault=4
;~         .CommandType = xlCmdSql=2
		.CommandType = 4
		.CommandText = "select * from T_TDDZDSX"
;~         .Name = "T_JHK.DB"
		.FieldNames = True
		.RowNumbers = False
		.FillAdjacentFormulas = False
		.PreserveFormatting = True
		.RefreshOnFileOpen = False
		.BackgroundQuery = True
;~         .RefreshStyle = $oExcel.xlInsertDeleteCells
		.RefreshStyle = 1
		.SavePassword = False
		.SaveData = True
		.AdjustColumnWidth = True
		.RefreshPeriod = 0
		.PreserveColumnInfo = True
;~         .SourceDataFile = "E:\NacuesC\Jiangxi\1040300\T_JHK.DB"
		.Refresh(False)
	EndWith
	$oExcel.ActiveWorkbook.Sheets(1).Select()
;~ 	EndIf
	$oExcel.Visible = True
EndFunc   ;==>converDb

Func MyErrFunc()

	MsgBox(0, "错误提示", "We intercepted a COM Error !" & @CRLF & @CRLF & _
			"err.description is: " & @TAB & $oMyError.description & @CRLF & _
			"err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
			"err.number is: " & @TAB & Hex($oMyError.number, 8) & @CRLF & _
			"err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
			"err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
			"err.source is: " & @TAB & $oMyError.source & @CRLF & _
			"err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
			"err.helpcontext is: " & @TAB & $oMyError.helpcontext)
	Local $err = $oMyError.number
	If $err = 0 Then $err = -1
	SetError($err)
	; to check for after this function returns
EndFunc   ;==>MyErrFunc


#CS
	Func _adoProvider()
	Local $oProvider = "Microsoft.Jet.OLEDB.4.0; "
	Local $objCheck = ObjCreate("Access.application")
	If IsObj($objCheck) Then
	Local $oVersion = $objCheck.Version
	If StringLeft($oVersion, 2) == "12" Then $oProvider="Microsoft.ACE.OLEDB.12.0; "
	EndIf
	Return $oProvider
	EndFunc
	Global Const $adoProvider = _adoProvider()
#CE

Func _OpenRecordset()
	Local $oRec = ObjCreate("ADODB.Recordset")
	Return $oRec
EndFunc   ;==>_OpenRecordset

Func _OpenConnToMdb($mdb_file)
	Local $oADO = ObjCreate("ADODB.Connection")
	$oADO.Provider = $mdb_adoProvider
	$oADO.Open($mdb_file)
	Return $oADO
EndFunc   ;==>_OpenConnToMdb
#CS Func _OpenConnToMdb($mdb_file)
	Local $oADO = ObjCreate("ADODB.Connection")
	;   	$oADO.Provider = $mdb_adoProvider
	$oADO.Open("Driver={Microsoft Access Driver (*.mdb)};Dbq="&$mdb_file)
	Return $oADO
	EndFunc
#CE


Func binToStr($s)
	Return _Iif(IsBinary($s), BinaryToString($s, 2), $s)
EndFunc   ;==>binToStr

Func closeOtherWin()
;~ 枚举进程的窗体
	Local $NacuesC_wins = _WinAPI_EnumProcessWindows(ProcessExists($exe_name))
	For $i = 1 To $NacuesC_wins[0][0]
		If $NacuesC_wins[$i][1] <> "TNCMainForm" And $NacuesC_wins[$i][1] <> "TApplication" Then WinClose($NacuesC_wins[$i][0])
	Next
EndFunc   ;==>closeOtherWin
