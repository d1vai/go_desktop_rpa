#include <GUIConstants.au3>

; 定义省份、用户名和密码
Global $sProvince = "江苏省"
Global $sUsername = "1040300"
Global $sPassword = "MGRecC92"

; 等待院校子系统登录窗口出现
WinWaitActive("登录系统 2020.5", "选择省、自治区或直辖市", 3)

; 获取省份下拉框中的所有选项
Local $hWnd = WinGetHandle("登录系统 2020.5", "选择省、自治区或直辖市")
Local $hComboBox = ControlGetHandle($hWnd, "", "TComboBox1")
;~ ConsoleWrite("省份选项 " ": " & ControlCommand($hComboBox, "GetListItem", $i) & @CRLF)
;~ Local $iCount = ControlCommand($hComboBox, "GetListCount")

;~ ; 打印所有选项
;~ For $i = 0 To $iCount - 1
;~     ConsoleWrite("省份选项 " & $i & ": " & ControlCommand($hComboBox, "GetListItem", $i) & @CRLF)
;~ Next

; 输入省份
ControlCommand("登录系统 2020.5", "选择省、自治区或直辖市", "TComboBox1", "SelectString", $sProvince)

; 输入用户名
ControlSetText("登录系统 2020.5", "选择省、自治区或直辖市", "TEdit2", $sUsername)

; 输入密码
ControlSetText("登录系统 2020.5", "选择省、自治区或直辖市", "TEdit1", $sPassword)

; 点击登录按钮，这里假设登录按钮的ControlID是"TButton1"
ControlClick("登录系统 2020.5", "选择省、自治区或直辖市", "TButton1")

; 以下是一些错误处理和等待操作，确保操作的稳定性
If @error Then
    MsgBox(0, "错误", "登录失败，请检查输入的信息是否正确，或者登录窗口是否已经打开。")
    Exit
EndIf

MsgBox(0, "成功", "登录信息已提交。")