package main

import (
	"fmt"
	"time"

	"github.com/go-vgo/robotgo"
)

func main() {
	// 定义省份、用户名和密码
	prov := "江苏省"
	username := "1040300"
	password := "MGRecC92"

	// 等待院校子系统登录窗口出现
	time.Sleep(3 * time.Second)
	robotgo.Type(username) // 输入用户名
	robotgo.Type(password) // 输入密码

	// 点击登录按钮，这里假设登录按钮的ControlID是"TButton1"

	// 移动鼠标到按钮位置并点击
	robotgo.MoveMouse(222, 222)
	robotgo.MouseClick("left", false)

	// 获取省份下拉框中的所有选项并打印
	// if err := robotgo.WaitActive("登录系统 2020.5", "用户信息"); err == nil {
	// 	// robotgo.Type(prov) // 输入省份
	// 	// robotgo.Type("\b") // 删除省份输入，因为RobotGo没有SelectString方法

	// 	// 打印所有选项（RobotGo没有直接获取下拉框选项的方法，需要使用Windows API或其他方法）
	// 	// 这里省略了打印所有选项的代码，因为RobotGo不直接支持获取下拉框选项列表

	// 	robotgo.Type(username) // 输入用户名
	// 	robotgo.Type(password) // 输入密码

	// 	// 点击登录按钮，这里假设登录按钮的ControlID是"TButton1"

	// 	// 移动鼠标到按钮位置并点击
	// 	robotgo.MoveMouse(x, y)
	// 	robotgo.MouseClick("left", false)

	// } else {
	// 	fmt.Println("登录窗口未激活")
	// 	return
	// }

	fmt.Println("登录信息已提交")
}
