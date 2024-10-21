package main

import (
	"fmt"
	"os/exec"
	"runtime"
	"strings"
	"time"

	"github.com/eyasliu/desktop"
	"github.com/eyasliu/desktop/tray"
)

//go:embed dog.ico
// var dogIco []byte
// var FallbackPage string = `<h1>出错了</h1>`

func CreateAndRunApp() {
	var app desktop.WebView
	var appTray *tray.Tray

	// 创建一个带有复选框的菜单项
	checkedMenu := &tray.TrayItem{
		Title:    "有勾选状态菜单项",
		Checkbox: true,
		Checked:  true,
	}
	checkedMenu.OnClick = func() {
		checkedMenu.Checked = !checkedMenu.Checked
		checkedMenu.Title = "未勾选"
		if checkedMenu.Checked {
			checkedMenu.Title = "已勾选"
		}
		checkedMenu.Update()
	}

	// 创建系统托盘
	appTray = &tray.Tray{
		Title:   "托盘演示",
		Tooltip: "提示文字，点击激活显示窗口",
		OnClick: func() {
			app.Show() // 显示窗口
		},
		Items: []*tray.TrayItem{
			checkedMenu,
			{
				Title: "修改托盘图标和文字",
				OnClick: func() {
					appTray.SetIconBytes(dogIco)
					appTray.SetTooltip("这是设置过后的托盘提示文字")
				},
			},
			{
				Title:   "触发错误页",
				OnClick: func() { app.Navigate("https://abcd.efgh.ijkl") },
			},
			{
				Title: "打开本地页面",
				OnClick: func() {
					app.SetHtml(`<h1>这是个本地页面</h1>
				<div style="-webkit-app-region: drag">设置css： -webkit-app-region: drag 可移动窗口</div>`)
				},
			},
			{
				Title: "JS 交互",
				Items: []*tray.TrayItem{
					{
						Title:   "执行 alert('hello')",
						OnClick: func() { app.Eval("alert('hello')") },
					},
					{
						Title: "显示所有桌面应用",
						OnClick: func() {
							fmt.Println("开始获取桌面应用列表")

							// 使用 PowerShell 获取所有桌面应用
							cmd := exec.Command("powershell", "-Command", `
								Get-StartApps | Where-Object { $_.AppID -notlike 'Microsoft.*' } | Select-Object -ExpandProperty Name
							`)
							output, err := cmd.Output()
							fmt.Println(string(output))
							if err != nil {
								fmt.Printf("执行 PowerShell 命令出错: %v\n", err)
								app.Eval(fmt.Sprintf("console.error('获取应用列表失败: %s'); alert('获取应用列表失败，请查看控制台');", err))
								return
							}

							fmt.Println("PowerShell 命令执行成功")

							// 处理输出
							apps := strings.Split(strings.TrimSpace(string(output)), "\n")
							appList := strings.Join(apps, "\n")

							// 准备用于 JavaScript 的应用列表字符串
							jsAppList := strings.Replace(appList, "\n", "\\n", -1)
							jsAppList = strings.Replace(jsAppList, "'", "\\'", -1) // 转义单引号

							// 显示应用列表
							app.Eval(fmt.Sprintf("alert('应用列表: %s')", jsAppList))

							// 显示应用数量
							app.Eval(fmt.Sprintf("alert('应用数量: %d')", len(apps)))

							// 创建一个用空格拼接的应用名称字符串
							appNamesString := strings.Join(apps, " ")

							// 显示所有应用名称（用空格拼接）
							app.Eval(fmt.Sprintf("alert('所有应用名称: %s')", appNamesString))

							fmt.Println("JavaScript 执行完毕")
							fmt.Println(appNamesString)
						},
					},
					{
						Title: "每次进入页面执行alert",
						OnClick: func() {
							app.Init("alert('每次进入页面都会执行一次')")
						},
					},
					{
						Title: "调用Go函数",
						OnClick: func() {
							app.Eval(`golangFn('tom').then(s => alert(s))`)
						},
					},
				},
			},
			{
				Title: "窗口操作",
				Items: []*tray.TrayItem{
					{
						Title: "无边框打开新窗口 im.qq.com",
						OnClick: func() {
							go func() {
								wpsai := desktop.New(&desktop.Options{
									StartURL:  "https://im.qq.com",
									Center:    true,
									Frameless: true, // 去掉边框
								})
								wpsai.Run()
							}()
						},
					},
					{
						Title: "显示窗口",
						OnClick: func() {
							app.Show()
						},
					},
					{
						Title: "隐藏窗口",
						OnClick: func() {
							app.Hide()
						},
					},
					{
						Title: "设置窗口标题",
						OnClick: func() {
							app.SetTitle("这是新的标题")
						},
					},
				},
			},
			{
				Title: "退出程序",
				OnClick: func() {
					app.Destroy()
				},
			},
		},
	}

	app = desktop.New(&desktop.Options{
		Debug:             true,
		AutoFocus:         true,
		Width:             1280,
		Height:            768,
		HideWindowOnClose: true,
		Center:            true,
		Title:             "basic 演示",
		StartURL:          "https://www.wps.cn",
		Tray:              appTray,
		DataPath:          "C:\\Users\\TK\\.envokv2\\webview2", // 请使用你的用户名替换 TK
		FallbackPage:      FallbackPage,
	})

	// 绑定 Go 函数到 JavaScript
	app.Bind("golangFn", func(name string) string {
		return fmt.Sprintf(`Hello %s, GOOS=%s`, name, runtime.GOOS)
	})

	// 等待应用程序准备就绪后显示窗口并导航
	go func() {
		<-time.After(100 * time.Millisecond)
		app.Show()
		app.Navigate("https://www.wps.cn")
	}()

	// 运行应用程序
	app.Run()
}
