这三个工具的获取方式如下：

1. Inspect.exe（Windows SDK 的一部分）:

   - 下载并安装 Windows SDK（Windows Software Development Kit）。
   - 访问 Microsoft 的 Windows SDK 下载页面：https://developer.microsoft.com/en-us/windows/downloads/windows-sdk/
   - 下载并运行安装程序。
   - 在安装选项中，你可以只选择 "Debugging Tools for Windows" 来获取 Inspect.exe。
   - 安装完成后，Inspect.exe 通常位于 `C:\Program Files (x86)\Windows Kits\10\bin\<version>\x64\` 目录下。

2. UISpy.exe（较旧版本的 Windows SDK）:

   - UISpy.exe 是较旧的工具，在新版本的 Windows SDK 中已被 Inspect.exe 替代。
   - 如果你仍然需要 UISpy.exe，你可能需要下载旧版本的 Windows SDK（如 Windows 7 SDK）。
   - 然而，不建议使用这个过时的工具，除非你有特殊需求。

3. Accessibility Insights for Windows:

   - 这是一个由 Microsoft 开发的现代化辅助功能测试工具。
   - 访问 Microsoft Store 并搜索 "Accessibility Insights for Windows"。
   - 或者直接使用这个链接：https://www.microsoft.com/store/productId/9PDZW0BQVJLR
   - 点击 "获取" 或 "安装" 按钮来下载和安装应用。

对于 UI 自动化任务，Accessibility Insights for Windows 是最新和最推荐的工具。它提供了更现代的界面和更多的功能。

如果你只是需要基本的 UI 检查功能，Windows 自带的 "检查" 工具也可以使用：

1. 按 Windows 键 + R 打开运行对话框。
2. 输入 "ms-settings:easeofaccess-narrator" 并按回车。
3. 在打开的设置页面中，找到并启用 "检查" 功能。

启用后，你可以使用 Ctrl + Windows 键 + Enter 来激活检查工具，然后悬停在 UI 元素上来查看其属性。

这些工具都能帮助你检查应用程序的 UI 结构和属性，对于开发自动化脚本非常有用。