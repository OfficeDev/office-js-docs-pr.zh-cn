---
title: 使用 Windows 上的 Microsoft Edge WebView2 （基于 Chromium）调试加载项
description: 了解如何在 VS 代码中使用适用于 Microsoft Edge 扩展的调试器来调试使用 Microsoft Edge WebView2（基于 Chromium）的 Office 加载项。
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 0908bb5040b49568006324600acacb5e36dbd1a5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238112"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>使用 Windows 上的 Microsoft Edge Chromium WebView2 调试加载项

在 Windows 上正在运行的 Office 加载项可以使用 VS 代码中适用于 Microsoft Edge 扩展的调试器来对 Edge Chromium WebView2 运行时进行调试。

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）
- [Node.js （版本 10+）](https://nodejs.org/)
- Windows 10
- [ 适用于 Windows Insiders 的 Microsoft Edge Chromium](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

1. 使用 [ 适用于 Office 加载项的 Yeoman 生成器 ](https://github.com/OfficeDev/generator-office) 创建项目。可以使用我们的任何一个快速入门指南，例如 [Outlook 加载项快速入门 ](../quickstarts/outlook-quickstart.md)，以做到这一点。

> [!TIP]
> 如果没有使用基于 Yeoman 生成器的加载项，需要调整一个注册表项。 在你的项目根目录下，在命令行中运行以下命令： `office-add-in-debugging start <your manifest path>`。

2. 在 VS 代码中打开项目。 在 VS 代码中，选择 **CTRL + SHIFT + X** 打开扩展栏。 搜索“适用于 Microsoft Edge 的调试器”扩展并安装。

3. 在你的项目 **.vscode** 文件夹中打开 **launch.json** 文件。 将以下代码添加到配置节：

```JSON
  {
      "name": "Debug Office Add-in (Edge Chromium)",
      "type": "edge",
      "request": "attach",
      "useWebView": "advanced",
      "port": 9229,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
    },
```

4. 下一步，选择 **View > Debug** 或者输入 **CTRL + SHIFT + D** 以切换到调试视图。

5. 从调试选项中，为你的主机应用程序选择 Microsoft Edge Chromium 选项，例如 **Excel 桌面版（Microsoft Edge Chromium）**。 选择 **F5** 或从菜单选择 **Debug > Start Debugging** 以开始调试。

6. 在主机应用程序（如 Excel）中，你的加载项现在可以使用了。 选择 **显示任务窗格** 或运行其他加载项命令。 此时将出现一个对话框，内容是：

> WebView 停止加载。 
> 要调试 webview，请使用适用于 Microsoft Edge 扩展的 Microsoft 调试器将 VS 代码附加到 webview 实例，然后单击“确定”以继续。 要防止今后出现此对话框，单击“取消”。

选择“**确定**”。

> [!NOTE]
> 如果选择“**取消**”，则当加载项的此实例正在运行时，将不会再次显示该对话框。 但如果重新启动加载项，则会再次看到该对话框。

7. 现在可以在你的项目代码中设置断点并进行调试。

## <a name="see-also"></a>另请参阅

* [测试和调试 Office 加载项](test-debug-office-add-ins.md)
* [适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展](debug-with-vs-extension.md)
* [从任务窗格附加调试器](attach-debugger-from-task-pane.md)