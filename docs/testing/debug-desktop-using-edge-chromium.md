---
title: 使用 Windows 上的 Microsoft Edge WebView2 （基于 Chromium）调试加载项
description: 了解如何在 VS 代码中使用适用于 Microsoft Edge 扩展的调试器来调试使用 Microsoft Edge WebView2（基于 Chromium）的 Office 加载项。
ms.date: 08/18/2021
localization_priority: Priority
ms.openlocfilehash: bbb475071660415f19b9a9fe5aaee0a6d735e20c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938767"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>使用 Windows 上的 Microsoft Edge Chromium WebView2 调试加载项

在 Windows 上正在运行的 Office 加载项可以使用 VS 代码中适用于 Microsoft Edge 扩展的调试器来对 Edge Chromium WebView2 运行时进行调试。

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）
- [Node.js （版本 10+）](https://nodejs.org/)
- Windows 10
- 支持包含 WebView2 的 Microsoft Edge（基于 Chromium）的平台和 Office 应用程序的组合，如 [ Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md) 中所述。如果 Microsoft 365 版本早于 2101，则需要安装 WebView2。 使用位于 [Microsoft Edge WebView2/在具有 Microsoft Edge Webview2 的...嵌入 Web 内容](https://developer.microsoft.com/microsoft-edge/webview2/) 的安装说明。

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

1. 使用 [ 适用于 Office 加载项的 Yeoman 生成器 ](https://github.com/OfficeDev/generator-office) 创建项目。可以使用我们的任何一个快速入门指南，例如 [Outlook 加载项快速入门 ](../quickstarts/outlook-quickstart.md)，以做到这一点。

    > [!TIP]
    > 如果没有使用基于 Yeoman 生成器的加载项，则系统可能提示需要调整一个注册表项。 在项目根文件夹下，在命令行中运行以下命令：`office-add-in-debugging start <your manifest path>`

1. 在 VS Code 中打开项目。 在 VS Code 中，选择 **CTRL+SHIFT+X** 打开扩展栏。 搜索“适用于 Microsoft Edge 的调试器”扩展并安装。

1. 下一步，选择“**视图”>“调试**”或者输入 **CTRL+SHIFT+D** 以切换到调试视图。

1. 从“**运行并调试**”选项中，为主机应用程序选择 Edge Chromium 选项，例如“**Excel 桌面版（Edge Chromium）**”。 选择 **F5** 或从菜单中选择“**运行”>“开始调试**”以开始调试。 此操作在节点窗口中自动启动本地服务器以托管加载项，然后自动打开主机应用程序，例如 Excel 或 Word。 这可能需要几秒钟的时间。

1. 在主机应用程序中，加载项现已可供使用。 选择 **显示任务窗格** 或运行其他加载项命令。 此时将出现一个对话框，内容是：

   > WebView 停止加载。
   > 要调试 webview，请使用适用于 Microsoft Edge 扩展的 Microsoft 调试器将 VS 代码附加到 webview 实例，然后单击“确定”以继续。 要防止今后出现此对话框，单击“取消”。

   选择“**确定**”。

   > [!NOTE]
   > 如果选择“**取消**”，则当加载项的此实例正在运行时，将不会再次显示该对话框。 但如果重新启动加载项，则会再次看到该对话框。

1. 现在可以在你的项目代码中设置断点并进行调试。

   > [!NOTE]
   > `Office.initialize` 或 `Office.onReady` 调用中的断点将被忽略。 有关这些方法的详细信息，请参阅 [初始化 Office 加载项](../develop/initialize-add-in.md)。

> [!IMPORTANT]
> 停止调试会话的最佳方式是选择 **Shift+F5** 或从菜单中选择“**运行”>“停止调试**”。 此操作应关闭节点服务器窗口并尝试关闭主机应用程序，但主机应用程序上会出现提示，询问是否保存文档。 请做出适当选择，让主机应用程序关闭。 避免手动关闭节点窗口或主机应用程序。 这样做可能会导致 bug，尤其是在重复停止和启动调试会话时。
>
> 如果调试停止工作；例如，如果忽略断点；停止调试。 然后，如有必要，关闭所有主机应用程序窗口和节点窗口。 最后，关闭 Visual Studio Code 并重新将其打开。

## <a name="see-also"></a>另请参阅

- [测试和调试 Office 加载项](test-debug-office-add-ins.md)
- [适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展](debug-with-vs-extension.md)
- [从任务窗格附加调试器](attach-debugger-from-task-pane.md)
