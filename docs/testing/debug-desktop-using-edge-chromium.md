---
title: 使用 Visual Studio Code 和 Microsoft Edge WebView2（基于 Chromium）在 Windows 上调试加载项
description: 了解如何在 VS Code 中调试使用 Microsoft Edge WebView2（基于 Chromium）的 Office 加载项。
ms.date: 02/18/2022
ms.localizationpriority: high
ms.openlocfilehash: 314799922b8d3687d8a24e93c49143cd3aa37e06
ms.sourcegitcommit: eef2064d7966db91f8401372dd255a32d76168c2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/31/2022
ms.locfileid: "67464816"
---
# <a name="debug-add-ins-on-windows-using-visual-studio-code-and-microsoft-edge-webview2-chromium-based"></a>使用 Visual Studio Code 和 Microsoft Edge WebView2（基于 Chromium）在 Windows 上调试加载项

在 Windows 上运行的 Office 加载项可以直接在 Visual Studio Code 中针对 Edge Chromium WebView2 运行时进行调试。

> [!IMPORTANT]
> 本文仅适用于 Office 在 Microsoft Edge Chromium WebView2 运行时中运行加载项时，如[ Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述。有关使用原始 WebView （EdgeHTML） 运行时针对Microsoft Edge 旧版进行Visual Studio Code调试的说明，请参阅 [适用于 Visual Studio Code 的调试器扩展](debug-with-vs-extension.md)。

> [!TIP]
> 如果不能或不希望使用内置于 Visual Studio Code 中的工具进行调试；或仅当加载项在 Visual Studio Code 外部运行时遇到问题，则可以使用 Edge（基于 Chromium）开发人员工具调试 Edge Chromium WebView2 运行时，如[使用 Microsoft Edge WebView2 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-chromium.md)中所述。

此调试模式是动态的，允许在代码运行时设置断点。 在附加调试器时立即查看代码中的更改，所有这些操作不会丢失调试会话。 代码更改也会持续存在，因此将看到对代码进行多次更改的结果。 下图显示此扩展正在运行。

![Office 加载项调试器扩展调试 Excel 加载项的一部分。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js （版本 10+）](https://nodejs.org/)
- Windows 10、11
- 支持包含 WebView2 的 Microsoft Edge（基于 Chromium）的平台和 Office 应用程序的组合，如 [ Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md) 中所述。如果 Microsoft 365 版本早于 2101，则需要安装 WebView2。 使用位于 [Microsoft Edge WebView2/在具有 Microsoft Edge Webview2 的...嵌入 Web 内容](https://developer.microsoft.com/microsoft-edge/webview2/) 的安装说明。

## <a name="use-the-visual-studio-code-debugger"></a>使用 Visual Studio Code 调试器

这些说明假定你在使用[适用于 Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)之前拥有使用命令行的经验，了解基本 JavaScript，并且已创建过 Office 加载项项目。如果你之前没有这样做过，请考虑访问我们的其中一个教程，例如 [Excel Office 加载项教程](../tutorials/excel-tutorial.md)。

1. 第一步取决于项目及其创建方式。

   - 如果要创建一个项目来试验Visual Studio Code中的调试，请使用 [Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)。若要执行此操作，请使用任何一个快速入门指南（例如 [Outlook 加载项快速入](../quickstarts/outlook-quickstart.md)门）。
   - 如果要调试使用 Yo Office 创建的现有项目，请跳到下一步。
   - 如果要调试未使用 Yo Office 创建的现有项目，请完成 [附录 A](#appendix-a) 中的过程，然后返回到此过程的下一步。

1. 打开 VS Code 并在其中打开项目。 

1. 选择“**视图”>“调试**”或者输入 **CTRL+SHIFT+D** 以切换到调试视图。

1. 从“**运行并调试**”选项中，为主机应用程序选择 Edge Chromium 选项，例如“**Outlook 桌面版（Edge Chromium）**”。 选择 **F5** 或从菜单中选择“**运行”>“开始调试**”以开始调试。 此操作在节点窗口中自动启动本地服务器以托管加载项，然后自动打开主机应用程序，例如 Excel 或 Word。 这可能需要几秒钟的时间。

   > [!TIP]
   > 如果不使用通过 Yo Office 创建的项目，系统可能会提示调整注册表项。 项目根文件夹下，在命令行中运行以下命令： 。
   >
   > ``` command&nbsp;line
   > npx office-addin-debugging start <your manifest path>
   > ```

   > [!IMPORTANT]
   > 如果项目是使用较旧版本的 Yo Office 创建的，则在开始调试大约 10 - 30 秒后，可能会看到以下错误对话框（此时可能已执行此过程中的另一步），并且可能隐藏在下一步中所述的对话框后面。
   >
   > ![显示"已配置的调试类型边缘不受支持"的错误。](../images/configured-debug-type-error.jpg)
   >
   > 完成 [附录 B](#appendix-b) 中的任务，然后重新启动此过程。
   
1. 在主机应用程序中，加载项现已可供使用。 选择 **显示任务窗格** 或运行其他加载项命令。 系统将显示一个对话框，其中包含类似于以下内容的文本：

   > WebView 停止加载。
   > 要调试 webview，请使用适用于 Microsoft Edge 扩展的 Microsoft 调试器将 VS 代码附加到 webview 实例，然后单击“确定”以继续。 要防止今后出现此对话框，单击“取消”。

   选择“**确定**”。

   [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

1. 现在可以在你的项目代码中设置断点并进行调试。 若要在 Visual Studio Code 中设置断点，请将鼠标悬停在代码行旁边，然后选择显示的红色圆圈。

    ![红色圆圈显示在 Visual Studio Code 中的代码行上。](../images/set-breakpoint.jpg)

1. 在加载项中运行调用断点行的功能。 你将看到已命中断点，可以检查局部变量。

   > [!NOTE]
   > `Office.initialize` 或 `Office.onReady` 调用中的断点将被忽略。 有关这些函数的详细信息，请参阅 [初始化 Office 加载项](../develop/initialize-add-in.md)。

> [!IMPORTANT]
> 停止调试会话的最佳方式是选择 **Shift+F5** 或从菜单中选择“**运行”>“停止调试**”。 此操作应关闭节点服务器窗口并尝试关闭主机应用程序，但主机应用程序上会出现提示，询问是否保存文档。 请做出适当选择，让主机应用程序关闭。 避免手动关闭节点窗口或主机应用程序。 这样做可能会导致 bug，尤其是在重复停止和启动调试会话时。
>
> 如果调试停止工作；例如，如果忽略断点；停止调试。 然后，如有必要，关闭所有主机应用程序窗口和节点窗口。 最后，关闭 Visual Studio Code 并重新将其打开。

### <a name="appendix-a"></a>附录 A

如果项目不是使用 Yo Office 创建的，则需要为 Visual Studio Code 创建调试配置。 

1. 在项目的 `\.vscode`文件夹中创建名为 `launch.json` 的文件（如果还没有文件夹）。 
1. 确保文件具有 `configurations` 数组。 下面是 `launch.json` 的简单示例。

   ```json
   {
     // other properities may be here.
   
     "configurations": [
   
       // configuration objects may be here.
   
     ]
   
     //other properies may be here.
   }
   ```

1. 将以下对象添加到 `configurations` 数组中。

   ```json
   {
      "name": "$HOST$ Desktop (Edge Chromium)",
      "type": "pwa-msedge",
      "request": "attach",
      "useWebView": true,
      "port": 9229,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: Excel Desktop",
      "postDebugTask": "Stop Debug"
   },
   ```

1. 将占位符 `$HOST$` 替换为加载项所运行的 Office 应用程序的名称，例如 `Outlook` 或 `Word`。
1. 保存并关闭此文件。

### <a name="appendix-b"></a>附录 B

1. 在错误对话框中，选择"**取消**"按钮。
1. 如果调试未自动停止，请选择 **Shift+F5** 或从菜单中选择"**运行>停止调试**"。 
1. 如果本地服务器未自动关闭，请关闭运行本地服务器的节点窗口。
1. 如果 Office 关闭应用程序，请关闭该应用程序。
1. 打开项目中的 `\.vscode\launch.json` 文件。 
1. 在 `configurations` 数组中，有多个配置对象。找到其名称具有模式 `$HOST$ Desktop (Edge Chromium)` 的对象，其中 $HOST$ 是加载项运行的 Office 应用程序；例如，`Outlook Desktop (Edge Chromium)` 或 `Word Desktop (Edge Chromium)`。 
1. 将 `"type"` 属性的值从 `"edge"` 更改为 `"pwa-msedge"`。
1. 将 `"useWebView"` 属性的值从字符串 `"advanced"` 更改为布尔值 `true` （请注意， `true` 周围没有引号）。
1. 保存文件。
1. 关闭 VS Code。

## <a name="see-also"></a>另请参阅

- [测试和调试 Office 加载项](test-debug-office-add-ins.md)
- [使用 Visual Studio Code 和 Microsoft Edge 旧版 WebView （EdgeHTML）在 Windows 上调试加载项](debug-with-vs-extension.md)
- [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
- [使用旧版 Edge 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-legacy.md)
- [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](debug-add-ins-using-devtools-edge-chromium.md)
- [从任务窗格附加调试器](attach-debugger-from-task-pane.md)
- [Office 加载项中的运行时](runtimes.md)
