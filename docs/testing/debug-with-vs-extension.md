---
title: 使用 Visual Studio Code 和 Microsoft Edge 旧版 WebView （EdgeHTML）在 Windows 上调试加载项
description: 了解如何使用 VS Code 中的 Office 加载项调试器扩展Office使用 Microsoft Edge 旧版 WebView (EdgeHTML) 的加载项。
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: e43afbe2ae600a7013a8d4872d4148c571c0861d
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744095"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展

Office在 Windows 上运行的外接程序可以使用 Visual Studio Code 中的 Office 外接程序调试器扩展，以通过原始 WebView (EdgeHTML) 运行时对 Microsoft Edge 旧版 进行调试。 

> [!IMPORTANT]
> 本文仅适用于 Office 在原始 WebView (EdgeHTML) 运行时中运行外接程序的情况，如 Office [Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) 使用的浏览器所介绍。有关针对基于 Microsoft Edge WebView2 Microsoft Edge (Chromium) 在 Visual Studio 代码中进行调试的说明，请参阅 [Microsoft Office Add-in Debugger Extension for Visual Studio Code](debug-desktop-using-edge-chromium.md)。

> [!TIP]
> 如果无法或不想使用 Visual Studio Code 中内置的工具进行调试;或者遇到仅在外接程序在 Visual Studio Code 外部运行时发生的问题，可以使用 Edge 旧版开发人员工具调试 Edge 旧版 (EdgeHTML) 运行时，如 使用开发人员工具调试外接程序中所述[Microsoft Edge 旧版](debug-add-ins-using-devtools-edge-legacy.md)。

此调试模式是动态的，允许在代码运行时设置断点。 在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。 代码更改也持续存在，因此可以看到对代码进行多次更改的结果。 下图显示此扩展正在运行。

![Office 加载项调试器扩展调试 Excel 加载项的一部分。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js （版本 10+）](https://nodejs.org/)
- Windows 10、11
- [Microsoft Edge](https://www.microsoft.com/edge)支持 Microsoft Edge 旧版 和原始 webview (EdgeHTML) 的平台和 Office 应用程序的组合，如 [Office 外接程序使用的浏览器部分所](../concepts/browsers-used-by-office-web-add-ins.md)说明。

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

这些说明假定你拥有使用命令行的经验，了解基本 JavaScript，并且已创建一个 Office 外接程序项目，然后才将 [Yeoman 生成器用于 Office 外接程序](../develop/yeoman-generator-overview.md)。如果你之前没有这样做，请考虑访问我们的其中一个教程，Excel Office[外接程序教程](../tutorials/excel-tutorial.md)。

1. 第一步取决于项目及其创建方式。

   - 如果要创建一个项目以在 Visual Studio Code 中试验调试，请使用 [适用于 Office 加载项的 Yeoman 生成器](../develop/yeoman-generator-overview.md)。若要执行此操作，请使用我们的任何快速入门指南（如 [Outlook 加载项快速入门](../quickstarts/outlook-quickstart.md)）。 
   - 如果要调试使用 Yo Office 创建的现有项目，请跳到下一步。
   - 如果要调试不是使用 Yo Office 创建的现有项目，请执行附录中的过程，然后返回到此过程的下一步。[](#appendix)


1. 打开 VS Code 并在其中打开项目。 

1. 在 VS Code 中，选择 **CTRL+SHIFT+X** 打开扩展栏。 搜索"Microsoft Office加载项调试器"扩展并安装它。

1. 选择“**视图”>“调试**”或者输入 **CTRL+SHIFT+D** 以切换到调试视图。

1. 从 **"运行和调试**"选项中，为主机应用程序选择"旧版边缘"选项，例如Outlook **桌面 (旧版)**。 选择 **F5** 或从菜单中选择“**运行”>“开始调试**”以开始调试。 此操作在节点窗口中自动启动本地服务器以托管加载项，然后自动打开主机应用程序，例如 Excel 或 Word。 这可能需要几秒钟的时间。

1. 在主机应用程序中，加载项现已可供使用。 选择 **显示任务窗格** 或运行其他加载项命令。 对话框将显示如下：

   > WebView 停止加载。
   > 若要调试 WebView，请将VS Code Microsoft Debugger for Edge 扩展附加到 WebView 实例，然后单击 **"确定"** 继续。 若要阻止将来显示此对话框，请单击"取消 **"**。

   选择“**确定**”。

   > [!NOTE]
   > 如果选择“**取消**”，则当加载项的此实例正在运行时，将不会再次显示该对话框。 但如果重新启动加载项，则会再次看到该对话框。

1. 在项目的任务窗格文件中设置断点。 若要在 Visual Studio Code 中设置断点，请将鼠标悬停在代码行旁边，然后选择显示的红色圆圈。

    ![红色圆圈显示在 Visual Studio Code 中的代码行上。](../images/set-breakpoint.jpg)

1. 在加载项中运行调用断点行的功能。 你将看到已命中断点，可以检查局部变量。

   > [!NOTE]
   > `Office.initialize` 或 `Office.onReady` 调用中的断点将被忽略。 有关这些方法的详细信息，请参阅 [初始化 Office 加载项](../develop/initialize-add-in.md)。

> [!IMPORTANT]
> 停止调试会话的最佳方式是选择 **Shift+F5** 或从菜单中选择“**运行”>“停止调试**”。 此操作应关闭节点服务器窗口并尝试关闭主机应用程序，但主机应用程序上会出现提示，询问是否保存文档。 请做出适当选择，让主机应用程序关闭。 避免手动关闭节点窗口或主机应用程序。 这样做可能会导致 bug，尤其是在重复停止和启动调试会话时。
>
> 如果调试停止工作；例如，如果忽略断点；停止调试。 然后，如有必要，关闭所有主机应用程序窗口和节点窗口。 最后，关闭 Visual Studio Code 并重新将其打开。

### <a name="appendix"></a>附录

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
      "name": "HOST Desktop (Edge Legacy)",
      "type": "office-addin",
      "request": "attach",
      "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
      "port": 9222,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: HOST Desktop",
      "postDebugTask": "Stop Debug"
    }
    ```

1. 将三`HOST`处占位符替换为外接程序Office应用程序的名称;例如 或 `Outlook` `Word`。
1. 保存并关闭此文件。

## <a name="see-also"></a>另请参阅

- [测试和调试 Office 加载项](test-debug-office-add-ins.md)
- [使用基于 WebView2 Windows Visual Studio Code Microsoft Edge 调试 (Chromium加载项](debug-desktop-using-edge-chromium.md)) 。
- [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
- [使用旧版 Edge 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-legacy.md)
- [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](debug-add-ins-using-devtools-edge-chromium.md)
- [从任务窗格附加调试器](attach-debugger-from-task-pane.md)
