---
title: '使用旧版 WebView Windows和 EdgeHTML Visual Studio Code Microsoft Edge调试 (加载项) '
description: 了解如何使用 VS Code 中的 Office 加载项调试器扩展Office使用 Microsoft Edge 旧版 WebView (EdgeHTML) 的加载项。
ms.date: 02/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 11b728f9b3f467017711c9d75cfd07767957deae
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467692"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展

Office运行在 Windows 上的外接程序可以使用 Visual Studio Code 中的 Office 外接程序调试器扩展，以针对原始 WebView (EdgeHTML) Microsoft Edge 旧版 进行调试。 

> [!IMPORTANT]
> 本文仅适用于 Office 在原始 WebView (EdgeHTML) 运行时中运行外接程序的情况，如 Office [外接程序](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器所说明。有关针对基于 Microsoft Edge WebView2 Microsoft Edge (Chromium) 在 Visual Studio 代码中进行调试的说明，请参阅 [Microsoft Office Add-in Debugger Extension for Visual Studio Code](debug-desktop-using-edge-chromium.md)。

> [!TIP]
> 如果无法或不想使用 Visual Studio Code 中内置的工具进行调试;或者遇到仅在外接程序在 Visual Studio Code 外部运行时发生的问题，可以使用 Edge 旧版开发人员工具调试 Edge 旧版 (EdgeHTML) 运行时，如 使用开发人员工具调试外接程序中所述[Microsoft Edge 旧版](debug-add-ins-using-devtools-edge-legacy.md)。

此调试模式是动态的，允许在代码运行时设置断点。 在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。 代码更改也持续存在，因此可以看到对代码进行多次更改的结果。 下图显示了此扩展的操作。

![Office加载项调试器扩展调试加载项Excel部分。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js （版本 10+）](https://nodejs.org/)
- Windows 10、11
- [Microsoft Edge](https://www.microsoft.com/edge)支持 Microsoft Edge 旧版 与原始 Webview (EdgeHTML) 的平台和 Office 应用程序的组合，如 [Office 外接程序](../concepts/browsers-used-by-office-web-add-ins.md)使用的浏览器部分所说明。

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

这些说明假定你拥有使用命令行的经验，了解基本 JavaScript，并且已创建一个 Office 加载项项目，然后才使用 Yo Office 生成器。 如果你之前没有这样做，请考虑访问我们的其中一个教程，Excel Office[外接程序教程](../tutorials/excel-tutorial.md)。

1. 第一步取决于项目及其创建方式。

   - 如果要创建一个项目来尝试在 Visual Studio Code 中调试，请使用适用于 Office [加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。为此，请使用我们的任一快速入门指南，Outlook[快速](../quickstarts/outlook-quickstart.md)入门。 
   - 如果要调试使用 Yo Office 创建的现有项目，请跳到下一步。
   - 如果要调试不是使用 Yo Office 创建的现有项目，请执行附录中的过程，然后返回到此过程的下一步。[](#appendix)


1. 打开VS Code，然后打开项目中的项目。 

1. 在 VS Code 中，选择 **CTRL+SHIFT+X** 打开扩展栏。 搜索"Microsoft Office加载项调试器"扩展并安装它。

1. Choose  **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.

1. 从 **"运行和调试**"选项中，为主机应用程序选择"旧版边缘"选项，Outlook **桌面 (旧版)**。 选择 **F5** 或从菜单中选择“**运行”>“开始调试**”以开始调试。 此操作在节点窗口中自动启动本地服务器以托管加载项，然后自动打开主机应用程序，例如 Excel 或 Word。 这可能需要几秒钟的时间。

1. 在主机应用程序中，加载项现已可供使用。 选择 **显示任务窗格** 或运行其他加载项命令。 对话框将显示如下：

   > WebView 停止加载。
   > 若要调试 WebView，请将VS Code Microsoft Debugger for Edge 扩展附加到 WebView 实例，然后单击 **"确定"** 继续。 若要阻止将来显示此对话框，请单击"取消 **"**。

   选择“**确定**”。

   > [!NOTE]
   > 如果选择“**取消**”，则当加载项的此实例正在运行时，将不会再次显示该对话框。 但如果重新启动加载项，则会再次看到该对话框。

1. 在项目的任务窗格文件中设置断点。 若要在代码Visual Studio Code断点，请将鼠标悬停在代码行旁边，然后选择出现的红色圆圈。

    ![在代码行上显示红色圆圈Visual Studio Code。](../images/set-breakpoint.jpg)

1. 在加载项中运行调用断点行的功能。 你将看到已命中断点，并且你可以检查本地变量。

   > [!NOTE]
   > `Office.initialize` 或 `Office.onReady` 调用中的断点将被忽略。 有关这些方法的详细信息，请参阅 [初始化 Office 加载项](../develop/initialize-add-in.md)。

> [!IMPORTANT]
> 停止调试会话的最佳方式是选择 **Shift+F5** 或从菜单中选择“**运行”>“停止调试**”。 此操作应关闭节点服务器窗口并尝试关闭主机应用程序，但主机应用程序上会出现提示，询问是否保存文档。 请做出适当选择，让主机应用程序关闭。 避免手动关闭节点窗口或主机应用程序。 这样做可能会导致 bug，尤其是在重复停止和启动调试会话时。
>
> 如果调试停止工作；例如，如果忽略断点；停止调试。 然后，如有必要，关闭所有主机应用程序窗口和节点窗口。 最后，关闭 Visual Studio Code 并重新将其打开。

### <a name="appendix"></a>附录

如果项目不是使用 Yo Office创建的，则需要为项目创建调试Visual Studio Code。 

1. 在项目文件夹中 `launch.json` 创建 `\.vscode` 一个名为 的文件（如果还没有）。 
1. 确保文件具有数组 `configurations` 。 下面是 一个简单示例 `launch.json`。

    ```json
    {
      // other properities may be here.

      "configurations": [

        // configuration objects may be here.

      ]

      //other properies may be here.
    }
    ```

1. 将以下对象添加到数组 `configurations` 。

    ```json
    {
      "name": "$HOST$ Desktop (Edge Legacy)",
      "type": "office-addin",
      "request": "attach",
      "url": "https://localhost:3000/taskpane.html?_host_Info=Excel$Win32$16.01$en-US$$$$0",
      "port": 9222,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: Excel Desktop",
      "postDebugTask": "Stop Debug"
    }
    ```

1. 将占位符`$HOST$`替换为`Outlook`外接程序Office应用程序的名称;例如 或 `Word`。
1. 保存并关闭此文件。

## <a name="see-also"></a>另请参阅

- [测试和调试 Office 加载项](test-debug-office-add-ins.md)
- [使用基于 WebView2 Windows的 Visual Studio Code Microsoft Edge ](debug-desktop-using-edge-chromium.md)调试 (Chromium加载项) 。
- [使用适用于 Internet Explorer 的开发人员工具调试加载项](debug-add-ins-using-f12-tools-ie.md)
- [使用旧版 Edge 开发人员工具调试加载项](debug-add-ins-using-devtools-edge-legacy.md)
- [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](debug-add-ins-using-devtools-edge-chromium.md)
- [从任务窗格附加调试器](attach-debugger-from-task-pane.md)
