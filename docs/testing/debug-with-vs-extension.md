---
title: 适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展
description: 使用Visual Studio Code调试Microsoft Office调试器中的扩展Office调试外接程序。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: d027e5937fa3a58623ce9e798fc683e5459e73b8b72606c0a006e465c9c1360c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088463"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展

Microsoft Office外接程序调试器扩展 for Visual Studio Code 允许你使用原始 webView Microsoft Edge EdgeHTML Microsoft Edge运行时调试 Office 外接程序 (调试) 外接程序。 有关针对基于 WebView2 Microsoft Edge (Chromium进行) 的说明，[请参阅本文](./debug-desktop-using-edge-chromium.md)

此调试模式是动态的，允许在代码运行时设置断点。 在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。 代码更改也持续存在，因此可以看到对代码进行多次更改的结果。 下图显示了此扩展的操作。

![Office加载项调试器扩展调试加载项Excel部分。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）
- [Node.js （版本 10+）](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

这些说明假定你拥有使用命令行的经验，了解基本 JavaScript，并且已创建一个 Office 加载项项目，然后才使用 Yo Office 生成器。 如果你之前没有这样做，请考虑访问我们的教程之一，Excel Office[外接程序教程](../tutorials/excel-tutorial.md)。

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

1. 如果需要创建加载项项目，请使用[Yo Office生成器创建一个](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。 按照命令行中的提示设置项目。 可以选择任何语言或项目类型以满足你的需求。

    > [!NOTE]
    > 如果已有项目，请跳过步骤 1 并移至步骤 2。

1. 以管理员角色打开命令提示符。
   ![命令提示符选项，包括"以管理员Windows 10。](../images/run-as-administrator-vs-code.jpg)

1. 导航到项目目录。

1. 运行以下命令以管理员Visual Studio Code打开项目。

    ```command&nbsp;line
    code .
    ```

  打开Visual Studio Code后，手动导航到项目文件夹。

  > [!TIP]
  > 若要以Visual Studio Code方式打开文件，请选择"以管理员方式运行"选项，Visual Studio Code中搜索后打开Windows。

1. 在 VS 代码中，选择 **CTRL + SHIFT + X** 打开扩展栏。 搜索"Microsoft Office加载项调试器"扩展并安装它。

1. 在你的项目 .vscode 文件夹中打开 **launch.json** 文件。 将以下代码添加到 `configurations` 部分。

    ```JSON
    {
      "type": "office-addin",
      "request": "attach",
      "name": "Attach to Office Add-ins",
      "port": 9222,
      "trace": "verbose",
      "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
      "webRoot": "${workspaceFolder}",
      "timeout": 45000
    }
    ```

1. 在刚刚复制的 JSON 部分中，找到"url"部分。 在此 URL 中，您需要将大写的 HOST 文本替换为托管您的外接程序Office应用程序。 例如，如果Office外接程序用于 Excel，则 URL 值将是 https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0"。

1. 打开命令提示符，并确保位于项目的根文件夹。 运行命令 `npm start` 以启动开发服务器。 当加载项在客户端Office时，打开任务窗格。

1. 返回到"Visual Studio Code并选择"查看 **>调试"** 或输入 **Ctrl + Shift + D** 以切换到调试视图。

1. 从"调试"选项中，选择"**附加到Office加载项"。** 从 **菜单中选择 F5** 或 **>** 调试 -开始调试"开始调试。

1. 在项目的任务窗格文件中设置断点。 通过将鼠标悬停在代码行Visual Studio Code并选择出现的红色圆圈，可以在代码中设置断点。

    ![在代码行中出现红色圆圈Visual Studio Code。](../images/set-breakpoint.jpg)

1. 运行加载项。 你将看到已命中的断点，并且你可以检查本地变量。

## <a name="see-also"></a>另请参阅

- [测试和调试 Office 加载项](test-debug-office-add-ins.md)

- [使用 Windows 10 上的开发人员工具调试加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [使用 Windows 上的 Microsoft Edge WebView2 （基于 Chromium）调试加载项](debug-desktop-using-edge-chromium.md)
