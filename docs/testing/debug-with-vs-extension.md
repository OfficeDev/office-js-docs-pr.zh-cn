---
title: 适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展
description: 使用Visual Studio Code调试Microsoft Office调试器中的扩展Office调试外接程序。
ms.date: 10/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1eb71ec1bd52198af32129882cb531451fff422a
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138637"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展

Microsoft Office 外接程序调试器扩展 for Visual Studio Code 允许你使用原始 webView (EdgeHTML) 运行时针对 Microsoft Edge 调试 Office 外接程序。 有关针对基于 WebView2 Microsoft Edge (Chromium进行) 的说明，[请参阅本文](./debug-desktop-using-edge-chromium.md)

此调试模式是动态的，允许在代码运行时设置断点。 在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。 代码更改也持续存在，因此可以看到对代码进行多次更改的结果。 下图显示了此扩展的操作。

![Office加载项调试器扩展调试加载项Excel部分。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）
- [Node.js （版本 10+）](https://nodejs.org/)
- Windows 10、11
- [Microsoft Edge](https://www.microsoft.com/edge)

这些说明假定你拥有使用命令行的经验，了解基本 JavaScript，并且已创建一个 Office 加载项项目，然后才使用 Yo Office 生成器。 如果你之前没有这样做，请考虑访问我们的其中一个教程，Excel Office[外接程序教程](../tutorials/excel-tutorial.md)。

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

1. 如果需要创建加载项项目，请使用[Yo Office生成器创建一个](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。 按照命令行中的提示设置项目。 可以选择任何语言或项目类型以满足你的需求。 本教程使用Excel窗格加载项。

    > [!NOTE]
    > 如果已有项目，请跳过步骤 1 并移至步骤 2。

1. 以管理员角色打开命令提示符。
   ![命令提示符选项，包括 Windows 10 和 11 中的"以管理员Windows 10"。](../images/run-as-administrator-vs-code.jpg)

1. 导航到项目目录。

1. 运行以下命令以管理员Visual Studio Code中打开项目。

    ```command&nbsp;line
    code .
    ```

  打开Visual Studio Code，手动导航到项目文件夹。

  > [!TIP]
  > 若要以Visual Studio Code方式打开网站，请选择"以管理员方式运行"选项，Visual Studio Code中搜索后打开Windows。

1. 在 VS Code 中，选择 **CTRL+SHIFT+X** 打开扩展栏。 搜索"Microsoft Office加载项调试器"扩展并安装它。

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

1. 在刚刚复制的 JSON 部分中，查找 `"url"` 属性。 在此 URL 中，您需要将大写的 **HOST** 文本替换为托管您的外接程序Office应用程序。 例如，如果你Office外接程序用于Excel，则你的 URL 值是 `"https://localhost:3000/taskpane.html?_host_Info=Excel$Win32$16.01$en-US$\$\$\$0"` 。

1. 打开命令提示符，并确保位于项目的根文件夹。 运行命令 `npm start` 以启动开发服务器。 当加载项在加载项应用程序中Office时，打开任务窗格。

1. 返回到"Visual Studio Code并选择"查看 **>调试"** 或输入 **Ctrl+Shift+D** 以切换到调试视图。

1. From the Debug options， choose **Attach to Office Add-ins**.从 **菜单中选择 F5** 或 **>开始调试**"以开始调试。

1. 在项目的任务窗格文件中设置断点。 通过将鼠标悬停在代码行Visual Studio Code并选择出现的红色圆圈，可以在代码行中设置断点。

    ![在代码行上显示红色圆圈Visual Studio Code。](../images/set-breakpoint.jpg)

1. 运行加载项。 你将看到已命中的断点，并且你可以检查本地变量。

## <a name="see-also"></a>另请参阅

- [测试和调试 Office 加载项](test-debug-office-add-ins.md)

- [在加载项上使用开发人员工具调试Windows](debug-add-ins-using-f12-developer-tools-on-windows.md)

- [使用 Windows 上的 Microsoft Edge WebView2 （基于 Chromium）调试加载项](debug-desktop-using-edge-chromium.md)
