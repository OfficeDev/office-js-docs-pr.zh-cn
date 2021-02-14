---
title: 适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展
description: 使用Visual Studio调试器Microsoft Office代码扩展来调试 Office 外接程序。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237929"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展

借助 Microsoft Office 外接程序调试器扩展 for Visual Studio Code，您可以使用原始 WebView (EdgeHTML) 运行时针对 Microsoft Edge 调试 Office 外接程序。 有关针对基于 Chromium (Microsoft Edge WebView2 进行) 的说明，请参阅 [本文](./debug-desktop-using-edge-chromium.md)

此调试模式是动态的，允许您在代码运行时设置断点。 在附加调试程序时，你可以立即在代码中看到更改，所有这些更改不会丢失调试会话。 代码更改也会持续存在，因此你可以看到对代码进行多次更改的结果。 下图显示了此扩展的操作。

![Office 加载项调试程序扩展调试 Excel 加载项的一部分](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>先决条件

- [Visual Studio必须](https://code.visualstudio.com/) (管理员角色运行代码) 
- [Node.js (版本 10+) ](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

这些说明假定你具有使用命令行的经验，了解基本 JavaScript，并且已使用 Yo Office 生成器之前创建了 Office 加载项项目。 如果之前尚未这样做，请考虑访问我们的教程之一，如本 Excel Office [加载项教程](../tutorials/excel-tutorial.md)。

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

1. 如果需要创建加载项项目，请使用 Yo Office 生成器 [创建一个](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。 按照命令行中的提示设置项目。 可以选择任何语言或项目类型来满足您的需求。

> [!NOTE]
> 如果已有项目，请跳过步骤 1 并移动到步骤 2。

2. 以管理员角色打开命令提示符。
   ![命令提示符选项，包括 Windows 10 中的"以管理员方式运行"](../images/run-as-administrator-vs-code.jpg)

3. 导航到项目目录。

4. 运行以下命令以管理员Visual Studio代码打开项目。

```command&nbsp;line
code .
```

打开Visual Studio后，手动导航到项目文件夹。

> [!TIP]
> 若要以Visual Studio方式打开代码，请在 Windows 中搜索代码后Visual Studio代码时选择"以管理员方式运行"选项。

5. 在 VS Code 中，选择 **Ctrl + Shift + X** 以打开扩展栏。 搜索"Microsoft Office调试器"扩展并安装它。

6. 在项目的 .vscode 文件夹中，打开launch.js **文件。** 将以下代码添加到 `configurations` 该部分：

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

7. 在刚复制的 JSON 部分中，查找"url"部分。 在此 URL 中，您需要将大写的 HOST 文本替换为托管 Office 外接程序的应用程序。 例如，如果 Office 外接程序适用于 Excel，则 URL 值为 https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0"。

8. 打开命令提示符，并确保你位于项目的根文件夹。 运行命令 `npm start` 以启动开发服务器。 当加载项在 Office 客户端中加载时，打开任务窗格。

9. 返回到Visual Studio代码，然后选择 **">调试** "或输入 **Ctrl + Shift + D** 以切换到调试视图。

10. 从"调试"选项中，选择 **"附加到 Office 加载项"。** 选择 **F5** 或从>开始 **调试** 以开始调试。

11. 在项目的任务窗格文件中设置断点。 通过在代码行旁边悬停并选择出现的红色圆圈，可以在 VS Code 中设置断点。

![VS Code 中的一行代码上显示一个红色圆圈](../images/set-breakpoint.jpg)

12. 运行加载项。 你将看到断点已命中，你可以检查本地变量。

## <a name="see-also"></a>另请参阅

* [测试和调试 Office 加载项](test-debug-office-add-ins.md)

* [使用 Windows 10 上的开发人员工具调试加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [使用 Microsoft Edge WebView2 和基于 Chromium (Windows 调试加载项) ](debug-desktop-using-edge-chromium.md)
