---
title: 适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展
description: 使用 Visual Studio Code extension Microsoft Office 加载项调试器调试 Office 外接程序。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 2439af12f30cef1b9d291578cbababe3ed601644
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530469"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>适用于 Visual Studio Code 的 Microsoft Office 加载项调试器扩展

通过 Visual Studio Code 的 Microsoft Office 外接程序调试器扩展，你可以针对边缘运行时调试 Office 外接程序。

此调试模式是动态的，允许您在代码运行时设置断点。 在调试器附加时，您可以立即看到代码中的更改，而不会丢失您的调试会话。 您的代码更改也会保留，以便您可以看到对代码进行多个更改的结果。 下图显示了此扩展在操作中。

![Office Addin 调试器扩展调试 Excel 外接程序的某个部分](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>先决条件

- [Visual Studio Code](https://code.visualstudio.com/) （必须以管理员身份运行）
- [Node.js （版本 10 +）](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

这些说明假定您有使用命令行的经验，了解基本 JavaScript，并已在使用 Yo Office 生成器之前创建了 Office 外接程序项目。 如果你之前未执行此操作，请考虑访问我们的一个教程，如此[Excel Office 外接教程教程](../tutorials/excel-tutorial.md)。

## <a name="install-and-use-the-debugger"></a>安装和使用调试器

1. 如果需要创建外接程序项目，请[使用 Yo Office 生成器创建一个](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator)外接程序项目。 按照命令行中的提示设置项目。 您可以根据需要选择任意语言或项目类型。

> [!NOTE]
> 如果已有一个项目，请跳过步骤1并转到步骤2。

2. 以管理员身份打开命令提示符。
   ![Windows 10 中的命令提示符选项，包括 "以管理员身份运行"](../images/run-as-administrator-vs-code.jpg)

3. 导航到您的项目目录。

4. 运行以下命令，以管理员身份在 Visual Studio Code 中打开项目。

```command&nbsp;line
code .
```

在 Visual Studio Code 打开后，手动导航到项目文件夹。

> [!TIP]
> 若要以管理员身份打开 Visual Studio Code，请选择 "以**管理员身份运行**" 选项，在 Windows 中搜索 Visual studio code 之后打开它。

5. 在 VS 代码中，选择**CTRL + SHIFT + X**打开扩展栏。 搜索 "Microsoft Office 外接程序调试器" 扩展并安装它。

6. 在项目的 ". vscode" 文件夹中，打开 " **launch.js"** 文件。 将以下代码添加到 `configurations` 部分：

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

7. 在刚刚复制的 JSON 部分中，找到 "url" 部分。 在此 URL 中，需要将大写的主机文本替换为 Office 加载项的主机应用程序。 例如，如果您的 Office 外接程序适用于 excel，则 URL 值将为 " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $ 16.01 $ en-us $ \$ \$ \$ 0"。

8. 打开命令提示符，并确保您在项目的根文件夹中。 运行命令 `npm start` 以启动开发服务器。 当加载项在 Office 客户端中加载时，打开任务窗格。

9. 返回到 Visual Studio Code，然后选择 "**查看 > 调试**" 或 enter **CTRL + SHIFT + D**切换到 "调试" 视图。

10. 从 "调试" 选项中，选择 "**附加到 Office 外接程序**"。从菜单中选择 " **F5** " 或选择 "**调试-> 启动调试**" 以开始调试。

11. 在项目的任务窗格文件中设置断点。 您可以通过悬停在代码行旁边并选择显示的红色圆圈，在 VS 代码中设置断点。

![对 VS 代码中的一行代码显示红色圆圈](../images/set-breakpoint.jpg)

12. 运行外接程序。 您将看到断点已命中，您可以检查局部变量。

## <a name="see-also"></a>另请参阅

* [测试和调试 Office 加载项](test-debug-office-add-ins.md)

* [使用 Windows 10 上的开发人员工具调试加载项](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [从任务窗格附加调试器](attach-debugger-from-task-pane.md)
