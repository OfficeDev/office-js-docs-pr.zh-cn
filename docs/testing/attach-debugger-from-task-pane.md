---
title: 从任务窗格附加调试器
description: ''
ms.date: 10/17/2018
localization_priority: Priority
ms.openlocfilehash: 1c2be843fcd911c1bfd87408d603bb750d6e738a
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386475"
---
# <a name="attach-a-debugger-from-the-task-pane"></a>从任务窗格附加调试器

在 Office 2016 for Windows 生成号 77xx.xxxx 或更高版本中，可以从任务窗格附加调试器。使用附加调试器功能，可直接将调试器附加到正确的 Internet Explorer 进程中。无论你使用的是 Yeoman 生成器、Visual Studio Code、node.js、Angular 还是其他任何工具，都可以附加调试器。 

若要启动“**附加调试器**”工具，选择任务窗格右上角来激活“**个性**”菜单，如下图红圈所示。   

> [!NOTE]
> - 目前，唯一受支持的调试器工具是 [Visual Studio 2015](https://www.visualstudio.com/downloads/) [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) 或更高版本。如果没有安装 Visual Studio，选择“附加调试器”**** 选项不会有任何结果。   
> - 只能使用“附加调试器”**** 工具调试客户端 JavaScript。 要调试服务器端代码（例如 Node.js 服务器），可选择多种方式。 有关如何使用 Visual Studio Code 进行调试的信息，请参阅 [VS Code 中的 Node.js 调试](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)。 如果没有使用 Visual Studio Code，请搜索“debug Node.js”或“debug {name-of-server}”。

![“附加调试器”菜单屏幕截图](../images/attach-debugger.png)

选择“**附加调试器**”。此操作将启动“**Visual Studio 实时调试器**”对话框，如下图所示。 

![“Visual Studio JIT 调试器”对话框屏幕截图](../images/visual-studio-debugger.png)

Visual Studio 中的“解决方案资源管理器”**** 会显示代码文件。   可以在要使用 Visual Studio 调试的代码行处设置断点。

> [!NOTE]
> 如果你没有看到“个性”菜单，则可以使用 Visual Studio 调试加载项。 确保你的任务窗格加载项已在 Office 中打开，然后按照以下步骤操作：

> 1. 在 Visual Studio 中，依次选择“**调试**” > “**附加到进程**”。
> 2. 在“**附加到进程**”对话框中，选择所有可用的 Iexplore.exe 进程，然后选择“**附加**”按钮。

若要详细了解如何在 Visual Studio 中进行调试，请参阅以下内容：

-   若要在 Visual Studio 中启动并使用 DOM 资源管理器，请参阅 [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates)（使用新项目模板为 Office 生成漂亮应用）博客文章中[提示和技巧](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks)部分的提示 4。
-   若要设置断点，请参阅[使用断点](https://docs.microsoft.com/visualstudio/debugger/using-breakpoints?view=vs-2015)。
-   若要使用 F12，请参阅[使用 F12 开发人员工具](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))。

## <a name="see-also"></a>另请参阅

- [在 Visual Studio 中创建和调试 Office 加载项](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [发布 Office 加载项](../publish/publish.md)
