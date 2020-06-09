---
ms.date: 05/17/2020
description: 了解如何调试不使用任务窗格的 Excel 自定义函数。
title: UI-较少的自定义函数调试
localization_priority: Normal
ms.openlocfilehash: c984c5a46f8418006fa64f167b0f5882b0cb8a17
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609324"
---
# <a name="ui-less-custom-functions-debugging"></a>UI-较少的自定义函数调试

不使用任务窗格或其他用户界面元素（无 UI 自定义函数）的自定义函数的调试可以通过多种方式来完成，具体取决于您使用的平台。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

在 Windows 上：
- [Excel Desktop 和 Visual Studio Code （VS Code）调试器](#use-the-vs-code-debugger-for-excel-desktop)
- [Web 上的 Excel 和 VS 代码调试器](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Web 和浏览器工具上的 Excel](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [命令行](#use-the-command-line-tools-to-debug)

在 Mac 上：
- [Web 和浏览器工具上的 Excel](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [命令行](#use-the-command-line-tools-to-debug)

> [!NOTE]
> 为简单起见，本文介绍了如何在使用 Visual Studio Code 编辑、运行任务以及某些情况下使用调试视图的上下文中进行调试。 如果使用的是其他编辑器或命令行工具，请参阅本文末尾的[命令行说明](#commands-for-building-and-running-your-add-in)。

## <a name="requirements"></a>Requirements

开始调试之前，应使用[Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建自定义函数项目。 有关如何创建自定义函数项目的指南，请参阅[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>对 Excel 桌面使用 VS 代码调试器

您可以使用 VS 代码在桌面的 Office Excel 中调试无 UI 的自定义函数。

> [!NOTE]
> 对 Mac 的桌面调试不可用，但可通过[使用浏览器工具和命令行来调试 web 上的 Excel 来](#use-the-command-line-tools-to-debug)实现。

### <a name="run-your-add-in-from-vs-code"></a>从 VS 代码运行外接程序

1. 打开[VS 代码](https://code.visualstudio.com/)中的自定义函数根项目文件夹。
2. 选择 "**终端 > 运行任务**" 并键入或选择 "**监视**"。 这将监视和重建任何文件更改。
3. 选择 "**终端 > 运行任务**"，然后键入或选择 " **Dev Server**"。

### <a name="start-the-vs-code-debugger"></a>启动 VS 代码调试器

4. 选择 "**查看 > 调试**" 或输入**Ctrl + Shift + D**以切换到 "调试" 视图。
5. 从 "调试" 选项中，选择 " **Excel 桌面**"。
6. 选择**F5** （或从菜单中选择**Debug-> 启动调试**）开始调试。 将打开一个新的 Excel 工作簿，您的外接程序已旁加载并可供使用。

### <a name="start-debugging"></a>开始调试

1. 在 VS 代码中，打开源代码脚本文件（**函数 .js**或**函数**）。
2. 在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。
3. 在 Excel 工作簿中，输入使用自定义函数的公式。

此时执行将在您设置断点的代码行处停止。 现在，您可以逐步完成您的代码、设置监视和使用所需的任何与代码调试功能。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>在 Microsoft Edge 中将 VS 代码调试程序与 Excel 一起使用

您可以使用 VS 代码在 Microsoft Edge 浏览器的 Excel 中调试不带 UI 的自定义函数。 若要将 VS 代码与 Microsoft Edge 结合使用，必须[为 Microsoft edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)扩展安装调试器。

### <a name="run-your-add-in-from-vs-code"></a>从 VS 代码运行外接程序

1. 打开[VS 代码](https://code.visualstudio.com/)中的自定义函数根项目文件夹。
2. 选择 "**终端 > 运行任务**" 并键入或选择 "**监视**"。 这将监视和重建任何文件更改。
3. 选择 "**终端 > 运行任务**"，然后键入或选择 " **Dev Server**"。

### <a name="start-the-vs-code-debugger"></a>启动 VS 代码调试器

4. 选择 "**查看 > 调试**" 或输入**Ctrl + Shift + D**以切换到 "调试" 视图。
5. 从 "调试" 选项中，选择 " **Office Online （Microsoft Edge）**"。
6. 在 Microsoft Edge 浏览器中打开 Excel 并创建一个新的工作簿。
7. 在功能区中选择 "**共享**"，然后复制此新工作簿的 URL 的链接。
8. 选择**F5** （或从菜单中选择 "**调试" > 启动调试**）以开始调试。 将显示提示，询问您的文档的 URL。
9. 粘贴到工作簿的 URL 中，然后按 Enter。

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 选择功能区上的 "**插入**" 选项卡，然后在 "**外接程序**" 部分，选择 " **Office 外接程序**"。
2. 在 " **Office 外接程序**" 对话框中，选择 "**我的外**接程序" 选项卡，选择 "**管理我的外接**程序"，然后**上传我的外接程序**。
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

3. **浏览**到加载项清单文件，然后选择 "**上传**"。
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>设置断点
1. 在 VS 代码中，打开源代码脚本文件（**函数 .js**或**函数**）。
2. 在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。
3. 在 Excel 工作簿中，输入使用自定义函数的公式。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>使用浏览器开发人员工具在 Excel 网页版中调试自定义函数

您可以使用浏览器开发人员工具在 Excel 网页版中调试无 UI 的自定义函数。 以下步骤适用于 Windows 和 macOS。

### <a name="run-your-add-in-from-visual-studio-code"></a>从 Visual Studio Code 运行外接程序

1. 在[Visual Studio Code （VS code）](https://code.visualstudio.com/)中打开您的自定义函数根项目文件夹。
2. 选择 "**终端 > 运行任务**" 并键入或选择 "**监视**"。 这将监视和重建任何文件更改。
3. 选择 "**终端 > 运行任务**"，然后键入或选择 " **Dev Server**"。

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 打开 [Microsoft Office 网页版](https://office.live.com/)。
2. 打开一个新的 Excel 工作簿。
3. 打开功能区上的 "**插入**" 选项卡，然后在 "**外接程序**" 部分中，选择 " **Office 外接程序**"。
4. 在 " **Office 外接程序**" 对话框中，选择 "**我的外**接程序" 选项卡，选择 "**管理我的外接**程序"，然后**上传我的外接程序**。
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5. **转到**加载项清单文件，再选择“上传”****。
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> 在旁加载文档后，每次打开文档时它都将保留旁加载。

### <a name="start-debugging"></a>开始调试

1. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器 F12 将打开开发人员工具。
2. 在开发人员工具中，使用**Cmd + p**或**Ctrl + p**打开源代码脚本文件（**函数 .js**或**函数**）。
3. 在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。 

如果您需要更改代码，您可以在 VS 代码中进行编辑并保存所做的更改。 刷新浏览器以查看加载的更改。

## <a name="use-the-command-line-tools-to-debug"></a>使用命令行工具进行调试

如果未使用 VS 代码，则可以使用命令行（如 bash 或 PowerShell）运行外接程序。 您需要使用浏览器开发人员工具在 Excel 网页版中调试代码。 无法使用命令行调试桌面版本的 Excel。

1. 在命令行中运行 `npm run watch` ，以便在发生代码更改时监视和重建。
2. 打开第二个命令行窗口（运行监视时将阻止第一个命令行窗口。）

3. 如果要在 Excel 的桌面版本中启动外接程序，请运行以下命令
    
    `npm run start:desktop`
    
    或者，如果您更愿意在 Excel 网页上启动您的外接程序，请运行以下命令
    
    `npm run start:web`
    
    对于 web 上的 Excel，您还需要旁加载您的外接程序。 按照[旁加载您的外接程序](#sideload-your-add-in)中的步骤，旁加载你的外接程序。 然后继续转到下一节以开始调试。
    
4. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器 F12 将打开开发人员工具。
5. 在开发人员工具中，打开源代码脚本文件（**函数 .js**或**函数**）。 您的自定义函数代码可能位于文件末尾附近。
6. 在自定义函数源代码中，通过选择一行代码来应用断点。

如果您需要更改代码，您可以在 Visual Studio 中进行编辑并保存所做的更改。 刷新浏览器以查看加载的更改。

### <a name="commands-for-building-and-running-your-add-in"></a>用于生成和运行外接程序的命令

有几个可用的生成任务：
- `npm run watch`：用于开发的构建，在保存源文件时自动重建
- `npm run build-dev`：开发一次开发版本
- `npm run build`：生产的内部版本
- `npm run dev-server`：运行用于开发的 web 服务器

您可以使用以下任务在桌面或联机时开始调试。
- `npm run start:desktop`：在桌面上启动 Excel 并将您的外接程序。
- `npm run start:web`：在 web 上启动 Excel 并将您的外接程序。
- `npm run stop`：停止 Excel 和调试。

## <a name="next-steps"></a>后续步骤
了解有关无[UI 的自定义函数的身份验证实践](custom-functions-authentication.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数疑难解答](custom-functions-troubleshooting.md)
* [在 Excel 中处理自定义函数时出错](custom-functions-errors.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
