---
ms.date: 07/10/2020
description: 了解如何调试不使用任务窗格的 Excel 自定义函数。
title: 无 UI 自定义函数调试
localization_priority: Normal
ms.openlocfilehash: 00065a465a22f83891dfb207943102b079e96a0f
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178074"
---
# <a name="ui-less-custom-functions-debugging"></a>无 UI 自定义函数调试

调试不使用任务窗格或其他用户界面元素的自定义函数 (无 UI 自定义函数) 可通过多种方法完成，具体取决于你使用的平台。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

在 Windows 上：
- [Excel Desktop and Visual Studio Code (VS Code) debugger](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel 网页和 VS 代码调试程序](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel 网页和浏览器工具](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [命令行](#use-the-command-line-tools-to-debug)

在 Mac 上：
- [Excel 网页和浏览器工具](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [命令行](#use-the-command-line-tools-to-debug)

> [!NOTE]
> 为简单起见，本文介绍在使用 Visual Studio 代码编辑、运行任务的情况下进行调试，在某些情况下，还使用调试视图。 如果使用的是其他编辑器或命令行工具，请参阅本文末尾的命令行说明[](#commands-for-building-and-running-your-add-in)。

## <a name="requirements"></a>要求

在开始调试之前，应该使用 Office 加载项 [的 Yeoman](https://github.com/OfficeDev/generator-office) 生成器创建自定义函数项目。 有关如何创建自定义函数项目的指南，请参阅 [自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>使用适用于 Excel Desktop 的 VS 代码调试程序

您可以使用 VS Code 在桌面上的 Office Excel 中调试无 UI 自定义函数。

> [!NOTE]
> 适用于 Mac 的桌面调试不可用，但可以使用浏览器工具和命令行来调试 [Excel 网页](#use-the-command-line-tools-to-debug) 版) 。

### <a name="run-your-add-in-from-vs-code"></a>从 VS Code 运行加载项

1. 在 VS Code 中打开自定义函数根项目 [文件夹](https://code.visualstudio.com/)。
2. 选择 **"终端>运行任务**"，然后键入或选择"**监视"。** 这将监视并重新生成任何文件更改。
3. 选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**

### <a name="start-the-vs-code-debugger"></a>启动 VS 代码调试程序

4. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
5. From the Run drop-down menu， choose **Excel Desktop (Edge Chromium)**.
6. 选择 **F5** (，或者从 **>开始调试** "菜单中选择") 开始调试"。 新的 Excel 工作簿将在外接程序已旁加载且可供使用时打开。

### <a name="start-debugging"></a>开始调试

1. 在 VS Code 中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。
2. [在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。
3. 在 Excel 工作簿中，输入使用自定义函数的公式。

此时，将在设置断点的代码行上停止执行。 现在，你可以逐步调试代码、设置监视以及使用所需的任何 VS 代码调试功能。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>在 Microsoft Edge 中为 Excel 使用 VS 代码调试程序

您可以使用 VS Code 在 Microsoft Edge 浏览器的 Excel 中调试无 UI 自定义函数。 若要将 VS Code 与 Microsoft Edge 一同使用，必须安装 [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) 扩展。

### <a name="run-your-add-in-from-vs-code"></a>从 VS Code 运行加载项

1. 在 VS Code 中打开自定义函数根项目 [文件夹](https://code.visualstudio.com/)。
2. 选择 **"终端>运行任务**"，然后键入或选择"**监视"。** 这将监视并重新生成任何文件更改。
3. 选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**

### <a name="start-the-vs-code-debugger"></a>启动 VS 代码调试程序

4. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
5. 从"调试"选项中，选择 **"Office Online (Edge Chromium) "。**
6. 在 Microsoft Edge 浏览器中打开 Excel 并创建新的工作簿。
7. 在 **功能** 区中选择"共享"，并复制此新工作簿的 URL 链接。
8. 选择 **F5** (**或从>** 开始调试"菜单中选择") 开始调试"。 将出现一个提示，询问文档的 URL。
9. 粘贴工作簿的 URL，然后按 Enter。

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 选择功能 **区** 上的"插入"选项卡，在 **"外接程序"** 部分，选择 **"Office 外接程序"。**
2. 在 **"Office 外接程序"** 对话框中，选择 **"我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，然后选择"**上载我的外接程序"。**
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

3. **浏览** 到外接程序清单文件， **然后选择上载**。
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>设置断点
1. 在 VS Code 中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。
2. [在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。
3. 在 Excel 工作簿中，输入使用自定义函数的公式。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>使用浏览器开发人员工具调试 Excel 网页版中的自定义函数

可以使用浏览器开发人员工具在 Excel 网页版中调试无 UI 自定义函数。 以下步骤适用于 Windows 和 macOS。

### <a name="run-your-add-in-from-visual-studio-code"></a>从代码运行Visual Studio加载项

1. 打开自定义函数根项目文件夹，Visual Studio [代码 ](https://code.visualstudio.com/) (VS Code) 。
2. 选择 **"终端>运行任务**"，然后键入或选择"**监视"。** 这将监视并重新生成任何文件更改。
3. 选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 在[Web 上打开 Office。](https://office.live.com/)
2. 打开一个新的 Excel 工作簿。
3. 打开功能 **区** 上的"插入"选项卡，在"**外接程序**"部分，选择 **"Office 外接程序"。**
4. 在 **"Office 外接程序"** 对话框中，选择 **"我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，然后选择"**上载我的外接程序"。**
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5. **转到** 加载项清单文件，再选择“上传”。
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> 旁加载文档后，每次打开文档时，文档都会保持旁加载状态。

### <a name="start-debugging"></a>开始调试

1. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。
2. 在开发人员工具中，使用 **Cmd+P** 或 **Ctrl+P** (functions.js或 **functions.ts**) 。
3. [在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。 

如果需要更改代码，可以在 VS Code 中编辑并保存更改。 刷新浏览器以查看已加载的更改。

## <a name="use-the-command-line-tools-to-debug"></a>使用命令行工具进行调试

如果不使用 VS Code，可以使用命令行命令 (Bash 或 PowerShell) 运行外接程序。 你需要使用浏览器开发人员工具在 Excel 网页版中调试代码。 不能使用命令行调试桌面版 Excel。

1. 从命令行运行 `npm run watch` 以观察代码发生更改时并重新生成代码。
2. 打开第二个命令行窗口 (运行 watch.) 

3. 如果要在桌面版 Excel 中启动加载项，请运行以下命令
    
    `npm run start:desktop`
    
    或者，如果你想要在 Excel 网页中启动加载项，请运行以下命令
    
    `npm run start:web`
    
    对于 Excel 网页应用，还需要旁加载加载项。 按照旁加载 [加载项中的步骤](#sideload-your-add-in) 旁加载加载项。 然后继续下一部分以开始调试。
    
4. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。
5. 在开发人员工具中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。  自定义函数代码可能位于文件的末尾附近。
6. 在自定义函数源代码中，通过选择一行代码来应用断点。

如果需要更改代码，可以在该代码中进行Visual Studio并保存更改。 刷新浏览器以查看已加载的更改。

### <a name="commands-for-building-and-running-your-add-in"></a>用于生成和运行加载项的命令

有几种可用的生成任务：
- `npm run watch`：用于开发内部版本，在保存源文件时自动重新生成
- `npm run build-dev`：生成一次用于开发
- `npm run build`：用于生产内部版本
- `npm run dev-server`：运行用于开发的 Web 服务器

可以使用以下任务在桌面或联机上开始调试。
- `npm run start:desktop`：在桌面上启动 Excel 并旁加载外接程序。
- `npm run start:web`：在 Web 上启动 Excel 并旁加载外接程序。
- `npm run stop`：停止 Excel 和调试。

## <a name="next-steps"></a>后续步骤
了解 [无 UI 自定义函数的身份验证做法](custom-functions-authentication.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数疑难解答](custom-functions-troubleshooting.md)
* [在 Excel 中处理自定义函数时出错](custom-functions-errors.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
