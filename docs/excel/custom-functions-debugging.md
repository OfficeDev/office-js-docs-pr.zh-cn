---
title: 无 UI 自定义函数调试
description: 了解如何调试不使用Excel窗格的自定义函数。
ms.date: 01/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 28f597d7a4dd6befac7b2df6d1d6c7f8bfc3033f
ms.sourcegitcommit: 33824aa3995a2e0bcc6d8e67ada46f296c224642
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/12/2022
ms.locfileid: "61765946"
---
# <a name="ui-less-custom-functions-debugging"></a>无 UI 自定义函数调试

本文仅讨论不使用任务窗格或其他用户界面元素的自定义函数的调试 (无 UI 自定义函数) 。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

在Windows：

- [Excel桌面和Visual Studio Code (VS Code) 调试器](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel web 版和VS Code调试器](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel web 版和浏览器工具](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [命令行](#use-the-command-line-tools-to-debug)

在 Mac 上：

- [Excel web 版和浏览器工具](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [命令行](#use-the-command-line-tools-to-debug)

> [!NOTE]
> 为简单起见，本文介绍在使用 Visual Studio Code编辑、运行任务以及在某些情况下使用调试视图的上下文中进行调试。 如果使用的是其他编辑器或命令行工具，请参阅本文末尾的命令行说明[](#commands-for-building-and-running-your-add-in)。

## <a name="requirements"></a>Requirements

此调试过程 **仅适用于无** UI 的自定义函数，这些函数不使用任务窗格或其他 UI 元素。 可以按照在[Excel](../tutorials/excel-tutorial-create-custom-functions.md)中创建自定义函数教程中的步骤创建无 UI 自定义函数，然后删除由 Office 加载项的[Yeoman](https://www.npmjs.com/package/generator-office)生成器安装的所有任务窗格和 UI 元素。

请注意，此调试过程与使用共享运行时 的自定义函数 [项目不兼容](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>使用 VS Code 桌面版Excel调试程序

可以使用VS Code调试桌面上的 Office Excel UI 无 UI 自定义函数。

> [!NOTE]
> 适用于 Mac 的桌面调试不可用，但可以使用浏览器工具和命令行[](#use-the-command-line-tools-to-debug)来调试Excel web 版) 。

### <a name="run-your-add-in-from-vs-code"></a>从应用程序运行VS Code

1. 打开自定义函数根项目[文件夹，VS Code。](https://code.visualstudio.com/)
1. 选择 **"终端>运行任务**"，然后键入或选择"监视 **"。** 这将监视并重新生成任何文件更改。
1. Choose **Terminal > Run Task** and type or select Dev **Server**.

### <a name="start-the-vs-code-debugger"></a>启动 VS Code调试程序

1. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
1. From the Run drop-down menu， choose **Excel Desktop (Custom Functions)**.
1. 选择 **F5** (，或者从>开始调试"菜单中选择") 开始调试"。 新的Excel工作簿将打开，并且外接程序已旁加载并可供使用。

### <a name="start-debugging"></a>开始调试

1. 在VS Code中，打开源代码脚本文件 **(functions.js****或 functions.ts**) 。
2. [在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。
3. 在Excel工作簿中，输入使用自定义函数的公式。

此时，将在设置断点的代码行上停止执行。 现在，你可以逐步调试代码、设置监视以及使用VS Code调试功能。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>使用 VS Code 中的Excel调试Microsoft Edge

可以使用VS Code在浏览器上的 Excel 中调试无 UI Microsoft Edge函数。 若要将 VS Code与 Microsoft Edge 一Microsoft Edge，必须安装适用于 Microsoft Edge[的 DevTools Visual Studio Code。](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension)

### <a name="run-your-add-in-from-vs-code"></a>从应用程序运行VS Code

1. 打开自定义函数根项目[文件夹，VS Code。](https://code.visualstudio.com/)
2. 选择 **"终端>运行任务**"，然后键入或选择"监视 **"。** 这将监视并重新生成任何文件更改。
3. Choose **Terminal > Run Task** and type or select Dev **Server**.

### <a name="start-the-vs-code-debugger"></a>启动 VS Code调试程序

1. Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.
1. 从"调试"选项中，选择"Office **Online (Edge Chromium) "。**
1. 在Excel中打开Microsoft Edge新建工作簿。
1. 在 **功能** 区中选择"共享"，并复制此新工作簿的 URL 链接。
1. 选择 **F5** (**或从>** 开始调试"菜单中选择") 开始调试"。 将出现一个提示，询问文档的 URL。
1. 粘贴工作簿的 URL，然后按 Enter。

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 选择功能 **区** 上的"插入"选项卡，在"外接程序"部分，选择"Office **外接程序"。**
2. 在 **"Office** 加载项"对话框中，选择"**我的** 加载项"选项卡，选择"管理我的加载项"，Upload"**我的加载项"。**
  
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

3. **浏览** 到外接程序清单文件，然后选择 **"Upload"。**
  
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>设置断点

1. 在VS Code中，打开源代码脚本文件 **(functions.js****或 functions.ts**) 。
2. [在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。
3. 在Excel工作簿中，输入使用自定义函数的公式。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>使用浏览器开发人员工具在浏览器中调试自定义Excel web 版

可以使用浏览器开发人员工具在浏览器中调试无 UI Excel web 版。 以下步骤适用于 Windows 和 macOS。

### <a name="run-your-add-in-from-visual-studio-code"></a>从应用程序运行Visual Studio Code

1. 打开自定义函数根项目[文件夹，Visual Studio Code (VS Code) 。 ](https://code.visualstudio.com/)
2. 选择 **"终端>运行任务**"，然后键入或选择"监视 **"。** 这将监视并重新生成任何文件更改。
3. Choose **Terminal > Run Task** and type or select Dev **Server**.

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 打开[Office web 版。](https://office.live.com/)
2. 打开一个新的Excel工作簿。
3. 打开功能 **区** 上的"插入"选项卡，在"外接程序"部分，选择"Office **外接程序"。**
4. 在 **"Office** 加载项"对话框中，选择"**我的** 加载项"选项卡，选择"管理我的加载项"，Upload"**我的加载项"。**
  
    ![The Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

5. **转到** 加载项清单文件，再选择“上传”。
  
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> 将文档旁加载后，每次打开文档时，文档都会保持旁加载状态。

### <a name="start-debugging"></a>开始调试

1. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。
2. 在开发人员工具中，使用 **Cmd+P** 或 **Ctrl+P** (functions.js **或 functions.ts**) 。 
3. [在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。 

如果需要更改代码，可以在该代码中进行VS Code并保存更改。 刷新浏览器以查看已加载的更改。

## <a name="use-the-command-line-tools-to-debug"></a>使用命令行工具进行调试

如果未使用 VS Code，可以使用命令行 (如 Bash 或 PowerShell) 运行外接程序。 你将需要使用浏览器开发人员工具在 Excel web 版 中调试代码。 不能使用命令行调试桌面Excel版本。

1. 从命令行运行 `npm run watch` 以观察代码发生更改时并重新生成代码。
2. 打开第二个命令行窗口 (运行 watch.) 

3. 如果要在桌面版本的外接程序中启动Excel，请运行以下命令。
  
    `npm run start:desktop`
  
    或者，如果你想要在外接程序中启动Excel web 版运行以下命令。
  
    `npm run start:web`
  
    例如Excel web 版你还需要旁加载你的外接程序。 按照旁加载 [加载项中的步骤](#sideload-your-add-in) 旁加载加载项。 然后继续下一部分以开始调试。
  
4. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。
5. 在开发人员工具中，打开源代码脚本文件 **(functions.js****或 functions.ts**) 。 自定义函数代码可能位于文件的末尾附近。
6. 在自定义函数源代码中，通过选择一行代码来应用断点。

如果需要更改代码，可以在该代码中进行Visual Studio并保存更改。 刷新浏览器以查看已加载的更改。

### <a name="commands-for-building-and-running-your-add-in"></a>用于生成和运行加载项的命令

有几个可用的生成任务。

- `npm run watch`：用于开发内部版本，在保存源文件时自动重新生成
- `npm run build-dev`：生成一次用于开发
- `npm run build`：用于生产内部版本
- `npm run dev-server`：运行用于开发的 Web 服务器

可以使用以下任务在桌面或联机上开始调试。

- `npm run start:desktop`：Excel启动加载项，并旁加载加载项。
- `npm run start:web`：Excel web 版并旁加载外接程序。
- `npm run stop`：停止Excel调试。

## <a name="next-steps"></a>后续步骤

了解 [无 UI 自定义函数的身份验证做法](custom-functions-authentication.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数疑难解答](custom-functions-troubleshooting.md)
* [在 Excel 中处理自定义函数时出错](custom-functions-errors.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
