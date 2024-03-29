---
title: 在非共享运行时中调试自定义函数
description: 了解如何调试不使用共享运行时的 Excel 自定义函数。
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4e9a1c7c521838b65d2df8d75e8eea5643b0a80b
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797636"
---
# <a name="custom-functions-debugging"></a>自定义函数调试

本文讨论仅针对 **不使用 [共享运行时的](../develop/configure-your-add-in-to-use-a-shared-runtime.md)** 自定义函数进行调试。 若要调试使用共享运行时的自定义函数加载项，请参阅 [配置 Office 外接程序以使用共享 JavaScript 运行时：调试](../develop/configure-your-add-in-to-use-a-shared-runtime.md#debug)。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

> [!TIP]
> 此调试过程不适用于使用 **Office 加载项项目创建的项目，该项目** 包含 Yeoman 生成器中仅包含清单选项。 本文后面提到的脚本未随该选项一起安装。 若要调试使用此选项创建的加载项，请根据需要查看以下文章之一中的说明。
>
> - [使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项](../testing/debug-add-ins-using-devtools-edge-chromium.md)
> - [在 Internet Explorer 中使用开发人员工具调试加载项](../testing/debug-add-ins-using-f12-tools-ie.md)
> - [在 Mac 上调试 Office 加载项](../testing/debug-office-add-ins-on-ipad-and-mac.md)

为不使用共享运行时的外接程序调试自定义函数的过程因目标平台 (Windows、Mac 或 Web) 而异，无论是使用Visual Studio Code还是使用其他 IDE，以及开发计算机的操作系统。 使用下表中的链接访问本文中与调试方案相关的部分。 在此表中，“CF-NSR”是指非共享运行时中的自定义函数。

| **目标平台** | **Visual Studio Code** | **其他 IDE** |
|--------------|-------------|-------------|
| Windows 版 Excel | [在 Windows 上使用适用于 Excel 的 VS Code 调试器](#use-the-vs-code-debugger-for-excel-on-windows) | 不支持在 VS Code 外部调试 CF-NSR。 针对Excel web 版进行调试。 |
| Excel 网页版 | Windows 开发计算机： [在 Microsoft Edge 中使用适用于 Excel 的 VS Code 调试器](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)</br>Mac 或 Windows 开发计算机： [使用 VS Code 和浏览器开发工具](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web) | [使用命令行工具](#use-the-command-line-tools-to-debug)|
| Mac 版 Excel |  不支持 CF-NSR 的 VS Code 调试。 针对Excel web 版进行调试。 | [使用命令行工具](#use-the-command-line-tools-to-debug)|

> [!NOTE]
> 本文主要介绍在使用Visual Studio Code编辑、运行任务和使用调试视图的上下文中进行调试。 如果使用的是其他编辑器或命令行工具，请参阅本文末尾 [用于生成和运行外接](#commands-for-building-and-running-your-add-in) 程序的命令。

## <a name="use-the-vs-code-debugger-for-excel-on-windows"></a>在 Windows 上使用适用于 Excel 的 VS Code 调试器

可以使用 VS Code 调试在桌面上的 Office Excel 中不使用共享运行时的自定义函数。

> [!IMPORTANT]
> 以下调试步骤存在已知问题。 这些步骤适用于在 Yeoman 生成器 **中安装了 Excel 自定义函数外接程序项目选项的项目** ，其中 **TypeScript** 已选中为脚本类型，但这些步骤不适用于已选中 **JavaScript** 作为脚本类型的项目。 有关其他信息，请参阅 [OfficeDev/office-js-docs-pr 问题 #3355](https://github.com/OfficeDev/office-js-docs-pr/issues/3355)。

### <a name="run-your-add-in-from-vs-code"></a>从 VS Code 运行加载项

1. 在 [VS Code](https://code.visualstudio.com/) 中打开自定义函数根项目文件夹。
1. 选择 **终端>运行任务** 并键入或选择 **“监视**”。 这将监视和重新生成任何文件更改。
1. 选择 **终端>运行任务** 并键入或选择 **开发服务器**。

### <a name="start-the-vs-code-debugger"></a>启动 VS Code 调试器

1. 选择 **“视图>运行** 或输入 **Ctrl+Shift+D** 以切换到调试视图。
1. 在 **“运行和调试** ”下拉菜单中，选择 **Excel Desktop (自定义函数)**。

    :::image type="content" source="../images/custom-functions-run-and-debug-menu.jpg" alt-text="显示“运行和调试”下拉菜单中 (自定义函数) 的 Excel Desktop 的屏幕截图。":::

1. 选择 **F5** (或从菜单中选择 **“运行>开始调试** ”，) 开始调试。 新的 Excel 工作簿将打开，加载项已旁加载并可供使用。

### <a name="start-debugging"></a>开始调试

1. 在 VS Code 中，打开源代码脚本文件 (**functions.js** 或 **functions.ts**) 。
2. 在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。
3. 在 Excel 工作簿中，输入使用自定义函数的公式。

此时，执行将停止在设置断点的代码行上。 现在，可以逐步完成代码，设置监视，并使用所需的任何 VS Code 调试功能。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>在 Microsoft Edge 中使用适用于 Excel 的 VS Code 调试器

可以使用 VS Code 调试在 Microsoft Edge 浏览器上的 Excel 中不使用共享运行时的自定义函数。 若要将 VS Code 与 Microsoft Edge 配合使用，必须安装[适用于 Visual Studio Code 的 Microsoft Edge DevTools 扩展](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension)。

### <a name="run-your-add-in-from-vs-code"></a>从 VS Code 运行加载项

1. 在 [VS Code](https://code.visualstudio.com/) 中打开自定义函数根项目文件夹。
1. 选择 **终端>运行任务** 并键入或选择 **“监视**”。 这将监视和重新生成任何文件更改。
1. 选择 **终端>运行任务** 并键入或选择 **开发服务器**。

### <a name="start-the-vs-code-debugger"></a>启动 VS Code 调试器

1. 选择 **“视图>运行** 或输入 **Ctrl+Shift+D** 以切换到调试视图。
1. 从调试选项中，选择 **Office Online (Edge Chromium)**。
1. 在 Microsoft Edge 浏览器中打开 Excel 并创建新的工作簿。
1. 选择功能区中的 **“共享** ”，并复制此新工作簿的 URL 链接。
1. 选择 **F5** (或从菜单中选择 **“运行>开始调试** ”，) 开始调试。 将显示一个提示，要求输入文档的 URL。
1. 粘贴工作簿的 URL，然后按 Enter。

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 选择功能区上的 **“插入** ”选项卡， **在“加载项”** 部分中，选择 **“Office 加载项**”。
2. 在 **“Office 加载项** ”对话框中，选择 **“我的外接程序”** 选项卡，选择 **“管理我的外接程序**”，然后 **上传“我的外接程序**”。
  
    ![Office 加载项对话框右上角有一个下拉列表，上面写着“管理我的加载项”，下面是一个下拉列表，选项为“上传我的加载项”。](../images/office-add-ins-my-account.png)

3. **浏览** 到加载项清单文件，然后选择 **“上传**”。
  
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>设置断点

1. 在 VS Code 中，打开源代码脚本文件 (**functions.js** 或 **functions.ts**) 。
2. 在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。
3. 在 Excel 工作簿中，输入使用自定义函数的公式。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>使用浏览器开发人员工具调试Excel web 版中的自定义函数

可以使用浏览器开发人员工具调试在Excel web 版中不使用共享运行时的自定义函数。 以下步骤适用于 Windows 和 macOS。

### <a name="run-your-add-in-from-visual-studio-code"></a>从Visual Studio Code运行加载项

1. 在 [VISUAL STUDIO CODE (VS Code) ](https://code.visualstudio.com/)中打开自定义函数根项目文件夹。
2. 选择 **终端>运行任务** 并键入或选择 **“监视**”。 这将监视和重新生成任何文件更改。
3. 选择 **终端>运行任务** 并键入或选择 **开发服务器**。

### <a name="sideload-your-add-in"></a>旁加载加载项

1. 打开[Office web 版](https://office.live.com/)。
2. 打开新的 Excel 工作簿。
3. 打开功能区上的 **“插入** ”选项卡， **并在“加载项”** 部分中选择 **“Office 加载项**”。
4. 在 **“Office 加载项** ”对话框中，选择 **“我的外接程序”** 选项卡，选择 **“管理我的外接程序**”，然后 **上传“我的外接程序**”。
  
    ![Office 加载项对话框右上角有一个下拉列表，上面写着“管理我的加载项”，下面是一个下拉列表，选项为“上传我的加载项”。](../images/office-add-ins-my-account.png)

5. **转到** 加载项清单文件，再选择“上传”。
  
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> 旁加载到文档后，每次打开文档时，文档都将保持旁加载状态。

### <a name="start-debugging"></a>开始调试

1. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器，F12 将打开开发人员工具。
2. 在开发人员工具中，使用 **Cmd+P** 或 **Ctrl+P** (functions.js或 **functions.ts** **)** 打开源代码脚本文件。
3. 在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。 

如果需要更改代码，可以在 VS Code 中进行编辑并保存更改。 刷新浏览器以查看已加载的更改。

## <a name="use-the-command-line-tools-to-debug"></a>使用命令行工具进行调试

如果不使用 VS Code，则可以使用命令行 (如 bash 或 PowerShell) 来运行加载项。 需要使用浏览器开发人员工具在Excel web 版中调试代码。 不能使用命令行调试桌面版本的 Excel。

1. 从命令行运行 `npm run watch` ，在发生代码更改时监视和重新生成。
2. 打开第二个命令行窗口 (第一个命令行窗口将在运行 watch 时被阻止。) 

3. 如果要在桌面版 Excel 中启动外接程序，请运行以下命令。
  
    `npm run start:desktop`
  
    或者，如果想要在 Excel web 版中启动外接程序，请运行以下命令。
  
    `npm run start:web -- --document {url}` (OneDrive 或 SharePoint) 上 Excel 文件的 URL 在哪里`{url}`
  
    如果加载项未旁加载文档，请按照旁 [加载加](#sideload-your-add-in) 载项中的步骤旁加载加载项。 然后继续下一部分开始调试。
  
4. 在浏览器中打开开发人员工具。 对于 Chrome 和大多数浏览器，F12 将打开开发人员工具。
5. 在开发人员工具中，打开源代码脚本文件 (**functions.js** 或 **functions.ts**) 。 自定义函数代码可能位于文件末尾附近。
6. 在自定义函数源代码中，通过选择代码行应用断点。

如果需要更改代码，可以在 Visual Studio 中进行编辑并保存更改。 刷新浏览器以查看已加载的更改。

### <a name="commands-for-building-and-running-your-add-in"></a>用于生成和运行外接程序的命令

有几个生成任务可用。

- `npm run watch`：用于开发的生成并在保存源文件时自动重新生成
- `npm run build-dev`：用于开发的生成一次
- `npm run build`：生产版本
- `npm run dev-server`：运行用于开发的 Web 服务器

可以使用以下任务在桌面或联机版上开始调试。

- `npm run start:desktop`：在桌面上启动 Excel 并旁加载加载项。
- `npm run start:web -- --document {url}` (OneDrive 或 SharePoint) 上 Excel 文件的 URL 的位置`{url}`：启动Excel web 版并旁加载加载项。
- `npm run stop`：停止 Excel 和调试。

## <a name="next-steps"></a>后续步骤

了解 [不使用共享运行时的自定义函数的身份验证](custom-functions-authentication.md)。

## <a name="see-also"></a>另请参阅

* [自定义函数故障排除](custom-functions-troubleshooting.md)
* [在 Excel 中处理自定义函数时出错](custom-functions-errors.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
