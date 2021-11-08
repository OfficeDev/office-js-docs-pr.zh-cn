---
title: 使用适用于加载项的开发人员工具调试Microsoft Edge 旧版
description: 使用开发人员工具在加载项中调试Microsoft Edge 旧版。
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: e3d0b77a6898dcefc7fba7c9d52eb739a2d685aa
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/03/2021
ms.locfileid: "60809075"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-legacy"></a>使用开发人员工具在加载项中调试Microsoft Edge 旧版

本文演示如何在满足以下条件时 (外接程序的 JavaScript 或 TypeScript) 调试客户端代码。

- 不能使用（或不希望）使用 IDE 中内置的工具进行调试;或者您遇到仅在外接程序在 IDE 外部运行时发生的问题。
- 您的计算机使用使用原始 Edge webview Windows EdgeHTML Office版本的组合。

> [!TIP]
> 有关在内部使用旧版边缘Visual Studio Code的信息，请参阅Microsoft Office[加载项调试器扩展Visual Studio Code。](debug-with-vs-extension.md)

若要确定你使用的浏览器，请参阅Office[使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。 

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> 若要安装使用Office旧版 Web 视图的旧版 webview 或强制当前版本的 Office 使用旧版 Edge，请参阅切换到旧版[边缘 Web 视图](#switch-to-the-edge-legacy-webview)。

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview"></a>使用开发人员工具预览版Microsoft Edge任务窗格加载项

1. 安装[Microsoft Edge DevTools Preview](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab)。  (历史原因，"预览"一词在名称中。 没有更新的版本.) 

   > [!NOTE]
   > 如果加载项具有执行函数的加载项[](../design/add-in-commands.md)命令，函数将在 Microsoft Edge DevTools 无法检测或附加到的隐藏浏览器进程中运行，因此本文中介绍的技术不能用于调试 函数中的代码。

1. [旁](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) 加载并运行外接程序。
1. 运行 Microsoft Edge 开发人员工具。
1. 在工具中，打开“**本地**”选项卡。加载项将按其名称列出。  (只有在 EdgeHTML 中运行的进程显示在选项卡上。该工具无法附加到在其他浏览器或 Web 视图（包括 Microsoft Edge (WebView2) 和 Internet Explorer (Trident) .) 

   :::image type="content" source="../images/edge-devtools-with-add-in-process.png" alt-text="Screenshot of Edge DevTools showing a process named legacy-edge-debugging.":::

1. 选择外接程序名称以在工具中打开它。
1. 打开“**调试器**”选项卡。
1. 通过以下步骤打开要调试的文件。

   1. 在调试器任务栏上，选择 **"在文件中显示查找"。** 这将打开搜索窗口。
   1. 在搜索框中输入要调试的文件的代码行。 它应该是不可能在任何其他文件中的内容。
   1. 选择刷新按钮。
   1. 在搜索结果中，选择行以在搜索结果上方的窗格中打开代码文件。

   :::image type="content" source="../images/open-file-in-edge-devtools.png" alt-text="Edge DevTools 调试选项卡的屏幕截图，其中 4 个部分标记为 A 到 D。":::

1. 若要设置断点，请选择代码文件中的代码行。 该断点在右下角 (调用堆栈) 注册。 代码文件的代码行可能也有一个红点，但无法可靠地显示。
1. 根据需要在加载项中执行函数以触发断点。

> [!TIP]
> 有关使用这些工具的信息，请参阅[EdgeHTML](/archive/microsoft-edge/legacy/developer/devtools-guide/)Microsoft Edge (开发人员) 工具。

## <a name="debug-a-dialog-in-an-add-in"></a>在加载项中调试对话框

如果加载项使用 Office 对话框 API，对话框将独立于任务窗格 (（如果有) 且工具必须附加到该流程）。 请按照以下步骤操作。

1. 运行加载项和工具。
1. 打开对话框，然后选择工具 **中的"** 刷新"按钮。 将显示对话框过程。 其名称来自 `<title>` 在对话框中打开的 HTML 文件的 元素。
1. 选择打开它并调试的过程，如使用开发人员工具预览调试任务窗格外接程序Microsoft Edge[所述](#debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview)。

   :::image type="content" source="../images/edge-devtools-with-add-in-and-dialog-processes.png" alt-text="Screenshot of Edge DevTools showing a process named My Dialog.":::

## <a name="switch-to-the-edge-legacy-webview"></a>切换到旧版边缘 Web 视图

有两种方法可以切换旧版边缘 Web 视图。 可以在命令提示符中运行一个简单的命令，也可以安装默认Office旧版边缘的客户端版本。 我们建议使用第一种方法。 但你应在以下方案中使用第二个。

- 您的项目是使用 Visual Studio IIS 开发的。 它不是基于node.js的。
- 你想要在测试中保持绝对可靠。
- 如果由于任何原因，命令行工具不起作用。

### <a name="switch-via-the-command-line"></a>通过命令行进行切换

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-edge-legacy"></a>安装使用旧版Office版本的客户端

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]
