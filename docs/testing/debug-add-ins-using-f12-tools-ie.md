---
title: 使用适用于 Internet Explorer 的开发人员工具调试加载项
description: 使用开发人员工具在加载项中调试Internet Explorer。
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: fb830f90c23b64e19420c73bee695bef669d93d1
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/03/2021
ms.locfileid: "60809084"
---
# <a name="debug-add-ins-using-developer-tools-in-internet-explorer"></a>使用开发人员工具在加载项中调试Internet Explorer

本文演示如何在满足以下条件时 (外接程序的 JavaScript 或 TypeScript) 调试客户端代码。

- 不能使用（或不希望）使用 IDE 中内置的工具进行调试;或者您遇到仅在外接程序在 IDE 外部运行时发生的问题。
- 您的计算机使用使用 Web 视图控件 Trident Windows Office和 Internet Explorer版本的组合。

若要确定计算机上所使用的浏览器，请参阅浏览器[Office外接程序。](../concepts/browsers-used-by-office-web-add-ins.md)

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> 若要安装使用 Office Webview 的 Internet Explorer 或强制当前版本使用 Internet Explorer，请参阅切换到[Internet Explorer 11 webview。](#switch-to-the-internet-explorer-11-webview)

## <a name="debug-a-task-pane-add-in-using-the-f12-tools"></a>使用 F12 工具调试任务窗格外接程序

Windows 10 11 包括一个称为"F12"的 Web 开发工具，因为它最初是按 F12 在 Internet Explorer。 F12 现在是一个独立的应用程序，用于在外接程序在 Web 视图控件 Trident 中运行时Internet Explorer调试外接程序。 应用程序在早期版本的 Windows 中不可用。

> [!NOTE]
> 如果加载项具有执行函数的加载项[](../design/add-in-commands.md)命令，函数将在 F12 工具无法检测或附加到的隐藏浏览器进程中运行，因此本文中所述的技术不能用于调试 函数中的代码。

以下步骤是调试外接程序的说明。 如果只想测试 F12 工具本身，请参阅示例加载项 [以测试 F12 工具](#example-add-in-to-test-the-f12-tools)。

1. [旁](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) 加载并运行外接程序。
1. 启动与版本对应的 F12 开发Office。

   - 对于 32 位版 Office，请使用 C:\Windows\System32\F12\IEChooser.exe
   - 对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\IEChooser.exe

   IEChooser 将打开一个名为 **"选择目标以调试"的窗口**。 加载项将显示在由加载项主页的文件名命名的窗口中。 在下面的屏幕截图中，它是 `Home.html` 。 只显示运行在 Internet Explorer 或 Trident 中的进程。 该工具无法附加到在其他浏览器或 Web 视图（包括 web 视图）中运行的进程Microsoft Edge。

    :::image type="content" source="../images/choose-target-to-debug.png" alt-text="IEChooser 屏幕，列出了Internet Explorer和 Trident 进程。一个名为 Home.html。":::

1. 选择加载项流程;即，其主页文件名。 此操作将 F12 工具附加到进程并打开主 F12 用户界面。
1. 打开“**调试器**”选项卡。
1. 在选项卡的左上角，调试器工具功能区正下方有一个小文件夹图标。 选择此选项可打开外接程序中的文件的下拉列表。 示例如下。

    :::image type="content" source="../images/f12-file-dropdown.png" alt-text="Screenshot of upper left corner of debugger tab with a folder drop down open and a list of files.":::

1. 选择要调试的文件，该文件将在"调试器"选项卡的 (左) **中** 打开。如果你使用的是更改文件名称的传输器、捆绑程序或微型程序，它将具有实际加载的最终名称，而不是原始源文件名。

1. 滚动到要设置断点的行，然后单击行号左侧的边距。 您将在该行左侧看到一个红点，相应的行将显示在右下窗格的"断点"选项卡中。 例如，下面的屏幕截图。

    :::image type="content" source="../images/debugger-home-js-02.png" alt-text="断点在文件home.js调试程序。":::

1. 根据需要在加载项中执行函数以触发断点。 命中断点时，断点的红点上将出现一个右箭头。 例如，下面的屏幕截图。

    :::image type="content" source="../images/debugger-home-js-01.png" alt-text="调试器，其结果来自触发的断点。":::

> [!TIP]
> 有关使用 F12 工具的信息，请参阅使用调试器检查正在运行的[JavaScript。](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))

### <a name="example-add-in-to-test-the-f12-tools"></a>测试 F12 工具的示例外接程序

此示例使用 Word 和从 AppSource 获取的免费加载项。

1. 打开 Word 并选择空白文档。
1. 在"**插入**"选项卡上的"外接程序"组中，选择"我的外接程序"以打开 **"Office 外接程序**"对话框，然后选择"**存储"** 选项卡。
1. 选择 **QR4Office** 外接程序。 它将在任务窗格中打开。
1. 启动与版本对应的 F12 开发工具Office如上一节中所述。
1. 在 F12 窗口中，选择 **"Home.html"。**
1. 在调试 **器** 选项卡中， **打开Home.js如** 上一节中所述。
1. 设置第 310 行和 312 行上的断点。
1. 在外接程序中，选择"插入 **"** 按钮。 命中一个或多个断点。

## <a name="debug-a-dialog-in-an-add-in"></a>在加载项中调试对话框

如果加载项使用 Office 对话框 API，对话框将独立于任务窗格 (（如果有) 且工具必须附加到该流程）。 请按照以下步骤操作。

1. 运行加载项和工具。 
1. 打开对话框，然后选择工具 **中的"** 刷新"按钮。 将显示对话框过程。 其名称是在对话框中打开的文件的文件名。
1. 选择打开并调试的过程，如使用 [F12](#debug-a-task-pane-add-in-using-the-f12-tools)工具调试任务窗格外接程序一节中所述。

## <a name="switch-to-the-internet-explorer-11-webview"></a>切换到 Internet Explorer 11 Webview

有两种方法可以切换 web Internet Explorer视图。 可以在命令提示符中运行一个简单的命令，也可以安装默认Office使用Internet Explorer版本。 我们建议使用第一种方法。 但你应在以下方案中使用第二个。

- 您的项目是使用 Visual Studio IIS 开发的。 它不是基于node.js的。
- 你想要在测试中保持绝对可靠。
- 如果由于任何原因，命令行工具不起作用。

### <a name="switch-via-the-command-line"></a>通过命令行进行切换

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>安装使用Office版本的Internet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>另请参阅

- [使用调试器检查正在运行的 JavaScript](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [使用 F12 开发人员工具](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
