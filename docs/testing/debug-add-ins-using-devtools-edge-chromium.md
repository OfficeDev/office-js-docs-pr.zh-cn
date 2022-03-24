---
title: 使用 WebView2 开发人员工具调试Microsoft Edge加载项
description: 使用基于 WebView2 的 webView2 Microsoft Edge工具 (Chromium加载项) 。
ms.date: 11/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7cd4e3d3279ef605c5a9ef5fc21a678984d978e5
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744687"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-chromium-based"></a>使用 Microsoft Edge（基于 Chromium）中的开发人员工具调试加载项

本文演示如何在满足以下条件时 (外接程序的 JavaScript 或 TypeScript) 调试客户端代码。

- 不能使用（或不希望）使用 IDE 中内置的工具进行调试;或者您遇到仅在外接程序在 IDE 外部运行时发生的问题。
- 您的计算机使用使用基于 Windows 的边缘Office WebView2 (Chromium的) 和) 版本。

> [!TIP]
> 有关在 Visual Studio Code 内使用 Edge WebView2 (Chromium) 进行调试的信息，请参阅使用 Visual Studio Code 调试 Windows 上的外接程序和 Microsoft Edge [WebView2 (Chromium基于) ](debug-desktop-using-edge-chromium.md)。

若要确定你使用的浏览器，请参阅浏览器[Office外接程序](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools"></a>使用基于加载项的开发人员工具Microsoft Edge (Chromium任务) 加载项

> [!NOTE]
> 如果加载项具有执行函数的加载项[](../design/add-in-commands.md)命令，函数将在隐藏的浏览器进程中运行，无法从该进程中启动基于 Microsoft Edge (Chromium 的) 开发人员工具，因此本文中介绍的技术不能用于调试 函数中的代码。

1. [旁](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) 加载并运行外接程序。
1. 通过Microsoft Edge (Chromium之一) 基于 Web 应用程序运行开发人员工具：

   - 确保加载项的任务窗格具有焦点，然后按 **Ctrl+Shift+I**。
   - 右键单击任务窗格以打开上下文菜单并选择"检查"，或打开 ["个性"菜单](../design/task-pane-add-ins.md#personality-menu)并选择"**附加调试器"**。

1. 打开" **源"** 选项卡。
1. 通过以下步骤打开要调试的文件。

   1. 在工具顶部菜单栏最右边，选择 **...按钮，****然后选择搜索。**
   1. 在搜索框中输入要调试的文件的代码行。 它应该是不可能在任何其他文件中的内容。
   1. 选择刷新按钮。
   1. 在搜索结果中，选择行以在搜索结果上方的窗格中打开代码文件。

   :::image type="content" source="../images/open-file-in-edge-chromium-devtools.png" alt-text="Edge Chromium开发人员工具源选项卡的屏幕截图，其中 4 个部分标记为 A 到 D。":::

1. 若要设置断点，请选择代码文件中行的行号。 代码文件的行将出现一个红点。 在右侧调试器窗口中，断点在" **断点** "下拉列表中注册。
1. 根据需要在加载项中执行函数以触发断点。

> [!TIP]
> 有关使用这些工具的信息，请参阅开发人员Microsoft Edge[概述](/microsoft-edge/devtools-guide-chromium/)。

## <a name="debug-a-dialog-in-an-add-in"></a>在加载项中调试对话框

如果加载项使用 Office 对话框 API，对话框将在任务窗格 (（如果有）中单独运行) 并且必须从该单独进程启动该工具。 请按照以下步骤操作。

1. 运行加载项。
1. 打开对话框并确保它具有焦点。
1. 通过Microsoft Edge (Chromium之) 打开基于 web 的开发人员工具：

   - 按 **Ctrl+Shift+I** 或 **F12**。
   - 右键单击对话框以打开上下文菜单，然后选择"检查 **"**。

1. 使用的工具与任务窗格中的代码相同。 请参阅[本文前面使用](#debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools)基于Microsoft Edge (Chromium的) 工具调试任务窗格外接程序。
