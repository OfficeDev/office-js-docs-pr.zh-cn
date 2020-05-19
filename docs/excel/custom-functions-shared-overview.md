---
ms.date: 05/17/2020
description: 了解如何在同一 JavaScript 运行时中运行自定义函数、功能区按钮和任务窗格代码，以便在加载项中协调方案。
title: 在共享 JavaScript 运行时中运行外接程序代码
localization_priority: Priority
ms.openlocfilehash: afb07c5223e26ba1e1adbf40c7a4b2e4f7c06349
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275929"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a>概述：在共享 JavaScript 运行时中运行外接程序代码

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

运行 Windows 版 Excel 或 Mac 版 Excel 时，加载项将在单独的 JavaScript 运行时环境中运行功能区按钮、自定义函数和任务窗格的代码。 这会产生一些局限性，例如无法轻松共享全局数据，也不能通过自定义函数访问所有 CORS 功能。

但是，你可以将 Excel 加载项配置为在同一 JavaScript 运行时（也称为共享运行时）中共享代码。 这可在加载项中实现更好的协调，并且可从加载项的所有部分访问任务窗格 DOM 和 CORS。

配置共享运行时可实现以下方案：

- 加载项将具有可供功能区、任务窗格和自定义函数访问的共享 DOM。
- 自定义函数将具有完整的 CORS 支持。
- 自定义函数可调用 Office.js API 以读取电子表格文档数据。
- 打开文档后，加载项即可运行代码。
- 关闭任务窗格后，加载项可以继续运行代码。

当使用任务窗格在共享运行时中运行自定义函数时，它将在不同平台上的浏览器实例中运行，如 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述。此外，Excel 加载项在功能区上显示的任何按钮都将在同一共享运行时中运行。 下图显示了自定义函数、功能区 UI 和任务窗格代码如何在同一 JavaScript 运行时中运行。

![在包含 Excel 中的功能区按钮和任务窗格的共享运行时中运行的自定义函数](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a>设置共享运行时

请参阅[配置共享运行时文章](./configure-your-add-in-to-use-a-shared-runtime.md)，了解如何将自定义函数设置为使用共享运行时。

### <a name="debugging"></a>调试

使用共享运行时时，目前不能使用 Visual Studio Code 在 Windows 版 Excel 中调试自定义函数。 而是需要使用开发人员工具。 有关详细信息，请参阅[使用 Windows 10 上的开发人员工具调试加载项](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)。

## <a name="give-us-feedback"></a>向我们提供反馈

我们非常乐意听取有关此功能的反馈。 如果你发现此功能存在任何 bug、问题或具有相关请求，请通过在 [office-js repo](https://github.com/OfficeDev/office-js) 中创建 GitHub 问题来告诉我们。

## <a name="see-also"></a>另请参阅

- [教程：在 Excel 自定义函数和任务窗格之间共享数据和事件](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [从自定义函数调用 Excel Api](call-excel-apis-from-custom-function.md)
