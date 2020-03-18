---
ms.date: 02/13/2020
description: 了解如何在同一 JavaScript 运行时中运行自定义函数、功能区按钮和任务窗格代码，以便在加载项中协调方案。
title: 在共享 JavaScript 运行时中运行加载项代码（预览版）
localization_priority: Priority
ms.openlocfilehash: 774990a9452d450bd5c4d968027bc64ebee858af
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719530"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtime-preview"></a>概述：在共享 JavaScript 运行时中运行加载项代码（预览版）

[!include[Running custom functions in shared JavaScript runtime note](../includes/excel-shared-runtime-preview-note.md)]

运行 Windows 版 Excel 或 Mac 版 Excel 时，加载项将在单独的 JavaScript 运行时环境中运行功能区按钮、自定义函数和任务窗格的代码。 这会产生一些局限性，例如无法轻松共享全局数据，也不能通过自定义函数访问所有 CORS 功能。

但是，你可以将 Excel 加载项配置为在同一 JavaScript 运行时（也称为共享运行时）中共享代码。 这可在加载项中实现更好的协调，并且可从加载项的所有部分访问任务窗格 DOM 和 CORS。

配置共享运行时可实现以下方案：

- 加载项将具有可供功能区、任务窗格和自定义函数访问的共享 DOM。
- 自定义函数将具有完整的 CORS 支持。
- 自定义函数可调用 Office.js API 以读取电子表格文档数据。
- 打开文档后，加载项即可运行代码。
- 关闭任务窗格后，加载项可以继续运行代码。

当使用任务窗格在共享运行时中运行自定义函数时，它将在不同平台上的浏览器实例中运行，如 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)中所述。此外，Excel 加载项在功能区上显示的任何按钮都将在同一共享运行时中运行。 下图显示了自定义函数、功能区 UI 和任务窗格代码如何在同一 JavaScript 运行时中运行。

![使用 Excel 中的功能区按钮和任务窗格在共享运行时中运行的自定义函数](../images/custom-functions-in-browser-runtime.png)

## <a name="differences-when-running-custom-functions-in-a-shared-runtime"></a>在共享运行时中运行自定义函数时的差异

将 Excel 加载项项目配置为在共享运行时中运行自定义函数时，与使用自定义函数运行时有一些不同。

### <a name="storage"></a>存储

无需再在任务窗格、自定义函数或功能区 UI 之间使用**存储** API 来共享数据。 可将全局变量置于 **window** 对象中，或使用自己的首选状态管理方法。

### <a name="authentication"></a>身份验证

如果在身份验证过程中收到令牌，无需使用**存储** API 在任务窗格、自定义函数和功能区 UI 之间共享它们。 你可以使用自己的首选存储技术和存储位置来共享它们，例如 `localStorage`。

### <a name="dialog-api"></a>对话框 API

无需再使用 **OfficeRuntime.Dialog** API 来显示来自自定义函数的对话框。 可以将同一[对话框 API](../develop/dialog-api-in-office-add-ins.md) 用于自定义函数、功能区按钮和任务窗格。

### <a name="debugging"></a>调试

使用共享运行时时，目前不能使用 Visual Studio Code 在 Windows 版 Excel 中调试自定义函数。 你需要使用开发人员工具。 有关详细信息，请参阅[使用 Windows 10 上的开发人员工具调试加载项](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)。

## <a name="get-started"></a>开始使用

若要将 Excel 加载项项目配置为在共享运行时中运行自定义函数，请参阅[将 Excel 加载项配置为使用共享 JavaScript 运行时（预览版）](configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="give-us-feedback"></a>向我们提供反馈

我们非常乐意听取有关此功能的反馈。 如果你发现此功能存在任何 bug、问题或具有相关请求，请通过在 [office-js repo](https://github.com/OfficeDev/office-js) 中创建 GitHub 问题来告诉我们。

## <a name="see-also"></a>另请参阅

共享运行时的相关文章列表
- [教程：在 Excel 自定义函数和任务窗格之间共享数据和事件（预览）](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [从自定义函数中调用 Excel API（预览版）](call-excel-apis-from-custom-function.md)