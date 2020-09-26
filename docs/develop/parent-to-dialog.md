---
title: 将邮件从其主页面传递到对话框的替代方法
description: 了解在 messageChild 方法不受支持时要使用的解决方法。
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 8f44f7f5c145b58d13e7387d01e28fd349a512fc
ms.sourcegitcommit: b47318a24a50443b0579e05e178b3bb5433c372f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2020
ms.locfileid: "48279484"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>将邮件从其主页面传递到对话框的替代方法

将来自父页面的数据和邮件传递到子对话框的建议方法是 `messageChild` 使用方法，如在 [office 外接程序中使用 OFFICE 对话框 API](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)中所述。如果外接程序在不支持 [DialogApi 1.2 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md)的平台或主机上运行，则可以通过两种其他方式将信息传递到该对话框：

- 向传递给 `displayDialogAsync` 的 URL 添加查询参数。
- 将信息存储在主机窗口和对话框都可访问的位置。 这两个窗口不共享一个公用会话存储 ([sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) ，但 *如果它们具有相同的域* (包括端口号（如果) 有），则它们共享一个公用的 [本地存储](https://www.w3schools.com/html/html5_webstorage.asp)。\*


> [!NOTE]
> \*有一个 bug 将影响你的令牌处理策略。 如果加载项正使用 Safari 或 Microsoft 浏览器在 **Office 网页版**上运行，则对话框和任务窗格不共享同一本地存储，因此该存储无法用于在它们之间通信。

## <a name="use-local-storage"></a>使用本地存储

若要使用本地存储，请 `setItem` `window.localStorage` 先在主机页中调用对象的方法，然后再 `displayDialogAsync` 调用，如以下示例所示：

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

对话框框中的代码会在需要时读取项，如以下示例所示：

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>使用查询参数

下面的示例展示了如何使用查询参数传递数据：

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

有关使用此技术的示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）。

对话框中的代码可以分析 URL，并读取参数值。

> [!IMPORTANT]
> Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。 （附加在自定义查询参数（若有）后面，不会附加到对话框导航到的任何后续 URL。 ）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。 将相同的值添加到对话框的 "会话存储" ([sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) 。 同样，*代码不得对此值执行读取和写入操作*。
