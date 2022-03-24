---
title: 将邮件从主机页传递到对话框的替代方法
description: 了解在 messageChild 方法不受支持时使用的解决方法。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: e17cb81ab781c6b9acf0ae76a29c601a61c9f931
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743366"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>将邮件从主机页传递到对话框的替代方法

建议通过 `messageChild` 方法将数据和消息从父页面传递到子对话框，如在 Office 外接程序中使用 [Office 对话框 API 中所述](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)。如果加载项在不支持 [DialogApi 1.2](../reference/requirement-sets/dialog-api-requirement-sets.md) 要求集的平台或主机上运行，可通过其他两种方法将信息传递到对话框。

- 向传递给 `displayDialogAsync` 的 URL 添加查询参数。
- 将信息存储在主机窗口和对话框都可访问的位置。 这两个窗口不共享 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性 (的常见会话存储) 但如果它们具有相同的 *域 (包括* 端口号，如果有) ，则它们共享一个公共 [本地 存储](https://www.w3schools.com/html/html5_webstorage.asp)。\*

> [!NOTE]
> \*有一个 bug 将影响你的令牌处理策略。 如果加载项正使用 Safari 或 Microsoft 浏览器在 **Office 网页版** 上运行，则对话框和任务窗格不共享同一本地存储，因此该存储无法用于在它们之间通信。

## <a name="use-local-storage"></a>使用本地存储

若要使用本地存储，在 `setItem` `window.localStorage` `displayDialogAsync` 调用前调用主机页中的 对象的 方法，如以下示例所示。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

对话框中的代码会根据需要读取项目，如以下示例所示。

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>使用查询参数

下面的示例展示了如何使用查询参数传递数据。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

有关使用此技术的示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）。

对话框中的代码可以分析 URL，并读取参数值。

> [!IMPORTANT]
> Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。 （附加在自定义查询参数（若有）后面，不会附加到对话框导航到的任何后续 URL。 ）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。 相同的值将添加到对话框的会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)) 。 同样，*代码不得对此值执行读取和写入操作*。
