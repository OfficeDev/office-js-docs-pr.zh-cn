---
title: 从其主机页将消息传递到对话框的替代方法
description: 了解在不支持 messageChild 方法时要使用的解决方法。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: f42a549a3c39866516cfd5395dd7589a890b0956
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889413"
---
# <a name="alternative-ways-of-passing-messages-to-a-dialog-box-from-its-host-page"></a>从其主机页将消息传递到对话框的替代方法

将数据和消息从父页面传递到子对话框的建议方法是使用 `messageChild` Office 外接程序中的 [“使用 Office”对话框 API 中](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)所述的方法。如果加载项在不支持 [DialogApi 1.2 要求集](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)的平台或主机上运行，可通过另外两种方式将信息传递到对话框。

- 向传递给 `displayDialogAsync` 的 URL 添加查询参数。
- 将信息存储在主机窗口和对话框都可访问的位置。 两个窗口不共享一个常用会话存储 ([Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性) ，但如果 *它们具有相同的域* (包括端口号，如果有任何) ，则它们共享一个常见的 [本地存储](https://www.w3schools.com/html/html5_webstorage.asp)。\*

> [!NOTE]
> \*有一个 bug 将影响你的令牌处理策略。 如果加载项正使用 Safari 或 Microsoft 浏览器在 **Office 网页版** 上运行，则对话框和任务窗格不共享同一本地存储，因此该存储无法用于在它们之间通信。

## <a name="use-local-storage"></a>使用本地存储

若要使用本地存储，请在调用前`displayDialogAsync`调用`setItem`主机页中的对象方法`window.localStorage`，如以下示例所示。

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

对话框中的代码在需要时读取项，如以下示例所示。

```js
const clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// const clientID = localStorage.clientID;
```

## <a name="use-query-parameters"></a>使用查询参数

下面的示例展示了如何使用查询参数传递数据。

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

有关使用此技术的示例，请参阅[Insert Excel charts using Microsoft Graph in a PowerPoint add-in](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)（在 PowerPoint 加载项中使用 Microsoft Graph 插入 Excel 图表）。

对话框中的代码可以分析 URL，并读取参数值。

> [!IMPORTANT]
> Office 会自动向传递给 `_host_info` 的 URL 添加查询参数 `displayDialogAsync`。 （附加在自定义查询参数（若有）后面，不会附加到对话框导航到的任何后续 URL。 ）Microsoft 可能会更改此值的内容，或者将来会将其全部删除，因此代码不得读取此值。 在 [Window.sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage) 属性)  (，将相同的值添加到对话框的会话存储中。 同样，*代码不得对此值执行读取和写入操作*。
