---
title: 处理 Office 对话框中的错误和事件
description: 介绍如何在打开和使用对话框时捕获和处理Office错误
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: be1fb8bcd30b47ac6399657d928d3cad7f857f39
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349894"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>处理 Office 对话框中的错误和事件

本文介绍如何捕获和处理打开对话框时发生的错误以及对话框内发生的错误。

> [!NOTE]
> 本文假定你熟悉使用 Office 对话框 API 的基础知识，如在 Office 加载项中使用 Office 对话框[API 中所述](dialog-api-in-office-add-ins.md)。
> 
> 另请参阅[Best practices and rules for the Office dialog API](dialog-best-practices.md)。

代码应处理两类事件：

- `displayDialogAsync` 调用返回的错误，因为无法创建对话框。
- 对话框中的错误和其他事件。

## <a name="errors-from-displaydialogasync"></a>DisplayDialogAsync 返回的错误

除了常规平台和系统错误之外，调用 还特定于四个错误 `displayDialogAsync` 。

|代码编号|含义|
|:-----|:-----|
|12004|传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。|
|12005|传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。 必须使用 HTTPS。  (在某些版本的 Office 中，返回 12005 的错误消息文本与为 12004.) |
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。|
|12009|用户已选择忽略对话框。 此错误可能发生在Office web 版，用户可能会选择不允许外接程序显示对话框。 有关详细信息，请参阅使用 Office web 版[处理弹出窗口阻止Office web 版。](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)|

调用 `displayDialogAsync` 时，它会将 [AsyncResult](/javascript/api/office/office.asyncresult) 对象传递给其回调函数。 调用成功后，对话框将打开，并且 `value` `AsyncResult` 对象的 属性是 [Dialog](/javascript/api/office/office.dialog) 对象。 有关此内容的示例，请参阅 [将信息从对话框发送到主机页](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)。 调用失败时，不会创建对话框，对象的 属性设置为 `displayDialogAsync` `status` `AsyncResult` `Office.AsyncResultStatus.Failed` ， `error` 并且填充对象的属性。 应始终提供回调，以在出现错误 `status` 时测试 并做出响应。 有关报告错误消息（无论其代码编号如何）的示例，请参阅以下代码。  (`showNotification` 本文中未定义的 函数将显示或记录错误。 有关如何在加载项中实现此函数的示例，请参阅Office[加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).) 

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a>对话框中的错误和事件

对话框中的三个错误和事件将在主机 `DialogEventReceived` 页中引发事件。 有关主机页的提醒，请参阅从主机页 [打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。

|代码编号|含义|
|:-----|:-----|
|12002|下列一种含义：<br> - 传递给 `displayDialogAsync` 的 URL 没有对应的页面。<br> - 传递到加载的页面，但对话框随后被重定向到找不到或加载的页面，或者已定向到具有无效语法的 `displayDialogAsync` URL。|
|12003|对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。|
|12006|对话框已关闭，通常是因为用户选择了 **关闭按钮****X**。|

代码可以在调用 `DialogEventReceived` 时为 `displayDialogAsync` 事件分配处理程序。下面展示了一个非常简单的示例。

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

有关为每个错误代码创建自定义错误消息的事件处理程序的示例， `DialogEventReceived` 请参阅以下示例。

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

有关这样处理错误的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。
