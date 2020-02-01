---
title: 处理 Office 对话框中的错误和事件
description: 介绍如何在打开和使用 Office 对话框时捕获和处理错误
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: a35131a46dc9f5edc18df37495abe5d8c2c5ad2a
ms.sourcegitcommit: 4c9e02dac6f8030efc7415e699370753ec9415c8
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/01/2020
ms.locfileid: "41650073"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a>处理 Office 对话框中的错误和事件

本文介绍如何在打开对话框以及对话框中发生的错误时捕获和处理错误。

> [!NOTE]
> 本文 presupposes 您熟悉使用 Office 对话框 API 的基础知识，如在[Office 外接程序中使用 office 对话框 api](dialog-api-in-office-add-ins.md)中所述。
> 
> 另请参阅[Office 对话框 API 的最佳实践和规则](dialog-best-practices.md)。

代码应处理两类事件：

- `displayDialogAsync` 调用返回的错误，因为无法创建对话框。
- 对话框中的错误和其他事件。

## <a name="errors-from-displaydialogasync"></a>DisplayDialogAsync 返回的错误

除了常规平台和系统错误之外，还有四个错误特定于调用`displayDialogAsync`。

|代码编号|含义|
|:-----|:-----|
|12004|传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。|
|12005|传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。 必须使用 HTTPS。 （在 Office 的某些版本中，返回12005的错误消息文本与为12004返回的文本相同。|
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。|
|12009|用户已选择忽略对话框。 此错误可能发生在 web 上的 Office 中，用户可以在其中选择不允许外接程序呈现对话框。 有关详细信息，请参阅[使用 web 上的 Office 处理弹出窗口阻止程序](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)。|

调用`displayDialogAsync`时，它会将[AsyncResult](/javascript/api/office/office.asyncresult)对象传递给它的回调函数。 当调用成功时，将打开对话框，并且`value` `AsyncResult`对象的属性是[dialog](/javascript/api/office/office.dialog)对象。 有关这种情况的示例，请参阅[将信息从对话框发送到主机页](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)。 当`displayDialogAsync`调用失败时，将不会创建对话框， `status` `AsyncResult`对象的属性设置为`Office.AsyncResultStatus.Failed`，并填充对象的`error`属性。 应始终提供一个回调，以在出错`status`时测试和响应。 有关报告错误消息（而不考虑其代码编号）的示例，请参阅以下代码。 （本文`showNotification`中未定义的函数可能显示或记录错误。 有关如何在外接程序中实现此函数的示例，请参阅[Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。）

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

对话框中的三个错误和事件将引发主机`DialogEventReceived`页中的事件。 有关主机页面的提示，请参阅[从主机页面打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。

|代码编号|含义|
|:-----|:-----|
|12002|下列一种含义：<br> - 传递给 `displayDialogAsync` 的 URL 没有对应的页面。<br> -已传递给`displayDialogAsync`加载的页面，但随后会将该对话框重定向到无法找到或加载的页面，或者已将其定向到语法无效的 URL。|
|12003|对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。|
|12006|对话框已关闭，通常是因为用户选择了 "**关闭**" 按钮**X**。|

代码可以在调用 `displayDialogAsync` 时分配 `DialogEventReceived` 事件处理程序。下面展示了一个简单示例：

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

有关为每个错误代码创建自定义错误消息的 `DialogEventReceived` 事件处理程序示例，请参阅下面的示例：

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
