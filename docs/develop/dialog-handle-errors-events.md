---
title: 处理 Office 对话框中的错误和事件
description: 了解如何在打开和使用 Office 对话框时捕获和处理错误。
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: d3bdae7d4dddcd92a54a46fec0d5854a1a18a0bc
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616033"
---
# <a name="handle-errors-and-events-in-the-office-dialog-box"></a>在“Office”对话框中处理错误和事件

本文介绍在打开对话框时如何捕获和处理错误以及对话框内发生的错误。

> [!NOTE]
> 本文假设你熟悉使用 Office 对话 API 的基础知识，如 Office 加载项中的 [“使用 Office”对话框 API 中](dialog-api-in-office-add-ins.md)所述。
>
> 另请参阅 [Office 对话框 API 的最佳做法和规则](dialog-best-practices.md)。

代码应处理两类事件。

- `displayDialogAsync` 调用返回的错误，因为无法创建对话框。
- 对话框中的错误和其他事件。

## <a name="errors-from-displaydialogasync"></a>DisplayDialogAsync 返回的错误

除了常规平台和系统错误外，还有四个错误是特定于调用 `displayDialogAsync`的。

|代码编号|含义|
|:-----|:-----|
|12004|传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。|
|12005|传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。 必须使用 HTTPS。  (在某些版本的 Office 中，返回的错误消息文本与 12005 返回的错误消息文本相同，为 12004.) |
|<span id="12007">12007</span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。|
|12009|用户已选择忽略对话框。 Office web 版中可能会出现此错误，用户可以选择不允许加载项显示对话框。 有关详细信息，请参阅[使用Office web 版处理弹出窗口阻止程序](dialog-best-practices.md#handle-pop-up-blockers-with-office-on-the-web)。|
|12011| 加载项在Office web 版中运行，用户的浏览器配置阻止弹出窗口。 当浏览器是 Edge Legacy 且加载项的域与对话尝试打开的域位于不同的安全区域时，最常发生这种情况。 触发此错误的另一种方案是，浏览器为 Safari，它配置为阻止所有弹出窗口。 请考虑响应此错误，并提示用户更改浏览器配置或使用其他浏览器。|

调用时 `displayDialogAsync` ，它会将 [AsyncResult](/javascript/api/office/office.asyncresult) 对象传递给其回调函数。 调用成功后，会打开对话框，并且 `value` 对象的 `AsyncResult` 属性是 [Dialog](/javascript/api/office/office.dialog) 对象。 有关此示例，请参阅将 [信息从对话框发送到主机页](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)。 调用 `displayDialogAsync` 失败时，将不创建对话框， `status` 将对象的 `AsyncResult` 属性设置为 `Office.AsyncResultStatus.Failed`，并 `error` 填充对象的属性。 应始终提供一个回调，用于在错误时测试 `status` 和响应。 有关报告错误消息的示例，无论其代码号如何，请参阅以下代码。 `showNotification` (本文中未定义的函数显示或记录错误。 有关如何在外接程序中实现此函数的示例，请参阅 [Office 加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)) 

```js
let dialog;
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

对话框中的三个错误和事件将在主机页中引发 `DialogEventReceived` 事件。 有关主机页的提醒，请参阅 [主机页中的“打开”对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。

|代码编号|含义|
|:-----|:-----|
|12002|下列一种含义：<br> - 传递给 `displayDialogAsync` 的 URL 没有对应的页面。<br> - 传递到 `displayDialogAsync` 加载的页面，但对话框随后被重定向到无法找到或加载的页面，或者已定向到语法无效的 URL。|
|12003|对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。|
|12006|对话框已关闭，通常是因为用户选择了 **“关闭** ”按钮 **X**。|

代码可以在调用 `DialogEventReceived` 时为 `displayDialogAsync` 事件分配处理程序。 下面展示了一个非常简单的示例。

```js
let dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

有关为 `DialogEventReceived` 每个错误代码创建自定义错误消息的事件的处理程序示例，请参阅以下示例。

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

## <a name="see-also"></a>另请参阅

有关这样处理错误的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。
