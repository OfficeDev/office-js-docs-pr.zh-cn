---
title: 将数据和邮件从其主机页传递到对话框
description: 了解如何使用 messageChild 和 DialogParentMessageReceived Api 将数据传递到主机页中的对话框。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 05220fa4cecad4fe412a5590605f774f92ef8f61
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093572"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a>将数据和邮件从其主机页传递到对话框 (预览) 

您的外接程序可以使用[dialog](/javascript/api/office/office.dialog)对象的[messageChild](/javascript/api/office/office.dialog#messagechild-message-)方法，将邮件从[主机页面](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)发送到对话框。

> [!Important]
>
> - 本文中介绍的 Api 处于预览阶段。 它们可供开发人员用来进行试验;但不应在生产外接中使用。 在发布此 API 之前，请使用将[信息传递到](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)生产外接程序的对话框中所述的技术。
> - 本文中所述的 Api 需要 Microsoft 365 订阅。 你应该使用来自预览体验成员频道的最新每月版本和内部版本。 你可能需要成为 Office 预览体验成员，才能获取此版本。 有关详细信息，请参阅[成为 Office 预览体验成员](https://insider.office.com)。 请注意，当内部版本毕业生到生产半年频道时，将对该内部版本禁用对预览功能的支持。
> - 在预览的初始阶段，Api 在 Excel、PowerPoint 和 Word 中受支持;而不是在 Outlook 中。
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a>`messageChild()`从主机页使用

调用 Office 对话框 API 打开对话框时，将返回[dialog](/javascript/api/office/office.dialog)对象。 应将其分配给变量，该变量的作用域通常大于[displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-)方法，因为该对象将被其他方法引用。 示例如下：

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

此 `Dialog` 对象具有向对话框发送任何字符串或字符串化数据的[messageChild](/javascript/api/office/office.dialog#messagechild-message-)方法。 这 `DialogParentMessageReceived` 将在对话框中引发事件。 您的代码应处理此事件，如下一节中所示。

假设对话框的 UI 应与当前活动的工作表关联，并且该工作表相对于其他工作表的位置。 在下面的示例中， `sheetPropertiesChanged` 将 Excel 工作表属性发送到对话框。 在这种情况下，当前工作表名为 "我的工作表"，而它是工作簿中的第二个工作表。 数据封装在字符串化对象中，以便可以将其传递给 `messageChild` 。

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>在对话框中处理 DialogParentMessageReceived

在对话框的 JavaScript 中， `DialogParentMessageReceived` 使用[addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-)方法为事件注册处理程序。 这通常是在[onReady 或 Office.initialize 方法](initialize-add-in.md)中完成的。 示例如下：

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

然后，定义该 `onMessageFromParent` 处理程序。 下面的代码将继续上一节中的示例。 请注意，Office 会将参数传递给处理程序，并确保 `message` argument 对象的属性包含主机页中的字符串。 在此示例中，邮件被 reconverted 到对象，jQuery 用于将对话框的顶部标题设置为与新工作表名称相匹配。

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

最佳做法是验证是否正确注册了处理程序。 为此，可以将回调传递给在 `addHandlerAsync` 注册处理程序的尝试完成时运行的方法。 如果未成功注册处理程序，请使用该处理程序记录或显示错误。 示例如下。 请注意，这 `reportError` 是未在此处定义的函数，它会记录或显示错误。

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a>条件消息

由于可以 `messageChild` 从主机页进行多次调用，但在该事件的对话框中只有一个处理程序 `DialogParentMessageReceived` ，因此处理程序必须使用条件逻辑来区分不同的消息。 您可以按照与[条件消息](dialog-api-in-office-add-ins.md#conditional-messaging)中所述的方式将消息发送到主机页时，精确地与构造条件消息传递的方式完全并行。
