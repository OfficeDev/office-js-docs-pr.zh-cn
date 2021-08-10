---
title: 在 Outlook 加载项中添加和删除附件
description: 可以使用各种附件 API 来管理文件或Outlook附加到用户正在撰写的项目的项目。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: ce2fddd9ed9df3f0119f1128d4a05a68a6b772a63c46f037d123a41d936e23e7
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088595"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>在邮件撰写窗体中管理项目的Outlook

JavaScript API Office JavaScript API 提供了多个 API，可用于在用户撰写时管理项目的附件。

## <a name="attach-a-file-or-outlook-item"></a>附加文件或Outlook项

您可以使用适用于附件Outlook的方法将文件或项目附加到撰写窗体。

- [addFileAttachmentAsync：](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)附加文件
- [addFileAttachmentFromBase64Async：](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)使用文件的 base64 字符串附加文件
- [addItemAttachmentAsync：](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)附加Outlook项

这些是异步方法，这意味着可以继续执行，而无需等待操作完成。 根据要添加的附件的原始位置和大小，异步调用可能需要一段时间才能完成。

如果有任务依赖于要完成的操作，则应在回调方法中执行这些任务。 此回调方法是可选的，在附件上载完成时调用此方法。 此回调方法使用 [AsyncResult](/javascript/api/office/office.asyncresult) 对象作为输出参数，提供添加附件操作的任何状态、错误和返回值。 如果此回调需要任何额外参数，则可以在可选的 `options.asyncContext` 参数中指定它们。 `options.asyncContext` 可以是回调方法所期望的任何类型。

例如，可以定义为包含一个或多个键值对的 `options.asyncContext` JSON 对象。 有关向 Office 外接程序平台中的异步方法传递可选参数的更多示例，Office外接程序中的异步[编程](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods)。以下示例演示如何使用 参数将 `asyncContext` 2 个参数传递给回调方法。

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

可以使用 `AsyncResult` 对象的 `status` 和 `error` 属性，检查回调方法中异步方法调用是成功还是出现错误。如果附加成功完成，可以使用 `AsyncResult.value` 属性获取附件 ID。附件 ID 是一个证书，稍后可使用附件 ID 删除附件。

> [!NOTE]
> 附件 ID 仅在同一会话中有效，不保证跨会话映射到同一附件。 会话结束的示例包括用户何时关闭外接程序，或者用户开始在内嵌表单中撰写，随后弹出内嵌表单以在单独的窗口中继续。

### <a name="attach-a-file"></a>附加文件

您可以使用 方法并指定文件的 URI，将文件附加到撰写窗体中的邮件或 `addFileAttachmentAsync` 约会。 您也可以使用 方法 `addFileAttachmentFromBase64Async` ，但将 base64 字符串指定为输入。 如果文件受保护，您可以包括相应的标识或身份验证令牌作为 URI 查询字符串参数。 Exchange 将向 URI 发出调用以获取附件，保护文件的 Web 服务将需要使用令牌作为进行身份验证的一种方式。

下面的 JavaScript 示例是从 Web 服务器将文件、picture.png 附加到正在撰写的邮件或约会的撰写加载项。回调方法将 `asyncResult` 作为参数，检查结果状态，并在方法成功的情况下获取附件 ID。

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback method as an argument to
        // the asyncContext parameter.
        Office.context.mailbox.item.addFileAttachmentAsync(
            `https://webserver/picture.png`,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                } else {
                    // Get the ID of the attached file.
                    var attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="attach-an-outlook-item"></a>附加Outlook项

您可以通过指定项目的 Exchange Web 服务 (EWS) ID 并使用此方法，将 Outlook 项目 (（例如，电子邮件、日历或联系人项目) ）附加到撰写窗体中的邮件或约会。 `addItemAttachmentAsync` 通过使用 [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法并访问 EWS 操作 [FindItem，](/exchange/client-developer/web-service-reference/finditem-operation)您可以获取用户邮箱中电子邮件、日历、联系人或任务项目的 EWS ID。 [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性还提供阅读窗体中某个现有项目的 EWS ID。

以下 JavaScript 函数 扩展了上述第一个示例，并将项目作为附件添加到正在撰写的电子邮件或 `addItemAttachment` 约会中。 此函数将要附加的项目的 EWS ID 作为实参。 如果附加成功，它将获取附件 ID 以进一步处理，包括在同一会话中删除该附件。

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> 可以使用撰写加载项在邮件中或移动设备上Outlook 网页版定期约会的实例。 但是，在Outlook客户端中，尝试附加实例将导致将定期系列 (父约会) 。

## <a name="get-attachments"></a>获取附件

在撰写模式下获取附件的 API 可从要求集 [1.8 获取](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)。

- [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

您可以使用 [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法获取正在撰写的邮件或约会的附件。

若要获取附件的内容，可以使用 [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法。 [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)枚举中列出了受支持的格式。

您应提供一个回调方法，以使用 output 参数对象检查状态 `AsyncResult` 和任何错误。 您还可以使用可选参数将任何其他参数传递给回调 `asyncContext` 方法。

以下 JavaScript 示例获取附件，并允许您针对每种受支持的附件格式设置不同的处理。

```js
var item = Office.context.mailbox.item;
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

## <a name="remove-an-attachment"></a>删除附件

使用 [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法时，可以通过指定相应的附件 ID 从撰写窗体中的邮件或约会项目中删除文件或项目附件。

> [!IMPORTANT]
> 如果使用的是要求集 1.7 或更早版本，则只应删除同一外接程序在同一会话中添加的附件。

与 、 `addFileAttachmentAsync` `addItemAttachmentAsync` 和 `getAttachmentsAsync` 方法类似 `removeAttachmentAsync` ，是一种异步方法。 您应提供一个回调方法，以使用 output 参数对象检查状态 `AsyncResult` 和任何错误。 您还可以使用可选参数将任何其他参数传递给回调 `asyncContext` 方法。

以下 JavaScript 函数 将继续扩展上述示例，并从正在撰写的电子邮件或约会中删除 `removeAttachment` 指定的附件。 此函数将要删除的附件的 ID 作为实参。 可以在成功调用 、 或 方法后获取附件的 ID，并可在后续方法调用 `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` `removeAttachmentAsync` 中使用它。 您还可以调用 (集 1.8 中引入) ，获取该外接程序会话的附件及其 `getAttachmentsAsync` ID。

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback method is invoked.
    // Here, the callback method uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback method as an argument to the asyncContext parameter.
    Office.context.mailbox.item.removeAttachmentAsync(
        attachmentId,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```

## <a name="see-also"></a>另请参阅

- [创建适用于撰写窗体的 Outlook 加载项](compose-scenario.md)
- [Office 外接程序中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md)
