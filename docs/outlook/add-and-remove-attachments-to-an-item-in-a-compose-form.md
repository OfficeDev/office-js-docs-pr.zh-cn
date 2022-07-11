---
title: 在 Outlook 加载项中添加和删除附件
description: 使用各种附件 API 管理附加到用户正在撰写的项目的文件或 Outlook 项目。
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6600fc5926cdecd95e4d232223f11dd9a7b1fc41
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712830"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>在 Outlook 的撰写窗体中管理项目的附件

Office JavaScript API 提供了多个 API，可用于在用户撰写时管理项目的附件。

## <a name="attach-a-file-or-outlook-item"></a>附加文件或 Outlook 项

可以使用适合附件类型的方法将文件或 Outlook 项附加到撰写窗体。

- [addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)：附加文件
- [addFileAttachmentFromBase64Async](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)：使用其 base64 字符串附加文件
- [addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)：附加 Outlook 项

这些是异步方法，这意味着执行可以继续，而无需等待操作完成。 异步调用可能需要一段时间才能完成，具体取决于要添加的附件的原始位置和大小。

如果有任务依赖于要完成的操作，则应在回调方法中执行这些任务。 此回调方法是可选的，在附件上传完成后调用。 此回调方法使用 [AsyncResult](/javascript/api/office/office.asyncresult) 对象作为输出参数，提供添加附件操作的任何状态、错误和返回值。 如果此回调需要任何额外参数，则可以在可选的 `options.asyncContext` 参数中指定它们。 `options.asyncContext` 可以是回调方法所期望的任何类型。

例如，可以定义 `options.asyncContext` 为包含一个或多个键值对的 JSON 对象。 可在 [Office 外](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods)接程序的异步编程中的 Office 外接程序平台中找到有关将可选参数传递到异步方法的更多示例。以下示例演示如何使用 `asyncContext` 参数将 2 个参数传递给回调方法。

```js
const options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

可以使用 `AsyncResult` 对象的 `status` 和 `error` 属性，检查回调方法中异步方法调用是成功还是出现错误。如果附加成功完成，可以使用 `AsyncResult.value` 属性获取附件 ID。附件 ID 是一个证书，稍后可使用附件 ID 删除附件。

> [!NOTE]
> 附件 ID 仅在同一会话中有效，不能保证跨会话映射到同一附件。 当用户关闭外接程序时，或者用户开始以内联窗体撰写，然后弹出内联窗体以在单独的窗口中继续时，会话何时结束的示例。

### <a name="attach-a-file"></a>附加文件

可以使用 `addFileAttachmentAsync` 该方法并指定文件的 URI，将文件附加到撰写窗体中的消息或约会。 还可以使用该方法， `addFileAttachmentFromBase64Async` 但将 base64 字符串指定为输入。 如果文件受保护，您可以包括相应的标识或身份验证令牌作为 URI 查询字符串参数。 Exchange 将向 URI 发出调用以获取附件，保护文件的 Web 服务将需要使用令牌作为进行身份验证的一种方式。

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
                    const attachmentID = asyncResult.value;
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

### <a name="attach-an-outlook-item"></a>附加 Outlook 项

可以通过指定项的 Exchange Web 服务 (EWS) ID 并使用 `addItemAttachmentAsync` 该方法，将 Outlook 项目 (（例如电子邮件、日历或联系人项）附加到撰写窗体中的消息或约会) 。 可以使用 [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法并访问 EWS 操作 [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)，获取用户邮箱中电子邮件、日历、联系人或任务项的 EWS ID。 [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性还提供阅读窗体中某个现有项目的 EWS ID。

下面的 JavaScript 函数 `addItemAttachment`扩展了上面的第一个示例，并将项目作为附件添加到正在撰写的电子邮件或约会中。 此函数将要附加的项目的 EWS ID 作为实参。 如果附加成功，它将获取用于进一步处理的附件 ID，包括在同一会话中删除该附件。

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
                const attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> 可以使用撰写加载项在Outlook 网页版或移动设备上附加定期约会的实例。 但是，在支持 Outlook 桌面客户端中，尝试附加实例将导致附加定期系列 (父约会) 。

## <a name="get-attachments"></a>获取附件

在撰写模式下获取附件的 API 可从 [要求集 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) 获取。

- [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

可以使用 [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法获取正在撰写的邮件或约会的附件。

若要获取附件的内容，可以使用 [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法。 受支持的格式列在 [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) 枚举中。

应提供回调方法，以使用 `AsyncResult` 输出参数对象检查状态和任何错误。 还可以使用可选 `asyncContext` 参数将任何其他参数传递给回调方法。

以下 JavaScript 示例获取附件，并允许为每个受支持的附件格式设置不同的处理。

```js
const item = Office.context.mailbox.item;
const options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (let i = 0 ; i < result.value.length ; i++) {
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

使用 [removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) 方法时，可以通过指定相应的附件 ID 从撰写窗体中的消息或约会项中删除文件或项附件。

> [!IMPORTANT]
> 如果使用的是要求集 1.7 或更低版本，则应仅删除同一加载项在同一会话中添加的附件。

`addFileAttachmentAsync``addItemAttachmentAsync`类似于异步方法和`getAttachmentsAsync`方法`removeAttachmentAsync`。 应提供回调方法，以使用 `AsyncResult` 输出参数对象检查状态和任何错误。 还可以使用可选 `asyncContext` 参数将任何其他参数传递给回调方法。

以下 JavaScript 函数 `removeAttachment`继续扩展上述示例，并从正在撰写的电子邮件或约会中删除指定的附件。 此函数将要删除的附件的 ID 作为实参。 可以在成功`addFileAttachmentAsync``addFileAttachmentFromBase64Async`或`addItemAttachmentAsync`方法调用后获取附件的 ID，并在后续`removeAttachmentAsync`方法调用中使用它。 还可以调用 `getAttachmentsAsync` 要求集 1.8) 中引入的 (，以获取该外接程序会话的附件及其 ID。

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
