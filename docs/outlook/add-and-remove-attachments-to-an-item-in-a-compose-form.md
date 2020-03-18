---
title: 在 Outlook 加载项中添加和删除附件
description: 您可以使用各种附件 Api 来管理附加到用户正在撰写的项目的文件或 Outlook 项目。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bb966ff80bae37fbaa781b5a428f6e26391aa9f4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720881"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a><span data-ttu-id="131c3-103">在 Outlook 的撰写窗体中管理项目的附件</span><span class="sxs-lookup"><span data-stu-id="131c3-103">Manage an item's attachments in a compose form in Outlook</span></span>

<span data-ttu-id="131c3-104">Office JavaScript API 提供了多个 Api，可用于在用户撰写时管理项目的附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-104">The Office JavaScript API provides several APIs you can use to manage an item's attachments when the user is composing.</span></span>

## <a name="attach-a-file-or-outlook-item"></a><span data-ttu-id="131c3-105">附加文件或 Outlook 项</span><span class="sxs-lookup"><span data-stu-id="131c3-105">Attach a file or Outlook item</span></span>

<span data-ttu-id="131c3-106">您可以使用适用于附件类型的方法，将文件或 Outlook 项目附加到撰写窗体中。</span><span class="sxs-lookup"><span data-stu-id="131c3-106">You can attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.</span></span>

- <span data-ttu-id="131c3-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)：附加文件</span><span class="sxs-lookup"><span data-stu-id="131c3-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file</span></span>
- <span data-ttu-id="131c3-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)：使用其 base64 字符串附加文件</span><span class="sxs-lookup"><span data-stu-id="131c3-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file using its base64 string</span></span>
- <span data-ttu-id="131c3-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)：附加 Outlook 项目</span><span class="sxs-lookup"><span data-stu-id="131c3-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach an Outlook item</span></span>

<span data-ttu-id="131c3-110">这些方法是异步方法，这意味着可以继续执行，而无需等待操作完成。</span><span class="sxs-lookup"><span data-stu-id="131c3-110">These are asynchronous methods, which means execution can go on without waiting for the action to complete.</span></span> <span data-ttu-id="131c3-111">异步调用可能需要一段时间才能完成，具体取决于要添加的附件的原始位置和大小。</span><span class="sxs-lookup"><span data-stu-id="131c3-111">Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.</span></span>

<span data-ttu-id="131c3-112">如果有任务依赖于要完成的操作，则应在回调方法中执行这些任务。</span><span class="sxs-lookup"><span data-stu-id="131c3-112">If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method.</span></span> <span data-ttu-id="131c3-113">此回调方法是可选的，并在附件上载完成时调用。</span><span class="sxs-lookup"><span data-stu-id="131c3-113">This callback method is optional and is invoked when the attachment upload has completed.</span></span> <span data-ttu-id="131c3-114">此回调方法使用 [AsyncResult](/javascript/api/office/office.asyncresult) 对象作为输出参数，提供添加附件操作的任何状态、错误和返回值。</span><span class="sxs-lookup"><span data-stu-id="131c3-114">The callback method takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment.</span></span> <span data-ttu-id="131c3-115">如果此回调需要任何额外参数，则可以在可选的 `options.asyncContext` 参数中指定它们。</span><span class="sxs-lookup"><span data-stu-id="131c3-115">If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter.</span></span> <span data-ttu-id="131c3-116">`options.asyncContext` 可以是回调方法所期望的任何类型。</span><span class="sxs-lookup"><span data-stu-id="131c3-116">`options.asyncContext` can be of any type that your callback method expects.</span></span>

<span data-ttu-id="131c3-p103">例如，可以将 `options.asyncContext` 定义为一个 JSON 对象，该对象包含一个或多个键值对。可以在 [Office 加载项中的异步编程](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods)中找到有关在 Office 加载项平台中将可选参数传递给异步方法 的更多示例。下面的示例演示了如何使用 `asyncContext` 参数将 2 个自变量传递给回调方法：</span><span class="sxs-lookup"><span data-stu-id="131c3-p103">For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs. You can find more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass 2 arguments to a callback method:</span></span>

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

<span data-ttu-id="131c3-p104">可以使用 `AsyncResult` 对象的 `status` 和 `error` 属性，检查回调方法中异步方法调用是成功还是出现错误。如果附加成功完成，可以使用 `AsyncResult.value` 属性获取附件 ID。附件 ID 是一个证书，稍后可使用附件 ID 删除附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-p104">You can check for success or error of an asynchronous method call in the callback method using the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, you can use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.</span></span>

> [!NOTE]
> <span data-ttu-id="131c3-122">最佳做法是，仅当同一个加载项在同一会话中添加了一个附件时，才应使用该附件 ID 来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-122">As a best practice, you should use the attachment ID to remove an attachment only if the same add-in has added that attachment in the same session.</span></span> <span data-ttu-id="131c3-123">在 web 和移动设备上的 Outlook 中，附件 ID 仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="131c3-123">In Outlook on the web and mobile devices, the attachment ID is valid only within the same session.</span></span> <span data-ttu-id="131c3-124">用户关闭外接程序时，或用户在内嵌窗体中开始撰写，然后弹出内嵌窗体以在单独的窗口中继续时，会话结束。</span><span class="sxs-lookup"><span data-stu-id="131c3-124">A session is over when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

### <a name="attach-a-file"></a><span data-ttu-id="131c3-125">附加文件</span><span class="sxs-lookup"><span data-stu-id="131c3-125">Attach a file</span></span>

<span data-ttu-id="131c3-126">您可以使用`addFileAttachmentAsync`方法并指定文件的 URI，在撰写窗体中将文件附加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="131c3-126">You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file.</span></span> <span data-ttu-id="131c3-127">您还可以使用方法`addFileAttachmentFromBase64Async` ，但将 base64 字符串指定为输入。</span><span class="sxs-lookup"><span data-stu-id="131c3-127">You can also use the `addFileAttachmentFromBase64Async` method but specify the base64 string as input.</span></span> <span data-ttu-id="131c3-128">如果文件受保护，您可以包括相应的标识或身份验证令牌作为 URI 查询字符串参数。</span><span class="sxs-lookup"><span data-stu-id="131c3-128">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span></span> <span data-ttu-id="131c3-129">Exchange 将向 URI 发出调用以获取附件，保护文件的 Web 服务将需要使用令牌作为进行身份验证的一种方式。</span><span class="sxs-lookup"><span data-stu-id="131c3-129">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span></span>

<span data-ttu-id="131c3-p107">下面的 JavaScript 示例是从 Web 服务器将文件、picture.png 附加到正在撰写的邮件或约会的撰写加载项。回调方法将 `asyncResult` 作为参数，检查结果状态，并在方法成功的情况下获取附件 ID。</span><span class="sxs-lookup"><span data-stu-id="131c3-p107">The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.</span></span>

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
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
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

### <a name="attach-an-outlook-item"></a><span data-ttu-id="131c3-132">附加 Outlook 项目</span><span class="sxs-lookup"><span data-stu-id="131c3-132">Attach an Outlook item</span></span>

<span data-ttu-id="131c3-p108">可以通过指定项的 Exchange Web Services (EWS) ID 并使用 `addItemAttachmentAsync` 方法，将 Outlook 项（例如，电子邮件、日历或联系人项）附加到撰写窗体中的邮件或约会。可以使用 [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法并访问 EWS 操作 [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)，获取用户邮箱中电子邮件、日历、联系人或任务项的 EWS ID。[item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性还提供阅读窗体中某个现有项的 EWS ID。</span><span class="sxs-lookup"><span data-stu-id="131c3-p108">You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the `addItemAttachmentAsync` method. You can get the EWS ID of an email, calendar, contact or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method and accessing the EWS operation [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). The [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property also provides the EWS ID of an existing item in a read form.</span></span>

<span data-ttu-id="131c3-136">下面的 JavaScript 函数`addItemAttachment`，扩展了上面的第一个示例，并将项添加为要撰写的电子邮件或约会的附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-136">The following JavaScript function, `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed.</span></span> <span data-ttu-id="131c3-137">此函数将要附加的项目的 EWS ID 作为实参。</span><span class="sxs-lookup"><span data-stu-id="131c3-137">The function takes as an argument the EWS ID of the item that is to be attached.</span></span> <span data-ttu-id="131c3-138">如果附加成功，它将获取附件 ID 以进行进一步处理，包括在同一会话中删除该附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-138">If attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.</span></span>

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
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> <span data-ttu-id="131c3-139">您可以使用撰写外接程序在 web 或移动设备上的 Outlook 中附加定期约会的实例。</span><span class="sxs-lookup"><span data-stu-id="131c3-139">You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web or mobile devices.</span></span> <span data-ttu-id="131c3-140">但是，在 Outlook 富客户端中，尝试附加实例将导致附加定期系列（主约会）。</span><span class="sxs-lookup"><span data-stu-id="131c3-140">However, in a supporting Outlook rich client, attempting to attach an instance would result in attaching the recurring series (the master appointment).</span></span>

## <a name="get-attachments"></a><span data-ttu-id="131c3-141">获取附件</span><span class="sxs-lookup"><span data-stu-id="131c3-141">Get attachments</span></span>

<span data-ttu-id="131c3-142">您可以使用[getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)方法来获取正在撰写的邮件或约会的附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-142">You can use the [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method to get the attachments of the message or appointment being composed.</span></span>

<span data-ttu-id="131c3-143">若要获取附件的内容，可以使用[getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)方法。</span><span class="sxs-lookup"><span data-stu-id="131c3-143">To get an attachment's content, you can use the [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="131c3-144">[AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)枚举中列出了受支持的格式。</span><span class="sxs-lookup"><span data-stu-id="131c3-144">The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.</span></span>

<span data-ttu-id="131c3-145">应通过使用`AsyncResult` output parameter 对象提供用于检查状态和任何错误的回调方法。</span><span class="sxs-lookup"><span data-stu-id="131c3-145">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="131c3-146">您还可以使用 optional `asyncContext`参数将任何其他参数传递给回调方法。</span><span class="sxs-lookup"><span data-stu-id="131c3-146">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="131c3-147">下面的 JavaScript 示例获取附件，并允许您为每个受支持的附件格式设置不同的处理。</span><span class="sxs-lookup"><span data-stu-id="131c3-147">The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.</span></span>

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

## <a name="remove-an-attachment"></a><span data-ttu-id="131c3-148">删除附件</span><span class="sxs-lookup"><span data-stu-id="131c3-148">Remove an attachment</span></span>

<span data-ttu-id="131c3-149">您可以指定相应的附件 ID，并使用 [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 方法在撰写窗体中从邮件或约会项目删除文件或项目附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-149">You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID and using the [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="131c3-150">您只应删除在同一会话中添加了同一外接程序的附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-150">You should only remove attachments that the same add-in has added in the same session.</span></span> <span data-ttu-id="131c3-151">与`addFileAttachmentAsync`和`addItemAttachmentAsync`方法类似， `removeAttachmentAsync`是一种异步方法。</span><span class="sxs-lookup"><span data-stu-id="131c3-151">Similar to the `addFileAttachmentAsync` and `addItemAttachmentAsync` methods, `removeAttachmentAsync` is an asynchronous method.</span></span> <span data-ttu-id="131c3-152">应通过使用`AsyncResult` output parameter 对象提供用于检查状态和任何错误的回调方法。</span><span class="sxs-lookup"><span data-stu-id="131c3-152">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="131c3-153">您还可以使用 optional `asyncContext`参数将任何其他参数传递给回调方法。</span><span class="sxs-lookup"><span data-stu-id="131c3-153">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="131c3-154">下面的 JavaScript 函数`removeAttachment`可继续扩展上面的示例，并从正在撰写的电子邮件或约会中删除指定的附件。</span><span class="sxs-lookup"><span data-stu-id="131c3-154">The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed.</span></span> <span data-ttu-id="131c3-155">此函数将要删除的附件的 ID 作为实参。</span><span class="sxs-lookup"><span data-stu-id="131c3-155">The function takes as an argument the ID of the attachment to be removed.</span></span> <span data-ttu-id="131c3-156">`addFileAttachmentAsync`可以在成功`addFileAttachmentFromBase64Async`、或`addItemAttachmentAsync`方法调用后获取附件的 ID，并将其存储为后续`removeAttachmentAsync`方法调用。</span><span class="sxs-lookup"><span data-stu-id="131c3-156">You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and store it for a subsequent `removeAttachmentAsync` method call.</span></span>

```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be
// removed.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
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

## <a name="see-also"></a><span data-ttu-id="131c3-157">另请参阅</span><span class="sxs-lookup"><span data-stu-id="131c3-157">See also</span></span>

- [<span data-ttu-id="131c3-158">创建适用于撰写窗体的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="131c3-158">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="131c3-159">Office 外接程序中的异步编程</span><span class="sxs-lookup"><span data-stu-id="131c3-159">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
