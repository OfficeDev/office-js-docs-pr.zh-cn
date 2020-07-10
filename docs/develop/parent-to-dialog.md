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
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a><span data-ttu-id="600ee-103">将数据和邮件从其主机页传递到对话框 (预览) </span><span class="sxs-lookup"><span data-stu-id="600ee-103">Passing data and messages to a dialog box from its host page (preview)</span></span>

<span data-ttu-id="600ee-104">您的外接程序可以使用[dialog](/javascript/api/office/office.dialog)对象的[messageChild](/javascript/api/office/office.dialog#messagechild-message-)方法，将邮件从[主机页面](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)发送到对话框。</span><span class="sxs-lookup"><span data-stu-id="600ee-104">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using the [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method of the [Dialog](/javascript/api/office/office.dialog) object.</span></span>

> [!Important]
>
> - <span data-ttu-id="600ee-105">本文中介绍的 Api 处于预览阶段。</span><span class="sxs-lookup"><span data-stu-id="600ee-105">The APIs described in this article are in preview.</span></span> <span data-ttu-id="600ee-106">它们可供开发人员用来进行试验;但不应在生产外接中使用。</span><span class="sxs-lookup"><span data-stu-id="600ee-106">They are available to developers for experimentation; but should not be used in a production add-in.</span></span> <span data-ttu-id="600ee-107">在发布此 API 之前，请使用将[信息传递到](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box)生产外接程序的对话框中所述的技术。</span><span class="sxs-lookup"><span data-stu-id="600ee-107">Until this API is released, use the techniques described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) for production add-ins.</span></span>
> - <span data-ttu-id="600ee-108">本文中所述的 Api 需要 Microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="600ee-108">The APIs described in this article require a Microsoft 365 subscription.</span></span> <span data-ttu-id="600ee-109">你应该使用来自预览体验成员频道的最新每月版本和内部版本。</span><span class="sxs-lookup"><span data-stu-id="600ee-109">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="600ee-110">你可能需要成为 Office 预览体验成员，才能获取此版本。</span><span class="sxs-lookup"><span data-stu-id="600ee-110">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="600ee-111">有关详细信息，请参阅[成为 Office 预览体验成员](https://insider.office.com)。</span><span class="sxs-lookup"><span data-stu-id="600ee-111">For more information, see [Be an Office Insider](https://insider.office.com).</span></span> <span data-ttu-id="600ee-112">请注意，当内部版本毕业生到生产半年频道时，将对该内部版本禁用对预览功能的支持。</span><span class="sxs-lookup"><span data-stu-id="600ee-112">Please note that when a build graduates to the production semi-annual channel, support for preview features is turned off for that build.</span></span>
> - <span data-ttu-id="600ee-113">在预览的初始阶段，Api 在 Excel、PowerPoint 和 Word 中受支持;而不是在 Outlook 中。</span><span class="sxs-lookup"><span data-stu-id="600ee-113">In the initial stage of the preview, the APIs are supported in Excel, PowerPoint, and Word; but not in Outlook.</span></span>
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="600ee-114">`messageChild()`从主机页使用</span><span class="sxs-lookup"><span data-stu-id="600ee-114">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="600ee-115">调用 Office 对话框 API 打开对话框时，将返回[dialog](/javascript/api/office/office.dialog)对象。</span><span class="sxs-lookup"><span data-stu-id="600ee-115">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="600ee-116">应将其分配给变量，该变量的作用域通常大于[displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-)方法，因为该对象将被其他方法引用。</span><span class="sxs-lookup"><span data-stu-id="600ee-116">It should be assigned to a variable, which typically has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="600ee-117">示例如下：</span><span class="sxs-lookup"><span data-stu-id="600ee-117">The following is an example:</span></span>

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

<span data-ttu-id="600ee-118">此 `Dialog` 对象具有向对话框发送任何字符串或字符串化数据的[messageChild](/javascript/api/office/office.dialog#messagechild-message-)方法。</span><span class="sxs-lookup"><span data-stu-id="600ee-118">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, or stringified data, to the dialog box.</span></span> <span data-ttu-id="600ee-119">这 `DialogParentMessageReceived` 将在对话框中引发事件。</span><span class="sxs-lookup"><span data-stu-id="600ee-119">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="600ee-120">您的代码应处理此事件，如下一节中所示。</span><span class="sxs-lookup"><span data-stu-id="600ee-120">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="600ee-121">假设对话框的 UI 应与当前活动的工作表关联，并且该工作表相对于其他工作表的位置。</span><span class="sxs-lookup"><span data-stu-id="600ee-121">Consider a scenario in which the UI of the dialog should correlate with the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="600ee-122">在下面的示例中， `sheetPropertiesChanged` 将 Excel 工作表属性发送到对话框。</span><span class="sxs-lookup"><span data-stu-id="600ee-122">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="600ee-123">在这种情况下，当前工作表名为 "我的工作表"，而它是工作簿中的第二个工作表。</span><span class="sxs-lookup"><span data-stu-id="600ee-123">In this case the current worksheet is named "My Sheet" and it is the 2nd sheet in the workbook.</span></span> <span data-ttu-id="600ee-124">数据封装在字符串化对象中，以便可以将其传递给 `messageChild` 。</span><span class="sxs-lookup"><span data-stu-id="600ee-124">The data is encapsulated in an object which is stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="600ee-125">在对话框中处理 DialogParentMessageReceived</span><span class="sxs-lookup"><span data-stu-id="600ee-125">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="600ee-126">在对话框的 JavaScript 中， `DialogParentMessageReceived` 使用[addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-)方法为事件注册处理程序。</span><span class="sxs-lookup"><span data-stu-id="600ee-126">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="600ee-127">这通常是在[onReady 或 Office.initialize 方法](initialize-add-in.md)中完成的。</span><span class="sxs-lookup"><span data-stu-id="600ee-127">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md).</span></span> <span data-ttu-id="600ee-128">示例如下：</span><span class="sxs-lookup"><span data-stu-id="600ee-128">The following is an example:</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="600ee-129">然后，定义该 `onMessageFromParent` 处理程序。</span><span class="sxs-lookup"><span data-stu-id="600ee-129">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="600ee-130">下面的代码将继续上一节中的示例。</span><span class="sxs-lookup"><span data-stu-id="600ee-130">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="600ee-131">请注意，Office 会将参数传递给处理程序，并确保 `message` argument 对象的属性包含主机页中的字符串。</span><span class="sxs-lookup"><span data-stu-id="600ee-131">Note that Office passes an argument to the handler and that the `message` property of argument object contains the string from the host page.</span></span> <span data-ttu-id="600ee-132">在此示例中，邮件被 reconverted 到对象，jQuery 用于将对话框的顶部标题设置为与新工作表名称相匹配。</span><span class="sxs-lookup"><span data-stu-id="600ee-132">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="600ee-133">最佳做法是验证是否正确注册了处理程序。</span><span class="sxs-lookup"><span data-stu-id="600ee-133">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="600ee-134">为此，可以将回调传递给在 `addHandlerAsync` 注册处理程序的尝试完成时运行的方法。</span><span class="sxs-lookup"><span data-stu-id="600ee-134">You can do this by passing a callback to the `addHandlerAsync` method that runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="600ee-135">如果未成功注册处理程序，请使用该处理程序记录或显示错误。</span><span class="sxs-lookup"><span data-stu-id="600ee-135">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="600ee-136">示例如下。</span><span class="sxs-lookup"><span data-stu-id="600ee-136">The following is an example.</span></span> <span data-ttu-id="600ee-137">请注意，这 `reportError` 是未在此处定义的函数，它会记录或显示错误。</span><span class="sxs-lookup"><span data-stu-id="600ee-137">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

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

## <a name="conditional-messaging"></a><span data-ttu-id="600ee-138">条件消息</span><span class="sxs-lookup"><span data-stu-id="600ee-138">Conditional messaging</span></span>

<span data-ttu-id="600ee-139">由于可以 `messageChild` 从主机页进行多次调用，但在该事件的对话框中只有一个处理程序 `DialogParentMessageReceived` ，因此处理程序必须使用条件逻辑来区分不同的消息。</span><span class="sxs-lookup"><span data-stu-id="600ee-139">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="600ee-140">您可以按照与[条件消息](dialog-api-in-office-add-ins.md#conditional-messaging)中所述的方式将消息发送到主机页时，精确地与构造条件消息传递的方式完全并行。</span><span class="sxs-lookup"><span data-stu-id="600ee-140">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](dialog-api-in-office-add-ins.md#conditional-messaging).</span></span>
