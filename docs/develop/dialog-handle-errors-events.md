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
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="34859-103">处理 Office 对话框中的错误和事件</span><span class="sxs-lookup"><span data-stu-id="34859-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="34859-104">本文介绍如何捕获和处理打开对话框时发生的错误以及对话框内发生的错误。</span><span class="sxs-lookup"><span data-stu-id="34859-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="34859-105">本文假定你熟悉使用 Office 对话框 API 的基础知识，如在 Office 加载项中使用 Office 对话框[API 中所述](dialog-api-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="34859-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="34859-106">另请参阅[Best practices and rules for the Office dialog API](dialog-best-practices.md)。</span><span class="sxs-lookup"><span data-stu-id="34859-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="34859-107">代码应处理两类事件：</span><span class="sxs-lookup"><span data-stu-id="34859-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="34859-108">`displayDialogAsync` 调用返回的错误，因为无法创建对话框。</span><span class="sxs-lookup"><span data-stu-id="34859-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="34859-109">对话框中的错误和其他事件。</span><span class="sxs-lookup"><span data-stu-id="34859-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="34859-110">DisplayDialogAsync 返回的错误</span><span class="sxs-lookup"><span data-stu-id="34859-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="34859-111">除了常规平台和系统错误之外，调用 还特定于四个错误 `displayDialogAsync` 。</span><span class="sxs-lookup"><span data-stu-id="34859-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="34859-112">代码编号</span><span class="sxs-lookup"><span data-stu-id="34859-112">Code number</span></span>|<span data-ttu-id="34859-113">含义</span><span class="sxs-lookup"><span data-stu-id="34859-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="34859-114">12004</span><span class="sxs-lookup"><span data-stu-id="34859-114">12004</span></span>|<span data-ttu-id="34859-p101">传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。</span><span class="sxs-lookup"><span data-stu-id="34859-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="34859-117">12005</span><span class="sxs-lookup"><span data-stu-id="34859-117">12005</span></span>|<span data-ttu-id="34859-118">传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。</span><span class="sxs-lookup"><span data-stu-id="34859-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="34859-119">必须使用 HTTPS。</span><span class="sxs-lookup"><span data-stu-id="34859-119">HTTPS is required.</span></span> <span data-ttu-id="34859-120"> (在某些版本的 Office 中，返回 12005 的错误消息文本与为 12004.) </span><span class="sxs-lookup"><span data-stu-id="34859-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="34859-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="34859-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="34859-p103">已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。</span><span class="sxs-lookup"><span data-stu-id="34859-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="34859-124">12009</span><span class="sxs-lookup"><span data-stu-id="34859-124">12009</span></span>|<span data-ttu-id="34859-125">用户已选择忽略对话框。</span><span class="sxs-lookup"><span data-stu-id="34859-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="34859-126">此错误可能发生在Office web 版，用户可能会选择不允许外接程序显示对话框。</span><span class="sxs-lookup"><span data-stu-id="34859-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="34859-127">有关详细信息，请参阅使用 Office web 版[处理弹出窗口阻止Office web 版。](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="34859-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="34859-128">调用 `displayDialogAsync` 时，它会将 [AsyncResult](/javascript/api/office/office.asyncresult) 对象传递给其回调函数。</span><span class="sxs-lookup"><span data-stu-id="34859-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="34859-129">调用成功后，对话框将打开，并且 `value` `AsyncResult` 对象的 属性是 [Dialog](/javascript/api/office/office.dialog) 对象。</span><span class="sxs-lookup"><span data-stu-id="34859-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="34859-130">有关此内容的示例，请参阅 [将信息从对话框发送到主机页](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)。</span><span class="sxs-lookup"><span data-stu-id="34859-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="34859-131">调用失败时，不会创建对话框，对象的 属性设置为 `displayDialogAsync` `status` `AsyncResult` `Office.AsyncResultStatus.Failed` ， `error` 并且填充对象的属性。</span><span class="sxs-lookup"><span data-stu-id="34859-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="34859-132">应始终提供回调，以在出现错误 `status` 时测试 并做出响应。</span><span class="sxs-lookup"><span data-stu-id="34859-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="34859-133">有关报告错误消息（无论其代码编号如何）的示例，请参阅以下代码。</span><span class="sxs-lookup"><span data-stu-id="34859-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="34859-134"> (`showNotification` 本文中未定义的 函数将显示或记录错误。</span><span class="sxs-lookup"><span data-stu-id="34859-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="34859-135">有关如何在加载项中实现此函数的示例，请参阅Office[加载项对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).) </span><span class="sxs-lookup"><span data-stu-id="34859-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

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

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="34859-136">对话框中的错误和事件</span><span class="sxs-lookup"><span data-stu-id="34859-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="34859-137">对话框中的三个错误和事件将在主机 `DialogEventReceived` 页中引发事件。</span><span class="sxs-lookup"><span data-stu-id="34859-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="34859-138">有关主机页的提醒，请参阅从主机页 [打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。</span><span class="sxs-lookup"><span data-stu-id="34859-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="34859-139">代码编号</span><span class="sxs-lookup"><span data-stu-id="34859-139">Code number</span></span>|<span data-ttu-id="34859-140">含义</span><span class="sxs-lookup"><span data-stu-id="34859-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="34859-141">12002</span><span class="sxs-lookup"><span data-stu-id="34859-141">12002</span></span>|<span data-ttu-id="34859-142">下列一种含义：</span><span class="sxs-lookup"><span data-stu-id="34859-142">One of the following:</span></span><br> <span data-ttu-id="34859-143">- 传递给 `displayDialogAsync` 的 URL 没有对应的页面。</span><span class="sxs-lookup"><span data-stu-id="34859-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="34859-144">- 传递到加载的页面，但对话框随后被重定向到找不到或加载的页面，或者已定向到具有无效语法的 `displayDialogAsync` URL。</span><span class="sxs-lookup"><span data-stu-id="34859-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="34859-145">12003</span><span class="sxs-lookup"><span data-stu-id="34859-145">12003</span></span>|<span data-ttu-id="34859-p107">对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。</span><span class="sxs-lookup"><span data-stu-id="34859-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="34859-148">12006</span><span class="sxs-lookup"><span data-stu-id="34859-148">12006</span></span>|<span data-ttu-id="34859-149">对话框已关闭，通常是因为用户选择了 **关闭按钮\*\*\*\*X**。</span><span class="sxs-lookup"><span data-stu-id="34859-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="34859-p108">代码可以在调用 `DialogEventReceived` 时为 `displayDialogAsync` 事件分配处理程序。下面展示了一个非常简单的示例。</span><span class="sxs-lookup"><span data-stu-id="34859-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example.</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="34859-152">有关为每个错误代码创建自定义错误消息的事件处理程序的示例， `DialogEventReceived` 请参阅以下示例。</span><span class="sxs-lookup"><span data-stu-id="34859-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example.</span></span>

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

<span data-ttu-id="34859-153">有关这样处理错误的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="34859-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>
