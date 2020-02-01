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
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="4390a-103">处理 Office 对话框中的错误和事件</span><span class="sxs-lookup"><span data-stu-id="4390a-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="4390a-104">本文介绍如何在打开对话框以及对话框中发生的错误时捕获和处理错误。</span><span class="sxs-lookup"><span data-stu-id="4390a-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="4390a-105">本文 presupposes 您熟悉使用 Office 对话框 API 的基础知识，如在[Office 外接程序中使用 office 对话框 api](dialog-api-in-office-add-ins.md)中所述。</span><span class="sxs-lookup"><span data-stu-id="4390a-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="4390a-106">另请参阅[Office 对话框 API 的最佳实践和规则](dialog-best-practices.md)。</span><span class="sxs-lookup"><span data-stu-id="4390a-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="4390a-107">代码应处理两类事件：</span><span class="sxs-lookup"><span data-stu-id="4390a-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="4390a-108">`displayDialogAsync` 调用返回的错误，因为无法创建对话框。</span><span class="sxs-lookup"><span data-stu-id="4390a-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="4390a-109">对话框中的错误和其他事件。</span><span class="sxs-lookup"><span data-stu-id="4390a-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="4390a-110">DisplayDialogAsync 返回的错误</span><span class="sxs-lookup"><span data-stu-id="4390a-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="4390a-111">除了常规平台和系统错误之外，还有四个错误特定于调用`displayDialogAsync`。</span><span class="sxs-lookup"><span data-stu-id="4390a-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="4390a-112">代码编号</span><span class="sxs-lookup"><span data-stu-id="4390a-112">Code number</span></span>|<span data-ttu-id="4390a-113">含义</span><span class="sxs-lookup"><span data-stu-id="4390a-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="4390a-114">12004</span><span class="sxs-lookup"><span data-stu-id="4390a-114">12004</span></span>|<span data-ttu-id="4390a-p101">传递给 `displayDialogAsync` 的 URL 的域不受信任。此域必须与主机页的域相同（包括协议和端口号）。</span><span class="sxs-lookup"><span data-stu-id="4390a-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="4390a-117">12005</span><span class="sxs-lookup"><span data-stu-id="4390a-117">12005</span></span>|<span data-ttu-id="4390a-118">传递给 `displayDialogAsync` 的 URL 使用 HTTP 协议。</span><span class="sxs-lookup"><span data-stu-id="4390a-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="4390a-119">必须使用 HTTPS。</span><span class="sxs-lookup"><span data-stu-id="4390a-119">HTTPS is required.</span></span> <span data-ttu-id="4390a-120">（在 Office 的某些版本中，返回12005的错误消息文本与为12004返回的文本相同。</span><span class="sxs-lookup"><span data-stu-id="4390a-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="4390a-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="4390a-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="4390a-p103">已从此主机窗口打开了一个对话框。主机窗口（如任务窗格）一次只能打开一个对话框。</span><span class="sxs-lookup"><span data-stu-id="4390a-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="4390a-124">12009</span><span class="sxs-lookup"><span data-stu-id="4390a-124">12009</span></span>|<span data-ttu-id="4390a-125">用户已选择忽略对话框。</span><span class="sxs-lookup"><span data-stu-id="4390a-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="4390a-126">此错误可能发生在 web 上的 Office 中，用户可以在其中选择不允许外接程序呈现对话框。</span><span class="sxs-lookup"><span data-stu-id="4390a-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="4390a-127">有关详细信息，请参阅[使用 web 上的 Office 处理弹出窗口阻止程序](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)。</span><span class="sxs-lookup"><span data-stu-id="4390a-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="4390a-128">调用`displayDialogAsync`时，它会将[AsyncResult](/javascript/api/office/office.asyncresult)对象传递给它的回调函数。</span><span class="sxs-lookup"><span data-stu-id="4390a-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="4390a-129">当调用成功时，将打开对话框，并且`value` `AsyncResult`对象的属性是[dialog](/javascript/api/office/office.dialog)对象。</span><span class="sxs-lookup"><span data-stu-id="4390a-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="4390a-130">有关这种情况的示例，请参阅[将信息从对话框发送到主机页](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)。</span><span class="sxs-lookup"><span data-stu-id="4390a-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="4390a-131">当`displayDialogAsync`调用失败时，将不会创建对话框， `status` `AsyncResult`对象的属性设置为`Office.AsyncResultStatus.Failed`，并填充对象的`error`属性。</span><span class="sxs-lookup"><span data-stu-id="4390a-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="4390a-132">应始终提供一个回调，以在出错`status`时测试和响应。</span><span class="sxs-lookup"><span data-stu-id="4390a-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="4390a-133">有关报告错误消息（而不考虑其代码编号）的示例，请参阅以下代码。</span><span class="sxs-lookup"><span data-stu-id="4390a-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="4390a-134">（本文`showNotification`中未定义的函数可能显示或记录错误。</span><span class="sxs-lookup"><span data-stu-id="4390a-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="4390a-135">有关如何在外接程序中实现此函数的示例，请参阅[Office 外接程序对话框 API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。）</span><span class="sxs-lookup"><span data-stu-id="4390a-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

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

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="4390a-136">对话框中的错误和事件</span><span class="sxs-lookup"><span data-stu-id="4390a-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="4390a-137">对话框中的三个错误和事件将引发主机`DialogEventReceived`页中的事件。</span><span class="sxs-lookup"><span data-stu-id="4390a-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="4390a-138">有关主机页面的提示，请参阅[从主机页面打开对话框](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)。</span><span class="sxs-lookup"><span data-stu-id="4390a-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="4390a-139">代码编号</span><span class="sxs-lookup"><span data-stu-id="4390a-139">Code number</span></span>|<span data-ttu-id="4390a-140">含义</span><span class="sxs-lookup"><span data-stu-id="4390a-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="4390a-141">12002</span><span class="sxs-lookup"><span data-stu-id="4390a-141">12002</span></span>|<span data-ttu-id="4390a-142">下列一种含义：</span><span class="sxs-lookup"><span data-stu-id="4390a-142">One of the following:</span></span><br> <span data-ttu-id="4390a-143">- 传递给 `displayDialogAsync` 的 URL 没有对应的页面。</span><span class="sxs-lookup"><span data-stu-id="4390a-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="4390a-144">-已传递给`displayDialogAsync`加载的页面，但随后会将该对话框重定向到无法找到或加载的页面，或者已将其定向到语法无效的 URL。</span><span class="sxs-lookup"><span data-stu-id="4390a-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="4390a-145">12003</span><span class="sxs-lookup"><span data-stu-id="4390a-145">12003</span></span>|<span data-ttu-id="4390a-p107">对话框定向到使用 HTTP 协议的 URL。必须使用 HTTPS。</span><span class="sxs-lookup"><span data-stu-id="4390a-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="4390a-148">12006</span><span class="sxs-lookup"><span data-stu-id="4390a-148">12006</span></span>|<span data-ttu-id="4390a-149">对话框已关闭，通常是因为用户选择了 "**关闭**" 按钮**X**。</span><span class="sxs-lookup"><span data-stu-id="4390a-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="4390a-p108">代码可以在调用 `displayDialogAsync` 时分配 `DialogEventReceived` 事件处理程序。下面展示了一个简单示例：</span><span class="sxs-lookup"><span data-stu-id="4390a-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="4390a-152">有关为每个错误代码创建自定义错误消息的 `DialogEventReceived` 事件处理程序示例，请参阅下面的示例：</span><span class="sxs-lookup"><span data-stu-id="4390a-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

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

<span data-ttu-id="4390a-153">有关这样处理错误的样本加载项，请参阅 [Office 加载项 Dialog API 示例](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)。</span><span class="sxs-lookup"><span data-stu-id="4390a-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>
