---
title: Outlook 加载项 API 要求集 1.5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d0489e4efa763b3963fcdc78ec894db46fa06362
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451813"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="280e2-102">Outlook 外接程序 API 要求集 1.5</span><span class="sxs-lookup"><span data-stu-id="280e2-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="280e2-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="280e2-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="280e2-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="280e2-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="280e2-105">1.5 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="280e2-105">What's new in 1.5?</span></span>

<span data-ttu-id="280e2-p101">要求集 1.5 包括[要求集 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) 的所有功能。它添加了下列功能。</span><span class="sxs-lookup"><span data-stu-id="280e2-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="280e2-108">添加了对[可固定任务窗格](/outlook/add-ins/pinnable-taskpane)的支持。</span><span class="sxs-lookup"><span data-stu-id="280e2-108">Added support for [pinnable task panes](/outlook/add-ins/pinnable-taskpane).</span></span>
- <span data-ttu-id="280e2-109">添加了对 [REST API](/outlook/add-ins/use-rest-api) 调用的支持。</span><span class="sxs-lookup"><span data-stu-id="280e2-109">Added support for calling [REST APIs](/outlook/add-ins/use-rest-api).</span></span>
- <span data-ttu-id="280e2-110">添加了将附件标记为内联的功能。</span><span class="sxs-lookup"><span data-stu-id="280e2-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="280e2-111">添加了关闭任务窗格或对话框的功能。</span><span class="sxs-lookup"><span data-stu-id="280e2-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="280e2-112">更改日志</span><span class="sxs-lookup"><span data-stu-id="280e2-112">Change log</span></span>

- <span data-ttu-id="280e2-113">添加了 [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback)：添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="280e2-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="280e2-114">添加了[removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): 删除受支持的事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="280e2-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#removehandlerasynceventtype-options-callback): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="280e2-115">添加了 [Office.EventType](office.md#eventtype-string)：指定与事件处理程序相关联的事件，并包括对 ItemChanged 事件的支持。</span><span class="sxs-lookup"><span data-stu-id="280e2-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="280e2-116">添加了 [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string)：获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="280e2-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="280e2-p102">修改了 [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback)：添加了此方法的新版本（具有新签名） (`getCallbackTokenAsync([options], callback)`)。原始版本仍可用且未更改。</span><span class="sxs-lookup"><span data-stu-id="280e2-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="280e2-119">添加了 [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--)。</span><span class="sxs-lookup"><span data-stu-id="280e2-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="280e2-120">修改了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback)：`options` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。</span><span class="sxs-lookup"><span data-stu-id="280e2-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="280e2-121">修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。</span><span class="sxs-lookup"><span data-stu-id="280e2-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="280e2-122">修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。</span><span class="sxs-lookup"><span data-stu-id="280e2-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="280e2-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="280e2-123">See also</span></span>

- [<span data-ttu-id="280e2-124">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="280e2-124">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="280e2-125">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="280e2-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="280e2-126">入门</span><span class="sxs-lookup"><span data-stu-id="280e2-126">Get started</span></span>](/outlook/add-ins/quick-start)
