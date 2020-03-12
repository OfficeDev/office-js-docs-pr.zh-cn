---
title: Outlook 加载项 API 要求集 1.5
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: f56cf4e13bdf3518ef14da6eca83b51abe82e50c
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42597045"
---
# <a name="outlook-add-in-api-requirement-set-15"></a><span data-ttu-id="750a9-102">Outlook 外接程序 API 要求集 1.5</span><span class="sxs-lookup"><span data-stu-id="750a9-102">Outlook add-in API requirement set 1.5</span></span>

<span data-ttu-id="750a9-103">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="750a9-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="750a9-104">本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="750a9-104">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-15"></a><span data-ttu-id="750a9-105">1.5 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="750a9-105">What's new in 1.5?</span></span>

<span data-ttu-id="750a9-p101">要求集 1.5 包括[要求集 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md) 的所有功能。它添加了下列功能。</span><span class="sxs-lookup"><span data-stu-id="750a9-p101">Requirement set 1.5 includes all of the features of [Requirement set 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). It added the following features.</span></span>

- <span data-ttu-id="750a9-108">添加了对[可固定任务窗格](../../../outlook/pinnable-taskpane.md)的支持。</span><span class="sxs-lookup"><span data-stu-id="750a9-108">Added support for [pinnable task panes](../../../outlook/pinnable-taskpane.md).</span></span>
- <span data-ttu-id="750a9-109">添加了对 [REST API](../../../outlook/use-rest-api.md) 调用的支持。</span><span class="sxs-lookup"><span data-stu-id="750a9-109">Added support for calling [REST APIs](../../../outlook/use-rest-api.md).</span></span>
- <span data-ttu-id="750a9-110">添加了将附件标记为内联的功能。</span><span class="sxs-lookup"><span data-stu-id="750a9-110">Added ability to mark an attachment as inline.</span></span>
- <span data-ttu-id="750a9-111">添加了关闭任务窗格或对话框的功能。</span><span class="sxs-lookup"><span data-stu-id="750a9-111">Added ability to close a task pane or dialog.</span></span>

### <a name="change-log"></a><span data-ttu-id="750a9-112">更改日志</span><span class="sxs-lookup"><span data-stu-id="750a9-112">Change log</span></span>

- <span data-ttu-id="750a9-113">添加了 [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods)：添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="750a9-113">Added [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#methods): Adds an event handler for a supported event.</span></span>
- <span data-ttu-id="750a9-114">添加了[removeHandlerAsync](office.context.mailbox.md#methods)：删除受支持的事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="750a9-114">Added [Office.context.mailbox.removeHandlerAsync](office.context.mailbox.md#methods): Removes the event handlers for a supported event type.</span></span>
- <span data-ttu-id="750a9-115">添加了 [Office.EventType](office.md#eventtype-string)：指定与事件处理程序相关联的事件，并包括对 ItemChanged 事件的支持。</span><span class="sxs-lookup"><span data-stu-id="750a9-115">Added [Office.EventType](office.md#eventtype-string): Specifies the event associated with an event handler and includes support for ItemChanged event.</span></span>
- <span data-ttu-id="750a9-116">添加了 [Office.context.mailbox.restUrl](office.context.mailbox.md#properties)：获取此电子邮件帐户的 REST 终结点的 URL。</span><span class="sxs-lookup"><span data-stu-id="750a9-116">Added [Office.context.mailbox.restUrl](office.context.mailbox.md#properties): Gets the URL of the REST endpoint for this email account.</span></span>
- <span data-ttu-id="750a9-p102">修改了 [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods)：添加了此方法的新版本（具有新签名） (`getCallbackTokenAsync([options], callback)`)。原始版本仍可用且未更改。</span><span class="sxs-lookup"><span data-stu-id="750a9-p102">Modified [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#methods): A new version of this method with a new signature (`getCallbackTokenAsync([options], callback)`) has been added. The original version is still available and is unchanged.</span></span>
- <span data-ttu-id="750a9-119">添加了 [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--)。</span><span class="sxs-lookup"><span data-stu-id="750a9-119">Added [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).</span></span>
- <span data-ttu-id="750a9-120">修改了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods)：`options` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。</span><span class="sxs-lookup"><span data-stu-id="750a9-120">Modified [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): A new value in the `options` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="750a9-121">修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。</span><span class="sxs-lookup"><span data-stu-id="750a9-121">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>
- <span data-ttu-id="750a9-122">修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods)：`formData.attachments` 字典中的新值调用 `isInline`，用于指定在邮件正文中内联使用了一个图像。</span><span class="sxs-lookup"><span data-stu-id="750a9-122">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): A new value in the `formData.attachments` dictionary called `isInline`, used to specify that an image is used inline in the message body.</span></span>

## <a name="see-also"></a><span data-ttu-id="750a9-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="750a9-123">See also</span></span>

- [<span data-ttu-id="750a9-124">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="750a9-124">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="750a9-125">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="750a9-125">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="750a9-126">入门</span><span class="sxs-lookup"><span data-stu-id="750a9-126">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="750a9-127">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="750a9-127">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
