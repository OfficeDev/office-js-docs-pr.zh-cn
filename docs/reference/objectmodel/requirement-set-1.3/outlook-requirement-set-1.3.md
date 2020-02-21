---
title: Outlook 外接程序 API 要求集 1.3
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 375fc5d7cce8592b8e4a270713c1f611129cc7d0
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165424"
---
# <a name="outlook-add-in-api-requirement-set-13"></a><span data-ttu-id="abf95-102">Outlook 外接程序 API 要求集 1.3</span><span class="sxs-lookup"><span data-stu-id="abf95-102">Outlook add-in API requirement set 1.3</span></span>

<span data-ttu-id="abf95-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="abf95-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="abf95-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="abf95-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-13"></a><span data-ttu-id="abf95-105">1.3 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="abf95-105">What's new in 1.3?</span></span>

<span data-ttu-id="abf95-p101">要求集 1.3 包括[要求集 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) 的所有功能。它添加了下列功能。</span><span class="sxs-lookup"><span data-stu-id="abf95-p101">Requirement set 1.3 includes all of the features of [Requirement set 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). It added the following features.</span></span>

- <span data-ttu-id="abf95-108">添加了对[外接程序命令](../../../outlook/add-in-commands-for-outlook.md)的支持。</span><span class="sxs-lookup"><span data-stu-id="abf95-108">Added support for [add-in commands](../../../outlook/add-in-commands-for-outlook.md).</span></span>
- <span data-ttu-id="abf95-109">添加了保存或关闭正在撰写的项目的功能。</span><span class="sxs-lookup"><span data-stu-id="abf95-109">Added ability to save or close an item being composed.</span></span>
- <span data-ttu-id="abf95-110">增强的[Body](/javascript/api/outlook/office.body?view=outlook-js-1.3)对象，允许外接程序获取或设置整个正文。</span><span class="sxs-lookup"><span data-stu-id="abf95-110">Enhanced [Body](/javascript/api/outlook/office.body?view=outlook-js-1.3) object to allow add-ins to get or set the entire body.</span></span>
- <span data-ttu-id="abf95-111">添加了在 EWS 和 REST 格式之间转换 ID 的转换方法。</span><span class="sxs-lookup"><span data-stu-id="abf95-111">Added conversion methods to convert IDs between EWS and REST formats.</span></span>
- <span data-ttu-id="abf95-112">添加了将通知邮件添加到项目的信息栏中的功能。</span><span class="sxs-lookup"><span data-stu-id="abf95-112">Added ability to add notification messages to the info bar on items.</span></span>

### <a name="change-log"></a><span data-ttu-id="abf95-113">更改日志</span><span class="sxs-lookup"><span data-stu-id="abf95-113">Change log</span></span>

- <span data-ttu-id="abf95-114">添加了 [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-)：使用指定格式返回当前正文。</span><span class="sxs-lookup"><span data-stu-id="abf95-114">Added [Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#getasync-coerciontype--options--callback-): Returns the current body in a specified format.</span></span>
- <span data-ttu-id="abf95-115">添加了 [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-)：将整个正文替换为指定文本。</span><span class="sxs-lookup"><span data-stu-id="abf95-115">Added [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-1.3#setasync-data--options--callback-): Replaces the entire body with the specified text.</span></span>
- <span data-ttu-id="abf95-p102">添加了 [Event](/javascript/api/office/office.addincommands.event) 对象：作为参数传递到 Outlook 外接程序中的无用户界面命令函数。用来表示处理已完成。</span><span class="sxs-lookup"><span data-stu-id="abf95-p102">Added [Event](/javascript/api/office/office.addincommands.event) object: Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.</span></span>
- <span data-ttu-id="abf95-118">添加了 [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods)：关闭正在撰写的当前项。</span><span class="sxs-lookup"><span data-stu-id="abf95-118">Added [Office.context.mailbox.item.close](office.context.mailbox.item.md#methods): Closes the current item that is being composed.</span></span>
- <span data-ttu-id="abf95-119">添加了 [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods)：异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="abf95-119">Added [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#methods): Asynchronously saves an item.</span></span>
- <span data-ttu-id="abf95-120">添加了 [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties)：获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="abf95-120">Added [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#properties): Gets the notification messages for an item.</span></span>
- <span data-ttu-id="abf95-121">添加了 [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods)：将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="abf95-121">Added [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#methods): Converts an item ID formatted for REST into EWS format.</span></span>
- <span data-ttu-id="abf95-122">添加了 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods)：将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="abf95-122">Added [Office.context.mailbox.convertToRestId](office.context.mailbox.md#methods): Converts an item ID formatted for EWS into REST format.</span></span>
- <span data-ttu-id="abf95-123">添加了 [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3)：为约会或邮件指定通知邮件类型。</span><span class="sxs-lookup"><span data-stu-id="abf95-123">Added [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype?view=outlook-js-1.3): Specifies the notification message type for an appointment or message.</span></span>
- <span data-ttu-id="abf95-124">添加了 [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)：指定对应于 REST 格式的项目 ID 的 REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="abf95-124">Added [Office.MailboxEnums.RestVersion](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3): Specifies the version of the REST API that corresponds to a REST-formatted item ID.</span></span>
- <span data-ttu-id="abf95-125">添加了 [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3) 对象：提供用于访问 Outlook 外接程序中的通知邮件的方法。</span><span class="sxs-lookup"><span data-stu-id="abf95-125">Added [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.3) object: Provides methods for accessing notification messages in an Outlook add-in.</span></span>
- <span data-ttu-id="abf95-126">添加了 [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3) 类型：由 `NotificationMessages.getAllAsync` 方法返回。</span><span class="sxs-lookup"><span data-stu-id="abf95-126">Added [NotificationMessageDetails](/javascript/api/outlook/office.notificationmessagedetails?view=outlook-js-1.3) type: Returned by the `NotificationMessages.getAllAsync` method.</span></span>

## <a name="see-also"></a><span data-ttu-id="abf95-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="abf95-127">See also</span></span>

- [<span data-ttu-id="abf95-128">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="abf95-128">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="abf95-129">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="abf95-129">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="abf95-130">入门</span><span class="sxs-lookup"><span data-stu-id="abf95-130">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="abf95-131">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="abf95-131">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
