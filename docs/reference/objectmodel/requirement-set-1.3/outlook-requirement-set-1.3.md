---
title: Outlook 外接程序 API 要求集 1.3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 276096870b128896e987bcb303b4cccdb77e0e50
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871274"
---
# <a name="outlook-add-in-api-requirement-set-13"></a><span data-ttu-id="802e6-102">Outlook 外接程序 API 要求集 1.3</span><span class="sxs-lookup"><span data-stu-id="802e6-102">Outlook add-in API requirement set 1.3</span></span>

<span data-ttu-id="802e6-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="802e6-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="802e6-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="802e6-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-13"></a><span data-ttu-id="802e6-105">1.3 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="802e6-105">What's new in 1.3?</span></span>

<span data-ttu-id="802e6-p101">要求集 1.3 包括[要求集 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) 的所有功能。它添加了下列功能。</span><span class="sxs-lookup"><span data-stu-id="802e6-p101">Requirement set 1.3 includes all of the features of [Requirement set 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md). It added the following features.</span></span>

- <span data-ttu-id="802e6-108">添加了对[外接程序命令](/outlook/add-ins/add-in-commands-for-outlook)的支持。</span><span class="sxs-lookup"><span data-stu-id="802e6-108">Added support for [add-in commands](/outlook/add-ins/add-in-commands-for-outlook).</span></span>
- <span data-ttu-id="802e6-109">添加了保存或关闭正在撰写的项目的功能。</span><span class="sxs-lookup"><span data-stu-id="802e6-109">Added ability to save or close an item being composed.</span></span>
- <span data-ttu-id="802e6-110">改进了 [Body](/javascript/api/outlook_1_3/office.body) 对象，允许外接程序获取或设置整个正文。</span><span class="sxs-lookup"><span data-stu-id="802e6-110">Enhanced [Body](/javascript/api/outlook_1_3/office.body) object to allow addins to get or set the entire body.</span></span>
- <span data-ttu-id="802e6-111">添加了在 EWS 和 REST 格式之间转换 ID 的转换方法。</span><span class="sxs-lookup"><span data-stu-id="802e6-111">Added conversion methods to convert IDs between EWS and REST formats.</span></span>
- <span data-ttu-id="802e6-112">添加了将通知邮件添加到项目的信息栏中的功能。</span><span class="sxs-lookup"><span data-stu-id="802e6-112">Added ability to add notification messages to the info bar on items.</span></span>

### <a name="change-log"></a><span data-ttu-id="802e6-113">更改日志</span><span class="sxs-lookup"><span data-stu-id="802e6-113">Change log</span></span>

- <span data-ttu-id="802e6-114">添加了 [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-)：使用指定格式返回当前正文。</span><span class="sxs-lookup"><span data-stu-id="802e6-114">Added [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-): Returns the current body in a specified format.</span></span>
- <span data-ttu-id="802e6-115">添加了 [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-)：将整个正文替换为指定文本。</span><span class="sxs-lookup"><span data-stu-id="802e6-115">Added [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-): Replaces the entire body with the specified text.</span></span>
- <span data-ttu-id="802e6-116">添加了 [Office.context.officeTheme](office.context.md#officetheme-object)：提供了对 Office 主题颜色的访问权限。</span><span class="sxs-lookup"><span data-stu-id="802e6-116">Added [Office.context.officeTheme](office.context.md#officetheme-object): Provides access to the Office theme colors.</span></span>
- <span data-ttu-id="802e6-p102">添加了 [Event](/javascript/api/office/office.addincommands.event) 对象：作为参数传递到 Outlook 外接程序中的无用户界面命令函数。用来表示处理已完成。</span><span class="sxs-lookup"><span data-stu-id="802e6-p102">Added [Event](/javascript/api/office/office.addincommands.event) object: Passed as a parameter to UI-less command functions in an Outlook add-in. Used to signal completion of processing.</span></span>
- <span data-ttu-id="802e6-119">添加了 [Office.context.mailbox.item.close](office.context.mailbox.item.md#close)：关闭正在撰写的当前项。</span><span class="sxs-lookup"><span data-stu-id="802e6-119">Added [Office.context.mailbox.item.close](office.context.mailbox.item.md#close): Closes the current item that is being composed.</span></span>
- <span data-ttu-id="802e6-120">添加了 [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback)：异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="802e6-120">Added [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback): Asynchronously saves an item.</span></span>
- <span data-ttu-id="802e6-121">添加了 [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessages)：获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="802e6-121">Added [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessages): Gets the notification messages for an item.</span></span>
- <span data-ttu-id="802e6-122">添加了 [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string)：将项目 ID 格式化（从 REST 转换为 EWS 格式）。</span><span class="sxs-lookup"><span data-stu-id="802e6-122">Added [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string): Converts an item ID formatted for REST into EWS format.</span></span>
- <span data-ttu-id="802e6-123">添加了 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)：将项目 ID 格式化（从 EWS 转换为 REST 格式）。</span><span class="sxs-lookup"><span data-stu-id="802e6-123">Added [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string): Converts an item ID formatted for EWS into REST format.</span></span>
- <span data-ttu-id="802e6-124">添加了 [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype)：为约会或邮件指定通知邮件类型。</span><span class="sxs-lookup"><span data-stu-id="802e6-124">Added [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype): Specifies the notification message type for an appointment or message.</span></span>
- <span data-ttu-id="802e6-125">添加了 [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion)：指定对应于 REST 格式的项目 ID 的 REST API 的版本。</span><span class="sxs-lookup"><span data-stu-id="802e6-125">Added [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion): Specifies the version of the REST API that corresponds to a REST-formatted item ID.</span></span>
- <span data-ttu-id="802e6-126">添加了 [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) 对象：提供用于访问 Outlook 外接程序中的通知邮件的方法。</span><span class="sxs-lookup"><span data-stu-id="802e6-126">Added [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) object: Provides methods for accessing notification messages in an Outlook add-in.</span></span>
- <span data-ttu-id="802e6-127">添加了 [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) 类型：由 `NotificationMessages.getAllAsync` 方法返回。</span><span class="sxs-lookup"><span data-stu-id="802e6-127">Added [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) type: Returned by the `NotificationMessages.getAllAsync` method.</span></span>

## <a name="see-also"></a><span data-ttu-id="802e6-128">另请参阅</span><span class="sxs-lookup"><span data-stu-id="802e6-128">See also</span></span>

- [<span data-ttu-id="802e6-129">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="802e6-129">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="802e6-130">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="802e6-130">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="802e6-131">入门</span><span class="sxs-lookup"><span data-stu-id="802e6-131">Get started</span></span>](/outlook/add-ins/quick-start)
