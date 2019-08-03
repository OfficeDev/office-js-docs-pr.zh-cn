---
title: Outlook 外接程序 API 要求集 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 537c99e8662d04ad006eb1bac6d76c0f753e8785
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064415"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="96193-102">Outlook 外接程序 API 要求集 1.1</span><span class="sxs-lookup"><span data-stu-id="96193-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="96193-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="96193-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="96193-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="96193-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-11"></a><span data-ttu-id="96193-105">1.1 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="96193-105">What's new in 1.1?</span></span>

<span data-ttu-id="96193-p101">要求集 1.1 包括要求集 1.0 的所有功能。它添加了外接程序访问邮件和约会的正文以及修改当前项的功能。</span><span class="sxs-lookup"><span data-stu-id="96193-p101">Requirement set 1.1 includes all of the features of Requirement set 1.0. It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="96193-108">更改日志</span><span class="sxs-lookup"><span data-stu-id="96193-108">Change log</span></span>

- <span data-ttu-id="96193-109">添加了 [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) 对象：提供用于在 Outlook 外接程序中添加和更新项目内容的方法。</span><span class="sxs-lookup"><span data-stu-id="96193-109">Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="96193-110">添加了 [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议地点的方法。</span><span class="sxs-lookup"><span data-stu-id="96193-110">Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="96193-111">添加了 [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="96193-111">Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="96193-112">添加了 [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的主题的方法。</span><span class="sxs-lookup"><span data-stu-id="96193-112">Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="96193-113">添加了 [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议开始或结束时间的方法。</span><span class="sxs-lookup"><span data-stu-id="96193-113">Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="96193-114">添加了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback)：将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="96193-114">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="96193-115">添加了 [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback)：将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="96193-115">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="96193-116">添加了 [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback)：将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="96193-116">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="96193-117">添加了 [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body)：获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="96193-117">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-body): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="96193-118">添加了邮件的["密件抄送"](office.context.mailbox.item.md#bcc-recipients)行。</span><span class="sxs-lookup"><span data-stu-id="96193-118">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipients) line of a message.</span></span>
- <span data-ttu-id="96193-119">添加了 [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1)：指定约会收件人的类型。</span><span class="sxs-lookup"><span data-stu-id="96193-119">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="96193-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="96193-120">See also</span></span>

- [<span data-ttu-id="96193-121">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="96193-121">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="96193-122">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="96193-122">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="96193-123">入门</span><span class="sxs-lookup"><span data-stu-id="96193-123">Get started</span></span>](/outlook/add-ins/quick-start)
