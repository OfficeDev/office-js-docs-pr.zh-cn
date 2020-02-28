---
title: Outlook 外接程序 API 要求集 1.1
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 159cfb223efff3893bce71687475c5e512b37ede
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324783"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="66ee8-102">Outlook 外接程序 API 要求集 1.1</span><span class="sxs-lookup"><span data-stu-id="66ee8-102">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="66ee8-103">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="66ee8-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span> <span data-ttu-id="66ee8-104">Outlook JavaScript API 1.1 （邮箱1.1）是第一个 API 版本。</span><span class="sxs-lookup"><span data-stu-id="66ee8-104">Outlook JavaScript API 1.1 (Mailbox 1.1) is the first version of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="66ee8-105">本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="66ee8-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-11"></a><span data-ttu-id="66ee8-106">1.1 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="66ee8-106">What's new in 1.1?</span></span>

<span data-ttu-id="66ee8-107">要求集1.1 包括在 Outlook 中支持的所有[通用 API 要求集](../../requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="66ee8-107">Requirement set 1.1 includes all of the [Common API requirement sets](../../requirement-sets/office-add-in-requirement-sets.md) supported in Outlook.</span></span> <span data-ttu-id="66ee8-108">它添加了外接程序访问邮件和约会的正文以及修改当前项的功能。</span><span class="sxs-lookup"><span data-stu-id="66ee8-108">It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="66ee8-109">更改日志</span><span class="sxs-lookup"><span data-stu-id="66ee8-109">Change log</span></span>

- <span data-ttu-id="66ee8-110">添加了 [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) 对象：提供用于在 Outlook 外接程序中添加和更新项目内容的方法。</span><span class="sxs-lookup"><span data-stu-id="66ee8-110">Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="66ee8-111">添加了 [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议地点的方法。</span><span class="sxs-lookup"><span data-stu-id="66ee8-111">Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="66ee8-112">添加了 [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="66ee8-112">Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="66ee8-113">添加了 [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的主题的方法。</span><span class="sxs-lookup"><span data-stu-id="66ee8-113">Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="66ee8-114">添加了 [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议开始或结束时间的方法。</span><span class="sxs-lookup"><span data-stu-id="66ee8-114">Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="66ee8-115">添加了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods)：将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="66ee8-115">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="66ee8-116">添加了 [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods)：将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="66ee8-116">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="66ee8-117">添加了 [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods)：将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="66ee8-117">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="66ee8-118">添加了 [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties)：获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="66ee8-118">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="66ee8-119">添加了邮件的["密件抄送"](office.context.mailbox.item.md#properties)行。</span><span class="sxs-lookup"><span data-stu-id="66ee8-119">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) line of a message.</span></span>
- <span data-ttu-id="66ee8-120">添加了 [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1)：指定约会收件人的类型。</span><span class="sxs-lookup"><span data-stu-id="66ee8-120">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="66ee8-121">另请参阅</span><span class="sxs-lookup"><span data-stu-id="66ee8-121">See also</span></span>

- [<span data-ttu-id="66ee8-122">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="66ee8-122">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="66ee8-123">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="66ee8-123">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="66ee8-124">入门</span><span class="sxs-lookup"><span data-stu-id="66ee8-124">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="66ee8-125">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="66ee8-125">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
