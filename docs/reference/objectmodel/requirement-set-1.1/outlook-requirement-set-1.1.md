---
title: Outlook 外接程序 API 要求集 1.1
description: 作为邮箱 API 1.1 的一部分引入的 Outlook 外接程序和 Office JavaScript Api 的功能和 Api。
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: a6d2d352b2882bf0e5de994c8924bbb99ebb9dfb
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610815"
---
# <a name="outlook-add-in-api-requirement-set-11"></a><span data-ttu-id="3d978-103">Outlook 外接程序 API 要求集 1.1</span><span class="sxs-lookup"><span data-stu-id="3d978-103">Outlook add-in API requirement set 1.1</span></span>

<span data-ttu-id="3d978-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="3d978-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span> <span data-ttu-id="3d978-105">Outlook JavaScript API 1.1 （邮箱1.1）是第一个 API 版本。</span><span class="sxs-lookup"><span data-stu-id="3d978-105">Outlook JavaScript API 1.1 (Mailbox 1.1) is the first version of the API.</span></span>

> [!NOTE]
> <span data-ttu-id="3d978-106">本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3d978-106">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-11"></a><span data-ttu-id="3d978-107">1.1 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="3d978-107">What's new in 1.1?</span></span>

<span data-ttu-id="3d978-108">要求集1.1 包括在 Outlook 中支持的所有[通用 API 要求集](../../requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3d978-108">Requirement set 1.1 includes all of the [Common API requirement sets](../../requirement-sets/office-add-in-requirement-sets.md) supported in Outlook.</span></span> <span data-ttu-id="3d978-109">它添加了外接程序访问邮件和约会的正文以及修改当前项的功能。</span><span class="sxs-lookup"><span data-stu-id="3d978-109">It added the ability for add-ins to access the body of messages and appointments and the ability to modify the current item.</span></span>

### <a name="change-log"></a><span data-ttu-id="3d978-110">更改日志</span><span class="sxs-lookup"><span data-stu-id="3d978-110">Change log</span></span>

- <span data-ttu-id="3d978-111">添加了 [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) 对象：提供用于在 Outlook 外接程序中添加和更新项目内容的方法。</span><span class="sxs-lookup"><span data-stu-id="3d978-111">Added [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1) object: Provides methods for adding and updating the content of an item in an Outlook add-in.</span></span>
- <span data-ttu-id="3d978-112">添加了 [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议地点的方法。</span><span class="sxs-lookup"><span data-stu-id="3d978-112">Added [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1) object: Provides methods to get and set the location of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="3d978-113">添加了 [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="3d978-113">Added [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1) object: Provides methods to get and set the recipients of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="3d978-114">添加了 [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的约会或邮件的主题的方法。</span><span class="sxs-lookup"><span data-stu-id="3d978-114">Added [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1) object: Provides methods to get and set the subject of an appointment or message in an Outlook add-in.</span></span>
- <span data-ttu-id="3d978-115">添加了 [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) 对象：提供用于获取和设置 Outlook 外接程序中的会议开始或结束时间的方法。</span><span class="sxs-lookup"><span data-stu-id="3d978-115">Added [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1) object: Provides methods to get and set the start or end time of a meeting in an Outlook add-in.</span></span>
- <span data-ttu-id="3d978-116">添加了 [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods)：将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3d978-116">Added [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods): Adds a file to a message or appointment as an attachment.</span></span>
- <span data-ttu-id="3d978-117">添加了 [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods)：将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3d978-117">Added [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods): Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>
- <span data-ttu-id="3d978-118">添加了 [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods)：将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="3d978-118">Added [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods): Removes an attachment from a message or appointment.</span></span>
- <span data-ttu-id="3d978-119">添加了 [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties)：获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="3d978-119">Added [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties): Gets an object that provides methods for manipulating the body of an item.</span></span>
- <span data-ttu-id="3d978-120">添加了邮件的["密件抄送"](office.context.mailbox.item.md#properties)行。</span><span class="sxs-lookup"><span data-stu-id="3d978-120">Added [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#properties) line of a message.</span></span>
- <span data-ttu-id="3d978-121">添加了 [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1)：指定约会收件人的类型。</span><span class="sxs-lookup"><span data-stu-id="3d978-121">Added [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1): Specifies the type of recipient for an appointment.</span></span>

## <a name="see-also"></a><span data-ttu-id="3d978-122">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3d978-122">See also</span></span>

- [<span data-ttu-id="3d978-123">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="3d978-123">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="3d978-124">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="3d978-124">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="3d978-125">入门</span><span class="sxs-lookup"><span data-stu-id="3d978-125">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="3d978-126">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="3d978-126">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
