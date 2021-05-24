---
title: Outlook 外接程序 API 要求集 1.2
description: 为邮箱 API 1.2 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: d643f0fdf07c5f22d8d863075b894cfc05b21363
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590398"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="9c7fa-103">Outlook 外接程序 API 要求集 1.2</span><span class="sxs-lookup"><span data-stu-id="9c7fa-103">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="9c7fa-104">Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9c7fa-105">本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-12"></a><span data-ttu-id="9c7fa-106">1.2 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="9c7fa-106">What's new in 1.2?</span></span>

<span data-ttu-id="9c7fa-107">要求集 1.2 包括要求集 [1.1 的所有功能](../requirement-set-1.1/outlook-requirement-set-1.1.md)。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-107">Requirement set 1.2 includes all of the features of [requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md).</span></span> <span data-ttu-id="9c7fa-108">它添加了外接程序在用户的游标中插入文本的功能，无论是在邮件的主题还是正文中皆可插入文本。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-108">It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="9c7fa-109">更改日志</span><span class="sxs-lookup"><span data-stu-id="9c7fa-109">Change log</span></span>

- <span data-ttu-id="9c7fa-110">添加了 [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods)：以异步方式返回邮件主题或正文中的选定数据。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-110">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="9c7fa-111">添加了 [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods)：以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-111">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="9c7fa-112">修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-112">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="9c7fa-113">修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="9c7fa-113">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="9c7fa-114">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9c7fa-114">See also</span></span>

- [<span data-ttu-id="9c7fa-115">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="9c7fa-115">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="9c7fa-116">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="9c7fa-116">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9c7fa-117">入门</span><span class="sxs-lookup"><span data-stu-id="9c7fa-117">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="9c7fa-118">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="9c7fa-118">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
