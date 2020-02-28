---
title: Outlook 外接程序 API 要求集 1.2
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: d4fa18f3ab12e22ff30ef841d921f5dac89fd064
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325211"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="1cd00-102">Outlook 外接程序 API 要求集 1.2</span><span class="sxs-lookup"><span data-stu-id="1cd00-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="1cd00-103">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="1cd00-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1cd00-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="1cd00-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="1cd00-105">1.2 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="1cd00-105">What's new in 1.2?</span></span>

<span data-ttu-id="1cd00-p101">要求集 1.2 包括[要求集 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) 的所有功能。它添加了外接程序在用户的游标中插入文本的功能，无论是在邮件的主题还是正文中皆可插入文本。</span><span class="sxs-lookup"><span data-stu-id="1cd00-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="1cd00-108">更改日志</span><span class="sxs-lookup"><span data-stu-id="1cd00-108">Change log</span></span>

- <span data-ttu-id="1cd00-109">添加了 [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods)：以异步方式返回邮件主题或正文中的选定数据。</span><span class="sxs-lookup"><span data-stu-id="1cd00-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="1cd00-110">添加了 [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods)：以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="1cd00-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="1cd00-111">修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="1cd00-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="1cd00-112">修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="1cd00-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="1cd00-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1cd00-113">See also</span></span>

- [<span data-ttu-id="1cd00-114">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="1cd00-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="1cd00-115">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="1cd00-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="1cd00-116">入门</span><span class="sxs-lookup"><span data-stu-id="1cd00-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="1cd00-117">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="1cd00-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
