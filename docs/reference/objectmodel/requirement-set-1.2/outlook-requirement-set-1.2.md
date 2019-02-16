---
title: Outlook 外接程序 API 要求集 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 1767b1b93f13de2c8a0731d2f08a1141b709b734
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068026"
---
# <a name="outlook-add-in-api-requirement-set-12"></a><span data-ttu-id="c90b0-102">Outlook 外接程序 API 要求集 1.2</span><span class="sxs-lookup"><span data-stu-id="c90b0-102">Outlook add-in API requirement set 1.2</span></span>

<span data-ttu-id="c90b0-103">适用于 Office 的 JavaScript API 的 Outlook 加载项 API 子集包括可以在 Outlook 加载项中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="c90b0-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="c90b0-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="c90b0-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span> 

## <a name="whats-new-in-12"></a><span data-ttu-id="c90b0-105">1.2 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="c90b0-105">What's new in 1.2?</span></span>

<span data-ttu-id="c90b0-p101">要求集 1.2 包括[要求集 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) 的所有功能。它添加了外接程序在用户的游标中插入文本的功能，无论是在邮件的主题还是正文中皆可插入文本。</span><span class="sxs-lookup"><span data-stu-id="c90b0-p101">Requirement set 1.2 includes all of the features of [Requirement set 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). It added the ability for add-ins to insert text at the user's cursor, either in the subject or the body of the message.</span></span>

### <a name="change-log"></a><span data-ttu-id="c90b0-108">更改日志</span><span class="sxs-lookup"><span data-stu-id="c90b0-108">Change log</span></span>

- <span data-ttu-id="c90b0-109">添加了 [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string)：以异步方式返回邮件主题或正文中的选定数据。</span><span class="sxs-lookup"><span data-stu-id="c90b0-109">Added [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string): Asynchronously returns selected data from the subject or body of a message.</span></span>
- <span data-ttu-id="c90b0-110">添加了 [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback)：以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="c90b0-110">Added [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback): Asynchronously inserts data into the body or subject of a message.</span></span>
- <span data-ttu-id="c90b0-111">修改了 [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="c90b0-111">Modified [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>
- <span data-ttu-id="c90b0-112">修改了 [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback)：将 `attachments` 属性添加到 `formData` 参数。</span><span class="sxs-lookup"><span data-stu-id="c90b0-112">Modified [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata-callback): Added `attachments` property to the `formData` parameter.</span></span>

## <a name="see-also"></a><span data-ttu-id="c90b0-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c90b0-113">See also</span></span>

- [<span data-ttu-id="c90b0-114">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="c90b0-114">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="c90b0-115">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="c90b0-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="c90b0-116">入门</span><span class="sxs-lookup"><span data-stu-id="c90b0-116">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
