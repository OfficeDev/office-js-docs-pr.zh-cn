---
title: Outlook 加载项 API 要求集 1.4
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: be700af413a041502cddd491f304a693c259da28
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432366"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="3bb6d-102">Outlook 外接程序 API 要求集 1.4</span><span class="sxs-lookup"><span data-stu-id="3bb6d-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="3bb6d-103">适用于 Office 的 JavaScript API 的 Outlook 加载项 API 子集包括可以在 Outlook 加载项中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="3bb6d-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="3bb6d-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="3bb6d-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="3bb6d-105">1.4 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="3bb6d-105">What's new in 1.4?</span></span>

<span data-ttu-id="3bb6d-p101">要求集 1.4 包括[要求集 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) 的所有功能。它添加了对 `Office.ui` 命名空间的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3bb6d-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="3bb6d-108">更改日志</span><span class="sxs-lookup"><span data-stu-id="3bb6d-108">Change log</span></span>

- <span data-ttu-id="3bb6d-109">添加了 [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)：在 Office 主机中显示一个对话框。</span><span class="sxs-lookup"><span data-stu-id="3bb6d-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="3bb6d-110">添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-)：将对话框中的消息传送到其父页/开始页。</span><span class="sxs-lookup"><span data-stu-id="3bb6d-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="3bb6d-111">添加了 [Dialog](/javascript/api/office/office.dialog) 对象：调用 [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) 方法时返回的对象。</span><span class="sxs-lookup"><span data-stu-id="3bb6d-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="3bb6d-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3bb6d-112">See also</span></span>

- [<span data-ttu-id="3bb6d-113">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="3bb6d-113">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="3bb6d-114">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="3bb6d-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="3bb6d-115">入门</span><span class="sxs-lookup"><span data-stu-id="3bb6d-115">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)