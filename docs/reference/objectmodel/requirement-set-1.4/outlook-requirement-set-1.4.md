---
title: Outlook 加载项 API 要求集 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: cb4c8eecd63604aa633ade1a40eb5391b3a62ef2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165403"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="17925-102">Outlook 外接程序 API 要求集 1.4</span><span class="sxs-lookup"><span data-stu-id="17925-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="17925-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="17925-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="17925-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="17925-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="17925-105">1.4 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="17925-105">What's new in 1.4?</span></span>

<span data-ttu-id="17925-p101">要求集 1.4 包括[要求集 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) 的所有功能。它添加了对 `Office.ui` 命名空间的访问权限。</span><span class="sxs-lookup"><span data-stu-id="17925-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="17925-108">更改日志</span><span class="sxs-lookup"><span data-stu-id="17925-108">Change log</span></span>

- <span data-ttu-id="17925-109">添加了 [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)：在 Office 主机中显示一个对话框。</span><span class="sxs-lookup"><span data-stu-id="17925-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="17925-110">添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-)：将对话框中的消息传送到其父页/开始页。</span><span class="sxs-lookup"><span data-stu-id="17925-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="17925-111">添加了 [Dialog](/javascript/api/office/office.dialog) 对象：调用 [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) 方法时返回的对象。</span><span class="sxs-lookup"><span data-stu-id="17925-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="17925-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="17925-112">See also</span></span>

- [<span data-ttu-id="17925-113">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="17925-113">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="17925-114">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="17925-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="17925-115">入门</span><span class="sxs-lookup"><span data-stu-id="17925-115">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="17925-116">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="17925-116">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
