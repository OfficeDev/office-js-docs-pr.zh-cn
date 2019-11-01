---
title: Outlook 加载项 API 要求集 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: ded3aa8cdcde1132f074f55356e64bb4468919ff
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901939"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="b41b2-102">Outlook 外接程序 API 要求集 1.4</span><span class="sxs-lookup"><span data-stu-id="b41b2-102">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="b41b2-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="b41b2-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b41b2-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="b41b2-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="b41b2-105">1.4 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="b41b2-105">What's new in 1.4?</span></span>

<span data-ttu-id="b41b2-p101">要求集 1.4 包括[要求集 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) 的所有功能。它添加了对 `Office.ui` 命名空间的访问权限。</span><span class="sxs-lookup"><span data-stu-id="b41b2-p101">Requirement set 1.4 includes all of the features of [Requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="b41b2-108">更改日志</span><span class="sxs-lookup"><span data-stu-id="b41b2-108">Change log</span></span>

- <span data-ttu-id="b41b2-109">添加了 [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)：在 Office 主机中显示一个对话框。</span><span class="sxs-lookup"><span data-stu-id="b41b2-109">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office host.</span></span>
- <span data-ttu-id="b41b2-110">添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-)：将对话框中的消息传送到其父页/开始页。</span><span class="sxs-lookup"><span data-stu-id="b41b2-110">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="b41b2-111">添加了 [Dialog](/javascript/api/office/office.dialog) 对象：调用 [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) 方法时返回的对象。</span><span class="sxs-lookup"><span data-stu-id="b41b2-111">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="b41b2-112">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b41b2-112">See also</span></span>

- [<span data-ttu-id="b41b2-113">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="b41b2-113">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="b41b2-114">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="b41b2-114">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="b41b2-115">入门</span><span class="sxs-lookup"><span data-stu-id="b41b2-115">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="b41b2-116">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="b41b2-116">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
