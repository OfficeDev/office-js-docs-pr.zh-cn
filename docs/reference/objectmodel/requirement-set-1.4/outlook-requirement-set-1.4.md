---
title: Outlook 加载项 API 要求集 1.4
description: 作为邮箱 API 1.4 Outlook外接程序和 Office JavaScript API 引入的功能和 API。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 19d77784926ac09d5620eb36242701da59b39f09
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591014"
---
# <a name="outlook-add-in-api-requirement-set-14"></a><span data-ttu-id="56b65-103">Outlook 外接程序 API 要求集 1.4</span><span class="sxs-lookup"><span data-stu-id="56b65-103">Outlook add-in API requirement set 1.4</span></span>

<span data-ttu-id="56b65-104">Outlook JavaScript API 的 Office 外接程序 API 子集包括可在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="56b65-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="56b65-105">本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="56b65-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-14"></a><span data-ttu-id="56b65-106">1.4 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="56b65-106">What's new in 1.4?</span></span>

<span data-ttu-id="56b65-107">要求集 1.4 包括要求集 [1.3 的所有功能](../requirement-set-1.3/outlook-requirement-set-1.3.md)。</span><span class="sxs-lookup"><span data-stu-id="56b65-107">Requirement set 1.4 includes all of the features of [requirement set 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md).</span></span> <span data-ttu-id="56b65-108">它添加了对 `Office.ui` 命名空间的访问权限。</span><span class="sxs-lookup"><span data-stu-id="56b65-108">It added access to the `Office.ui` namespace.</span></span>

### <a name="change-log"></a><span data-ttu-id="56b65-109">更改日志</span><span class="sxs-lookup"><span data-stu-id="56b65-109">Change log</span></span>

- <span data-ttu-id="56b65-110">添加了[Office.context.ui.displayDialogAsync：](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-)在应用程序Office对话框。</span><span class="sxs-lookup"><span data-stu-id="56b65-110">Added [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Displays a dialog box in an Office application.</span></span>
- <span data-ttu-id="56b65-111">添加了 [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-)：将对话框中的消息传送到其父页/开始页。</span><span class="sxs-lookup"><span data-stu-id="56b65-111">Added [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): Delivers a message from the dialog box to its parent/opener page.</span></span>
- <span data-ttu-id="56b65-112">添加了 [Dialog](/javascript/api/office/office.dialog) 对象：调用 [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) 方法时返回的对象。</span><span class="sxs-lookup"><span data-stu-id="56b65-112">Added [Dialog](/javascript/api/office/office.dialog) object: The object that is returned when the [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method is called.</span></span>

## <a name="see-also"></a><span data-ttu-id="56b65-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="56b65-113">See also</span></span>

- [<span data-ttu-id="56b65-114">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="56b65-114">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="56b65-115">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="56b65-115">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="56b65-116">入门</span><span class="sxs-lookup"><span data-stu-id="56b65-116">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="56b65-117">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="56b65-117">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
