---
title: 清单文件中运行时
description: 运行时元素将加载项配置为将共享的 JavaScript 运行时用于其各种组件，例如功能区、任务窗格、自定义函数。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789182"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="2a830-103">运行时元素 (预览) </span><span class="sxs-lookup"><span data-stu-id="2a830-103">Runtime element (preview)</span></span>

<span data-ttu-id="2a830-104">将加载项配置为使用共享的 JavaScript 运行时，以便各种组件都在同一运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="2a830-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="2a830-105">元素的 [`<Runtimes>`](runtimes.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="2a830-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="2a830-106">在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。</span><span class="sxs-lookup"><span data-stu-id="2a830-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="2a830-107">有关详细信息，请参阅配置 [Excel 加载项以使用共享的 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="2a830-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="2a830-108">在 Outlook 中，此元素启用基于事件的外接程序激活。</span><span class="sxs-lookup"><span data-stu-id="2a830-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="2a830-109">有关详细信息，请参阅配置 [Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="2a830-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="2a830-110">**加载项类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="2a830-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2a830-111">**Outlook：** 基于事件的激活目前处于 [预览阶段，](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 仅在 Outlook 网页版中可用。</span><span class="sxs-lookup"><span data-stu-id="2a830-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="2a830-112">有关详细信息，请参阅 [如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="2a830-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="2a830-113">语法</span><span class="sxs-lookup"><span data-stu-id="2a830-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="2a830-114">包含于</span><span class="sxs-lookup"><span data-stu-id="2a830-114">Contained in</span></span>

- [<span data-ttu-id="2a830-115">运行时</span><span class="sxs-lookup"><span data-stu-id="2a830-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="2a830-116">属性</span><span class="sxs-lookup"><span data-stu-id="2a830-116">Attributes</span></span>

|  <span data-ttu-id="2a830-117">属性</span><span class="sxs-lookup"><span data-stu-id="2a830-117">Attribute</span></span>  |  <span data-ttu-id="2a830-118">必需</span><span class="sxs-lookup"><span data-stu-id="2a830-118">Required</span></span>  |  <span data-ttu-id="2a830-119">说明</span><span class="sxs-lookup"><span data-stu-id="2a830-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2a830-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="2a830-120">**resid**</span></span>  |  <span data-ttu-id="2a830-121">是</span><span class="sxs-lookup"><span data-stu-id="2a830-121">Yes</span></span>  | <span data-ttu-id="2a830-122">指定外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="2a830-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="2a830-123">`resid`不能超过 32 个字符，并且必须与元素 `id` `Url` 中的元素属性 `Resources` 匹配。</span><span class="sxs-lookup"><span data-stu-id="2a830-123">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="2a830-124">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="2a830-124">**lifetime**</span></span>  |  <span data-ttu-id="2a830-125">否</span><span class="sxs-lookup"><span data-stu-id="2a830-125">No</span></span>  | <span data-ttu-id="2a830-126">默认值是 `lifetime` `short` ，不需要指定。</span><span class="sxs-lookup"><span data-stu-id="2a830-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="2a830-127">Outlook 外接程序仅使用 `short` 该值。</span><span class="sxs-lookup"><span data-stu-id="2a830-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="2a830-128">如果要在 Excel 加载项中使用共享运行时，请显式将值设置为 `long` 。</span><span class="sxs-lookup"><span data-stu-id="2a830-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2a830-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2a830-129">See also</span></span>

- [<span data-ttu-id="2a830-130">运行时</span><span class="sxs-lookup"><span data-stu-id="2a830-130">Runtimes</span></span>](runtimes.md)
