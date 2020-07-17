---
title: 清单文件中的运行时
description: Runtime 元素将您的外接程序配置为对其各个组件使用共享的 JavaScript 运行时，例如，功能区、任务窗格、自定义函数。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 9e6e13f83db363fb5485c8d8defbc381c80e32d6
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159365"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="34519-103">Runtime 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="34519-103">Runtime element (preview)</span></span>

<span data-ttu-id="34519-104">将您的外接程序配置为使用共享的 JavaScript 运行时，以便在同一运行时中运行各种组件。</span><span class="sxs-lookup"><span data-stu-id="34519-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="34519-105">元素的子 [`<Runtimes>`](runtimes.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="34519-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="34519-106">在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。</span><span class="sxs-lookup"><span data-stu-id="34519-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="34519-107">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="34519-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="34519-108">在 Outlook 中，此元素启用基于事件的加载项激活。</span><span class="sxs-lookup"><span data-stu-id="34519-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="34519-109">有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="34519-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="34519-110">**外接类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="34519-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="34519-111">**Outlook**：基于事件的激活当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅适用于 web 上的 outlook。</span><span class="sxs-lookup"><span data-stu-id="34519-111">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="34519-112">有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="34519-112">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="34519-113">语法</span><span class="sxs-lookup"><span data-stu-id="34519-113">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="34519-114">包含于</span><span class="sxs-lookup"><span data-stu-id="34519-114">Contained in</span></span>

- [<span data-ttu-id="34519-115">运行时</span><span class="sxs-lookup"><span data-stu-id="34519-115">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="34519-116">属性</span><span class="sxs-lookup"><span data-stu-id="34519-116">Attributes</span></span>

|  <span data-ttu-id="34519-117">属性</span><span class="sxs-lookup"><span data-stu-id="34519-117">Attribute</span></span>  |  <span data-ttu-id="34519-118">必需</span><span class="sxs-lookup"><span data-stu-id="34519-118">Required</span></span>  |  <span data-ttu-id="34519-119">说明</span><span class="sxs-lookup"><span data-stu-id="34519-119">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="34519-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="34519-120">**resid**</span></span>  |  <span data-ttu-id="34519-121">是</span><span class="sxs-lookup"><span data-stu-id="34519-121">Yes</span></span>  | <span data-ttu-id="34519-122">指定您的外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="34519-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="34519-123">`resid`必须与 `id` `Url` 元素中元素的属性相匹配 `Resources` 。</span><span class="sxs-lookup"><span data-stu-id="34519-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="34519-124">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="34519-124">**lifetime**</span></span>  |  <span data-ttu-id="34519-125">否</span><span class="sxs-lookup"><span data-stu-id="34519-125">No</span></span>  | <span data-ttu-id="34519-126">的默认值 `lifetime` 是 `short` ，不需要指定。</span><span class="sxs-lookup"><span data-stu-id="34519-126">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="34519-127">Outlook 外接程序仅使用 `short` 值。</span><span class="sxs-lookup"><span data-stu-id="34519-127">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="34519-128">如果要在 Excel 外接程序中使用共享运行时，请将值显式设置为 `long` 。</span><span class="sxs-lookup"><span data-stu-id="34519-128">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="34519-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="34519-129">See also</span></span>

- [<span data-ttu-id="34519-130">运行时</span><span class="sxs-lookup"><span data-stu-id="34519-130">Runtimes</span></span>](runtimes.md)
