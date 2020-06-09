---
title: 清单文件中的运行时
description: Runtime 元素将您的外接程序配置为对其各个组件使用共享的 JavaScript 运行时，例如，功能区、任务窗格、自定义函数。
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: e81bd7222585bfa7d5f0f34fe5d9b32e4d45a71e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608102"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="b9bf5-103">Runtime 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="b9bf5-103">Runtime element (preview)</span></span>

<span data-ttu-id="b9bf5-104">将您的外接程序配置为使用共享的 JavaScript 运行时，以便在同一运行时中运行各种组件。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-104">Configures your add-in to use a shared JavaScript runtime so that various components all run in the same runtime.</span></span> <span data-ttu-id="b9bf5-105">元素的子 [`<Runtimes>`](runtimes.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-105">Child of the [`<Runtimes>`](runtimes.md) element.</span></span>

<span data-ttu-id="b9bf5-106">在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-106">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="b9bf5-107">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-107">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="b9bf5-108">在 Outlook 中，此元素启用基于事件的加载项激活。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-108">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="b9bf5-109">有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-109">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="b9bf5-110">**外接类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="b9bf5-110">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9bf5-111">**Excel**：共享运行时目前仅适用于 Windows 中的 Excel。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-111">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="b9bf5-112">**Outlook**：基于事件的激活当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅适用于 web 上的 outlook。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-112">**Outlook**: Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="b9bf5-113">有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="b9bf5-114">语法</span><span class="sxs-lookup"><span data-stu-id="b9bf5-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="b9bf5-115">包含于</span><span class="sxs-lookup"><span data-stu-id="b9bf5-115">Contained in</span></span>

- [<span data-ttu-id="b9bf5-116">运行时</span><span class="sxs-lookup"><span data-stu-id="b9bf5-116">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="b9bf5-117">属性</span><span class="sxs-lookup"><span data-stu-id="b9bf5-117">Attributes</span></span>

|  <span data-ttu-id="b9bf5-118">属性</span><span class="sxs-lookup"><span data-stu-id="b9bf5-118">Attribute</span></span>  |  <span data-ttu-id="b9bf5-119">必需</span><span class="sxs-lookup"><span data-stu-id="b9bf5-119">Required</span></span>  |  <span data-ttu-id="b9bf5-120">Description</span><span class="sxs-lookup"><span data-stu-id="b9bf5-120">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="b9bf5-121">**resid**</span><span class="sxs-lookup"><span data-stu-id="b9bf5-121">**resid**</span></span>  |  <span data-ttu-id="b9bf5-122">是</span><span class="sxs-lookup"><span data-stu-id="b9bf5-122">Yes</span></span>  | <span data-ttu-id="b9bf5-123">指定您的外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-123">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="b9bf5-124">`resid`必须与 `id` `Url` 元素中元素的属性相匹配 `Resources` 。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-124">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |
|  <span data-ttu-id="b9bf5-125">**lifetime**</span><span class="sxs-lookup"><span data-stu-id="b9bf5-125">**lifetime**</span></span>  |  <span data-ttu-id="b9bf5-126">否</span><span class="sxs-lookup"><span data-stu-id="b9bf5-126">No</span></span>  | <span data-ttu-id="b9bf5-127">的默认值 `lifetime` 是 `short` ，不需要指定。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-127">The default value for `lifetime` is `short` and doesn't need to be specified.</span></span> <span data-ttu-id="b9bf5-128">Outlook 外接程序仅使用 `short` 值。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-128">Outlook add-ins use only the `short` value.</span></span> <span data-ttu-id="b9bf5-129">如果要在 Excel 外接程序中使用共享运行时，请将值显式设置为 `long` 。</span><span class="sxs-lookup"><span data-stu-id="b9bf5-129">If you want to use a shared runtime in an Excel add-in, explicitly set the value to `long`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b9bf5-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b9bf5-130">See also</span></span>

- [<span data-ttu-id="b9bf5-131">运行时</span><span class="sxs-lookup"><span data-stu-id="b9bf5-131">Runtimes</span></span>](runtimes.md)
