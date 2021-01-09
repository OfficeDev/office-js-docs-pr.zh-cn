---
title: 清单文件中运行时
description: Runtimes 元素指定加载项的运行时。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: afbcc6a909c51d2ed56292ef1541193f7f698d28
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789161"
---
# <a name="runtimes-element"></a><span data-ttu-id="b1df5-103">Runtimes 元素</span><span class="sxs-lookup"><span data-stu-id="b1df5-103">Runtimes element</span></span>

<span data-ttu-id="b1df5-104">指定加载项的运行时。</span><span class="sxs-lookup"><span data-stu-id="b1df5-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="b1df5-105">元素的 [`<Host>`](host.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="b1df5-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="b1df5-106">When running in Office on Windows， your add-in uses the Internet Explorer 11 browser.</span><span class="sxs-lookup"><span data-stu-id="b1df5-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="b1df5-107">在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。</span><span class="sxs-lookup"><span data-stu-id="b1df5-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="b1df5-108">有关详细信息，请参阅配置 [Excel 加载项以使用共享的 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="b1df5-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="b1df5-109">在 Outlook 中，此元素启用基于事件的外接程序激活。</span><span class="sxs-lookup"><span data-stu-id="b1df5-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="b1df5-110">有关详细信息，请参阅配置 [Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="b1df5-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="b1df5-111">**加载项类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="b1df5-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b1df5-112">**Outlook：** 基于事件的激活 [功能目前处于](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 预览阶段，仅在 Outlook 网页版中可用。</span><span class="sxs-lookup"><span data-stu-id="b1df5-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="b1df5-113">有关详细信息，请参阅 [如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="b1df5-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="b1df5-114">语法</span><span class="sxs-lookup"><span data-stu-id="b1df5-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="b1df5-115">包含于</span><span class="sxs-lookup"><span data-stu-id="b1df5-115">Contained in</span></span>

[<span data-ttu-id="b1df5-116">Host</span><span class="sxs-lookup"><span data-stu-id="b1df5-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="b1df5-117">子元素</span><span class="sxs-lookup"><span data-stu-id="b1df5-117">Child elements</span></span>

|  <span data-ttu-id="b1df5-118">元素</span><span class="sxs-lookup"><span data-stu-id="b1df5-118">Element</span></span> |  <span data-ttu-id="b1df5-119">必需</span><span class="sxs-lookup"><span data-stu-id="b1df5-119">Required</span></span>  |  <span data-ttu-id="b1df5-120">说明</span><span class="sxs-lookup"><span data-stu-id="b1df5-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="b1df5-121">运行时</span><span class="sxs-lookup"><span data-stu-id="b1df5-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="b1df5-122">是</span><span class="sxs-lookup"><span data-stu-id="b1df5-122">Yes</span></span> |  <span data-ttu-id="b1df5-123">加载项的运行时。</span><span class="sxs-lookup"><span data-stu-id="b1df5-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b1df5-124">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b1df5-124">See also</span></span>

- [<span data-ttu-id="b1df5-125">运行时</span><span class="sxs-lookup"><span data-stu-id="b1df5-125">Runtime</span></span>](runtime.md)
