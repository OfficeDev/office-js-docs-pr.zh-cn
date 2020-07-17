---
title: 清单文件中的运行时
description: 运行时元素指定外接程序的运行时。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 082491befc6b9dbdc474b0e40f9defd90a4ef75f
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159358"
---
# <a name="runtimes-element"></a><span data-ttu-id="af3cf-103">运行时元素</span><span class="sxs-lookup"><span data-stu-id="af3cf-103">Runtimes element</span></span>

<span data-ttu-id="af3cf-104">指定外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="af3cf-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="af3cf-105">元素的子 [`<Host>`](host.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="af3cf-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="af3cf-106">在 Windows 上的 Office 中运行时，外接程序使用 Internet Explorer 11 浏览器。</span><span class="sxs-lookup"><span data-stu-id="af3cf-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="af3cf-107">在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。</span><span class="sxs-lookup"><span data-stu-id="af3cf-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="af3cf-108">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="af3cf-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="af3cf-109">在 Outlook 中，此元素启用基于事件的加载项激活。</span><span class="sxs-lookup"><span data-stu-id="af3cf-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="af3cf-110">有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="af3cf-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="af3cf-111">**外接类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="af3cf-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="af3cf-112">**Outlook**：基于事件的激活功能当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅适用于 web 上的 Outlook。</span><span class="sxs-lookup"><span data-stu-id="af3cf-112">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="af3cf-113">有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="af3cf-113">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="af3cf-114">语法</span><span class="sxs-lookup"><span data-stu-id="af3cf-114">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="af3cf-115">包含于</span><span class="sxs-lookup"><span data-stu-id="af3cf-115">Contained in</span></span>

[<span data-ttu-id="af3cf-116">Host</span><span class="sxs-lookup"><span data-stu-id="af3cf-116">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="af3cf-117">子元素</span><span class="sxs-lookup"><span data-stu-id="af3cf-117">Child elements</span></span>

|  <span data-ttu-id="af3cf-118">元素</span><span class="sxs-lookup"><span data-stu-id="af3cf-118">Element</span></span> |  <span data-ttu-id="af3cf-119">必需</span><span class="sxs-lookup"><span data-stu-id="af3cf-119">Required</span></span>  |  <span data-ttu-id="af3cf-120">说明</span><span class="sxs-lookup"><span data-stu-id="af3cf-120">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="af3cf-121">运行时</span><span class="sxs-lookup"><span data-stu-id="af3cf-121">Runtime</span></span>](runtime.md) | <span data-ttu-id="af3cf-122">是</span><span class="sxs-lookup"><span data-stu-id="af3cf-122">Yes</span></span> |  <span data-ttu-id="af3cf-123">外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="af3cf-123">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="af3cf-124">另请参阅</span><span class="sxs-lookup"><span data-stu-id="af3cf-124">See also</span></span>

- [<span data-ttu-id="af3cf-125">运行时</span><span class="sxs-lookup"><span data-stu-id="af3cf-125">Runtime</span></span>](runtime.md)
