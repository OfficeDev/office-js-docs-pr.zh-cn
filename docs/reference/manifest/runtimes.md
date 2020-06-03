---
title: 清单文件中的运行时
description: 运行时元素指定外接程序的运行时。
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 95549d88df24a7d7c54cf27c92c15693491bdf29
ms.sourcegitcommit: 9229102c16a1864e3a8724aaf9b0dc68b1428094
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/03/2020
ms.locfileid: "44520337"
---
# <a name="runtimes-element"></a><span data-ttu-id="96ecd-103">运行时元素</span><span class="sxs-lookup"><span data-stu-id="96ecd-103">Runtimes element</span></span>

<span data-ttu-id="96ecd-104">指定外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="96ecd-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="96ecd-105">元素的子 [`<Host>`](host.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="96ecd-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="96ecd-106">在 Windows 上的 Office 中运行时，外接程序使用 Internet Explorer 11 浏览器。</span><span class="sxs-lookup"><span data-stu-id="96ecd-106">When running in Office on Windows, your add-in uses the Internet Explorer 11 browser.</span></span>

<span data-ttu-id="96ecd-107">在 Excel 中，此元素使功能区、任务窗格和自定义函数能够使用相同的运行时。</span><span class="sxs-lookup"><span data-stu-id="96ecd-107">In Excel, this element enables the ribbon, task pane, and custom functions to use the same runtime.</span></span> <span data-ttu-id="96ecd-108">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="96ecd-108">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="96ecd-109">在 Outlook 中，此元素启用基于事件的加载项激活。</span><span class="sxs-lookup"><span data-stu-id="96ecd-109">In Outlook, this element enables event-based add-in activation.</span></span> <span data-ttu-id="96ecd-110">有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="96ecd-110">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="96ecd-111">**外接类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="96ecd-111">**Add-in type:** Task pane, Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="96ecd-112">**Excel**：共享运行时目前仅适用于 Windows 中的 Excel。</span><span class="sxs-lookup"><span data-stu-id="96ecd-112">**Excel**: Shared runtime is currently only available in Excel on Windows.</span></span>
>
> <span data-ttu-id="96ecd-113">**Outlook**：基于事件的激活功能当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅适用于 web 上的 Outlook。</span><span class="sxs-lookup"><span data-stu-id="96ecd-113">**Outlook**: The event-based activation feature is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="96ecd-114">有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="96ecd-114">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="96ecd-115">语法</span><span class="sxs-lookup"><span data-stu-id="96ecd-115">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="96ecd-116">包含于</span><span class="sxs-lookup"><span data-stu-id="96ecd-116">Contained in</span></span>

[<span data-ttu-id="96ecd-117">Host</span><span class="sxs-lookup"><span data-stu-id="96ecd-117">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="96ecd-118">子元素</span><span class="sxs-lookup"><span data-stu-id="96ecd-118">Child elements</span></span>

|  <span data-ttu-id="96ecd-119">元素</span><span class="sxs-lookup"><span data-stu-id="96ecd-119">Element</span></span> |  <span data-ttu-id="96ecd-120">必需</span><span class="sxs-lookup"><span data-stu-id="96ecd-120">Required</span></span>  |  <span data-ttu-id="96ecd-121">Description</span><span class="sxs-lookup"><span data-stu-id="96ecd-121">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="96ecd-122">运行时</span><span class="sxs-lookup"><span data-stu-id="96ecd-122">Runtime</span></span>](runtime.md) | <span data-ttu-id="96ecd-123">是</span><span class="sxs-lookup"><span data-stu-id="96ecd-123">Yes</span></span> |  <span data-ttu-id="96ecd-124">外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="96ecd-124">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="96ecd-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="96ecd-125">See also</span></span>

- [<span data-ttu-id="96ecd-126">运行时</span><span class="sxs-lookup"><span data-stu-id="96ecd-126">Runtime</span></span>](runtime.md)
