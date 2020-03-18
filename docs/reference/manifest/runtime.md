---
title: 清单文件中的运行时（预览）
description: Runtime 元素将您的外接程序配置为对其功能区、任务窗格和自定义函数使用共享的 JavaScript 运行时。
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 6237f64fec47ed22b0105bf74c8eb7e2b7c38afe
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717927"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="63aa2-103">Runtime 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="63aa2-103">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="63aa2-104">[`<Runtimes>`](runtimes.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="63aa2-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="63aa2-105">此元素将您的外接程序配置为使用共享的 JavaScript 运行时，以便功能区、任务窗格和自定义函数在同一运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="63aa2-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="63aa2-106">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="63aa2-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="63aa2-107">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="63aa2-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="63aa2-108">共享运行时当前处于预览阶段，仅适用于 Windows 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="63aa2-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="63aa2-109">若要尝试预览功能，你需要加入[Office 预览体验成员](https://insider.office.com/)。</span><span class="sxs-lookup"><span data-stu-id="63aa2-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="63aa2-110">语法</span><span class="sxs-lookup"><span data-stu-id="63aa2-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="63aa2-111">包含于</span><span class="sxs-lookup"><span data-stu-id="63aa2-111">Contained in</span></span>

- [<span data-ttu-id="63aa2-112">运行时</span><span class="sxs-lookup"><span data-stu-id="63aa2-112">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="63aa2-113">属性</span><span class="sxs-lookup"><span data-stu-id="63aa2-113">Attributes</span></span>

|  <span data-ttu-id="63aa2-114">属性</span><span class="sxs-lookup"><span data-stu-id="63aa2-114">Attribute</span></span>  |  <span data-ttu-id="63aa2-115">必需</span><span class="sxs-lookup"><span data-stu-id="63aa2-115">Required</span></span>  |  <span data-ttu-id="63aa2-116">说明</span><span class="sxs-lookup"><span data-stu-id="63aa2-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="63aa2-117">**生存时间 = "long"**</span><span class="sxs-lookup"><span data-stu-id="63aa2-117">**lifetime="long"**</span></span>  |  <span data-ttu-id="63aa2-118">是</span><span class="sxs-lookup"><span data-stu-id="63aa2-118">Yes</span></span>  | <span data-ttu-id="63aa2-119">应始终是`long` ，如果您想要为 Excel 加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="63aa2-119">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="63aa2-120">**resid**</span><span class="sxs-lookup"><span data-stu-id="63aa2-120">**resid**</span></span>  |  <span data-ttu-id="63aa2-121">是</span><span class="sxs-lookup"><span data-stu-id="63aa2-121">Yes</span></span>  | <span data-ttu-id="63aa2-122">指定您的外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="63aa2-122">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="63aa2-123">`resid`必须与`Resources`元素中`id` `Url`元素的属性相匹配。</span><span class="sxs-lookup"><span data-stu-id="63aa2-123">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="63aa2-124">另请参阅</span><span class="sxs-lookup"><span data-stu-id="63aa2-124">See also</span></span>

- [<span data-ttu-id="63aa2-125">运行时</span><span class="sxs-lookup"><span data-stu-id="63aa2-125">Runtimes</span></span>](runtimes.md)
