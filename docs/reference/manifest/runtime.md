---
title: 清单文件中的运行时
description: Runtime 元素将您的外接程序配置为对其功能区、任务窗格和自定义函数使用共享的 JavaScript 运行时。
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217758"
---
# <a name="runtime-element"></a><span data-ttu-id="6b7eb-103">Runtime 元素</span><span class="sxs-lookup"><span data-stu-id="6b7eb-103">Runtime element</span></span>

<span data-ttu-id="6b7eb-104">元素的子元素 [`<Runtimes>`](runtimes.md) 。</span><span class="sxs-lookup"><span data-stu-id="6b7eb-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="6b7eb-105">此元素将您的外接程序配置为使用共享的 JavaScript 运行时，以便功能区、任务窗格和自定义函数在同一运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="6b7eb-105">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="6b7eb-106">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="6b7eb-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="6b7eb-107">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="6b7eb-107">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="6b7eb-108">语法</span><span class="sxs-lookup"><span data-stu-id="6b7eb-108">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="6b7eb-109">包含于</span><span class="sxs-lookup"><span data-stu-id="6b7eb-109">Contained in</span></span>

- [<span data-ttu-id="6b7eb-110">运行时</span><span class="sxs-lookup"><span data-stu-id="6b7eb-110">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="6b7eb-111">属性</span><span class="sxs-lookup"><span data-stu-id="6b7eb-111">Attributes</span></span>

|  <span data-ttu-id="6b7eb-112">属性</span><span class="sxs-lookup"><span data-stu-id="6b7eb-112">Attribute</span></span>  |  <span data-ttu-id="6b7eb-113">必需</span><span class="sxs-lookup"><span data-stu-id="6b7eb-113">Required</span></span>  |  <span data-ttu-id="6b7eb-114">说明</span><span class="sxs-lookup"><span data-stu-id="6b7eb-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6b7eb-115">**生存时间 = "long"**</span><span class="sxs-lookup"><span data-stu-id="6b7eb-115">**lifetime="long"**</span></span>  |  <span data-ttu-id="6b7eb-116">是</span><span class="sxs-lookup"><span data-stu-id="6b7eb-116">Yes</span></span>  | <span data-ttu-id="6b7eb-117">应始终是 `long` ，如果您想要为 Excel 加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="6b7eb-117">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="6b7eb-118">**resid**</span><span class="sxs-lookup"><span data-stu-id="6b7eb-118">**resid**</span></span>  |  <span data-ttu-id="6b7eb-119">是</span><span class="sxs-lookup"><span data-stu-id="6b7eb-119">Yes</span></span>  | <span data-ttu-id="6b7eb-120">指定您的外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="6b7eb-120">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="6b7eb-121">`resid`必须与 `id` `Url` 元素中元素的属性相匹配 `Resources` 。</span><span class="sxs-lookup"><span data-stu-id="6b7eb-121">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="6b7eb-122">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6b7eb-122">See also</span></span>

- [<span data-ttu-id="6b7eb-123">运行时</span><span class="sxs-lookup"><span data-stu-id="6b7eb-123">Runtimes</span></span>](runtimes.md)
