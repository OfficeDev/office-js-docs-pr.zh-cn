---
title: 清单文件中的运行时（预览）
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: dd51c5b317700f92ee74c94835e68523371789f8
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561826"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="316de-102">Runtime 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="316de-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="316de-103">[`<Runtimes>`](runtimes.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="316de-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="316de-104">此元素将您的外接程序配置为使用共享的 JavaScript 运行时，以便功能区、任务窗格和自定义函数在同一运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="316de-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="316de-105">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="316de-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="316de-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="316de-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="316de-107">共享运行时当前处于预览阶段，仅适用于 Windows 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="316de-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="316de-108">若要尝试预览功能，你需要加入[Office 预览体验成员](https://insider.office.com/)。</span><span class="sxs-lookup"><span data-stu-id="316de-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="316de-109">语法</span><span class="sxs-lookup"><span data-stu-id="316de-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="316de-110">包含于</span><span class="sxs-lookup"><span data-stu-id="316de-110">Contained in</span></span>

- [<span data-ttu-id="316de-111">运行时</span><span class="sxs-lookup"><span data-stu-id="316de-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="316de-112">属性</span><span class="sxs-lookup"><span data-stu-id="316de-112">Attributes</span></span>

|  <span data-ttu-id="316de-113">属性</span><span class="sxs-lookup"><span data-stu-id="316de-113">Attribute</span></span>  |  <span data-ttu-id="316de-114">必需</span><span class="sxs-lookup"><span data-stu-id="316de-114">Required</span></span>  |  <span data-ttu-id="316de-115">说明</span><span class="sxs-lookup"><span data-stu-id="316de-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="316de-116">**生存时间 = "long"**</span><span class="sxs-lookup"><span data-stu-id="316de-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="316de-117">是</span><span class="sxs-lookup"><span data-stu-id="316de-117">Yes</span></span>  | <span data-ttu-id="316de-118">应始终是`long` ，如果您想要为 Excel 加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="316de-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="316de-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="316de-119">**resid**</span></span>  |  <span data-ttu-id="316de-120">是</span><span class="sxs-lookup"><span data-stu-id="316de-120">Yes</span></span>  | <span data-ttu-id="316de-121">指定您的外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="316de-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="316de-122">`resid`必须与`Resources`元素中`id` `Url`元素的属性相匹配。</span><span class="sxs-lookup"><span data-stu-id="316de-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="316de-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="316de-123">See also</span></span>

- [<span data-ttu-id="316de-124">运行时</span><span class="sxs-lookup"><span data-stu-id="316de-124">Runtimes</span></span>](runtimes.md)
