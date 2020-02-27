---
title: 清单文件中的运行时（预览）
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 26702896604f9ecf4c69296e5110efe5cdf4218b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283882"
---
# <a name="runtime-element-preview"></a><span data-ttu-id="64ce7-102">Runtime 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="64ce7-102">Runtime element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="64ce7-103">[`<Runtimes>`](runtimes.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="64ce7-103">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="64ce7-104">此元素将您的外接程序配置为使用共享的 JavaScript 运行时，以便功能区、任务窗格和自定义函数在同一运行时中运行。</span><span class="sxs-lookup"><span data-stu-id="64ce7-104">This element configures your add-in to use a shared JavaScript runtime so that your ribbon, task pane, and custom functions, all run in the same runtime.</span></span> <span data-ttu-id="64ce7-105">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="64ce7-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="64ce7-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="64ce7-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
<span data-ttu-id="64ce7-107"><<<<<<< 头共享运行时当前处于预览阶段，仅在 Windows 上的 Excel 中可用。</span><span class="sxs-lookup"><span data-stu-id="64ce7-107"><<<<<<< HEAD Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="64ce7-108">若要尝试预览功能，你需要加入[Office 预览体验成员](https://insider.office.com/)。</span><span class="sxs-lookup"><span data-stu-id="64ce7-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="64ce7-109">语法</span><span class="sxs-lookup"><span data-stu-id="64ce7-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="64ce7-110">包含于</span><span class="sxs-lookup"><span data-stu-id="64ce7-110">Contained in</span></span>

- [<span data-ttu-id="64ce7-111">运行时</span><span class="sxs-lookup"><span data-stu-id="64ce7-111">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="64ce7-112">属性</span><span class="sxs-lookup"><span data-stu-id="64ce7-112">Attributes</span></span>

|  <span data-ttu-id="64ce7-113">属性</span><span class="sxs-lookup"><span data-stu-id="64ce7-113">Attribute</span></span>  |  <span data-ttu-id="64ce7-114">必需</span><span class="sxs-lookup"><span data-stu-id="64ce7-114">Required</span></span>  |  <span data-ttu-id="64ce7-115">说明</span><span class="sxs-lookup"><span data-stu-id="64ce7-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="64ce7-116">**生存时间 = "long"**</span><span class="sxs-lookup"><span data-stu-id="64ce7-116">**lifetime="long"**</span></span>  |  <span data-ttu-id="64ce7-117">是</span><span class="sxs-lookup"><span data-stu-id="64ce7-117">Yes</span></span>  | <span data-ttu-id="64ce7-118">应始终是`long` ，如果您想要为 Excel 加载项使用共享运行时。</span><span class="sxs-lookup"><span data-stu-id="64ce7-118">Should always be `long` if you want to use a shared runtime for the Excel add-in.</span></span> |
|  <span data-ttu-id="64ce7-119">**resid**</span><span class="sxs-lookup"><span data-stu-id="64ce7-119">**resid**</span></span>  |  <span data-ttu-id="64ce7-120">是</span><span class="sxs-lookup"><span data-stu-id="64ce7-120">Yes</span></span>  | <span data-ttu-id="64ce7-121">指定您的外接程序的 HTML 页面的 URL 位置。</span><span class="sxs-lookup"><span data-stu-id="64ce7-121">Specifies the URL location of the HTML page for your add-in.</span></span> <span data-ttu-id="64ce7-122">`resid`必须与`Resources`元素中`id` `Url`元素的属性相匹配。</span><span class="sxs-lookup"><span data-stu-id="64ce7-122">The `resid` must match an `id` attribute of a `Url` element in the `Resources` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="64ce7-123">另请参阅</span><span class="sxs-lookup"><span data-stu-id="64ce7-123">See also</span></span>

- [<span data-ttu-id="64ce7-124">运行时</span><span class="sxs-lookup"><span data-stu-id="64ce7-124">Runtimes</span></span>](runtimes.md)
