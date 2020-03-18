---
title: 清单文件中的运行时（预览）
description: 运行时元素指定外接程序的运行时。
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 5797aa78ae3667461de48de481ff44f14c307ced
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720419"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="aa1c4-103">运行时元素（预览）</span><span class="sxs-lookup"><span data-stu-id="aa1c4-103">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="aa1c4-104">指定外接程序的运行时，并启用自定义函数、功能区按钮和任务窗格，以使用相同的 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="aa1c4-104">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="aa1c4-105">清单文件中`<Host>`的元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="aa1c4-105">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="aa1c4-106">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="aa1c4-106">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="aa1c4-107">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="aa1c4-107">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="aa1c4-108">共享运行时当前处于预览阶段，仅适用于 Windows 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="aa1c4-108">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="aa1c4-109">若要尝试预览功能，你需要加入[Office 预览体验成员](https://insider.office.com/)。</span><span class="sxs-lookup"><span data-stu-id="aa1c4-109">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="aa1c4-110">语法</span><span class="sxs-lookup"><span data-stu-id="aa1c4-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="aa1c4-111">包含于</span><span class="sxs-lookup"><span data-stu-id="aa1c4-111">Contained in</span></span> 
[<span data-ttu-id="aa1c4-112">Host</span><span class="sxs-lookup"><span data-stu-id="aa1c4-112">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="aa1c4-113">子元素</span><span class="sxs-lookup"><span data-stu-id="aa1c4-113">Child elements</span></span>

|  <span data-ttu-id="aa1c4-114">元素</span><span class="sxs-lookup"><span data-stu-id="aa1c4-114">Element</span></span> |  <span data-ttu-id="aa1c4-115">必需</span><span class="sxs-lookup"><span data-stu-id="aa1c4-115">Required</span></span>  |  <span data-ttu-id="aa1c4-116">说明</span><span class="sxs-lookup"><span data-stu-id="aa1c4-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="aa1c4-117">**运行时**</span><span class="sxs-lookup"><span data-stu-id="aa1c4-117">**Runtime**</span></span>     | <span data-ttu-id="aa1c4-118">是</span><span class="sxs-lookup"><span data-stu-id="aa1c4-118">Yes</span></span> |  <span data-ttu-id="aa1c4-119">外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="aa1c4-119">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="aa1c4-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="aa1c4-120">See also</span></span>

- [<span data-ttu-id="aa1c4-121">运行时</span><span class="sxs-lookup"><span data-stu-id="aa1c4-121">Runtime</span></span>](runtime.md)
