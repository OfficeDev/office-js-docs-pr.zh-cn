---
title: 清单文件中的运行时（预览）
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283870"
---
# <a name="runtimes-element-preview"></a><span data-ttu-id="fddc2-102">运行时元素（预览）</span><span class="sxs-lookup"><span data-stu-id="fddc2-102">Runtimes element (preview)</span></span>

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

<span data-ttu-id="fddc2-103">指定外接程序的运行时，并启用自定义函数、功能区按钮和任务窗格，以使用相同的 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="fddc2-103">Specifies the runtime of your add-in and enables custom functions, ribbon buttons, and the task pane to use the same JavaScript runtime.</span></span> <span data-ttu-id="fddc2-104">清单文件中`<Host>`的元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="fddc2-104">Child of the `<Host>` element in your manifest file.</span></span> <span data-ttu-id="fddc2-105">有关详细信息，请参阅[Configure Excel 外接程序以使用共享的 JavaScript 运行时](../../excel/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="fddc2-105">For more information, see [Configure your Excel add-in to use a shared JavaScript runtime](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="fddc2-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="fddc2-106">**Add-in type:** Task pane</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fddc2-107">共享运行时当前处于预览阶段，仅适用于 Windows 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="fddc2-107">Shared runtime is currently in preview and are only available on Excel on Windows.</span></span> <span data-ttu-id="fddc2-108">若要尝试预览功能，你需要加入[Office 预览体验成员](https://insider.office.com/)。</span><span class="sxs-lookup"><span data-stu-id="fddc2-108">To try the preview features, you will need to join [Office Insider](https://insider.office.com/).</span></span>

## <a name="syntax"></a><span data-ttu-id="fddc2-109">语法</span><span class="sxs-lookup"><span data-stu-id="fddc2-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="fddc2-110">包含于</span><span class="sxs-lookup"><span data-stu-id="fddc2-110">Contained in</span></span> 
[<span data-ttu-id="fddc2-111">Host</span><span class="sxs-lookup"><span data-stu-id="fddc2-111">Host</span></span>](./host.md)

## <a name="child-elements"></a><span data-ttu-id="fddc2-112">子元素</span><span class="sxs-lookup"><span data-stu-id="fddc2-112">Child elements</span></span>

|  <span data-ttu-id="fddc2-113">元素</span><span class="sxs-lookup"><span data-stu-id="fddc2-113">Element</span></span> |  <span data-ttu-id="fddc2-114">必需</span><span class="sxs-lookup"><span data-stu-id="fddc2-114">Required</span></span>  |  <span data-ttu-id="fddc2-115">说明</span><span class="sxs-lookup"><span data-stu-id="fddc2-115">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="fddc2-116">**运行时**</span><span class="sxs-lookup"><span data-stu-id="fddc2-116">**Runtime**</span></span>     | <span data-ttu-id="fddc2-117">是</span><span class="sxs-lookup"><span data-stu-id="fddc2-117">Yes</span></span> |  <span data-ttu-id="fddc2-118">外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="fddc2-118">The runtime for your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="fddc2-119">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fddc2-119">See also</span></span>

- [<span data-ttu-id="fddc2-120">运行时</span><span class="sxs-lookup"><span data-stu-id="fddc2-120">Runtime</span></span>](runtime.md)
