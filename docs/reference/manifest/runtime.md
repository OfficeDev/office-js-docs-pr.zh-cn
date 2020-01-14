---
title: 清单文件中的运行时
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111168"
---
# <a name="runtime-element"></a><span data-ttu-id="6151f-102">Runtime 元素</span><span class="sxs-lookup"><span data-stu-id="6151f-102">Runtime element</span></span>

<span data-ttu-id="6151f-103">此功能处于预览阶段。</span><span class="sxs-lookup"><span data-stu-id="6151f-103">This feature is in preview.</span></span> <span data-ttu-id="6151f-104">[`<Runtimes>`](runtime.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="6151f-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="6151f-105">此元素有助于在 Excel 自定义函数和外接程序的任务窗格之间共享全局数据和函数调用。</span><span class="sxs-lookup"><span data-stu-id="6151f-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span> 

## <a name="contained-in"></a><span data-ttu-id="6151f-106">包含于</span><span class="sxs-lookup"><span data-stu-id="6151f-106">Contained in</span></span>

<span data-ttu-id="6151f-107">-[时](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="6151f-107">-[Runtimes](runtimes.md)</span></span>

<span data-ttu-id="6151f-108">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="6151f-108">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="6151f-109">语法</span><span class="sxs-lookup"><span data-stu-id="6151f-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a><span data-ttu-id="6151f-110">属性</span><span class="sxs-lookup"><span data-stu-id="6151f-110">Attributes</span></span>

|  <span data-ttu-id="6151f-111">属性</span><span class="sxs-lookup"><span data-stu-id="6151f-111">Attribute</span></span>  |  <span data-ttu-id="6151f-112">必需</span><span class="sxs-lookup"><span data-stu-id="6151f-112">Required</span></span>  |  <span data-ttu-id="6151f-113">说明</span><span class="sxs-lookup"><span data-stu-id="6151f-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="6151f-114">**生存时间 = "long"**</span><span class="sxs-lookup"><span data-stu-id="6151f-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="6151f-115">是</span><span class="sxs-lookup"><span data-stu-id="6151f-115">Yes</span></span>  | <span data-ttu-id="6151f-116">如果希望 Excel 自定义函数在外接程序的任务窗格关闭时正常工作，应始终将其列为长。</span><span class="sxs-lookup"><span data-stu-id="6151f-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="6151f-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="6151f-117">**resid**</span></span>  |  <span data-ttu-id="6151f-118">是</span><span class="sxs-lookup"><span data-stu-id="6151f-118">Yes</span></span>  | <span data-ttu-id="6151f-119">如果用于 Excel 自定义函数，则`resid`应指向`TaskPaneAndCustomFunction.Url`。</span><span class="sxs-lookup"><span data-stu-id="6151f-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="6151f-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6151f-120">See also</span></span>

<span data-ttu-id="6151f-121">-[语言](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="6151f-121">-[Runtime](runtime.md)</span></span>
