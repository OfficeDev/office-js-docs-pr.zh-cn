---
title: 清单文件中的运行时
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 945a30527632b23a594d7bfb82cec94e74754249
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120633"
---
# <a name="runtime-element"></a><span data-ttu-id="ff4ba-102">Runtime 元素</span><span class="sxs-lookup"><span data-stu-id="ff4ba-102">Runtime element</span></span>

<span data-ttu-id="ff4ba-103">此功能处于预览阶段。</span><span class="sxs-lookup"><span data-stu-id="ff4ba-103">This feature is in preview.</span></span> <span data-ttu-id="ff4ba-104">[`<Runtimes>`](runtime.md)元素的子元素。</span><span class="sxs-lookup"><span data-stu-id="ff4ba-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="ff4ba-105">此元素有助于在 Excel 自定义函数和外接程序的任务窗格之间共享全局数据和函数调用。</span><span class="sxs-lookup"><span data-stu-id="ff4ba-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="ff4ba-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="ff4ba-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="ff4ba-107">语法</span><span class="sxs-lookup"><span data-stu-id="ff4ba-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="ff4ba-108">包含于</span><span class="sxs-lookup"><span data-stu-id="ff4ba-108">Contained in</span></span>

<span data-ttu-id="ff4ba-109">-[时](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="ff4ba-109">-[Runtimes](runtimes.md)</span></span>

## <a name="attributes"></a><span data-ttu-id="ff4ba-110">属性</span><span class="sxs-lookup"><span data-stu-id="ff4ba-110">Attributes</span></span>

|  <span data-ttu-id="ff4ba-111">属性</span><span class="sxs-lookup"><span data-stu-id="ff4ba-111">Attribute</span></span>  |  <span data-ttu-id="ff4ba-112">必需</span><span class="sxs-lookup"><span data-stu-id="ff4ba-112">Required</span></span>  |  <span data-ttu-id="ff4ba-113">Description</span><span class="sxs-lookup"><span data-stu-id="ff4ba-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="ff4ba-114">**生存时间 = "long"**</span><span class="sxs-lookup"><span data-stu-id="ff4ba-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="ff4ba-115">是</span><span class="sxs-lookup"><span data-stu-id="ff4ba-115">Yes</span></span>  | <span data-ttu-id="ff4ba-116">如果希望 Excel 自定义函数在外接程序的任务窗格关闭时正常工作，应始终将其列为长。</span><span class="sxs-lookup"><span data-stu-id="ff4ba-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="ff4ba-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="ff4ba-117">**resid**</span></span>  |  <span data-ttu-id="ff4ba-118">是</span><span class="sxs-lookup"><span data-stu-id="ff4ba-118">Yes</span></span>  | <span data-ttu-id="ff4ba-119">如果用于 Excel 自定义函数，则`resid`应指向`TaskPaneAndCustomFunction.Url`。</span><span class="sxs-lookup"><span data-stu-id="ff4ba-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ff4ba-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ff4ba-120">See also</span></span>

<span data-ttu-id="ff4ba-121">-[语言](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="ff4ba-121">-[Runtime](runtime.md)</span></span>
