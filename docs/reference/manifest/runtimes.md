---
title: 清单文件中的运行时
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111175"
---
# <a name="runtimes-element"></a><span data-ttu-id="c09c5-102">运行时元素</span><span class="sxs-lookup"><span data-stu-id="c09c5-102">Runtimes element</span></span>

<span data-ttu-id="c09c5-103">此功能处于预览阶段。</span><span class="sxs-lookup"><span data-stu-id="c09c5-103">This feature is in preview.</span></span> <span data-ttu-id="c09c5-104">指定外接程序的运行时，并允许自定义函数和任务窗格共享全局数据，并使函数相互调用。</span><span class="sxs-lookup"><span data-stu-id="c09c5-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="c09c5-105">应遵循清单`<Host>`文件中的元素。</span><span class="sxs-lookup"><span data-stu-id="c09c5-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="c09c5-106">**外接程序类型：** 任务窗格</span><span class="sxs-lookup"><span data-stu-id="c09c5-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="c09c5-107">语法</span><span class="sxs-lookup"><span data-stu-id="c09c5-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="c09c5-108">子元素</span><span class="sxs-lookup"><span data-stu-id="c09c5-108">Child elements</span></span>

|  <span data-ttu-id="c09c5-109">元素</span><span class="sxs-lookup"><span data-stu-id="c09c5-109">Element</span></span> |  <span data-ttu-id="c09c5-110">必需</span><span class="sxs-lookup"><span data-stu-id="c09c5-110">Required</span></span>  |  <span data-ttu-id="c09c5-111">说明</span><span class="sxs-lookup"><span data-stu-id="c09c5-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="c09c5-112">**运行时**</span><span class="sxs-lookup"><span data-stu-id="c09c5-112">**Runtime**</span></span>     | <span data-ttu-id="c09c5-113">是</span><span class="sxs-lookup"><span data-stu-id="c09c5-113">Yes</span></span> |  <span data-ttu-id="c09c5-114">外接程序的运行时通常与 Excel 自定义函数一起使用。</span><span class="sxs-lookup"><span data-stu-id="c09c5-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="c09c5-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c09c5-115">See also</span></span>

<span data-ttu-id="c09c5-116">-[时](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="c09c5-116">-[Runtimes](runtimes.md)</span></span>
