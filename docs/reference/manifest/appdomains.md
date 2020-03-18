---
title: 清单文件中的 AppDomains 元素
description: 列出除 Office 加载项将用于加载页面的`SourceLocation`元素中指定的域之外的所有域。
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: f60579d773e81a7e8006bafcf1c151874af42aeb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720699"
---
# <a name="appdomains-element"></a><span data-ttu-id="d3ab9-103">AppDomains 元素</span><span class="sxs-lookup"><span data-stu-id="d3ab9-103">AppDomains element</span></span>

<span data-ttu-id="d3ab9-104">列出除 Office 加载项将用于加载页面的`SourceLocation`元素中指定的域之外的所有域。</span><span class="sxs-lookup"><span data-stu-id="d3ab9-104">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="d3ab9-105">此外，它还列出了可以从加载项内的 Iframe 中进行的 Office .js API 调用的受信任域。</span><span class="sxs-lookup"><span data-stu-id="d3ab9-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="d3ab9-106">对于每个其他域，指定 AppDomain 元素。</span><span class="sxs-lookup"><span data-stu-id="d3ab9-106">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="d3ab9-107">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="d3ab9-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d3ab9-108">语法</span><span class="sxs-lookup"><span data-stu-id="d3ab9-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="d3ab9-109">每个 **AppDomain** 元素的值都必须包括协议（如 `<AppDomain>https://myappdomain<AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="d3ab9-109">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="d3ab9-110">包含于</span><span class="sxs-lookup"><span data-stu-id="d3ab9-110">Contained in</span></span>

[<span data-ttu-id="d3ab9-111">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d3ab9-111">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d3ab9-112">可以包含</span><span class="sxs-lookup"><span data-stu-id="d3ab9-112">Can contain</span></span>

[<span data-ttu-id="d3ab9-113">AppDomain</span><span class="sxs-lookup"><span data-stu-id="d3ab9-113">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="d3ab9-114">注释</span><span class="sxs-lookup"><span data-stu-id="d3ab9-114">Remarks</span></span>

<span data-ttu-id="d3ab9-115">默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。</span><span class="sxs-lookup"><span data-stu-id="d3ab9-115">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="d3ab9-116">要加载与外接程序位于不同域中的页面，可以使用 **AppDomains** 和 **AppDomain** 元素来指定域。</span><span class="sxs-lookup"><span data-stu-id="d3ab9-116">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="d3ab9-117">此元素不能为空。</span><span class="sxs-lookup"><span data-stu-id="d3ab9-117">This element can't be empty.</span></span>
