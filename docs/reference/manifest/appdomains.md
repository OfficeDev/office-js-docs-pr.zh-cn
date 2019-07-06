---
title: 清单文件中的 AppDomains 元素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: b6db3d46d004021f25edd5733566544010abb457
ms.sourcegitcommit: c3673cc693fa7070e1b397922bd735ba3f9342f3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/05/2019
ms.locfileid: "35575329"
---
# <a name="appdomains-element"></a><span data-ttu-id="c9d0d-102">AppDomains 元素</span><span class="sxs-lookup"><span data-stu-id="c9d0d-102">AppDomains element</span></span>

<span data-ttu-id="c9d0d-103">列出除 Office 加载项将用于加载页面的`SourceLocation`元素中指定的域之外的所有域。</span><span class="sxs-lookup"><span data-stu-id="c9d0d-103">Lists any domains in addition to the domain specified in the `SourceLocation` element that your Office Add-in will use to load pages.</span></span> <span data-ttu-id="c9d0d-104">此外, 它还列出了可以从加载项内的 Iframe 中进行的 Office .js API 调用的受信任域。</span><span class="sxs-lookup"><span data-stu-id="c9d0d-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span> <span data-ttu-id="c9d0d-105">对于每个其他域，指定 AppDomain 元素。</span><span class="sxs-lookup"><span data-stu-id="c9d0d-105">For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="c9d0d-106">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="c9d0d-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c9d0d-107">语法</span><span class="sxs-lookup"><span data-stu-id="c9d0d-107">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="c9d0d-108">每个 **AppDomain** 元素的值都必须包括协议（如 `<AppDomain>https://myappdomain<AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="c9d0d-108">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="c9d0d-109">包含于</span><span class="sxs-lookup"><span data-stu-id="c9d0d-109">Contained in</span></span>

[<span data-ttu-id="c9d0d-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="c9d0d-110">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="c9d0d-111">可以包含</span><span class="sxs-lookup"><span data-stu-id="c9d0d-111">Can contain</span></span>

[<span data-ttu-id="c9d0d-112">AppDomain</span><span class="sxs-lookup"><span data-stu-id="c9d0d-112">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="c9d0d-113">注释</span><span class="sxs-lookup"><span data-stu-id="c9d0d-113">Remarks</span></span>

<span data-ttu-id="c9d0d-114">默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。</span><span class="sxs-lookup"><span data-stu-id="c9d0d-114">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="c9d0d-115">要加载与外接程序位于不同域中的页面，可以使用 **AppDomains** 和 **AppDomain** 元素来指定域。</span><span class="sxs-lookup"><span data-stu-id="c9d0d-115">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="c9d0d-116">此元素不能为空。</span><span class="sxs-lookup"><span data-stu-id="c9d0d-116">This element can't be empty.</span></span>
