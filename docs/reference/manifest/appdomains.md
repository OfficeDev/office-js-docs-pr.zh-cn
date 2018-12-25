---
title: 清单文件中的 AppDomains 元素
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: cc2f5ade0bdda214c85490f8e474b42f921edbe8
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433676"
---
# <a name="appdomains-element"></a><span data-ttu-id="67e5d-102">AppDomains 元素</span><span class="sxs-lookup"><span data-stu-id="67e5d-102">AppDomains element</span></span>

<span data-ttu-id="67e5d-p101">列出了除 Office 外接程序用于加载页面的 SourceLocation 元素中指定的域之外的所有域。对于每个其他域，指定 AppDomain 元素。</span><span class="sxs-lookup"><span data-stu-id="67e5d-p101">Lists any domains in addition to the domain specified in the SourceLocation element that your Office Add-in will use to load pages. For each additional domain, specify an AppDomain element.</span></span>

 <span data-ttu-id="67e5d-105">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="67e5d-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="67e5d-106">语法</span><span class="sxs-lookup"><span data-stu-id="67e5d-106">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="67e5d-107">每个 **AppDomain** 元素的值都必须包括协议（如 `<AppDomain>https://myappdomain<AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="67e5d-107">The value of each **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="67e5d-108">包含于</span><span class="sxs-lookup"><span data-stu-id="67e5d-108">Contained in</span></span>

[<span data-ttu-id="67e5d-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="67e5d-109">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="67e5d-110">可以包含</span><span class="sxs-lookup"><span data-stu-id="67e5d-110">Can contain</span></span>

[<span data-ttu-id="67e5d-111">AppDomain</span><span class="sxs-lookup"><span data-stu-id="67e5d-111">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="67e5d-112">注释</span><span class="sxs-lookup"><span data-stu-id="67e5d-112">Remarks</span></span>

<span data-ttu-id="67e5d-113">默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。</span><span class="sxs-lookup"><span data-stu-id="67e5d-113">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="67e5d-114">要加载与外接程序位于不同域中的页面，可以使用 **AppDomains** 和 **AppDomain** 元素来指定域。</span><span class="sxs-lookup"><span data-stu-id="67e5d-114">To load pages that are not in the same domain as the add-in, specify their domains by using the **AppDomains** and **AppDomain** elements.</span></span> <span data-ttu-id="67e5d-115">此元素不能为空。</span><span class="sxs-lookup"><span data-stu-id="67e5d-115">This element can't be empty.</span></span>
