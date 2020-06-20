---
title: 清单文件中的 AppDomains 元素
description: 列出除 Office 外接程序将使用的元素中指定的域之外的所有域 `SourceLocation` ，以及 office 应信任的域。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778653"
---
# <a name="appdomains-element"></a><span data-ttu-id="e3896-103">AppDomains 元素</span><span class="sxs-lookup"><span data-stu-id="e3896-103">AppDomains element</span></span>

<span data-ttu-id="e3896-104">列出 `SourceLocation` 您的 Office 外接程序将使用且应受 office 信任的任何域（除了元素中指定的域）。</span><span class="sxs-lookup"><span data-stu-id="e3896-104">Lists any domains, in addition to the domain specified in the `SourceLocation` element, that your Office Add-in will use and that should be trusted by Office.</span></span> <span data-ttu-id="e3896-105">这使域中的页面可以调用来自加载项中的 Iframe 的 Office.js Api，并具有其他效果。</span><span class="sxs-lookup"><span data-stu-id="e3896-105">This enables pages in the domains to make calls to Office.js APIs from IFrames within the add-in and has other effects.</span></span> <span data-ttu-id="e3896-106">对于每个其他域，指定 **AppDomain** 元素。</span><span class="sxs-lookup"><span data-stu-id="e3896-106">For each additional domain, specify an **AppDomain** element.</span></span>

 <span data-ttu-id="e3896-107">**外接程序类型：** 内容、任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="e3896-107">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e3896-108">语法</span><span class="sxs-lookup"><span data-stu-id="e3896-108">Syntax</span></span>

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> <span data-ttu-id="e3896-109">对可以成为**AppDomain**元素的值的值有一些限制。</span><span class="sxs-lookup"><span data-stu-id="e3896-109">There are restrictions on what can be the value of a **AppDomain** element.</span></span> <span data-ttu-id="e3896-110">有关详细信息，请参阅[AppDomain](appdomain.md)。</span><span class="sxs-lookup"><span data-stu-id="e3896-110">For more information, see [AppDomain](appdomain.md).</span></span>

## <a name="contained-in"></a><span data-ttu-id="e3896-111">包含于</span><span class="sxs-lookup"><span data-stu-id="e3896-111">Contained in</span></span>

[<span data-ttu-id="e3896-112">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="e3896-112">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="e3896-113">可以包含</span><span class="sxs-lookup"><span data-stu-id="e3896-113">Can contain</span></span>

[<span data-ttu-id="e3896-114">AppDomain</span><span class="sxs-lookup"><span data-stu-id="e3896-114">AppDomain</span></span>](appdomain.md)

## <a name="remarks"></a><span data-ttu-id="e3896-115">注释</span><span class="sxs-lookup"><span data-stu-id="e3896-115">Remarks</span></span>

<span data-ttu-id="e3896-116">默认情况下，外接程序可以加载与 [SourceLocation](sourcelocation.md) 元素中指定的位置位于同一个域中的任何页面。</span><span class="sxs-lookup"><span data-stu-id="e3896-116">By default, your add-in can load any page that is in the same domain as the location specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="e3896-117">此元素不能为空。</span><span class="sxs-lookup"><span data-stu-id="e3896-117">This element can't be empty.</span></span>
