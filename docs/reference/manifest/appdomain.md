---
title: 清单文件中的 AppDomain 元素
description: ''
ms.date: 03/21/2019
localization_priority: Normal
ms.openlocfilehash: 8216603c87a7dcafde84d25a82f068c9aa86ed96
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870406"
---
# <a name="appdomain-element"></a><span data-ttu-id="c63ea-102">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="c63ea-102">AppDomain element</span></span>

<span data-ttu-id="c63ea-103">指定将用于在外接程序窗口中加载页面的其他域。</span><span class="sxs-lookup"><span data-stu-id="c63ea-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="c63ea-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="c63ea-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c63ea-105">语法</span><span class="sxs-lookup"><span data-stu-id="c63ea-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="c63ea-106">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain</AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="c63ea-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="c63ea-107">不要\*\* 在值上添加一个结束斜杠 "/"。</span><span class="sxs-lookup"><span data-stu-id="c63ea-107">Do *not* put a closing slash, "/", on the the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="c63ea-108">包含于</span><span class="sxs-lookup"><span data-stu-id="c63ea-108">Contained in</span></span>

[<span data-ttu-id="c63ea-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="c63ea-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="c63ea-110">注释</span><span class="sxs-lookup"><span data-stu-id="c63ea-110">Remarks</span></span>

<span data-ttu-id="c63ea-111">**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。</span><span class="sxs-lookup"><span data-stu-id="c63ea-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="c63ea-112">有关详细信息，请参阅 [Office 加载项 XML 清单](/office/dev/add-ins/develop/add-in-manifests)。</span><span class="sxs-lookup"><span data-stu-id="c63ea-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
