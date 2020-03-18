---
title: 清单文件中的 AppDomain 元素
description: 指定在外接程序窗口中加载页面的其他域。
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 6990f759df806f24b1d617c036bc1a452e6da38f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718452"
---
# <a name="appdomain-element"></a><span data-ttu-id="2da9b-103">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="2da9b-103">AppDomain element</span></span>

<span data-ttu-id="2da9b-104">指定在外接程序窗口中加载页面的其他域。</span><span class="sxs-lookup"><span data-stu-id="2da9b-104">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="2da9b-105">此外，它还列出了可以从加载项内的 Iframe 中进行的 Office .js API 调用的受信任域。</span><span class="sxs-lookup"><span data-stu-id="2da9b-105">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="2da9b-106">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="2da9b-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2da9b-107">语法</span><span class="sxs-lookup"><span data-stu-id="2da9b-107">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="2da9b-108">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain</AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="2da9b-108">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="2da9b-109">不要*在值上添加一个*结束斜杠 "/"。</span><span class="sxs-lookup"><span data-stu-id="2da9b-109">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="2da9b-110">包含于</span><span class="sxs-lookup"><span data-stu-id="2da9b-110">Contained in</span></span>

[<span data-ttu-id="2da9b-111">AppDomains</span><span class="sxs-lookup"><span data-stu-id="2da9b-111">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="2da9b-112">注释</span><span class="sxs-lookup"><span data-stu-id="2da9b-112">Remarks</span></span>

<span data-ttu-id="2da9b-113">**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。</span><span class="sxs-lookup"><span data-stu-id="2da9b-113">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="2da9b-114">有关详细信息，请参阅 [Office 加载项 XML 清单](../../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="2da9b-114">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
