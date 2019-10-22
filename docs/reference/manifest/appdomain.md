---
title: 清单文件中的 AppDomain 元素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 2f65302d1ac3d85f2867cd13501bc67606cd00b5
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/21/2019
ms.locfileid: "35575637"
---
# <a name="appdomain-element"></a><span data-ttu-id="c1875-102">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="c1875-102">AppDomain element</span></span>

<span data-ttu-id="c1875-103">指定在外接程序窗口中加载页面的其他域。</span><span class="sxs-lookup"><span data-stu-id="c1875-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="c1875-104">此外，它还列出了可以从加载项内的 Iframe 中进行的 Office .js API 调用的受信任域。</span><span class="sxs-lookup"><span data-stu-id="c1875-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="c1875-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="c1875-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c1875-106">语法</span><span class="sxs-lookup"><span data-stu-id="c1875-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="c1875-107">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain</AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="c1875-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="c1875-108">不要*在值上添加一个*结束斜杠 "/"。</span><span class="sxs-lookup"><span data-stu-id="c1875-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="c1875-109">包含于</span><span class="sxs-lookup"><span data-stu-id="c1875-109">Contained in</span></span>

[<span data-ttu-id="c1875-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="c1875-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="c1875-111">注释</span><span class="sxs-lookup"><span data-stu-id="c1875-111">Remarks</span></span>

<span data-ttu-id="c1875-112">**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。</span><span class="sxs-lookup"><span data-stu-id="c1875-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="c1875-113">有关详细信息，请参阅 [Office 加载项 XML 清单](/office/dev/add-ins/develop/add-in-manifests)。</span><span class="sxs-lookup"><span data-stu-id="c1875-113">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
