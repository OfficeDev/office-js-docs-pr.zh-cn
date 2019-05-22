---
title: 清单文件中的 AppDomain 元素
description: ''
ms.date: 05/15/2019
localization_priority: Normal
ms.openlocfilehash: b1d71648cc7646eec246f3d0a8113c843eed2e74
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337193"
---
# <a name="appdomain-element"></a><span data-ttu-id="18ae9-102">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="18ae9-102">AppDomain element</span></span>

<span data-ttu-id="18ae9-103">指定将用于在外接程序窗口中加载页面的其他域。</span><span class="sxs-lookup"><span data-stu-id="18ae9-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="18ae9-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="18ae9-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="18ae9-105">语法</span><span class="sxs-lookup"><span data-stu-id="18ae9-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="18ae9-106">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain</AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="18ae9-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="18ae9-107">不要\*\* 在值上添加一个结束斜杠 "/"。</span><span class="sxs-lookup"><span data-stu-id="18ae9-107">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="18ae9-108">包含于</span><span class="sxs-lookup"><span data-stu-id="18ae9-108">Contained in</span></span>

[<span data-ttu-id="18ae9-109">AppDomains</span><span class="sxs-lookup"><span data-stu-id="18ae9-109">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="18ae9-110">注释</span><span class="sxs-lookup"><span data-stu-id="18ae9-110">Remarks</span></span>

<span data-ttu-id="18ae9-111">**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。</span><span class="sxs-lookup"><span data-stu-id="18ae9-111">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="18ae9-112">有关详细信息，请参阅 [Office 加载项 XML 清单](/office/dev/add-ins/develop/add-in-manifests)。</span><span class="sxs-lookup"><span data-stu-id="18ae9-112">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
