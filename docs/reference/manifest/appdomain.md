---
title: 清单文件中的 AppDomain 元素
description: ''
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: da28b3b4dec5d669462a781db3c0628bd32c7182
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596786"
---
# <a name="appdomain-element"></a><span data-ttu-id="0f852-102">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="0f852-102">AppDomain element</span></span>

<span data-ttu-id="0f852-103">指定在外接程序窗口中加载页面的其他域。</span><span class="sxs-lookup"><span data-stu-id="0f852-103">Specifies additional domains that load pages in the add-in window.</span></span> <span data-ttu-id="0f852-104">此外，它还列出了可以从加载项内的 Iframe 中进行的 Office .js API 调用的受信任域。</span><span class="sxs-lookup"><span data-stu-id="0f852-104">It also lists trusted domains from which Office.js API calls can be made from IFrames within the add-in.</span></span>

<span data-ttu-id="0f852-105">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="0f852-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0f852-106">语法</span><span class="sxs-lookup"><span data-stu-id="0f852-106">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="0f852-107">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain</AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="0f852-107">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain</AppDomain>`).</span></span>
> 2. <span data-ttu-id="0f852-108">不要*在值上添加一个*结束斜杠 "/"。</span><span class="sxs-lookup"><span data-stu-id="0f852-108">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="0f852-109">包含于</span><span class="sxs-lookup"><span data-stu-id="0f852-109">Contained in</span></span>

[<span data-ttu-id="0f852-110">AppDomains</span><span class="sxs-lookup"><span data-stu-id="0f852-110">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="0f852-111">注释</span><span class="sxs-lookup"><span data-stu-id="0f852-111">Remarks</span></span>

<span data-ttu-id="0f852-112">**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。</span><span class="sxs-lookup"><span data-stu-id="0f852-112">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="0f852-113">有关详细信息，请参阅 [Office 加载项 XML 清单](../../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="0f852-113">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
