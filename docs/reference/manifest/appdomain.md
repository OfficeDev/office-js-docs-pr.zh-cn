---
title: 清单文件中的 AppDomain 元素
description: ''
ms.date: 12/13/2018
ms.openlocfilehash: 2b55f2c1ea7a2a3dc7dec42c913d74006c0f2e3b
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433066"
---
# <a name="appdomain-element"></a><span data-ttu-id="b99a7-102">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="b99a7-102">AppDomain element</span></span>

<span data-ttu-id="b99a7-103">指定将用于在外接程序窗口中加载页面的其他域。</span><span class="sxs-lookup"><span data-stu-id="b99a7-103">Specifies an additional domain that will be used to load pages in the add-in window.</span></span>

<span data-ttu-id="b99a7-104">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="b99a7-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b99a7-105">语法</span><span class="sxs-lookup"><span data-stu-id="b99a7-105">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> <span data-ttu-id="b99a7-106">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain<AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="b99a7-106">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain<AppDomain>`).</span></span>

## <a name="contained-in"></a><span data-ttu-id="b99a7-107">包含于</span><span class="sxs-lookup"><span data-stu-id="b99a7-107">Contained in</span></span>

[<span data-ttu-id="b99a7-108">AppDomains</span><span class="sxs-lookup"><span data-stu-id="b99a7-108">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="b99a7-109">注释</span><span class="sxs-lookup"><span data-stu-id="b99a7-109">Remarks</span></span>

<span data-ttu-id="b99a7-110">**AppDomain** 元素用于指定除在 [SourceLocation](sourcelocation.md) 元素中指定的域之外的任何其他域。</span><span class="sxs-lookup"><span data-stu-id="b99a7-110">**AppDomain** elements should be used to specify any additional domains other than the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="b99a7-111">有关详细信息，请参阅 [Office 加载项 XML 清单](/office/dev/add-ins/develop/add-in-manifests)。</span><span class="sxs-lookup"><span data-stu-id="b99a7-111">For more information, see [Office Add-ins XML manifest](/office/dev/add-ins/develop/add-in-manifests).</span></span>
