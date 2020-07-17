---
title: 清单文件中的 AppDomain 元素
description: 指定加载项使用的其他域，并且应受 Office 信任。
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: ae49944afceada559b39353cd119e26a21fd3d15
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778646"
---
# <a name="appdomain-element"></a><span data-ttu-id="c3d9f-103">AppDomain 元素</span><span class="sxs-lookup"><span data-stu-id="c3d9f-103">AppDomain element</span></span>

<span data-ttu-id="c3d9f-104">指定除了在[SourceLocation 元素](sourcelocation.md)中指定的域之外，Office 应信任的其他域。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-104">Specifies an additional domain that Office should trust, in addition to the one specified in the [SourceLocation element](sourcelocation.md).</span></span> <span data-ttu-id="c3d9f-105">指定域具有以下效果：</span><span class="sxs-lookup"><span data-stu-id="c3d9f-105">Specifying a domain has these effects:</span></span>

- <span data-ttu-id="c3d9f-106">它允许在桌面 Office 平台上的加载项的根任务窗格中直接打开页面、路由或域中的其他资源。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-106">It enables pages, routes, or other resources in the domain to be opened directly in the root task pane of the add-in on desktop Office platforms.</span></span> <span data-ttu-id="c3d9f-107">（为 web 上的 Office**指定不需要的域**或在 IFrame 中打开资源，也需要在使用[对话框 API](../../develop/dialog-api-in-office-add-ins.md)打开的对话框中打开资源时。）</span><span class="sxs-lookup"><span data-stu-id="c3d9f-107">(Specifying a domain in an **AppDomain** isn't necessary for Office on the web or to open a resource in an IFrame, nor it is necessary for opening a resource in a dialog opened with the [Dialog API](../../develop/dialog-api-in-office-add-ins.md).)</span></span>
- <span data-ttu-id="c3d9f-108">它使域中的页面可以从加载项中的 Iframe 进行 Office.js API 调用。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-108">It enables pages in the domain to make Office.js API calls from IFrames within the add-in.</span></span>

<span data-ttu-id="c3d9f-109">**加载项类型：** 内容、任务窗格和邮件</span><span class="sxs-lookup"><span data-stu-id="c3d9f-109">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="c3d9f-110">语法</span><span class="sxs-lookup"><span data-stu-id="c3d9f-110">Syntax</span></span>

```XML
<AppDomain>string</AppDomain>
```

> [!IMPORTANT]
> 1. <span data-ttu-id="c3d9f-111">**AppDomain** 元素的值必须包括协议（如 `<AppDomain>https://myappdomain.com</AppDomain>`）。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-111">The value of the **AppDomain** element must include the protocol (e.g., `<AppDomain>https://myappdomain.com</AppDomain>`).</span></span>
> 2. <span data-ttu-id="c3d9f-112">如果有域的显式端口，请将其包括在内（例如， `<AppDomain>https://myappdomain.com:9999</AppDomain>` ）。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-112">If there is an explicit port for the domain, include it (e.g.,`<AppDomain>https://myappdomain.com:9999</AppDomain>`).</span></span>
> 3. <span data-ttu-id="c3d9f-113">如果需要信任某个子域，请将其包括在内（例如， `<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>` ）。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-113">If a subdomain needs to be trusted, include it (e.g.,`<AppDomain>https://mysubdomain.myappdomain.com</AppDomain>`).</span></span> <span data-ttu-id="c3d9f-114">子域 `mysubdomain.mydomain.com` 和 `mydomain.com` 不同的域。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-114">The subdomain `mysubdomain.mydomain.com` and `mydomain.com` are different domains.</span></span> <span data-ttu-id="c3d9f-115">如果两者都需要信任，则这两个元素都需要位于单独的**AppDomain**元素中。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-115">If both need to be trusted, then both need to be in separate **AppDomain** elements.</span></span>
> 4. <span data-ttu-id="c3d9f-116">列出与[SourceLocation 元素](sourcelocation.md)中指定的域相同的域不起作用，并且可能会引起误导。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-116">Listing the same domain as the one specified in the [SourceLocation element](sourcelocation.md) has no effect and may be misleading.</span></span> <span data-ttu-id="c3d9f-117">特别是在上进行开发时 `localhost` ，不需要为创建**AppDomain**元素 `localhost` 。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-117">In particular, when you are developing on `localhost`, you don't need to create an **AppDomain** element for `localhost`.</span></span>
> 5. <span data-ttu-id="c3d9f-118">不要将任何段的 URL 包含在域之后。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-118">Don't include any segments of a URL past the domain.</span></span> <span data-ttu-id="c3d9f-119">例如，不要包含页面的完整 URL。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-119">For example, don't include the full URL of a page.</span></span>
> 6. <span data-ttu-id="c3d9f-120">不要*在值上添加一个*结束斜杠 "/"。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-120">Do *not* put a closing slash, "/", on the value.</span></span>

## <a name="contained-in"></a><span data-ttu-id="c3d9f-121">包含于</span><span class="sxs-lookup"><span data-stu-id="c3d9f-121">Contained in</span></span>

[<span data-ttu-id="c3d9f-122">AppDomains</span><span class="sxs-lookup"><span data-stu-id="c3d9f-122">AppDomains</span></span>](appdomains.md)

## <a name="remarks"></a><span data-ttu-id="c3d9f-123">注解</span><span class="sxs-lookup"><span data-stu-id="c3d9f-123">Remarks</span></span>

<span data-ttu-id="c3d9f-124">有关详细信息，请参阅 [Office 外接程序 XML 清单](../../develop/add-in-manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="c3d9f-124">For more information, see [Office Add-ins XML manifest](../../develop/add-in-manifests.md).</span></span>
