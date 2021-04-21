---
title: 清单文件中运行时
description: Runtimes 元素指定外接程序的运行时。
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917084"
---
# <a name="runtimes-element"></a><span data-ttu-id="bd8b5-103">Runtimes 元素</span><span class="sxs-lookup"><span data-stu-id="bd8b5-103">Runtimes element</span></span>

<span data-ttu-id="bd8b5-104">指定外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="bd8b5-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="bd8b5-105">元素的 [`<Host>`](host.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="bd8b5-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8b5-106">When running in Office on Windows， an add-in that has a element in its manifest does notnecessarily `<Runtimes>` run in the same webview control as it otherwise would.</span><span class="sxs-lookup"><span data-stu-id="bd8b5-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="bd8b5-107">有关 Windows 和 Office 版本如何确定正常使用的 Webview 控件的信息，请参阅 [Office 外接程序使用的浏览器](../../concepts/browsers-used-by-office-web-add-ins.md)。如果满足针对将 Microsoft Edge 与 WebView2 一 (基于 Chromium) 的条件，则无论外接程序是否具有 元素，外接程序都使用该 `<Runtimes>` 浏览器。</span><span class="sxs-lookup"><span data-stu-id="bd8b5-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="bd8b5-108">但是，当不满足这些条件时，具有 元素的外接程序始终使用 `<Runtimes>` Internet Explorer 11，无论 Windows 或 Microsoft 365 版本如何。</span><span class="sxs-lookup"><span data-stu-id="bd8b5-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="bd8b5-109">**外接程序类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="bd8b5-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="bd8b5-110">语法</span><span class="sxs-lookup"><span data-stu-id="bd8b5-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="bd8b5-111">包含于</span><span class="sxs-lookup"><span data-stu-id="bd8b5-111">Contained in</span></span>

[<span data-ttu-id="bd8b5-112">Host</span><span class="sxs-lookup"><span data-stu-id="bd8b5-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="bd8b5-113">子元素</span><span class="sxs-lookup"><span data-stu-id="bd8b5-113">Child elements</span></span>

|  <span data-ttu-id="bd8b5-114">元素</span><span class="sxs-lookup"><span data-stu-id="bd8b5-114">Element</span></span> |  <span data-ttu-id="bd8b5-115">必需</span><span class="sxs-lookup"><span data-stu-id="bd8b5-115">Required</span></span>  |  <span data-ttu-id="bd8b5-116">说明</span><span class="sxs-lookup"><span data-stu-id="bd8b5-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="bd8b5-117">运行时</span><span class="sxs-lookup"><span data-stu-id="bd8b5-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="bd8b5-118">是</span><span class="sxs-lookup"><span data-stu-id="bd8b5-118">Yes</span></span> |  <span data-ttu-id="bd8b5-119">加载项的运行时。</span><span class="sxs-lookup"><span data-stu-id="bd8b5-119">The runtime for your add-in.</span></span> |

## <a name="see-also"></a><span data-ttu-id="bd8b5-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bd8b5-120">See also</span></span>

- [<span data-ttu-id="bd8b5-121">运行时</span><span class="sxs-lookup"><span data-stu-id="bd8b5-121">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="bd8b5-122">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="bd8b5-122">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="bd8b5-123">配置 Outlook 外接程序进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="bd8b5-123">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
