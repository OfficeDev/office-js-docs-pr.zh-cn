---
title: 清单文件中运行时
description: Runtimes 元素指定外接程序的运行时。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555295"
---
# <a name="runtimes-element"></a><span data-ttu-id="58aa4-103">Runtimes 元素</span><span class="sxs-lookup"><span data-stu-id="58aa4-103">Runtimes element</span></span>

<span data-ttu-id="58aa4-104">指定外接程序的运行时。</span><span class="sxs-lookup"><span data-stu-id="58aa4-104">Specifies the runtime of your add-in.</span></span> <span data-ttu-id="58aa4-105">元素的 [`<Host>`](host.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="58aa4-105">Child of the [`<Host>`](host.md) element.</span></span>

> [!NOTE]
> <span data-ttu-id="58aa4-106">当在 Office Windows 中运行时，其清单中具有 元素的加载项不必像否则一样在同一 `<Runtimes>` Webview 控件中运行。</span><span class="sxs-lookup"><span data-stu-id="58aa4-106">When running in Office on Windows, an add-in that has a `<Runtimes>` element in its manifest does not necessarily run in the same webview control as it otherwise would.</span></span> <span data-ttu-id="58aa4-107">有关 web 视图和 Windows Office版本如何确定通常使用的 Webview 控件Office[请参阅](../../concepts/browsers-used-by-office-web-add-ins.md)浏览器。如果满足将 Microsoft Edge与 WebView2 (Chromium) 的条件，则无论外接程序是否具有 元素，外接程序都使用该 `<Runtimes>` 浏览器。</span><span class="sxs-lookup"><span data-stu-id="58aa4-107">For more information about how the versions of Windows and Office determine what webview control is normally used, see [Browsers used by Office Add-ins](../../concepts/browsers-used-by-office-web-add-ins.md). If the conditions described there for using Microsoft Edge with WebView2 (Chromium-based) are met, then the add-in uses that browser whether or not it has a `<Runtimes>` element.</span></span> <span data-ttu-id="58aa4-108">但是，当不满足这些条件时，具有 元素的外接程序始终使用 Internet Explorer 11，而不考虑 Windows `<Runtimes>` 或 Microsoft 365 版本。</span><span class="sxs-lookup"><span data-stu-id="58aa4-108">However, when those conditions are not met, an add-in with a `<Runtimes>` element always uses Internet Explorer 11 regardless of the Windows or Microsoft 365 version.</span></span>

<span data-ttu-id="58aa4-109">**外接程序类型：** 任务窗格、邮件</span><span class="sxs-lookup"><span data-stu-id="58aa4-109">**Add-in type:** Task pane, Mail</span></span>

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a><span data-ttu-id="58aa4-110">语法</span><span class="sxs-lookup"><span data-stu-id="58aa4-110">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="58aa4-111">包含于</span><span class="sxs-lookup"><span data-stu-id="58aa4-111">Contained in</span></span>

[<span data-ttu-id="58aa4-112">Host</span><span class="sxs-lookup"><span data-stu-id="58aa4-112">Host</span></span>](host.md)

## <a name="child-elements"></a><span data-ttu-id="58aa4-113">子元素</span><span class="sxs-lookup"><span data-stu-id="58aa4-113">Child elements</span></span>

|  <span data-ttu-id="58aa4-114">元素</span><span class="sxs-lookup"><span data-stu-id="58aa4-114">Element</span></span> |  <span data-ttu-id="58aa4-115">必需</span><span class="sxs-lookup"><span data-stu-id="58aa4-115">Required</span></span>  |  <span data-ttu-id="58aa4-116">说明</span><span class="sxs-lookup"><span data-stu-id="58aa4-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="58aa4-117">运行时</span><span class="sxs-lookup"><span data-stu-id="58aa4-117">Runtime</span></span>](runtime.md) | <span data-ttu-id="58aa4-118">是</span><span class="sxs-lookup"><span data-stu-id="58aa4-118">Yes</span></span> |  <span data-ttu-id="58aa4-119">加载项的运行时。</span><span class="sxs-lookup"><span data-stu-id="58aa4-119">The runtime for your add-in.</span></span> <span data-ttu-id="58aa4-120">**重要** 提示：目前，只能定义一 `<Runtime>` 个元素。</span><span class="sxs-lookup"><span data-stu-id="58aa4-120">**Important**: At present, you can only define one `<Runtime>` element.</span></span> |

## <a name="see-also"></a><span data-ttu-id="58aa4-121">另请参阅</span><span class="sxs-lookup"><span data-stu-id="58aa4-121">See also</span></span>

- [<span data-ttu-id="58aa4-122">运行时</span><span class="sxs-lookup"><span data-stu-id="58aa4-122">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="58aa4-123">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="58aa4-123">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="58aa4-124">配置Outlook加载项进行基于事件的激活</span><span class="sxs-lookup"><span data-stu-id="58aa4-124">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
