---
title: 清单文件中 LaunchEvents
description: LaunchEvents 元素将外接程序配置为基于支持的事件进行激活。
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 16d721ca6d9402d2bd5d19787707e146358044f0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590913"
---
# <a name="launchevents-element"></a><span data-ttu-id="78d78-103">LaunchEvents 元素</span><span class="sxs-lookup"><span data-stu-id="78d78-103">LaunchEvents element</span></span>

<span data-ttu-id="78d78-104">将加载项配置为基于支持的事件进行激活。</span><span class="sxs-lookup"><span data-stu-id="78d78-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="78d78-105">元素的 [`<ExtensionPoint>`](extensionpoint.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="78d78-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="78d78-106">有关详细信息，请参阅[Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="78d78-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="78d78-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="78d78-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="78d78-108">语法</span><span class="sxs-lookup"><span data-stu-id="78d78-108">Syntax</span></span>

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a><span data-ttu-id="78d78-109">包含于</span><span class="sxs-lookup"><span data-stu-id="78d78-109">Contained in</span></span>

<span data-ttu-id="78d78-110">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** 邮件外接程序) </span><span class="sxs-lookup"><span data-stu-id="78d78-110">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="78d78-111">子元素</span><span class="sxs-lookup"><span data-stu-id="78d78-111">Child elements</span></span>

|  <span data-ttu-id="78d78-112">元素</span><span class="sxs-lookup"><span data-stu-id="78d78-112">Element</span></span> |  <span data-ttu-id="78d78-113">必需</span><span class="sxs-lookup"><span data-stu-id="78d78-113">Required</span></span>  |  <span data-ttu-id="78d78-114">说明</span><span class="sxs-lookup"><span data-stu-id="78d78-114">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="78d78-115">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="78d78-115">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="78d78-116">是</span><span class="sxs-lookup"><span data-stu-id="78d78-116">Yes</span></span> |  <span data-ttu-id="78d78-117">将受支持的事件映射到 JavaScript 文件中用于外接程序激活的函数。</span><span class="sxs-lookup"><span data-stu-id="78d78-117">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="78d78-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="78d78-118">See also</span></span>

- [<span data-ttu-id="78d78-119">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="78d78-119">LaunchEvent</span></span>](launchevent.md)
