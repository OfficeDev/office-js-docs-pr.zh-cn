---
title: 清单文件中的 LaunchEvents （预览）
description: LaunchEvents 元素将你的外接程序配置为根据受支持的事件进行激活。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 2e1ad56d405fca0f85fad500a113fba7d0448caf
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278527"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="f2f64-103">LaunchEvents 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="f2f64-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="f2f64-104">将你的外接程序配置为根据受支持的事件进行激活。</span><span class="sxs-lookup"><span data-stu-id="f2f64-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="f2f64-105">元素的子 [`<ExtensionPoint>`](extensionpoint.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="f2f64-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="f2f64-106">有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="f2f64-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="f2f64-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="f2f64-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f2f64-108">基于事件的激活当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅在 Outlook 网页版中可用。</span><span class="sxs-lookup"><span data-stu-id="f2f64-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="f2f64-109">有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="f2f64-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="f2f64-110">语法</span><span class="sxs-lookup"><span data-stu-id="f2f64-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f2f64-111">包含于</span><span class="sxs-lookup"><span data-stu-id="f2f64-111">Contained in</span></span>

<span data-ttu-id="f2f64-112">[ExtensionPoint](extensionpoint.md) （**LaunchEvent**邮件外接端）</span><span class="sxs-lookup"><span data-stu-id="f2f64-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="f2f64-113">子元素</span><span class="sxs-lookup"><span data-stu-id="f2f64-113">Child elements</span></span>

|  <span data-ttu-id="f2f64-114">元素</span><span class="sxs-lookup"><span data-stu-id="f2f64-114">Element</span></span> |  <span data-ttu-id="f2f64-115">必需</span><span class="sxs-lookup"><span data-stu-id="f2f64-115">Required</span></span>  |  <span data-ttu-id="f2f64-116">说明</span><span class="sxs-lookup"><span data-stu-id="f2f64-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="f2f64-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="f2f64-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="f2f64-118">是</span><span class="sxs-lookup"><span data-stu-id="f2f64-118">Yes</span></span> |  <span data-ttu-id="f2f64-119">将支持的事件映射到其在外接程序激活的 JavaScript 文件中的功能。</span><span class="sxs-lookup"><span data-stu-id="f2f64-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f2f64-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f2f64-120">See also</span></span>

- [<span data-ttu-id="f2f64-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="f2f64-121">LaunchEvent</span></span>](launchevent.md)
