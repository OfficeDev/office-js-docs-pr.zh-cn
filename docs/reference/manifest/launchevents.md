---
title: '清单文件中的 LaunchEvents (预览) '
description: LaunchEvents 元素将加载项配置为基于支持的事件进行激活。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 9df059879018d79a61f1c900888c8d197e0b9880
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104810"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="b05f5-103">LaunchEvents 元素 (预览) </span><span class="sxs-lookup"><span data-stu-id="b05f5-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="b05f5-104">配置加载项以根据支持的事件激活。</span><span class="sxs-lookup"><span data-stu-id="b05f5-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="b05f5-105">元素的 [`<ExtensionPoint>`](extensionpoint.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="b05f5-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="b05f5-106">有关详细信息，请参阅配置 [Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="b05f5-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="b05f5-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="b05f5-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b05f5-108">基于事件的激活目前处于 [预览阶段，](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 仅在 Outlook 网页版和 Windows 版中可用。</span><span class="sxs-lookup"><span data-stu-id="b05f5-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web and Windows.</span></span> <span data-ttu-id="b05f5-109">有关详细信息，请参阅 [如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="b05f5-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="b05f5-110">语法</span><span class="sxs-lookup"><span data-stu-id="b05f5-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b05f5-111">包含于</span><span class="sxs-lookup"><span data-stu-id="b05f5-111">Contained in</span></span>

<span data-ttu-id="b05f5-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** 邮件外接程序) </span><span class="sxs-lookup"><span data-stu-id="b05f5-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="b05f5-113">子元素</span><span class="sxs-lookup"><span data-stu-id="b05f5-113">Child elements</span></span>

|  <span data-ttu-id="b05f5-114">元素</span><span class="sxs-lookup"><span data-stu-id="b05f5-114">Element</span></span> |  <span data-ttu-id="b05f5-115">必需</span><span class="sxs-lookup"><span data-stu-id="b05f5-115">Required</span></span>  |  <span data-ttu-id="b05f5-116">说明</span><span class="sxs-lookup"><span data-stu-id="b05f5-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="b05f5-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="b05f5-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="b05f5-118">是</span><span class="sxs-lookup"><span data-stu-id="b05f5-118">Yes</span></span> |  <span data-ttu-id="b05f5-119">将支持的事件映射到 JavaScript 文件中用于加载项激活的函数。</span><span class="sxs-lookup"><span data-stu-id="b05f5-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="b05f5-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b05f5-120">See also</span></span>

- [<span data-ttu-id="b05f5-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="b05f5-121">LaunchEvent</span></span>](launchevent.md)
