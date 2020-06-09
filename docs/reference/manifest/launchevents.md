---
title: 清单文件中的 LaunchEvents （预览）
description: LaunchEvents 元素将你的外接程序配置为根据受支持的事件进行激活。
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 92416f8c646326410a8cd9ee7831e17a5c5f1ffc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611769"
---
# <a name="launchevents-element-preview"></a><span data-ttu-id="f71de-103">LaunchEvents 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="f71de-103">LaunchEvents element (preview)</span></span>

<span data-ttu-id="f71de-104">将你的外接程序配置为根据受支持的事件进行激活。</span><span class="sxs-lookup"><span data-stu-id="f71de-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="f71de-105">元素的子 [`<ExtensionPoint>`](extensionpoint.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="f71de-105">Child of the [`<ExtensionPoint>`](extensionpoint.md) element.</span></span> <span data-ttu-id="f71de-106">有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="f71de-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="f71de-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="f71de-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f71de-108">基于事件的激活当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅在 Outlook 网页版中可用。</span><span class="sxs-lookup"><span data-stu-id="f71de-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="f71de-109">有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="f71de-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="f71de-110">语法</span><span class="sxs-lookup"><span data-stu-id="f71de-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f71de-111">包含于</span><span class="sxs-lookup"><span data-stu-id="f71de-111">Contained in</span></span>

<span data-ttu-id="f71de-112">[ExtensionPoint](extensionpoint.md) （**LaunchEvent**邮件外接端）</span><span class="sxs-lookup"><span data-stu-id="f71de-112">[ExtensionPoint](extensionpoint.md) (**LaunchEvent** mail add-in)</span></span>

## <a name="child-elements"></a><span data-ttu-id="f71de-113">子元素</span><span class="sxs-lookup"><span data-stu-id="f71de-113">Child elements</span></span>

|  <span data-ttu-id="f71de-114">元素</span><span class="sxs-lookup"><span data-stu-id="f71de-114">Element</span></span> |  <span data-ttu-id="f71de-115">必需</span><span class="sxs-lookup"><span data-stu-id="f71de-115">Required</span></span>  |  <span data-ttu-id="f71de-116">Description</span><span class="sxs-lookup"><span data-stu-id="f71de-116">Description</span></span>  |
|:-----|:-----|:-----|
| [<span data-ttu-id="f71de-117">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="f71de-117">LaunchEvent</span></span>](launchevent.md) | <span data-ttu-id="f71de-118">是</span><span class="sxs-lookup"><span data-stu-id="f71de-118">Yes</span></span> |  <span data-ttu-id="f71de-119">将支持的事件映射到其在外接程序激活的 JavaScript 文件中的功能。</span><span class="sxs-lookup"><span data-stu-id="f71de-119">Map supported event to its function in the JavaScript file for add-in activation.</span></span> |

## <a name="see-also"></a><span data-ttu-id="f71de-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f71de-120">See also</span></span>

- [<span data-ttu-id="f71de-121">LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="f71de-121">LaunchEvent</span></span>](launchevent.md)
