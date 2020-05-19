---
title: 清单文件中的 LaunchEvent （预览）
description: LaunchEvent 元素将你的外接程序配置为根据受支持的事件进行激活。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: a4f5208ec7f735d926c3a878cae34973c3992cf9
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278528"
---
# <a name="launchevent-element-preview"></a><span data-ttu-id="98e32-103">LaunchEvent 元素（预览）</span><span class="sxs-lookup"><span data-stu-id="98e32-103">LaunchEvent element (preview)</span></span>

<span data-ttu-id="98e32-104">将你的外接程序配置为根据受支持的事件进行激活。</span><span class="sxs-lookup"><span data-stu-id="98e32-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="98e32-105">元素的子 [`<LaunchEvents>`](launchevents.md) 元素。</span><span class="sxs-lookup"><span data-stu-id="98e32-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="98e32-106">有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="98e32-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="98e32-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="98e32-107">**Add-in type:** Mail</span></span>

> [!IMPORTANT]
> <span data-ttu-id="98e32-108">基于事件的激活当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅在 Outlook 网页版中可用。</span><span class="sxs-lookup"><span data-stu-id="98e32-108">Event-based activation is currently [in preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) and only available in Outlook on the web.</span></span> <span data-ttu-id="98e32-109">有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。</span><span class="sxs-lookup"><span data-stu-id="98e32-109">For more information, see [How to preview the event-based activation feature](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).</span></span>

## <a name="syntax"></a><span data-ttu-id="98e32-110">语法</span><span class="sxs-lookup"><span data-stu-id="98e32-110">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="98e32-111">包含于</span><span class="sxs-lookup"><span data-stu-id="98e32-111">Contained in</span></span>

- [<span data-ttu-id="98e32-112">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="98e32-112">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="98e32-113">属性</span><span class="sxs-lookup"><span data-stu-id="98e32-113">Attributes</span></span>

|  <span data-ttu-id="98e32-114">属性</span><span class="sxs-lookup"><span data-stu-id="98e32-114">Attribute</span></span>  |  <span data-ttu-id="98e32-115">必需</span><span class="sxs-lookup"><span data-stu-id="98e32-115">Required</span></span>  |  <span data-ttu-id="98e32-116">说明</span><span class="sxs-lookup"><span data-stu-id="98e32-116">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="98e32-117">**Type**</span><span class="sxs-lookup"><span data-stu-id="98e32-117">**Type**</span></span>  |  <span data-ttu-id="98e32-118">是</span><span class="sxs-lookup"><span data-stu-id="98e32-118">Yes</span></span>  | <span data-ttu-id="98e32-119">指定受支持的事件类型。</span><span class="sxs-lookup"><span data-stu-id="98e32-119">Specifies a supported event type.</span></span> <span data-ttu-id="98e32-120">可用的类型有 `OnNewMessageCompose` 和 `OnNewAppointmentOrganizer` 。</span><span class="sxs-lookup"><span data-stu-id="98e32-120">Available types are `OnNewMessageCompose` and `OnNewAppointmentOrganizer`.</span></span> |
|  <span data-ttu-id="98e32-121">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="98e32-121">**FunctionName**</span></span>  |  <span data-ttu-id="98e32-122">是</span><span class="sxs-lookup"><span data-stu-id="98e32-122">Yes</span></span>  | <span data-ttu-id="98e32-123">指定用于处理属性中指定的事件的 JavaScript 函数的名称 `Type` 。</span><span class="sxs-lookup"><span data-stu-id="98e32-123">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="98e32-124">另请参阅</span><span class="sxs-lookup"><span data-stu-id="98e32-124">See also</span></span>

- [<span data-ttu-id="98e32-125">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="98e32-125">LaunchEvents</span></span>](launchevents.md)
