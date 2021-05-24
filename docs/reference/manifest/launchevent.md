---
title: 清单文件中 LaunchEvent
description: LaunchEvent 元素将外接程序配置为基于支持的事件激活。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: c866a085ed6b7a33c8d7bf02d25e6ec748629e07
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591077"
---
# <a name="launchevent-element"></a><span data-ttu-id="2ec83-103">LaunchEvent 元素</span><span class="sxs-lookup"><span data-stu-id="2ec83-103">LaunchEvent element</span></span>

<span data-ttu-id="2ec83-104">将加载项配置为基于支持的事件进行激活。</span><span class="sxs-lookup"><span data-stu-id="2ec83-104">Configures your add-in to activate based on supported events.</span></span> <span data-ttu-id="2ec83-105">元素的 [`<LaunchEvents>`](launchevents.md) 子元素。</span><span class="sxs-lookup"><span data-stu-id="2ec83-105">Child of the [`<LaunchEvents>`](launchevents.md) element.</span></span> <span data-ttu-id="2ec83-106">有关详细信息，请参阅[Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="2ec83-106">For more information, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>

<span data-ttu-id="2ec83-107">**外接程序类型：** 邮件</span><span class="sxs-lookup"><span data-stu-id="2ec83-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2ec83-108">语法</span><span class="sxs-lookup"><span data-stu-id="2ec83-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="2ec83-109">包含于</span><span class="sxs-lookup"><span data-stu-id="2ec83-109">Contained in</span></span>

- [<span data-ttu-id="2ec83-110">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="2ec83-110">LaunchEvents</span></span>](launchevents.md)

## <a name="attributes"></a><span data-ttu-id="2ec83-111">属性</span><span class="sxs-lookup"><span data-stu-id="2ec83-111">Attributes</span></span>

|  <span data-ttu-id="2ec83-112">属性</span><span class="sxs-lookup"><span data-stu-id="2ec83-112">Attribute</span></span>  |  <span data-ttu-id="2ec83-113">必需</span><span class="sxs-lookup"><span data-stu-id="2ec83-113">Required</span></span>  |  <span data-ttu-id="2ec83-114">说明</span><span class="sxs-lookup"><span data-stu-id="2ec83-114">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2ec83-115">**类型**</span><span class="sxs-lookup"><span data-stu-id="2ec83-115">**Type**</span></span>  |  <span data-ttu-id="2ec83-116">是</span><span class="sxs-lookup"><span data-stu-id="2ec83-116">Yes</span></span>  | <span data-ttu-id="2ec83-117">指定支持的事件类型。</span><span class="sxs-lookup"><span data-stu-id="2ec83-117">Specifies a supported event type.</span></span> <span data-ttu-id="2ec83-118">有关受支持的类型集，请参阅[配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md#supported-events)。</span><span class="sxs-lookup"><span data-stu-id="2ec83-118">For the set of supported types, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md#supported-events).</span></span> |
|  <span data-ttu-id="2ec83-119">**FunctionName**</span><span class="sxs-lookup"><span data-stu-id="2ec83-119">**FunctionName**</span></span>  |  <span data-ttu-id="2ec83-120">是</span><span class="sxs-lookup"><span data-stu-id="2ec83-120">Yes</span></span>  | <span data-ttu-id="2ec83-121">指定要处理属性中指定的事件的 JavaScript 函数 `Type` 的名称。</span><span class="sxs-lookup"><span data-stu-id="2ec83-121">Specifies the name of the JavaScript function to handle the event specified in the `Type` attribute.</span></span> |

## <a name="see-also"></a><span data-ttu-id="2ec83-122">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2ec83-122">See also</span></span>

- [<span data-ttu-id="2ec83-123">LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="2ec83-123">LaunchEvents</span></span>](launchevents.md)
