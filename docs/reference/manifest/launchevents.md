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
# <a name="launchevents-element"></a>LaunchEvents 元素

将加载项配置为基于支持的事件进行激活。 元素的 [`<ExtensionPoint>`](extensionpoint.md) 子元素。 有关详细信息，请参阅[Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)。

**外接程序类型：** 邮件

## <a name="syntax"></a>语法

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

## <a name="contained-in"></a>包含于

[ExtensionPoint](extensionpoint.md) (**LaunchEvent** 邮件外接程序) 

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | 是 |  将受支持的事件映射到 JavaScript 文件中用于外接程序激活的函数。 |

## <a name="see-also"></a>另请参阅

- [LaunchEvent](launchevent.md)
