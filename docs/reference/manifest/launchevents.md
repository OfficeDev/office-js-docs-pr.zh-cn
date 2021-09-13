---
title: 清单文件中 LaunchEvents
description: LaunchEvents 元素将外接程序配置为基于支持的事件进行激活。
ms.date: 05/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 02e0b21d65733492a783ffb099caf9e76225e53f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149444"
---
# <a name="launchevents-element"></a>LaunchEvents 元素

配置加载项以根据支持的事件激活。 元素的 [`<ExtensionPoint>`](extensionpoint.md) 子元素。 有关详细信息，请参阅[Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)。

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
