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
# <a name="launchevent-element"></a>LaunchEvent 元素

将加载项配置为基于支持的事件进行激活。 元素的 [`<LaunchEvents>`](launchevents.md) 子元素。 有关详细信息，请参阅[Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)。

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **类型**  |  是  | 指定支持的事件类型。 有关受支持的类型集，请参阅[配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md#supported-events)。 |
|  **FunctionName**  |  是  | 指定要处理属性中指定的事件的 JavaScript 函数 `Type` 的名称。 |

## <a name="see-also"></a>另请参阅

- [LaunchEvents](launchevents.md)
