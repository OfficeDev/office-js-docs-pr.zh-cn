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
# <a name="launchevents-element-preview"></a>LaunchEvents 元素 (预览) 

配置加载项以根据支持的事件激活。 元素的 [`<ExtensionPoint>`](extensionpoint.md) 子元素。 有关详细信息，请参阅配置 [Outlook 外接程序进行基于事件的激活](../../outlook/autolaunch.md)。

**外接程序类型：** 邮件

> [!IMPORTANT]
> 基于事件的激活目前处于 [预览阶段，](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 仅在 Outlook 网页版和 Windows 版中可用。 有关详细信息，请参阅 [如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

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
| [LaunchEvent](launchevent.md) | 是 |  将支持的事件映射到 JavaScript 文件中用于加载项激活的函数。 |

## <a name="see-also"></a>另请参阅

- [LaunchEvent](launchevent.md)
