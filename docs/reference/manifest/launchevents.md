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
# <a name="launchevents-element-preview"></a>LaunchEvents 元素（预览）

将你的外接程序配置为根据受支持的事件进行激活。 元素的子 [`<ExtensionPoint>`](extensionpoint.md) 元素。 有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。

**外接程序类型：** 邮件

> [!IMPORTANT]
> 基于事件的激活当前[处于预览阶段](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)，仅在 Outlook 网页版中可用。 有关详细信息，请参阅[如何预览基于事件的激活功能](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)。

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

[ExtensionPoint](extensionpoint.md) （**LaunchEvent**邮件外接端）

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  Description  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | 是 |  将支持的事件映射到其在外接程序激活的 JavaScript 文件中的功能。 |

## <a name="see-also"></a>另请参阅

- [LaunchEvent](launchevent.md)
