---
title: 清单文件中的 LaunchEvent （预览）
description: LaunchEvent 元素将你的外接程序配置为根据受支持的事件进行激活。
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 4874b9f4c14e3a999f41ec3fa20a15393b031ea6
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611776"
---
# <a name="launchevent-element-preview"></a>LaunchEvent 元素（预览）

将你的外接程序配置为根据受支持的事件进行激活。 元素的子 [`<LaunchEvents>`](launchevents.md) 元素。 有关详细信息，请参阅[Configure Outlook 外接程序以进行基于事件的激活](../../outlook/autolaunch.md)。

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>属性

|  属性  |  必需  |  说明  |
|:-----|:-----|:-----|
|  **类型**  |  是  | 指定受支持的事件类型。 可用的类型有 `OnNewMessageCompose` 和 `OnNewAppointmentOrganizer` 。 |
|  **FunctionName**  |  是  | 指定用于处理属性中指定的事件的 JavaScript 函数的名称 `Type` 。 |

## <a name="see-also"></a>另请参阅

- [LaunchEvents](launchevents.md)
