---
title: 清单文件中 LaunchEvent
description: LaunchEvent 元素将你的外接程序配置为基于支持的事件进行激活。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: a8ab75633d87284e02e9db9b1a71f7a8436f7daf
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681707"
---
# <a name="launchevent-element"></a>LaunchEvent 元素

配置加载项以根据支持的事件激活。 元素的 [`<LaunchEvents>`](launchevents.md) 子元素。 有关详细信息，请参阅[Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md)。

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
|  **类型**  |  是  | 指定支持的事件类型。 有关受支持的类型集，请参阅配置Outlook[加载项进行基于事件的激活](../../outlook/autolaunch.md#supported-events)。 |
|  **FunctionName**  |  是  | 指定用于处理属性中指定的事件的 JavaScript 函数 `Type` 的名称。 |
|  **SendMode** (预览)  |  否  | 和 `OnMessageSend` 事件 `OnAppointmentSend` 是必需的。 指定外接程序停止发送项目时可供用户使用的选项。 有关可用选项，请参阅 [可用 SendMode 选项](#available-sendmode-options-preview)。 |

## <a name="available-sendmode-options-preview"></a>预览版中可用的 SendMode () 

在清单中包括 `OnMessageSend` 或 `OnAppointmentSend` 事件时，还必须设置 **SendMode** 属性。 以下是可用选项。 根据加载项查找的条件，如果加载项在正在发送的项中发现问题，用户会发出警报。

| SendMode 选项 | 说明 |
|---|---|
|`PromptUser`|在警报中，用户可以选择"继续发送"，或解决问题，然后再次尝试发送该项目。|
|`SoftBlock`|用户必须先修复此问题，然后才能尝试再次发送该项目。|

## <a name="see-also"></a>另请参阅

- [LaunchEvents](launchevents.md)
- [配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md#supported-events)
- [在加载项中使用智能警报Outlook OnMessageSend 事件](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
