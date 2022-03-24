---
title: 清单文件中 LaunchEvent
description: LaunchEvent 元素将你的外接程序配置为基于支持的事件进行激活。
ms.date: 03/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 71469693bff7213455582a3247778cabf92c2aa3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745808"
---
# <a name="launchevent-element"></a>LaunchEvent 元素

配置加载项以根据支持的事件激活。 元素的 [`<LaunchEvents>`](launchevents.md) 子元素。 有关详细信息，请参阅[配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md)。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

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
|  **类型**  |  是  | 指定支持的事件类型。 有关受支持的类型集，请参阅配置[Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md#supported-events)。 |
|  **FunctionName**  |  是  | 指定要处理属性中指定的事件的 JavaScript 函数 `Type` 的名称。 |
|  **SendMode** (预览)  |  否  | 由 和 `OnMessageSend` `OnAppointmentSend` 事件使用。 指定您的外接程序停止发送项目或外接程序不可用时可供用户使用的选项。 如果 **不包括 SendMode** 属性，则 `SoftBlock` 默认情况下设置该选项。 有关可用选项，请参阅 [可用 SendMode 选项](#available-sendmode-options-preview)。 |

## <a name="available-sendmode-options-preview"></a>可用的 SendMode 选项 (预览) 

在清单中包括 `OnMessageSend` 或 `OnAppointmentSend` 事件时，还应设置 **SendMode** 属性。 如果 **不包括 SendMode** 属性，则 `SoftBlock` 默认情况下设置该选项。 以下是可用选项。 根据加载项查找的条件，如果加载项在正在发送的项中发现问题，用户会发出警报。

| SendMode 选项 | 说明 |
|---|---|
|`PromptUser`|如果项目不满足加载项的条件，用户可以在通知中选择"继续发送"，或解决问题，然后再次尝试发送该项目。 如果加载项处理项目的时间很长，系统将提示用户选择停止运行加载项，然后选择"继续 **发送"**。 例如，如果加载项在加载项 (，加载加载项时出错) ，将发送项目。|
|`SoftBlock`|如果未包括 **SendMode** 属性，则默认选项。 用户收到通知，提醒他们发送的项目不符合外接程序的条件，他们必须在尝试再次发送项目之前解决该问题。 但是，如果外接程序不可用 (例如，加载外接程序时) ，该项目将发送。|
|`Block`|如果发生以下任一情况，不发送该项目。<br>- 项不符合加载项的条件。<br>- 加载项无法连接到服务器。<br>- 加载加载项时出错。|

## <a name="see-also"></a>另请参阅

- [LaunchEvents](launchevents.md)
- [配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md#supported-events)
- [在加载项中使用智能警报Outlook OnMessageSend 事件](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
