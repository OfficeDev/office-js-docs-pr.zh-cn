---
title: 清单文件中 LaunchEvent
description: LaunchEvent 元素将你的外接程序配置为基于支持的事件进行激活。
ms.date: 02/02/2022
ms.localizationpriority: medium
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
|  **SendMode** (预览)  |  否  | 和 事件`OnMessageSend``OnAppointmentSend`是必需的。 指定外接程序停止发送项目时可供用户使用的选项。 有关可用选项，请参阅 [可用 SendMode 选项](#available-sendmode-options-preview)。 |

## <a name="available-sendmode-options-preview"></a>可用的 SendMode 选项 (预览) 

在清单中包括 `OnMessageSend` 或 `OnAppointmentSend` 事件时，还必须设置 **SendMode** 属性。 以下是可用选项。 根据加载项查找的条件，如果加载项在正在发送的项中发现问题，用户会发出警报。

| SendMode 选项 | 说明 |
|---|---|
|`PromptUser`|In the alert， the user can choose to **Send Anyway**， or address the issue then try to send the item again.|
|`SoftBlock`|用户必须先修复此问题，然后才能尝试再次发送该项目。|

## <a name="see-also"></a>另请参阅

- [LaunchEvents](launchevents.md)
- [配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md#supported-events)
- [在加载项中使用智能警报Outlook OnMessageSend 事件](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
