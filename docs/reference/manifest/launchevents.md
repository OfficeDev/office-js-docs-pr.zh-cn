---
title: 清单文件中 LaunchEvents
description: LaunchEvents 元素将外接程序配置为基于支持的事件进行激活。
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="launchevents-element"></a>LaunchEvents 元素

配置加载项以根据支持的事件激活。 元素的 [`<ExtensionPoint>`](extensionpoint.md) 子元素。 有关详细信息，请参阅[配置Outlook加载项进行基于事件的激活](../../outlook/autolaunch.md)。

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

[ExtensionPoint](extensionpoint.md) (**LaunchEvent** 邮件外接程序) 

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | 是 |  将受支持的事件映射到 JavaScript 文件中用于外接程序激活的函数。 |

## <a name="see-also"></a>另请参阅

- [LaunchEvent](launchevent.md)
