---
title: 清单文件中的 MobileFormFactor 元素
description: MobileFormFactor 元素指定外接程序的移动外形设置。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 88ed8a351cdb2e52dab79c30315123ad33550500
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340125"
---
# <a name="mobileformfactor-element"></a>MobileFormFactor 元素

指定对移动外形规格的外接程序的设置。它包含移动外形规格的所有外接程序信息（**资源** 节点的信息除外）。

每个 **MobileFormFactor** 定义都包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。 有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。

在 VersionOverrides 架构 1.1 中定义了 **MobileFormFactor** 元素。包含  [VersionOverrides](versionoverrides.md) 元素的 `xsi:type` 属性值必须为 `VersionOverridesV1_1`。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

## <a name="child-elements"></a>子元素

| 元素                             | 必需 | 说明  |
|:------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | 是      | 定义外接程序公开功能的位置。 |
| [FunctionFile](functionfile.md)     | 是      | 包含 JavaScript 函数的文件的 URL。|

## <a name="mobileformfactor-example"></a>MobileFormFactor 示例

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
