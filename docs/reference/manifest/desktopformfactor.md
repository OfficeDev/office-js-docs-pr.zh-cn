---
title: 清单文件中的 DesktopFormFactor 元素
description: 指定对桌面外形规格的外接程序的设置。
ms.date: 06/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: f89dff5626867258c8df93d5f047e3d08103e71b
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149569"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 元素

指定对桌面外形规格的外接程序的设置。 桌面设备包括 Office web 版、Windows 和 Mac。 它包含桌面设备类型的所有外接程序信息，"资源"节点 **除外** 。

每个 DesktopFormFactor 定义都包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。 有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。

## <a name="child-elements"></a>子元素

| 元素                               | 必需 | 说明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | 是      | 定义外接程序公开功能的位置。 |
| [FunctionFile](functionfile.md)       | 是      | 包含 JavaScript 函数的文件的 URL。|
| [GetStarted](getstarted.md)           | 否       | 定义在 Word、加载项或加载项中安装加载项时Excel标注PowerPoint。 |
| [SupportsSharedFolders](supportssharedfolders.md) | 否 | 定义Outlook外接程序是否可用于共享邮箱 (预览) 和共享文件夹 (即委派访问权限) 方案。 默认情况下设置为 *false。* |

## <a name="desktopformfactor-example"></a>DesktopFormFactor 示例

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
