---
title: 清单文件中的 DesktopFormFactor 元素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: d1f09203518a38f1568b13e6c1a9c70752697152
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128515"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 元素

指定对桌面外形规格的外接程序的设置。 桌面外形规格包括 web、Windows 和 Mac 上的 Office。 它包含该外形规格的所有外接程序信息（**资源**节点的信息除外）。

每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。

## <a name="child-elements"></a>子元素

| 元素                               | 必需 | 说明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | 是      | 定义外接程序公开功能的位置。 |
| [FunctionFile](functionfile.md)       | 是      | 包含 JavaScript 函数的文件的 URL。|
| [GetStarted](getstarted.md)           | 否       | 定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。 |
| [SupportsSharedFolders](supportssharedfolders.md) | 否 | 定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。<br><br>**重要说明**: 由于 Outlook 外接程序的代理访问当前处于预览阶段, 使用`SupportSharedFolders`元素的外接程序不能发布到 AppSource 或通过集中部署进行部署。 |

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
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
