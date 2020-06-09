---
title: 清单文件中的 DesktopFormFactor 元素
description: 指定对桌面外形规格的外接程序的设置。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 46de234f2d97a9e6c7645c17a0f0a61d0c3e1a80
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612281"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 元素

指定对桌面外形规格的外接程序的设置。 桌面外形规格包括 web、Windows 和 Mac 上的 Office。 除了 "**资源**" 节点外，它还包含桌面外形规格的所有外接程序信息。

每个 DesktopFormFactor 定义都包含**FunctionFile**元素和一个或多个**ExtensionPoint**元素。 有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。

## <a name="child-elements"></a>子元素

| 元素                               | 必需 | Description  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | 是      | 定义外接程序公开功能的位置。 |
| [FunctionFile](functionfile.md)       | 是      | 包含 JavaScript 函数的文件的 URL。|
| [GetStarted](getstarted.md)           | 否       | 定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。 |
| [SupportsSharedFolders](supportssharedfolders.md) | 否 | 定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。 |

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
