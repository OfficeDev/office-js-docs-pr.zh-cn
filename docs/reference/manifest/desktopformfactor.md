---
title: 清单文件中的 DesktopFormFactor 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433738"
---
# <a name="desktopformfactor-element"></a>DesktopFormFactor 元素

指定对桌面外形规格的外接程序的设置。桌面外形规格包括 Office for Windows、Office for Mac 和 Office Online。它包含该外形规格的所有外接程序信息（**资源**节点的信息除外）。

每个 DesktopFormFactor 定义均包含 **FunctionFile** 元素和一个或多个 **ExtensionPoint** 元素。有关详细信息，请参阅 [FunctionFile 元素](functionfile.md) 和 [ExtensionPoint 元素](extensionpoint.md)。

## <a name="child-elements"></a>子元素

| 元素                               | 必需 | 说明  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md)   | 是      | 定义外接程序公开功能的位置。 |
| [FunctionFile](functionfile.md)       | 是      | 包含 JavaScript 函数的文件的 URL。|
| [GetStarted](getstarted.md)           | 否       | 定义在 Word、Excel 或 PowerPoint 主机中安装外接程序时将显示的标注。 |
| [SupportsSharedFolders](supportssharedfolders.md) | 否 | 定义 Outlook 外接程序在代理应用场景中是否可用，默认设置为 *false*。<br><br>**重要说明**：此元素仅适用于针对 Exchange Online 的 Outlook 外接程序预览要求集。 使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。 |

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
