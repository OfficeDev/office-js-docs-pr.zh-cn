---
title: 清单文件中的 SupportsSharedFolders 元素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 776d44ec66c4e27a72e5487051bed1edf4b3dcaf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432681"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 元素

定义 Outlook 加载项在代理应用场景中是否可用。 **SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。 默认情况下，此元素设置为 *false*。

> [!IMPORTANT]
> 此元素仅适用于针对 Exchange Online 的 [Outlook 加载项预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。 使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。

以下是 **SupportsSharedFolders** 元素的示例。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
