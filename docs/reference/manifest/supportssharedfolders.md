---
title: 清单文件中的 SupportsSharedFolders 元素
description: ''
ms.date: 03/01/2019
localization_priority: Normal
ms.openlocfilehash: bfbce42c7d1aa5eefab40b528c5b622aa7d2d54f
ms.sourcegitcommit: 7ebd383f16ae5809bb6980a5f213b695d410e62c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/06/2019
ms.locfileid: "30413613"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 元素

定义 Outlook 加载项在代理应用场景中是否可用。 **SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。 默认情况下，此元素设置为 *false*。

> [!IMPORTANT]
> Outlook 外接程序的委派访问权限当前处于预览阶段, 仅在对 Exchange Online 运行的客户端中受支持。 使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。

以下是 **SupportsSharedFolders** 元素的示例。

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="MessageReadCommandSurface">
    <!-- configure selected extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
