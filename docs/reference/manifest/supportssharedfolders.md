---
title: 清单文件中的 SupportsSharedFolders 元素
description: SupportsSharedFolders 元素定义 Outlook 加载项在委托方案中是否可用。
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 3835f7060cc52a72ff0a5ed4dbdb9f1e09258669
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608710"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 元素

定义 Outlook 加载项在代理应用场景中是否可用。 **SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。 默认情况下，此元素设置为 *false*。

> [!IMPORTANT]
> 只有 Outlook 网页和 Windows 支持**SupportsSharedFolders**元素。
>
> 对此元素的支持是在要求集1.8 中引入的。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

下面是**SupportsSharedFolders**元素的一个示例。

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
