---
title: 清单文件中的 SupportsSharedFolders 元素
description: SupportsSharedFolders 元素定义Outlook外接程序在共享文件夹和共享邮箱方案中是否可用。
ms.date: 06/15/2021
ms.localizationpriority: medium
ms.openlocfilehash: f6a7cbe1c6549c8a93a6ecceab9e4bdaba07001f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152505"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 元素

定义Outlook外接程序是否可用于共享邮箱 (预览) 和共享文件夹 (即委派访问权限) 方案。 **SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。 默认情况下，此元素设置为 *false*。

> [!IMPORTANT]
> 要求集 1.8 中引入了对此元素的支持。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

下面是 **SupportsSharedFolders 元素的一** 个示例。

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
