---
title: 清单文件中的 SupportsSharedFolders 元素
description: SupportsSharedFolders 元素定义 Outlook外接程序在共享文件夹和共享邮箱方案中是否可用。
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e13393f10b12e0a3c5ca1b004b202eb2970d264
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138721"
---
# <a name="supportssharedfolders-element"></a>SupportsSharedFolders 元素

定义外接程序Outlook外接程序是否可用于共享邮箱 (预览) 和共享文件夹 (即委派访问权限) 方案。 **SupportsSharedFolders** 元素是 [DesktopFormFactor](desktopformfactor.md) 的子元素。 默认情况下，此元素设置为 *false*。

> [!IMPORTANT]
> 要求集 1.8 中引入了对此元素的支持。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [Mailbox 1.8](../../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)

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
