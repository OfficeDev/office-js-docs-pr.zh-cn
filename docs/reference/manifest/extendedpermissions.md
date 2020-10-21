---
title: 清单文件中的 ExtendedPermissions 元素
description: 定义加载项访问关联的 Api 或功能所需的扩展权限的集合。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626440"
---
# <a name="extendedpermissions-element"></a>ExtendedPermissions 元素

定义加载项访问关联的 Api 或功能所需的扩展权限的集合。 `ExtendedPermissions`元素是[VersionOverrides](versionoverrides.md)的子元素。

> [!IMPORTANT]
> 对此元素的支持是在要求集1.9 中引入的。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  否   | 定义外接程序访问关联的 API 或功能所需的扩展权限。 |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` 示例

以下是元素的示例 `ExtendedPermissions` 。

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
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a>包含于

[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>可以包含

[ExtendedPermission](extendedpermission.md)
