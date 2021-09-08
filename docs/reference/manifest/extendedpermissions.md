---
title: 清单文件中 ExtendedPermissions 元素
description: 定义外接程序访问关联 API 或功能所需的扩展权限的集合。
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938848"
---
# <a name="extendedpermissions-element"></a>ExtendedPermissions 元素

定义外接程序访问关联 API 或功能所需的扩展权限的集合。 元素 `ExtendedPermissions` 是 [VersionOverrides 的子元素](versionoverrides.md)。

> [!IMPORTANT]
> 要求集 1.9 中引入了对此元素的支持。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  否   | 定义外接程序访问关联的 API 或功能所需的扩展权限。 |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` 示例

下面是 元素 `ExtendedPermissions` 的一个示例。

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
