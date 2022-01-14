---
title: 清单文件中 ExtendedPermissions 元素
description: 定义外接程序访问关联 API 或功能所需的扩展权限的集合。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 46ca6e3e2fb992755d9067b4251200073f07ade1
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042124"
---
# <a name="extendedpermissions-element"></a>ExtendedPermissions 元素

定义外接程序访问关联 API 或功能所需的扩展权限的集合。 元素 `ExtendedPermissions` 是 [VersionOverrides 的子元素](versionoverrides.md)。

> [!IMPORTANT]
> 要求集 1.9 中引入了对此元素的支持。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

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
