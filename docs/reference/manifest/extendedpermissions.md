---
title: 清单文件中的 ExtendedPermissions 元素
description: 定义加载项访问关联的 Api 或功能所需的扩展权限的集合。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 86d898052af6ba0e6f6bc8b341fff9f0f8408967
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718221"
---
# <a name="extendedpermissions-element"></a>ExtendedPermissions 元素

定义加载项访问关联的 Api 或功能所需的扩展权限的集合。 `ExtendedPermissions`元素是[VersionOverrides](versionoverrides.md)的子元素。

> [!IMPORTANT]
> 此元素仅适用于针对 Exchange Online 的[Outlook 外接程序预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。 使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  否   | 定义外接程序访问关联的 API 或功能所需的扩展权限。 |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions`示例

以下是`ExtendedPermissions`元素的示例。

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
