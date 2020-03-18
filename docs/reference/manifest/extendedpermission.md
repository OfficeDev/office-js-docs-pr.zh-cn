---
title: 清单文件中的 ExtendedPermission 元素
description: 定义外接程序访问关联的 API 或功能所需的扩展权限。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 7ff17312ae487d20f4d7af0ed4405cedd8820253
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720601"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission`网元

定义外接程序访问关联的 API 或功能所需的扩展权限。 `ExtendedPermission`元素是[ExtendedPermissions](extendedpermissions.md)的子元素。

> [!IMPORTANT]
> 此元素仅适用于针对 Exchange Online 的[Outlook 外接程序预览要求集](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)。 使用此元素的外接程序无法发布到 AppSource 或通过集中部署进行部署。

## <a name="available-extended-permissions"></a>可用扩展权限

以下是可用的值。

|可用值|说明|主机|
|---|---|---|
|`AppendOnSend`|声明外接程序使用的是[appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API。|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission`示例

以下是`ExtendedPermission`元素的示例。

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

[ExtendedPermissions](extendedpermissions.md)
