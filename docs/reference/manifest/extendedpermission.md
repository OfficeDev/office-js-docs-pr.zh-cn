---
title: 清单文件中 ExtendedPermission 元素
description: 定义加载项访问关联 API 或功能所需的扩展权限。
ms.date: 10/15/2020
ms.localizationpriority: medium
ms.openlocfilehash: 3655ff43088e4506c8bf4125a4a9fae1ef721b4e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152304"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` 元素

定义加载项访问关联 API 或功能所需的扩展权限。 元素 `ExtendedPermission` 是 [ExtendedPermissions 的子元素](extendedpermissions.md)。

> [!IMPORTANT]
> 要求集 1.9 中引入了对此元素的支持。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

## <a name="available-extended-permissions"></a>可用的扩展权限

以下是可用值。

|可用值|描述|Hosts|
|---|---|---|
|`AppendOnSend`|声明外接程序正在使用[Office。Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendOnSendAsync_data__options__callback_) API。|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` 示例

下面是 元素 `ExtendedPermission` 的一个示例。

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
