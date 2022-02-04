---
title: 清单文件中 ExtendedPermission 元素
description: 定义加载项访问关联 API 或功能所需的扩展权限。
ms.date: 01/04/2022
ms.localizationpriority: medium
---

# <a name="extendedpermission-element"></a>`ExtendedPermission` 元素

定义加载项访问关联 API 或功能所需的扩展权限。 元素 `ExtendedPermission` 是 [ExtendedPermissions 的子元素](extendedpermissions.md)。

> [!IMPORTANT]
> 要求集 1.9 中引入了对此元素的支持。 请查看支持此要求集的[客户端和平台](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)。

**外接程序类型：** 邮件

**仅在以下 VersionOverrides 架构中有效**：

- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

## <a name="available-extended-permissions"></a>可用的扩展权限

以下是可用值。

|可用值|说明|Hosts|
|---|---|---|
|`AppendOnSend`|声明外接程序正在使用 Office[。Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-appendonsendasync-member(1)) API。|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` 示例

下面是 元素的一 `ExtendedPermission` 个示例。

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
