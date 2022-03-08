---
title: 清单文件中的 WebApplicationInfo 元素
description: WebApplicationInfo 元素的参考文档Office外接程序清单 (XML) 文件。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa74c4fc19d060f92c8c0ac2fe723c42f6ad9cdd
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340657"
---
# <a name="webapplicationinfo-element"></a>WebApplicationInfo 元素

支持 Office 外接程序中的单一登录 (SSO)。此元素包含外接程序中的信息，如下所示：

- OAuth 2.0 *资源*，Office应用程序可能需要权限。
- OAuth 2.0 *客户端*，可能需要访问 Microsoft Graph 的权限。

**外接程序类型：** 任务窗格、邮件、内容

**仅在以下 VersionOverrides 架构中有效**：

- 任务窗格 1.0
- 内容 1.0
- 邮件 1.0
- 邮件 1.1

有关详细信息，请参阅清单 [中的版本替代](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**与以下要求集相关联**：

- [IdentityAPI 1.3](../requirement-sets/identity-api-requirement-sets.md)

> [!NOTE]
> Word、Excel、Outlook 和 PowerPoint 目前支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../requirement-sets/identity-api-requirement-sets.md)。 如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。 要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

**WebApplicationInfo** 是清单中的 [VersionOverrides](versionoverrides.md) 元素的子元素。  

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Id**    |  是   |  在 Azure Active Directory v2.0 终结点中注册的加载项关联服务的 **应用程序 ID**。|
|  **Resource**  |  是   |  指定在 Azure Active Directory v2.0 终结点中注册的加载项的 **应用程序 ID URI**。|
|  [Scopes](scopes.md)                |  是  |  指定加载项对资源（如 Microsoft 加载项）所需的Graph。  |

## <a name="webapplicationinfo-example"></a>WebApplicationInfo 示例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://contoso.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
