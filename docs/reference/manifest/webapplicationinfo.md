---
title: 清单文件中的 WebApplicationInfo 元素
description: 外接程序清单的 WebApplicationInfo 元素的参考Office XML (XML) 文档。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 037de49320a6d1a1ca7dce3446b4f4008a2f1331
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937087"
---
# <a name="webapplicationinfo-element"></a>WebApplicationInfo 元素

支持 Office 外接程序中的单一登录 (SSO)。此元素包含外接程序中的信息，如下所示：

- OAuth 2.0 *资源*，Office应用程序可能需要权限。
- OAuth 2.0 *客户端*，可能需要访问 Microsoft Graph 的权限。

> [!NOTE]
> Word、Excel、Outlook 和 PowerPoint 目前支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../requirement-sets/identity-api-requirement-sets.md)。 如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。 要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

**WebApplicationInfo** 是清单中的 [VersionOverrides](versionoverrides.md) 元素的子元素。  

## <a name="child-elements"></a>子元素

|  元素 |  必需  |  说明  |
|:-----|:-----|:-----|
|  **Id**    |  是   |  在 Azure Active Directory v2.0 终结点中注册的加载项关联服务的 **应用程序 ID**。|
|  **MsaId**    |  否   |  外接程序 Web 应用程序的 MSA 客户端 ID，如 msm.live.com 中注册。|
|  **Resource**  |  是   |  指定在 Azure Active Directory v2.0 终结点中注册的加载项的 **应用程序 ID URI**。|
|  [Scopes](scopes.md)                |  是  |  指定外接程序对资源（如 Microsoft Graph）所需的权限。  |
|  [授权](authorizations.md)  |  否   | 指定加载项的 Web 应用程序需要授权的外部资源和所需权限。|

## <a name="webapplicationinfo-example"></a>WebApplicationInfo 示例

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    ...
    <WebApplicationInfo>
      <Id>12345678-abcd-1234-efab-123456789abc</Id>
      <Resource>api://myDomain.com/12345678-abcd-1234-efab-123456789abc</Resource>
      <Scopes>
        <Scope>Files.Read.All</Scope>
        <Scope>offline_access</Scope>
        <Scope>openid</Scope>
        <Scope>profile</Scope>
      </Scopes>
      <Authorizations>
        <Authorization>
          <Resource>https://api.contoso.com</Resource>
            <Scopes>
              <Scope>profile</Scope>
          </Scopes>
        </Authorization>
      </Authorizations>
    </WebApplicationInfo>
  </VersionOverrides>
...
</OfficeApp>
```
