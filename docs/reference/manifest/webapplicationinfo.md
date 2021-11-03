---
title: 清单文件中的 WebApplicationInfo 元素
description: 外接程序清单的 WebApplicationInfo 元素的参考Office XML (XML) 文档。
ms.date: 10/25/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb21c584f516fc9e50bdd881a383fb03f01c753c
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681547"
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
