---
title: 使用单一登录令牌对用户进行身份验证
description: 了解如何使用 Outlook 外接程序提供的单一登录令牌为服务实现 SSO。
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 23b7936cc0ba4453a2a10cbfe0731941a913c118
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607441"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>在 Outlook 加载项中使用单一登录令牌对用户进行身份验证

使用单一登录 (SSO)，加载项可以无缝方式验证用户（并根据需要获取访问令牌来调用 [Microsoft Graph API](/graph/overview)）。

借助此方法，加载项可以获取范围限定为服务器后端 API 的访问令牌。 加载项将此令牌用作 `Authorization` 标头中的持有者令牌，来对 API 回调进行身份验证。 （可选）还可以使用服务器端代码。

- 完成“代表”流来获取作用域为 Microsoft Graph API 的访问令牌
- 使用令牌中的标识信息，以创建用户标识并验证自己的后端服务

有关 Office 外接程序中的 SSO 的概述，请参阅[为 Office 外接程序启用单一登录](../develop/sso-in-office-add-ins.md)和[在 Office 外接程序中授予对 Microsoft Graph 的访问权限](../develop/authorize-to-microsoft-graph.md)。

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>在 Microsoft 365 租户中启用新式身份验证

若要将 SSO 与 Outlook 加载项配合使用，必须为 Microsoft 365 租户启用新式身份验证。 若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。

## <a name="register-your-add-in"></a>注册外接程序

若要使用 SSO，Outlook 外接程序需要有已向 Azure Active Directory (AAD) v2.0 注册的服务器端 Web API。 有关详细信息，请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 外接程序](../develop/register-sso-add-in-aad-v2.md)。

### <a name="provide-consent-when-sideloading-an-add-in"></a>旁加载加载项时授予许可

开发加载项时，必须事先提供许可。 有关详细信息，请参阅 [向加载项授予管理员许可](../develop/grant-admin-consent-to-an-add-in.md)。

## <a name="update-the-add-in-manifest"></a>更新加载项清单

在加载项中启用 SSO 的下一步是从外接程序的Microsoft 标识平台注册中向清单添加一些信息。 标记因清单类型而异。

- **XML 清单**：在 [VersionOverrides](/javascript/api/manifest/versionoverrides) 元素的`VersionOverridesV1_1`末尾添加`WebApplicationInfo`元素。 然后，添加其所需的子元素。 有关标记的详细信息，请参阅 [配置加载项](../develop/sso-in-office-add-ins.md#configure-the-add-in)。
- **Teams 清单 (预览)**：向清单中的根 `{ ... }` 对象添加“webApplicationInfo”属性。 向此对象提供一个子“id”属性，该属性设置为加载项的 Web 应用的应用程序 ID，因为它是在注册加载项时在Azure 门户中生成的。  (请参阅本文前面的“ [注册加载](#register-your-add-in) 项”部分。) 还为其提供一个子“resource”属性，该属性设置为注册加载项时设置的同一应用程序 **ID URI** 。 此 URI 应具有表单 `api://<fully-qualified-domain-name>/<application-id>`。 示例如下。

   ```json
   "webApplicationInfo": {
        "id": "a661fed9-f33d-4e95-b6cf-624a34a2f51d",
        "resource": "api://addin.contoso.com/a661fed9-f33d-4e95-b6cf-624a34a2f51d"
    },
   ```

  > [!NOTE]
  > 可以旁加载使用 Teams 清单的已启用 SSO 的加载项，但目前无法以任何其他方式部署。

## <a name="get-the-sso-token"></a>获取 SSO 令牌

加载项使用客户端脚本获取 SSO 令牌。 有关详细信息，请参阅[添加客户端代码](../develop/sso-in-office-add-ins.md#add-client-side-code)。

## <a name="use-the-sso-token-at-the-back-end"></a>在后端使用 SSO 令牌

大多数情况下，如果加载项没有将访问令牌传递到服务器端并在其中使用它，那么获取访问令牌的意义就不大。 若要详细了解服务器端可以和应该执行的操作，请参阅[添加服务器端代码](../develop/sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code)。

> [!IMPORTANT]
> 若要将 SSO 令牌用作 *Outlook* 加载项中的标识，建议还 [使用 Exchange 标识令牌](authenticate-a-user-with-an-identity-token.md)作为备用标识。 加载项用户可能使用多个客户端，而有些客户端可能不支持提供 SSO 令牌。 通过将 Exchange 标识令牌用作备用令牌，就不用多次提示这些用户输入凭据了。 有关详细信息，请参阅[应用场景：在 Outlook 外接程序中对服务实现单一登录](implement-sso-in-outlook-add-in.md)。

## <a name="sso-for-event-based-activation"></a>基于事件的激活的 SSO

如果加载项使用基于事件的激活，则需要执行其他步骤。 有关详细信息，请参阅 [使用基于事件的激活的 Outlook 外接程序中启用单一登录 (SSO) ](use-sso-in-event-based-activation.md)。

## <a name="see-also"></a>另请参阅

- [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1))
- 有关使用 SSO 令牌访问 Microsoft 图形 API的示例 Outlook 外接程序，请参阅 [Outlook 外接程序 SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)。
- [SSO API 参考](/javascript/api/office/office.auth#office-office-auth-getaccesstoken-member(1))
- [IdentityAPI 要求集](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [在使用基于事件的激活的 Outlook 加载项中启用单一登录 (SSO) ](use-sso-in-event-based-activation.md)
