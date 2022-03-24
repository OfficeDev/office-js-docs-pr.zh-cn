---
title: 使用 SSO 对 Microsoft Graph 授权
description: 了解加载项Office如何使用 SSO (单一登录) Microsoft Graph。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: dfdfda7ff01f07873da7bd5dd32a5878c29a88b1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743550"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>使用 SSO 对 Microsoft Graph 授权

用户可以使用自己的个人 Microsoft 帐户或 Microsoft 365 教育或工作帐户，登录 Office（在线、移动和桌面平台）。 在 Office 加载项中获取对 [Microsoft Graph](https://developer.microsoft.com/graph/docs) 的访问权限的最佳方式是使用用户的 Office 登录凭据。 这使用户能够访问其 Microsoft Graph 数据，而无需再次登录。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO 和 Microsoft Graph 的加载项体系结构

除了托管 Web 应用程序的页面和 JavaScript 之外，外接程序还必须以相同的[完全限定的域名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)托管一个或多个 Web API，这些 API 可获取 Microsoft Graph 的访问令牌，并向它发出请求。

外接程序清单包含 **一个 WebApplicationInfo** 元素，该元素向 Office 提供重要的 Azure 应用注册信息，包括外接程序所需的 Microsoft Graph 权限。

### <a name="how-it-works-at-runtime"></a>运行时的工作方式

下图显示了登录和访问 Microsoft Graph。 整个过程使用 OAuth 2.0 和 JWT 访问令牌。

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="显示 SSO 过程的图表。" border="false":::

1. 外接程序的客户端代码调用 Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1))。 这将告知Office主机获取外接程序的访问令牌。

    如果用户未登录，Office主机与 Microsoft 标识平台会为用户提供 UI 进行登录和同意。

2. 该Office主机从应用程序请求访问Microsoft 标识平台。
3. 该Microsoft 标识平台将访问令牌 *A* 返回到 Office 主机。 访问 *令牌 A* 仅提供对外接程序自己的服务器端 API 的访问。 它不提供对 Microsoft Graph。
4. Office主机将访问令牌 *A* 返回到外接程序的客户端代码。 现在，客户端代码可以调用经过身份验证的服务器端 API。
5. 客户端代码向需要身份验证的服务器端 Web API 发送 HTTP 请求。 它包括访问令牌 *A* 作为授权证明。 服务器端代码验证访问令牌 *A*。
6. 服务器端代码使用 OAuth 2.0 代表流 (OBO) 请求具有 Microsoft Graph 权限的新访问令牌。
7. 如果Microsoft 标识平台请求获取 Microsoft Graph (权限，则返回新的访问令牌 *B* 和刷新 *offline_access刷新)*。 服务器可以选择缓存访问令牌 *B*。
8. 服务器端代码向 Microsoft Graph API 提出请求，并包含对 Microsoft Graph 具有权限的访问令牌 *B*。
9. Microsoft Graph将数据返回给服务器端代码。
10. 服务器端代码将数据返回回客户端代码。

在后续请求中，客户端代码在向服务器端代码执行经过身份验证的调用时将始终传递访问令牌 *A* 。 服务器端代码可以缓存令牌 *B* ，因此它无需在将来的 API 调用时再次请求它。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>开发可访问 Microsoft Graph 的 SSO 加载项

您开发一个可以访问 Microsoft Graph的外接程序，就像使用 SSO 的其他任何应用程序一样。 有关完整的说明，请参阅为加载项启用Office[登录](../develop/sso-in-office-add-ins.md)。区别在于，加载项必须拥有服务器端 Web API。

根据所用的语言和框架，可能存在一些库，可简化必须编写的服务器端代码。 代码应执行以下操作：

* 每次从客户端代码传递访问令牌 A 时，验证令牌 *A* 。 有关详细信息，请参阅[验证访问令牌](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code)。
* 通过调用 Microsoft 标识平台 启动 OAuth 2.0 代表流 (OBO) ，其中包括访问令牌、有关用户的一些元数据以及外接程序的凭据 (其 ID 和密码) 。 有关 OBO 流详细信息，请参阅 Microsoft 标识平台 [和 OAuth 2.0 代表流](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)。
* （可选）在流完成后，缓存返回的访问令牌 *B*，并授予对 Microsoft Graph。 如果加载项对 Microsoft Graph 进行多次调用，则需要执行此操作。 有关详细信息，请参阅使用 MICROSOFT 身份验证库获取和缓存令牌 ([MSAL) ](/azure/active-directory/develop/msal-acquire-cache-tokens)
* 创建一个或多个 Web API 方法，用于将可能缓存Graph的 Web API (*访问令牌 B*) Microsoft Graph。

有关详细演练和应用场景的示例，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)
* [应用场景：为 Outlook 加载项中的服务实现单一登录](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>在 Microsoft AppSource 中分发支持 SSO 的加载项

当Microsoft 365从 [AppSource](https://appsource.microsoft.com) 获取加载项时，管理员可以通过集成应用重新分发它，并授予加载项管理员同意以访问 [](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) Microsoft Graph作用域。 但是，最终用户也可以直接从 AppSource 获取加载项，在这种情况下，用户必须同意加载项。 这可以创建一个潜在的性能问题，我们提供了一个解决方案。

`allowConsentPrompt` `getAccessToken``OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`如果代码在 调用 中传递 选项，例如 ，Office 如果 Microsoft 标识平台 向 Office报告尚未向加载项授予同意，Office 可以提示用户同意。 但是，出于安全Office，用户只能提示用户同意 Microsoft `profile` Graph作用域。 *Office无法提示同意其他 Microsoft Graph范围，* 甚至无法提示`User.Read`。 这意味着，如果用户对提示授予同意，Office返回访问令牌。 但是，尝试将访问令牌交换为具有其他 Microsoft Graph 作用域的新访问令牌失败，并出现错误 AADSTS65001，这意味着尚未授予 (同意 Microsoft Graph 作用域) 。

> [!NOTE]
> 即使管理员已关闭`{ allowConsentPrompt: true }``profile`最终用户同意，同意请求仍可能会失败，即使范围失败。 有关详细信息，请参阅 [Configure how end-users consent to applications using Azure Active Directory](/azure/active-directory/manage-apps/configure-user-consent)。

代码可以并且应该通过回滚到备用身份验证系统来处理此错误，这将提示用户同意 Microsoft Graph作用域。 有关代码示例，请参阅[](create-sso-office-add-ins-nodejs.md)创建Node.js Office单一登录的加载项和创建使用单一登录的 [ASP.NET Office](create-sso-office-add-ins-aspnet.md) 加载项及其链接到的示例。 整个过程需要多次往返于Microsoft 标识平台。 若要避免这种性能损失，在 `forMSGraphAccess` `getAccessToken`调用中包括 选项;例如， `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`。 这表示Office加载项需要 Microsoft Graph作用域。 Office请求Microsoft 标识平台验证是否Graph向加载项授予了对 Microsoft 许可范围。 如果已返回，则返回访问令牌。 如果没有，则 `getAccessToken` 调用 将返回错误 13012。 代码可以通过立即回滚到备用身份验证系统来处理此错误，而无需尝试与代理交换Microsoft 标识平台。

最佳做法是，始终传递到`forMSGraphAccess``getAccessToken`加载项在 AppSource 中分发并需要 Microsoft Graph范围。

## <a name="details-on-sso-with-an-outlook-add-in"></a>有关使用加载项Outlook SSO 的详细信息

如果您开发使用 SSO 的 Outlook 外接程序并旁加载它进行测试，Office 在传递到时将始终返回错误 13012  `forMSGraphAccess` `getAccessToken`，即使已授予管理员同意。 因此，在开发加载项时`forMSGraphAccess`**，** 应该Outlook选项。 请确保在部署用于生产时取消对选项的注释。 只有在旁加载时，才能使用 Outlook。

对于Outlook，请务必为租户启用新式Microsoft 365身份验证。 若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。

## <a name="see-also"></a>另请参阅

* [OAuth2 令牌Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Microsoft 标识平台和 OAuth 2.0 代表流](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)
