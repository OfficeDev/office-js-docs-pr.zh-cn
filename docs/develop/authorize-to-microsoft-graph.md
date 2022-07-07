---
title: 使用 SSO 对 Microsoft Graph 授权
description: 了解 Office 外接程序的用户如何使用单一登录 (SSO) 从 Microsoft Graph 提取数据。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4ecb945dcd99400065fde3e80e8b60d67266e0b1
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659645"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>使用 SSO 对 Microsoft Graph 授权

用户使用其个人 Microsoft 帐户或其Microsoft 365 教育版或工作帐户登录到 Office。 在 Office 加载项中获取对 [Microsoft Graph](https://developer.microsoft.com/graph/docs) 的访问权限的最佳方式是使用用户的 Office 登录凭据。 这使用户能够访问其 Microsoft Graph 数据，而无需再次登录。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO 和 Microsoft Graph 的加载项体系结构

除了托管 Web 应用程序的页面和 JavaScript 之外，外接程序还必须以相同的[完全限定的域名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)托管一个或多个 Web API，这些 API 可获取 Microsoft Graph 的访问令牌，并向它发出请求。

外接程序清单包含一个 **\<WebApplicationInfo\>** 元素，该元素向 Office 提供重要的 Azure 应用注册信息，包括加载项所需的 Microsoft Graph 权限。

### <a name="how-it-works-at-runtime"></a>运行时的工作方式

下图显示了登录和访问 Microsoft Graph 所涉及的步骤。 整个过程使用 OAuth 2.0 和 JWT 访问令牌。

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="显示 SSO 进程的示意图。" border="false":::

1. 外接程序的客户端代码调用 Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1))。 这会告知 Office 主机获取加载项的访问令牌。

    如果用户未登录，则 Office 主机与Microsoft 标识平台一起为用户提供登录和同意的 UI。

2. Office 主机从Microsoft 标识平台请求访问令牌。
3. Microsoft 标识平台向 Office 主机返回访问令牌 *A*。 访问令牌 *A* 仅提供对加载项自己的服务器端 API 的访问权限。 它不提供对 Microsoft Graph 的访问权限。
4. Office 主机将访问令牌 *A* 返回到加载项的客户端代码。 现在，客户端代码可以对服务器端 API 进行经过身份验证的调用。
5. 客户端代码向服务器端需要身份验证的 Web API 发出 HTTP 请求。 它包括访问令牌 *A* 作为授权证明。 服务器端代码验证访问令牌 *A*。
6. 服务器端代码使用 OAuth 2.0 代理流 (OBO) 请求具有 Microsoft Graph 权限的新访问令牌。
7. 如果加载 *项请求offline_access* 权限) ，则Microsoft 标识平台返回具有 Microsoft Graph (权限和刷新令牌的新访问令牌 *B*。 服务器可以选择性地缓存访问令牌 *B*。
8. 服务器端代码向 Microsoft 图形 API发出请求，并包含具有 Microsoft Graph 权限的访问令牌 *B*。
9. Microsoft Graph 将数据返回到服务器端代码。
10. 服务器端代码将数据返回到客户端代码。

在后续请求中，对服务器端代码进行经过身份验证的调用时，客户端代码将始终传递访问令牌 *A* 。 服务器端代码可以缓存令牌 *B* ，这样就不需要在将来的 API 调用中再次请求它。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>开发可访问 Microsoft Graph 的 SSO 加载项

开发一个可访问 Microsoft Graph 的加载项，就像使用 SSO 的任何其他应用程序一样。 有关详细说明，请参阅 [为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)。区别在于，外接程序必须具有服务器端 Web API。

根据所用的语言和框架，可能存在一些库，可简化必须编写的服务器端代码。 代码应执行以下操作：

* 每次从客户端代码传递访问令牌 *A* 时，都会对其进行验证。 有关详细信息，请参阅[验证访问令牌](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code)。
* 启动 OAuth 2.0 代理流 (OBO) ，并调用Microsoft 标识平台，其中包括访问令牌、有关用户的一些元数据，以及加载项的凭据 (其 ID 和机密) 。 有关 OBO 流的详细信息，请参阅 [Microsoft 标识平台 和 OAuth 2.0 代理流](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)。
* （可选）流完成后，缓存具有 Microsoft Graph 权限的返回访问令牌 *B* 。 如果加载项对 Microsoft Graph 进行多次调用，则需要执行此操作。 有关详细信息，请参阅 [使用 Microsoft 身份验证库 (MSAL) 获取和缓存令牌 ](/azure/active-directory/develop/msal-acquire-cache-tokens)
* 通过将可能缓存的 () 访问令牌 *B* 传递给 Microsoft Graph，创建一个或多个用于获取 Microsoft Graph 数据的 Web API 方法。

有关详细演练和应用场景的示例，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)
* [应用场景：为 Outlook 加载项中的服务实现单一登录](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>在 Microsoft AppSource 中分发已启用 SSO 的加载项

当 Microsoft 365 管理员从 [AppSource](https://appsource.microsoft.com) 获取外接程序时，管理员可以通过 [集成应用](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) 重新分发该加载项，并向加载项授予管理员访问 Microsoft Graph 范围的许可。 但是，最终用户也可以直接从 AppSource 获取加载项，在这种情况下，用户必须同意外接程序。 这可能会造成潜在的性能问题，我们为此提供了解决方案。

如果代码在调用时传递`allowConsentPrompt`了该选项（例如`OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`，如果Microsoft 标识平台向 Office 报告尚未向加载项授予同意，则 Office 可以提示用户`getAccessToken`同意。 但是，出于安全原因，Office 只能提示用户同意 Microsoft Graph `profile` 范围。 *Office 无法提示同意其他 Microsoft Graph 范围*，甚至 `User.Read`不能。 这意味着，如果用户在提示时授予同意，Office 将返回访问令牌。 但是，尝试将访问令牌交换为具有其他 Microsoft Graph 范围的新访问令牌失败，并出现错误 AADSTS65001，这意味着尚未授予对 Microsoft Graph 作用域) 的许可 (。

> [!NOTE]
> 即使管理员已关闭最终用户同意，该范围的同意 `{ allowConsentPrompt: true }` 请求仍可能失败 `profile` 。 有关详细信息，请参阅 [配置最终用户如何使用 Azure Active Directory 同意应用程序](/azure/active-directory/manage-apps/configure-user-consent)。

代码可以通过回退到另一个身份验证系统来处理此错误，这会提示用户同意 Microsoft Graph 范围。 有关代码示例，请参阅 [创建使用单一登录的 Node.js Office 加](create-sso-office-add-ins-nodejs.md) 载项，并 [创建使用单一登录的 ASP.NET Office 加](create-sso-office-add-ins-aspnet.md) 载项及其链接到的示例。 整个过程需要多次往返Microsoft 标识平台。 若要避免此性能损失，请在调用`getAccessToken`中包括`forMSGraphAccess`选项，例如。 `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` 这会向 Office 发出信号，表明您的外接程序需要 Microsoft Graph 范围。 Office 将要求Microsoft 标识平台验证是否已向加载项授予对 Microsoft Graph 范围的许可。 如果有，则返回访问令牌。 如果没有，则调用 `getAccessToken` 返回错误 13012。 代码可以通过立即回退到备用的身份验证系统来处理此错误，而无需尝试将令牌与Microsoft 标识平台交换。

最佳做法是始终传递 `forMSGraphAccess` 到 `getAccessToken` 加载项将在 AppSource 中分发并需要 Microsoft Graph 作用域。

## <a name="details-on-sso-with-an-outlook-add-in"></a>有关使用 Outlook 加载项的 SSO 的详细信息

如果开发使用 SSO 的 Outlook 外接程序并旁加载它进行测试，则即使已授予管理员同意，在传递给 `getAccessToken` Office 时`forMSGraphAccess`*，* Office 始终会返回错误 13012。 因此，在开发 Outlook 加载项 **时**，应注释掉`forMSGraphAccess`该选项。 部署生产时，请务必取消注释该选项。 仅当您在 Outlook 中旁加载时，才会发生虚假的 13012。

对于 Outlook 加载项，请务必为 Microsoft 365 租户启用新式身份验证。 若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。

## <a name="see-also"></a>另请参阅

* [OAuth2 令牌交换](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Microsoft 标识平台和 OAuth 2.0 代表流](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [IdentityAPI 要求集](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
