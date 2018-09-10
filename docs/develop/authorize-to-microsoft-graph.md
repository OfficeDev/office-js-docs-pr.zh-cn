---
title: 向 Office 加载项中的 Microsoft Graph 授权
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: f6e7de146d2f03256aa673a0653c1e03f9340d86
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876590"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>向 Office 加载项（预览版）中的 Microsoft Graph 授权

用户可以使用自己的个人 Microsoft 帐户/工作或学校 (Office 365) 帐户，登录 Office（在线、移动和桌面平台）。 Office 加载项获得 [Microsoft Graph](https://developer.microsoft.com/graph/docs) 授权访问权限的最佳方式是使用用户的 Office 登录凭据。 如此一来，Office 加载项能够访问其 Microsoft Graph 数据，而无需再次登录。 

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 预览版支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets)。
> 如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。 若要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>用于 SSO 和 Microsoft Graph 的加载项体系结构

除了托管 Web 应用程序的页面和 JavaScript 之外，外接程序还必须以相同的[完全限定的域名](https://msdn.microsoft.com/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly)托管一个或多个 Web API，这些 API 可获取 Microsoft Graph 的访问令牌，并向它发出请求。

外接程序清单包含标记，用于指定外接程序在 Azure Active Directory (Azure AD) v2.0 终结点中的注册方式，并指定外接程序需要的 Microsoft Graph 的任何权限。

### <a name="how-it-works-at-runtime"></a>运行时的工作方式

以下关系图显示了登录 Microsoft Graph 和获取访问的工作原理。

![SSO 过程关系图](../images/sso-access-to-microsoft-graph.png)

1. 在外接程序，JavaScript 调用新的 Office.js API [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)。它会让 Office 主机应用程序获取外接程序的访问令牌（以后，这称为**引导程序访问令牌**，因为它在此过程的后面将替换为另一个令牌。有关解码引导程序访问令牌的示例，请参阅 [示例访问令牌](sso-in-office-add-ins.md#example-access-token)。）
1. 如果用户未登录，Office 主机应用程序会打开弹出窗口，以供用户登录。
1. 如果当前用户是首次使用加载项，则会看到同意提示。
1. Office 主机应用程序从当前用户的 Azure AD v2.0 端点请求获取**引导程序访问令牌**。
1. Azure AD 将引导程序令牌发送给 Office 主机应用程序。
1. Office 主机应用程序向加载项发送 **引导程序访问令牌**，这是 `getAccessTokenAsync` 调用返回对象结果的一部分。
1. 加载项中的 JavaScript 向 Web API（与加载项托管在同一完全限定的域中）发出 HTTP 请求，并添加**引导程序访问令牌**作为授权证明。  
1. 服务器端代码验证传入**引导程序访问令牌**。
1. 服务器端代码使用“代表”流程（在 [OAuth2 令牌交换](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)和 [ Web API Azure 方案的守护进程或服务器应用程序](https://docs.microsoft.com/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)中定义）获取 Microsoft Graph 的访问令牌来交换引导程序访问令牌。
1. Azure AD 将 Microsoft Graph 访问令牌 （如果外接程序请求获取 *offline_access*权限，则同时返回刷新令牌）返回给加载项。
1. 服务器端代码将相应访问令牌缓存到 Microsoft Graph。
1. 服务器端代码向 Microsoft Graph 发出请求，并将访问令牌包含到 Microsoft Graph 中。
1. Microsoft Graph 将数据返回给加载项，从而将数据传递到加载项 UI。
1. 当 Microsoft Graph 的访问令牌到期时，服务器端代码可以使用其刷新令牌获取新的 Microsoft Graph 访问令牌。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>开发用于访问 Microsoft Graph 的 SSO 加载项

正如开发任何其他使用 SSO 的加载项一样，可以开发一个访问 Microsoft Graph 的加载项。 有关详细说明，请参阅[为 Office 加载项启用单一登录](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)。不同之处在于，加载项必须具有服务器端 Web API，并且该文章中所谓的访问令牌称为“引导程序访问令牌”。 

根据您的语言和框架，可能会提供库，这将简化您必须编写的服务器端代码。 您的代码应该执行以下操作：

* 验证从之前创建的令牌处理程序收到的加载项引导程序访问令牌。 有关更多信息，请参阅[验证访问令牌](sso-in-office-add-ins.md#validate-the-access-token)。 
* 通过调用 Azure AD v2.0 端点启动“代表”流程，该端点包括引导程序访问令牌、关于用户的一些元数据以及外接程序的凭据（相应 ID 和密钥）。
* 将返回的访问令牌缓存到 Microsoft Graph。 有关此流程的更多信息，请参阅 [Azure Active Directory v2.0 和 OAuth 2.0 代表流程](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。
* 通过将缓存的访问令牌传递给 Microsoft Graph，创建一个或多个获取 Microsoft Graph 数据的 Web API 方法。

> [!NOTE]
> 有关“代表”流程获取的已解码的 Microsoft Graph 访问令牌示例，请参阅 [Azure Active Directory v2.0 和 OAuth 2.0 代表流程](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。

有关详细演练和应用场景的示例，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)
* [应用场景：在 Outlook 加载项中为服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



