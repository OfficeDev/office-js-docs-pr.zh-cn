---
title: 排查单一登录 (SSO) 错误消息
description: 有关如何解决 Office 加载项中的单一登录（SSO）问题，以及如何处理特殊条件或错误的指南。
ms.date: 04/13/2020
localization_priority: Normal
ms.openlocfilehash: da42b3b3d9b5cf1fede999a18bbe36c5532bd866
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609697"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>排查单一登录 (SSO) 错误消息（预览）

本文提供了一些指南，介绍了如何排查 Office 加载项中出现的单一登录 (SSO) 问题，以及如何让已启用 SSO 的加载项可靠地处理特殊条件或错误。

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 在预览版中支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。 若要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

## <a name="debugging-tools"></a>调试工具

开发时，强烈建议使用具有以下功能的工具：能够截获并显示加载项 Web 服务发出的 HTTP 请求和发送给它的响应。最热门的两个工具是：

- [Fiddler](https://www.telerik.com/fiddler)：免费使用（[文档](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler)）
- [Charles](https://www.charlesproxy.com)：免费使用 30 天。 （[文档](https://www.charlesproxy.com/documentation/)）

## <a name="causes-and-handling-of-errors-from-getaccesstoken"></a>导致 getAccessToken 生成错误的原因和处理方法

有关此部分中介绍的错误处理示例，请参阅：
- [Office-Add-in-ASPNET-SSO 中的 HomeES6.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js)
- [ssoAuthES6.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js)

### <a name="13000"></a>13000

加载项或 Office 版本不支持 [getAccessToken](../develop/sso-in-office-add-ins.md#sso-api-reference) API。

- Office 版本不支持 SSO。 版本必须为任何月度频道中的 Office 365（Office 的订阅版本）。
- 加载项清单缺少适当的 [WebApplicationInfo](../reference/manifest/webapplicationinfo.md) 部分。

加载项应该通过回退到用户身份验证备用系统来响应此错误。 有关详细信息，请参阅[要求和最佳做法](../develop/sso-in-office-add-ins.md#requirements-and-best-practices)。

### <a name="13001"></a>13001

用户未登录 Office。 在大多数情况下，应通过在 `AuthOptions` 参数中传递选项 `allowSignInPrompt: true` 来防止出现此错误。

但是，可能会出现异常情况。 例如，你希望加载项打开要求用户登录的功能；但*前提*是该用户*已经*登录 Office。 如果用户未登录，则你希望该加载项打开不要求用户登录的备用功能集。 在这种情况下，当加载项启动时运行的逻辑将调用不具有 `allowSignInPrompt: true` 的 `getAccessToken`。 使用 13001 错误作为标志，告诉加载项显示备用功能集。

另一个选项是通过回退到用户身份验证备用系统来响应 13001。 这会使用户登录到 AAD，但不会使用户登录到 Office。

**Office 网页版**中绝不会出现此错误。 如果用户的 Cookie 到期，**Office 网页版**将返回错误 13006。

### <a name="13002"></a>13002

用户中止登录或同意（例如，在同意对话框中选择**取消**）。

- 如果加载项提供的功能无需用户登录（或授予许可），代码应捕获此错误，并让加载项继续正常运行。
- 如果加载项需要登录用户授予许可，则代码应显示一个登录按钮。

### <a name="13003"></a>13003

用户类型不受支持。 用户未使用有效的 Microsoft 帐户或 Office 365（工作或学校）帐户登录 Office。 例如，当使用本地域帐户运行 Office 时，可能会生成此错误。 代码应回退到用户身份验证备用系统。 有关详细信息，请参阅[要求和最佳做法](../develop/sso-in-office-add-ins.md#requirements-and-best-practices)。

### <a name="13004"></a>13004

资源无效。 （应仅在开发过程中看到此错误。）尚未正确配置加载项清单。 请更新此清单。 有关详细信息，请参阅[验证 Office 加载项的清单](../testing/troubleshoot-manifest.md)。 最常见的问题是**资源**元素（在 **WebApplicationInfo** 元素中）具有与加载项的域不匹配的域。 虽然资源值的协议部分应该是“api”而不是“https”；域名的所有其他部分（包括端口，如果有）应与加载项域名的相应部分相同。

### <a name="13005"></a>13005

授权无效。 这通常意味着，Office 尚未获得对加载项 Web 服务的预授权。 有关详细信息，请参阅[创建服务应用程序](sso-in-office-add-ins.md#create-the-service-application)和[向 Azure AD v2.0 终结点注册加载项](register-sso-add-in-aad-v2.md)。 如果用户未授权服务应用程序访问其 `profile`，或已吊销许可，也可能会生成此错误。 代码应回退到用户身份验证备用系统。

另一个可能的原因是，在开发过程中，加载项使用的是 Internet Explorer，并且你使用的是自签名证书。 （若要确定加载项使用的浏览器，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。）

### <a name="13006"></a>13006

客户端错误。 此错误仅出现在 **Office 网页版**中。 代码应提示用户注销，然后重启 Office 浏览器会话。

### <a name="13007"></a>13007

Office 主机无法获取对加载项 Web 服务的访问令牌。

- 如果在开发过程中发生此错误，请确保加载项注册和加载项清单指定 `profile` 权限（和 `openid` 权限 - 如果你使用的是 MSAL.NET）。 如需了解更多信息，请参阅[向 Azure AD v2.0 终结点注册加载项](register-sso-add-in-aad-v2.md)。
- 在生产中，有几种情况可能导致此错误。 其中一些是：
    - 用户拥有一个 Microsoft 帐户 (MSA) 标识。
    - 使用 MSA 时，某些导致某个其他工作或学校帐户 13xxx 错误的情况将导致 13007。

  对于所有这些情况，代码应回退到用户身份验证备用系统。

### <a name="13008"></a>13008

用户在上一次调用 `getAccessToken` 完成前触发了调用 `getAccessToken` 的操作。 此错误仅出现在 **Office 网页版**中。 代码应提示用户在上一次操作完成后再重复此操作。

### <a name="13010"></a>13010

用户正在 Microsoft Edge 或 Internet Explorer 上的 Office 中运行加载项。 用户的 Office 365 域和 `login.microsoftonline.com` 域在浏览器设置中位于不同的安全区域中。 此错误仅出现在 **Office 网页版**中。 如果此错误返回，用户将已看到对此进行解释的错误，并链接到关于如何更改区域配置的页面。 如果加载项提供的功能无需用户登录，代码应捕获此错误，并让加载项继续正常运行。

### <a name="13012"></a>13012

存在几种可能的原因：

- 加载项在不支持 `getAccessToken` API的平台上运行。 例如，在 iPad 上它不受支持。 另请参阅[标识 API 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。
- `forMSGraphAccess` 选项在调用中传递给 `getAccessToken`，并且用户从 AppSource 获得了加载项。 在此方案中，对于所需的 Microsoft Graph 范围（权限），租户管理员未向加载项授予许可。 撤回具有 `allowConsentPrompt` 的 `getAccessToken` 将无法解决此问题，因为允许 Office 提示用户仅同意 AAD `profile` 范围。

代码应回退到用户身份验证备用系统。

在开发中，该加载项在 Outlook 中旁加载，并且在 `getAccessToken` 调用中传递了 `forMSGraphAccess` 选项。

### <a name="13013"></a>13013

`getAccessToken`在短时间内调用次数过多，因此 Office 限制了最近的呼叫。 这通常是由对方法的调用的无限循环引起的。 在某些情况下，建议撤回该方法。 但是，您的代码应使用计数器或标志变量以确保不会重复回调该方法。 如果相同的 "重试" 代码路径再次运行，则代码应回退到用户身份验证的备用系统。 有关代码示例，请参阅如何 `retryGetAccessToken` 在[HomeES6](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js)或[ssoAuthES6](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js)中使用变量。

### <a name="50001"></a>50001

此错误（未特定于 `getAccessToken`）可能表示浏览器已缓存 office.js 文件的旧副本。 在开发时，清除浏览器的缓存。 另一种可能是 Office 的版本不够新，不足以支持 SSO。 在 Windows 上，最低版本是 16.0.12215.20006。 在 Mac 上，它是 16.32.19102902。

在生产加载项中，加载项应该通过回退到用户身份验证备用系统来响应此错误。 有关详细信息，请参阅[要求和最佳做法](../develop/sso-in-office-add-ins.md#requirements-and-best-practices)。

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Azure Active Directory 服务器端错误

有关此部分中介绍的错误处理示例，请参阅：
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)

### <a name="conditional-access--multifactor-authentication-errors"></a>条件访问/多重身份验证错误

在特定 AAD 和 Office 365 标识配置中，一些可通过 Microsoft Graph 访问的资源可以要求进行多重身份验证 (MFA)，即使用户的 Office 365 租赁并不要求此验证。 通过代表流收到对 MFA 保护资源的令牌请求时，AAD 会向加载项 Web 服务返回包含 `claims` 属性的 JSON 消息。 claims 属性指明需要进一步执行哪几重身份验证。

代码应对此 `claims` 属性进行测试。 根据加载项的体系结构，你可以在客户端进行测试，也可以在服务器端进行测试并将其中继到客户端。 客户端需要此信息，因为 Office 处理 SSO 加载项的身份验证。如果从服务器端进行中继，则发送到客户端的消息可以是错误（如 `500 Server Error` 或 `401 Unauthorized`），也可以是成功响应的正文部分（如 `200 OK`）。 无论属于上述哪种情况，代码对加载项 Web API 的客户端 AJAX 调用的（失败或成功）回调都应测试此响应是否有错。 

无论采用何种体系结构，如果声明值已从 AAD 发送，则代码应重新调用 `getAccessToken` 并 `authChallenge: CLAIMS-STRING-HERE` 在参数中传递选项 `options` 。 如果 AAD 看到此字符串，它会先提示用户进行多重身份验证，再返回将在代表流中接受的新访问令牌。

### <a name="consent-missing-errors"></a>缺少许可错误

如果 AAD 未记录用户（或租户管理员）已授权加载项访问 Microsoft Graph 资源，AAD 会向 Web 服务发送错误消息。 代码必须指示客户端（例如，在 `403 Forbidden` 响应的正文中）。

如果加载项需要只能由管理员许可的 Microsoft Graph 范围，则代码应该会引发错误。 如果用户只能许可所需的范围，则代码应回退到用户身份验证备用系统。

### <a name="invalid-or-missing-scope-permission-errors"></a>范围（权限）无效或缺失错误

应该只会在开发中看到此类错误。

- 服务器端代码应向客户端发送 `403 Forbidden` 响应，它应该会在控制台或日志中记录此错误。
- 请确保加载项清单 [Scopes](../reference/manifest/scopes.md) 部分指定了所需的全部权限。 此外，还请确保加载项 Web 服务注册指定了相同的权限。 同时检查是否有拼写错误。 如需了解更多信息，请参阅[向 Azure AD v2.0 终结点注册加载项](register-sso-add-in-aad-v2.md)。

### <a name="invalid-audience-error-in-the-access-token-not-the-bootstrap-token"></a>访问令牌（而非启动令牌）中的无效受众错误

服务器端代码应向客户端发送 `403 Forbidden` 响应，向用户显示易记消息，并尽可能在控制台或日志中记录此错误。
