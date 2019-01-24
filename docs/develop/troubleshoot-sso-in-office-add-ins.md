---
title: 排查单一登录 (SSO) 错误消息
description: ''
ms.date: 12/08/2017
localization_priority: Priority
ms.openlocfilehash: d25461b3d3bdb6db96cb897575efe0ecb5fe1439
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386853"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>排查单一登录 (SSO) 错误消息（预览）

本文提供了一些指南，介绍了如何排查 Office 加载项中出现的单一登录 (SSO) 问题，以及如何让已启用 SSO 的加载项可靠地处理特殊条件或错误。

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 在预览版中支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅[IdentityAPI 要求集]https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。
> 若要使用 SSO，必须从加载项的启动 HTML 页面中的 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 加载 Office JavaScript 库的 Beta 版。
> 如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。 若要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

## <a name="debugging-tools"></a>调试工具

开发时，强烈建议使用具有以下功能的工具：能够截获并显示加载项 Web 服务发出的 HTTP 请求和发送给它的响应。最热门的两个工具是：

- [Fiddler](https://www.telerik.com/fiddler)：免费使用（[文档](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler)）
- [Charles](https://www.charlesproxy.com/)：免费使用 30 天。 （[文档](https://www.charlesproxy.com/documentation/)）

开发服务 API 时，不妨还尝试使用：

- [Postman](https://www.getpostman.com/postman)：免费使用（[文档](https://www.getpostman.com/docs/)）

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>导致 getAccessTokenAsync 生成错误的原因和处理方法

有关此部分中介绍的错误处理示例，请参阅：
- [Office-Add-in-ASPNET-SSO 中的 Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [Office-Add-in-NodeJS-SSO 中的 program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> 除了本节中提出的建议外，Outlook 加载项还有一种方法可以响应任何 13*nnn* 错误。 有关详细信息，请参阅[应用场景：在 Outlook 加载项中对服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)和 [AttachmentsDemo 示例加载项](https://github.com/OfficeDev/outlook-add-in-attachments-demo)。

### <a name="13000"></a>13000

加载项或 Office 版本不支持 [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) API。

- Office 版本不支持 SSO。必须为 Office 2016 版本 1710（生成号 8629.nnnn）或更高版本（Office 365 订阅版本，有时亦称为“即点即用”）。可能必须成为 Office 预览体验成员，才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。 
- 加载项清单缺少适当的 [WebApplicationInfo](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/webapplicationinfo) 部分。

加载项应该通过回退到用户身份验证备用系统来响应此错误。 有关详细信息，请参阅[要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)。

### <a name="13001"></a>13001

用户未登录 Office。 代码应重新调用 `getAccessTokenAsync` 方法，并在 [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 参数中传递选项 `forceAddAccount: true`。 但是不要多次执行此操作。 用户可能已决定不登录。

Office Online 中绝不会出现此错误。 如果用户的 Cookie 到期，Office Online 返回的是错误 13006。

### <a name="13002"></a>13002

用户中止登录或同意（例如，在同意对话框中选择**取消**）。

- 如果加载项提供的功能无需用户登录（或授予许可），代码应捕获此错误，并让加载项继续正常运行。
- 如果加载项需要登录用户授予许可，代码应提示用户重复执行操作，但只能重复一次。

### <a name="13003"></a>13003

用户类型不受支持。 用户未使用有效的 Microsoft 帐户或 Office 365（工作或学校）帐户登录 Office。 例如，当使用本地域帐户运行 Office 时，可能会生成此错误。 代码应该要求用户登录 Office 或退回到用户身份验证备用系统。 有关详细信息，请参阅[要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)。

### <a name="13004"></a>13004

资源无效。 加载项清单未正确配置。 请更新此清单。 有关详细信息，请参阅[验证并排查清单问题](../testing/troubleshoot-manifest.md)。 最常见的问题是**资源**元素（在 **WebApplicationInfo** 元素中）具有与加载项的域不匹配的域。 虽然资源值的协议部分应该是“api”而不是“https”；域名的所有其他部分（包括端口，如果有）应与加载项域名的相应部分相同。

### <a name="13005"></a>13005

授权无效。 这通常意味着，Office 尚未获得对加载项 Web 服务的预授权。 有关详细信息，请参阅[创建服务应用程序](sso-in-office-add-ins.md#create-the-service-application)和[向 Azure AD v2.0 终结点注册加载项](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) 或[向 Azure AD v2.0 终结点注册加载项](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS)。 如果用户未授权服务应用程序访问其 `profile`，也可能会生成此错误。

### <a name="13006"></a>13006

客户端错误。代码应提示用户注销并重启 Office，或重启 Office Online 会话。

### <a name="13007"></a>13007

Office 主机无法获取对加载项 Web 服务的访问令牌。

- 如果在开发过程中发生此错误，请确保加载项注册和加载项清单指定 `openid` 和 `profile` 权限。 有关详细信息，请参阅[向 Azure AD v2.0 终结点注册加载项](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) 或[向 Azure AD v2.0 终结点注册加载项](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS)，以及[配置加载项](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) 或[配置加载项](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS)。
- 在生产中，有几种情况可能导致此错误。 其中一些是：
    - 用户之前表示同意，之后撤销同意。 代码应该使用选项 `forceConsent: true` 调用 `getAccessTokenAsync` 方法，但不能超过一次。
    - 用户拥有一个 Microsoft 帐户 (MSA) 标识。 使用 MSA 时，某些导致某个其他工作或学校帐户 13nnn 错误的情况将导致 13007。 

  对于所有这些情况，如果已尝试一次 `forceConsent` 选项，则代码可能会建议用户稍后重试该操作。

### <a name="13008"></a>13008

用户在上一次调用 `getAccessTokenAsync` 完成前触发了调用 `getAccessTokenAsync` 的操作。 代码应提示用户在上一次操作完成后再重复此操作。

### <a name="13009"></a>13009

加载项使用选项 `forceConsent: true` 调用了 `getAccessTokenAsync` 方法，但加载项清单部署到的目录类型不支持强制许可。 代码应重新调用 `getAccessTokenAsync` 方法，并在 [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 参数中传递选项 `forceConsent: false`。 不过，调用 `getAccessTokenAsync` 与 `forceConsent: true` 本身可能是对调用 `getAccessTokenAsync` 和 `forceConsent: false` 失败的自动响应，因此代码应该跟踪是否已调用 `getAccessTokenAsync` 和 `forceConsent: false`。 如果已尝试该选项，则代码应该告知用户退出 Office 并再次登录，或者应该回退到用户身份验证备用系统。 有关详细信息，请参阅[要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)。

> [!NOTE]
> Microsoft 不一定将这个限制强加给任何类型的加载项目录。 如果没有，则绝不会出现这个错误。

### <a name="13010"></a>13010

用户正在 Office Online 上运行加载项，但使用的是 Edge 或 Internet Explorer。 用户的 Office 365 域和 login.microsoftonline.com 域位于浏览器设置中的不同安全区域。 如果此错误返回，用户将已看到对此进行解释的错误，并链接到关于如何更改区域配置的页面。 如果加载项提供的功能无需用户登录，代码应捕获此错误，并让加载项继续正常运行。

### <a name="13012"></a>13012

加载项在不支持 `getAccessTokenAsync` API的平台上运行。 例如，在 iPad 上它不受支持。 另请参阅[标识 API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。

### <a name="50001"></a>50001

此错误（未特定于 `getAccessTokenAsync`）可能表示浏览器已缓存 office.js 文件的旧副本。 在开发时，清除浏览器的缓存。 另一种可能是 Office 的版本不够新，不足以支持 SSO。 请参阅[先决条件](create-sso-office-add-ins-aspnet.md#prerequisites)。

在生产加载项中，加载项应该通过回退到用户身份验证备用系统来响应此错误。 有关详细信息，请参阅[要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)。


## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Azure Active Directory 服务器端错误

有关此部分中介绍的错误处理示例，请参阅：
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>条件访问/多重身份验证错误
 
在特定 AAD 和 Office 365 标识配置中，一些可通过 Microsoft Graph 访问的资源可以要求进行多重身份验证 (MFA)，即使用户的 Office 365 租赁并不要求此验证。 通过代表流收到对 MFA 保护资源的令牌请求时，AAD 会向加载项 Web 服务返回包含 `claims` 属性的 JSON 消息。 claims 属性指明需要进一步执行哪几重身份验证。

服务器端代码应测试此消息是否有错，并将 claims 值中继到客户端代码。 客户端需要此信息，因为 Office 处理 SSO 加载项的身份验证。发送到客户端的消息可以是错误（如 `500 Server Error` 或 `401 Unauthorized`），也可以是成功响应的正文部分（如 `200 OK`）。 无论属于上述哪种情况，代码对加载项 Web API 的客户端 AJAX 调用的（失败或成功）回调都应测试此响应是否有错。 如果已中继 claims 值，代码应重新调用 `getAccessTokenAsync`，并在 [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 参数中传递选项 `authChallenge: CLAIMS-STRING-HERE`。 如果 AAD 看到此字符串，它会先提示用户进行多重身份验证，再返回将在代表流中接受的新访问令牌。

### <a name="consent-missing-errors"></a>缺少许可错误

如果 AAD 未记录用户（或租户管理员）已授权加载项访问 Microsoft Graph 资源，AAD 会向 Web 服务发送错误消息。 代码必须指示客户端（例如，在 `403 Forbidden` 响应的正文中）通过 `forceConsent: true` 选项重新调用 `getAccessTokenAsync`。

### <a name="invalid-or-missing-scope-permission-errors"></a>范围（权限）无效或缺失错误

- 服务器端代码应向客户端发送 `403 Forbidden` 响应，向用户显示易记消息。如果可能，请在控制台或日志中记录此错误。
- 请确保加载项清单 [Scopes](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/scopes) 部分指定了所需的全部权限。 此外，还请确保加载项 Web 服务注册指定了相同的权限。 同时检查是否有拼写错误。 有关详细信息，请参阅[向 Azure AD v2.0 终结点注册加载项](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) 或[向 Azure AD v2.0 终结点注册加载项](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) (Node JS)，以及[配置加载项](create-sso-office-add-ins-aspnet.md#configure-the-add-in) (ASP.NET) 或[配置加载项](create-sso-office-add-ins-nodejs.md#configure-the-add-in) (Node JS)。

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>调用 Microsoft Graph 时令牌过期或无效错误

一些身份验证和授权库（包括 MSAL）在必要时使用缓存的刷新令牌，防止出现令牌过期错误。 也可以编码自己的令牌缓存系统。 有关如何这样做的示例，请参阅 [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)，并重点参阅 [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) 文件。

不过，如果收到了令牌过期或令牌无效错误，代码必须指示客户端（例如，在 `401 Unauthorized` 响应的正文中）重新调用 `getAccessTokenAsync`，并重复调用加载项 Web API 终结点，这会重复执行代表流来获取新 Microsoft Graph 令牌。

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>调用 Microsoft Graph 时令牌无效错误

按照处理令牌到期错误的方法处理此错误。请参阅上一部分。

### <a name="invalid-audience-error"></a>受众无效错误

服务器端代码应向客户端发送 `403 Forbidden` 响应，向用户显示易记消息，并尽可能在控制台或日志中记录此错误。

若要详细了解如何添加令牌验证的多租户支持，请参阅 [Azure 多租户示例](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect)。
