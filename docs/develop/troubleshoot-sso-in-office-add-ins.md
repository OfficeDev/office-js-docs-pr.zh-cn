---
title: 排查单一登录 (SSO) 错误消息
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: 5abf10d8281ea54be9a172c3f45b742fb33991df
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506068"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>排查单一登录 (SSO) 错误消息（预览）

本文提供了一些指南，介绍了如何排查 Office 加载项中出现的单一登录 (SSO) 问题，以及如何让已启用 SSO 的加载项可靠地处理特殊条件或错误。

> [!NOTE]
> 在预览的 Word、 Excel、 Outlook 和 PowerPoint 当前支持单一登录 API。有关其中当前支持单一登录 API 的详细信息，请参阅 [ IdentityAPI 要求集]https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。若要使用 SSO，必须加载 beta 版中的 Office JavaScript 库 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 中的加载项的启动 HTML 页。如果正在使用 Outlook 加载项，请务必为 Office 365 租户启用现代的身份验证。欲知如何执行此操作的信息，请参阅 [在线交流： 如何启用租户的现代身份验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

## <a name="debugging-tools"></a>调试工具

开发时，强烈建议使用具有以下功能的工具：能够截获并显示加载项 Web 服务发出的 HTTP 请求和发送给它的响应。最热门的两个工具是： 

- [Fiddler](http://www.telerik.com/fiddler)：免费使用（[文档](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler)）
- [Charles](https://www.charlesproxy.com/)：免费使用 30 天。（[文档](https://www.charlesproxy.com/documentation/)）

开发服务 API 时，不妨还尝试使用：

- [Postman](http://www.getpostman.com/postman)：免费使用（[文档](https://www.getpostman.com/docs/)）

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>导致 getAccessTokenAsync 生成错误的原因和处理方法

有关此部分中介绍的错误处理示例，请参阅：
- [Office-Add-in-ASPNET-SSO 中的 Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [Office-Add-in-NodeJS-SSO 中的 program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> 除了本节中所提出的建议，Outlook 加载项具有任何 13*nnn* 错误响应的其他方法。欲知详请，请参阅 [方案： 在 Outlook 加载项中实现单一登录服务器](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) 和 [AttachmentsDemo 示例加载项](https://github.com/OfficeDev/outlook-add-in-attachments-demo)。 

### <a name="13000"></a>13000

加载项或 Office 版本不支持 [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) API。 

- Office 版本不支持 SSO。必须为 Office 2016 版本 1710（生成号 8629.nnnn）或更高版本（Office 365 订阅版本，有时亦称为“即点即用”）。可能必须成为 Office 预览体验成员，才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。 
- 加载项清单缺少适当的 [WebApplicationInfo](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/webapplicationinfo?view=office-js) 部分。

加载项应该通过回退到备用的用户身份验证系统来响应此错误。欲知详情，请参阅 [要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)。

### <a name="13001"></a>13001

用户未登录 Office。代码应调回 `getAccessTokenAsync` 方法并传递 `forceAddAccount: true` [ options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 参数中的选项。但此操作不能执行一次以上。用户可能已决定不登录。

Office Online 永远不会出现此错误。如果用户的 cookie 过期，Office Online 返回错误 13006。 

### <a name="13002"></a>13002

用户放弃了登录或许可；例如，在许可对话框上 **取消**。 

- 如果加载项提供的功能无需用户登录（或授予许可），代码应捕获此错误，并让加载项继续正常运行。
- 如果加载项需要登录用户授予许可，代码应提示用户重复执行操作，但只能重复一次。 

### <a name="13003"></a>13003

用户类型不受支持。用户未使用有效的 Microsoft 帐户或 Office 365 （"工作或学校"） 帐户登录 Office 。如果 Office 使用本地域帐户运行，则可能会发生这种情况，例如，代码应要求用户登录 Office 或退回到备用的用户身份验证系统。欲知详情，请参阅 [要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)。


### <a name="13004"></a>13004

无效的资源。加载项清单尚未正确配置。更新清单。欲知详情，请参阅 [与清单一同验证和解决问题](../testing/troubleshoot-manifest.md)。最常见的问题是 （在 **WebApplicationInfo** 元素中） 的 **Resource** 元素具有与加载项的域不匹配的域。尽管 Resource 值的协议部分应为" api "不是" https "; 域名的所有其他部分（包括端口，如果有的话）应与加载项的相同。

### <a name="13005"></a>13005

授予无效。这通常意味着 Office 仍未事先授权加载项的 web 服务。欲知详情，请参阅 [创建服务应用程序](sso-in-office-add-ins.md#create-the-service-application) 和 [使用Azure AD v2.0端点注册加载项](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) ( ASP.NET ) 或 [使用Azure AD v2.0端点注册加载项](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) ( Node JS )。如果用户未将服务应用程序权限授予其 `profile`，也可能发生这种情况。

### <a name="13006"></a>13006

客户端错误。代码应提示用户注销并重启 Office，或重启 Office Online 会话。

### <a name="13007"></a>13007

Office 主机无法获取对加载项 Web 服务的访问令牌。

- 如果在开发过程中发生此错误，请确保加载项注册和加载项指令清单指定 `openid` 和 `profile` 权限，欲知详情，请参阅 [使用Azure AD v2.0端点注册加载项](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) 或 [使用Azure AD v2.0端点注册加载项](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)  ( Node JS )，以及 [配置加载项](create-sso-office-add-ins-aspnet.md#configure-the-add-in) ( ASP.NET ) 或 [配置加载项](create-sso-office-add-ins-nodejs.md#configure-the-add-in) ( Node JS )。
- 生成时，有几种情况会导致这个错误。其中有一些是：
    - 用户在事先授予许可后已撤销同意。代码应使用选项 `forceConsent: true`选项调回 `getAccessTokenAsync` 方法，但是不能超过一次。
    - 用户具备 Microsoft 帐户 ( MSA ) 标识。当使用MSA时将导致13007，某些情况会导致工作或学校帐户中出现其他13nnn错误之一。 

  对于所有这些情况，如果已经尝试过一次 `forceConsent` 选项，那么代码可能会提示用户稍后重试操作。

### <a name="13008"></a>13008

在先前的 `getAccessTokenAsync` 调用完成之前，用户触发了一个调用 `getAccessTokenAsync` 的操作。代码会要求用户在上一个操作完成后重复操作。

### <a name="13009"></a>13009

加载项调用 `getAccessTokenAsync` 方法选项 `forceConsent: true`，但加载项的清单将部署到一种类型的不支持强制同意的目录。代码应调回在 [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 参数中 `getAccessTokenAsync` 方法并传递选项 `forceConsent: false` 。但是，调用 `getAccessTokenAsync` 与 `forceConsent: true` 可能本身已失败的呼叫自动响应 `getAccessTokenAsync` 与 `forceConsent: false`，因此代码应跟踪的是否 `getAccessTokenAsync` 与 `forceConsent: false` 已被调用。如果是，则代码应也会告知用户注销 Office 和再次登录或它应回退到备用系统的用户身份验证。欲知详情，请参阅 [要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)。

> [!NOTE]
> Microsoft 不一定会对任何类型的加载项目录实施此限制。如果没有，则永远不会出现此错误。

### <a name="13010"></a>13010

用户正在 Office Online 上运行加载项，并且正在使用 Edge 或 Internet Explorer，用户的 Office 365 域和 login.microsoftonline.com 域位于浏览器设置中的不同安全区域中。如果返回此错误，则用户将看到解释此错误并链接到有关如何更改区域配置的页面的错误。如果加载项提供的功能不需要用户登录，那么代码应该捕获此错误并允许加载项保持运行。

### <a name="13012"></a>13012

加载项不支持的平台上运行 `getAccessTokenAsync` API。例如，它不支持 iPad 。另请参阅 [标识 API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。

### <a name="50001"></a>50001

此错误 (不是特定于 `getAccessTokenAsync`) 可能指示在浏览器已缓存的 office.js 文件的旧副本。开发时, 清除浏览器的缓存。另一种可能是 Office 的版本不够最新，不足以支持 SSO 。请参阅 [先决条件](create-sso-office-add-ins-aspnet.md#prerequisites)。

在生产加载项时，加载项应该通过回退到备用的用户身份验证系统来响应此错误。欲知详情，请参阅 [要求和最佳做法](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)。


## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Azure Active Directory 服务器端错误

有关此部分中介绍的错误处理示例，请参阅：
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>条件访问/多重身份验证错误
 
在 AAD 和 Office 365 中的特定身份配置中， Microsoft Graph 可访问的某些资源可能需要多因素身份验证（MFA），即使用户的 Office 365 租户不具备。 当 AAD 通过代表流收到对受 MFA 保护的资源的令牌请求时，它会向您的加载项的 Web 服务返回包含 `claims` 属性的 JSON 消息。声明属性包含有关需要进一步的身份验证因素的信息。 

服务器端代码应测试此消息并将声明值转发到客户端代码。由于 Office 处理 SSO 加载项的身份验证，需要在客户端中使用此信息。发送给客户端的消息可以是错误 (如 `500 Server Error` 或 `401 Unauthorized`) 或也可以是成功响应的正文 (如 `200 OK`)。在任一情况下，代码的客户端 AJAX （故障或成功） 回调调用加载项的 web API 应该测试此响应。如果声明值已被转达，则代码应调回 `getAccessTokenAsync` 和传递在 [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 参数中的选项 `authChallenge: CLAIMS-STRING-HERE`。当 AAD 看到此字符串时，它会提示用户输入其他因素，然后返回一个新的访问令牌，该令牌将在代表流中被接受。

### <a name="consent-missing-errors"></a>缺少许可证错误

如果AAD没有记录用户（或租户管理员）已授权加载项（对Microsoft Graph资源），AAD会向网络服务器发送错误消息。代码必须告知客户端 (例如，正文中的 `403 Forbidden` 响应）通过 `forceConsent: true` 选项调回 `getAccessTokenAsync` 。

### <a name="invalid-or-missing-scope-permission-errors"></a>范围（权限）无效或缺失错误

- 服务器端代码应向客户端发送 `403 Forbidden` 响应，该响应应向用户呈现友好的信息。如果可能，请在控制台或日志中记录此错误。
- 确保加载项清单[Scopes](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/scopes?view=office-js)部分指定所有所需的权限。并确保加载项的 Web 服务的注册指定了相同的权限。同时检查拼写错误。欲知详情，请参阅 [使用Azure AD v2.0端点注册加载项](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint) (ASP.NET) 或 [使用Azure AD v2.0端点注册加载项](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint) ( Node JS )，以及 [配置加载项](create-sso-office-add-ins-aspnet.md#configure-the-add-in) ( ASP.NET ) 或 [配置加载项](create-sso-office-add-ins-nodejs.md#configure-the-add-in)( Node JS )。

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>调用 Microsoft Graph 时令牌过期或无效错误

某些身份验证和授权的库，包括 MSAL，必要时通过使用缓存的刷新令牌来防止过期令牌错误。还可以编写代码自己的令牌缓存系统。执行此操作的示例，请参阅 [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)，尤其是文件 [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts)。

不过，如果收到了令牌过期或令牌无效错误，代码必须指示客户端（例如，在 `401 Unauthorized` 响应的正文中）重新调用 `getAccessTokenAsync`，并重复调用加载项的Web API的端点，这会重复执行代表流来获取 Microsoft Graph 的新令牌。 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>调用 Microsoft Graph 时令牌无效错误

按照处理令牌到期错误的方法处理此错误。请参阅上一部分。

### <a name="invalid-audience-error"></a>受众无效错误

服务器端代码应向客户端发送 `403 Forbidden` 响应，向用户显示易记消息，并尽可能在控制台或日志中记录此错误。

若要详细了解如何添加令牌验证的多租户支持，请参阅 [Azure 多租户示例](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect)。
