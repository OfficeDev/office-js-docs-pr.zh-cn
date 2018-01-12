# <a name="authorize-external-services-in-your-office-add-in"></a>在 Office 加载项中授权外部服务

使用常用在线服务（包括 Office 365、Google、Facebook、LinkedIn、SalesForce 和 GitHub），开发者可以让用户登录他/她在其他应用程序中的帐户。 这样，便可以在 Office 加载项中添加这些服务。

>**注意：**如果外部服务（如 Office 365 或 OneDrive）可以通过 Microsoft Graph 进行访问，便能使用[为 Office 加载项启用单一登录](http://dev.office.com/docs/add-ins/develop/sso-in-office-add-ins)及其相关文章中介绍的单一登录系统，既可以为用户提供最佳体验，也能够最大限度地简化自己的开发体验。 本文介绍的技术非常适用于不可通过 Microsoft Graph 访问的外部服务。 不过，这些服务*可*用于访问 Microsoft Graph，就这一点而言，可能会更青睐它们，而无视单一登录具有的优势。 例如，单一登录系统需要使用服务器端代码，因此无法用于真正的单页应用程序。 此外，并不是所有平台都支持单一登录系统。

启用 Web 应用程序对在线服务的访问权限的行业标准框架是 **OAuth 2.0**。 在大多数情况下，无需详细了解此框架的具体工作原理，即可在加载项中使用它。 可以使用许多库来简化需要了解的详细信息。

OAuth 的基本概念是，应用程序本身可以是一个安全主体，就像一个用户或组，拥有其自己的标识和权限集。 在最典型的应用场景中，当用户在需要联机服务的 Office 外接程序中进行操作时，外接程序会向服务发送请求，请求为用户帐户提供一组特定权限。 然后，该服务会提示用户向外接程序授予这些权限。 授予权限之后，该服务会向外接程序发送一个小的编码*访问令牌*。 外接程序可以通过在其向服务 API 发送的所有请求中包含令牌来使用该服务。 但外接程序只能在用户授予它的权限范围内进行操作。 令牌还会在某个指定时间后过期。

几种称为*流*或*授权类型*的 OAuth 模式专为不同方案而设计。 以下两种模式最常实现：

- **隐式流**：外接程序和联机服务之间的通信通过客户端 JavaScript 实现。
- **授权代码流**：外接程序的 Web 应用程序和联机服务之间的通信是*服务器到服务器*。 因此，它是通过服务器端代码实现。

OAuth 流旨在保护应用程序的标识和授权。 授权代码流中提供了需要一直隐藏的*客户端密码*。 由于单页应用程序 (SPA) 无法保护密码，因此建议在 SPA 中使用隐式流。

应熟悉隐式流和授权代码流的利与弊。 若要详细了解这两个流，请参阅[授权代码](https://tools.ietf.org/html/rfc6749#section-1.3.1)和[隐式](https://tools.ietf.org/html/rfc6749#section-1.3.2)。

>**注意：**还可以视需要使用中间人服务，执行授权操作，并将访问令牌传递给加载项。 有关此方案的详细信息，请参阅本文稍后介绍的**中间人服务**部分。

## <a name="using-the-implicit-flow-in-office-add-ins"></a>在 Office 外接程序中使用隐式流
若要确定在线服务是否支持隐式流，最好是查阅服务文档。 对于支持隐式流的服务，可以使用 **Office-js-helpers** JavaScript 库，完成所有细节工作：

- [Office-js-helpers](https://github.com/OfficeDev/office-js-helpers)

若要了解支持隐式流的其他库，请参阅本文稍后介绍的**库**部分。

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>在 Office 加载项中使用授权代码流

许多库都可通过各种语言和框架实现授权代码流。 若要详细了解其中部分库，请参阅本文稍后介绍的**库**部分。

下面为实现授权代码流的加载项示例：

- [Office-Add-in-Nodejs-ServerAuth](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth) (NodeJS)
- [PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart) (ASP.NET MVC)

### <a name="relayproxy-functions"></a>中继/代理函数

通过存储托管在如 [Azure 函数](https://azure.microsoft.com/en-us/services/functions)或 [Amazon Lambda](https://aws.amazon.com/lambda) 服务的简单函数中的**客户端 ID** 和**客户端密码**值，甚至可以在无服务器的 Web 应用程序上使用授权代码流。
此函数用给定代码交换**访问令牌**，并将它中继回客户端。 这种方法的安全性取决于对函数访问的保护程度。

为了使用此技术，加载项会显示 UI/弹出窗口，提供在线服务（如 Google、Facebook 等）的登录屏幕。 如果用户登录并授权加载项访问在线服务中的资源，加载项就会收到代码，然后可以将它发送给在线函数。 在本文稍后介绍的**中间人服务**部分中，涉及的服务就是使用类似的流。

## <a name="libraries"></a>库

库适用于许多语言和平台，既可用于隐式流，也可用于授权代码流。 一些库用于一般用途，另一些库专用于特定在线服务。

**将 Azure Active Directory 用作授权提供程序的 Office 365 及其他服务**：[Azure Active Directory 授权库](https://azure.microsoft.com/en-us/documentation/articles/active-directory-authentication-libraries/)。 预览也适用于 [Microsoft 身份验证库](https://www.nuget.org/packages/Microsoft.Identity.Client)。

**Google**：在 [GitHub.com/Google](https://github.com/google) 中搜索 "auth" 或你语言的相应名称。 大部分的相关存储库被命名为 `google-auth-library-[name of language]`。

**Facebook**：在 [Facebook 开发者](https://developers.facebook.com) 中搜索 "library" 或 "sdk"。

**常规 OAuth 2.0**：指向十几种语言库的链接页面由 IETF OAuth 工作组在以下位置进行维护：[OAuth 代码](http://oauth.net/code/)。 请注意，其中一些库可用于实现符合 OAuth 标准的服务。 作为外接程序开发人员，你所感兴趣的库就是此页上称为*客户端*的库，因为 Web 服务器是 OAuth 兼容服务的客户端。

## <a name="middleman-services"></a>中间人服务

加载项可以使用中间人服务（如 OAuth.io 或 Auth0）执行授权操作。 中间人服务可以提供热门在线服务的访问令牌，和/或简化为加载项启用社交登录的过程。 通过极少量的代码，加载项可以使用客户端脚本或服务器端代码连接到中间人服务，它会向加载项发送在线服务所需的全部令牌。 所有授权实现代码都位于中间人服务中。

有关使用中间人服务进行授权的加载项示例，请参阅以下示例：

- [Office-Add-in-Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0) 使用 Auth0 启用 Facebook、Google 和 Microsoft 帐户社交登录。

- [Office-Add-in-OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io) 使用 OAuth.io 从 Facebook 和 Google 获取访问令牌。

## <a name="what-is-cors"></a>什么是 CORS？

CORS 的全称是[“Cross Origin Resource Sharing”（即跨域资源共享）](https://developer.mozilla.org/en-US/docs/Web/HTTP/Access_control_CORS)。 若要了解如何在加载项内使用 CORS，请参阅[解决 Office 加载项中的同域策略限制](http://dev.office.com/docs/add-ins/develop/addressing-same-origin-policy-limitations)。
