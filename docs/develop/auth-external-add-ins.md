---
title: 在 Office 加载项中授权外部服务
description: 获得对非 Microsoft 数据的授权，如 Google、Facebook、LinkedIn、SalesForce 和使用 OAuth 2.0、授权代码和隐式流的 GitHub。
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 3b6ba291912f9d6284ca5e71a7bf634261a0bd24
ms.sourcegitcommit: 24303ca235ebd7144a1d913511d8e4fb7c0e8c0d
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2019
ms.locfileid: "36838541"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>在 Office 加载项中授权外部服务

借助热门在线服务（包括 Office 365、Google、Facebook、领英、SalesForce 和 GitHub），开发人员可以授权用户访问自己在其他应用中的帐户。这样一来，就可以在 Office 加载项中添加这些服务了。

> [!NOTE]
> 本文的其余部分涉及的是访问非 Microsoft 服务。 有关访问 Microsoft Graph（包括Office 365）的信息，请参阅[使用 SSO 访问 Microsoft Graph](overview-authn-authz.md#access-to-microsoft-graph-with-sso) 和[不使用 SSO 访问 Microsoft Graph](overview-authn-authz.md#access-to-microsoft-graph-without-sso)。

授权 Web 应用访问在线服务的行业标准框架为 **OAuth 2.0**。大多数情况下，无需了解框架的详细工作原理，即可在加载项中使用它。许多库都可用来化繁为简。

OAuth 的基本概念是，应用程序本身可以是一个[安全主体](/windows/security/identity-protection/access-control/security-principals)，就像一个用户或组，拥有其自己的标识和权限集。 在最典型的应用场景中，当用户在需要联机服务的 Office 加载项中进行操作时，加载项会向服务发送请求，请求为用户帐户提供一组特定权限。 然后，该服务会提示用户向加载项授予这些权限。 授予权限之后，该服务会向外接程序发送一个小的编码*访问令牌*。 外接程序可以通过在其向服务 API 发送的所有请求中包含令牌来使用该服务。 但外接程序只能在用户授予它的权限范围内进行操作。 令牌还会在某个指定时间后过期。

几种称为*流*或*授权类型*的 OAuth 模式专为不同方案而设计。 以下两种模式最常实现：

- **隐式流**：加载项与在线服务的通信是通过客户端 JavaScript 实现。 此流常用于单页应用程序 (SPA)。
- **授权代码流**：外接程序的 Web 应用程序和联机服务之间的通信是*服务器到服务器*。 因此，它是通过服务器端代码实现。

OAuth 流旨在保护应用程序的标识和授权。 授权代码流中提供了需要一直隐藏的*客户端密码*。 由于没有服务器端后端的应用程序（如 SPA）无法保护密码，因此建议在 SPA 中使用隐式流。

应熟悉隐式流和授权代码流的利与弊。 若要详细了解这两个流，请参阅[授权代码流](https://tools.ietf.org/html/rfc6749#section-1.3.1)和[隐式流](https://tools.ietf.org/html/rfc6749#section-1.3.2)。

> [!NOTE]
> 还可以视需要使用中间人服务，从而执行授权操作，并将访问令牌传递给加载项。 有关此方案的详细信息，请参阅本文稍后介绍的**中间人服务**部分。

## <a name="using-the-implicit-flow-in-office-add-ins"></a>在 Office 外接程序中使用隐式流

若要确定在线服务是否支持隐式流，最好是查阅服务文档。

有关支持隐式流的库的信息，请参阅本文后面的**库**部分。

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>在 Office 加载项中使用授权代码流

许多库都可用于在各种语言和框架中实现授权代码流。若要详细了解其中某些库，请参阅本文稍后将介绍的**库**部分。

## <a name="libraries"></a>库

库适用于许多语言和平台，既可用于隐式流，也可用于授权代码流。 一些库是通用的，而另一些库则为在线服务专用。

**Google**：在 [GitHub.com/Google](https://github.com/google) 中搜索 "auth" 或你语言的相应名称。大部分的相关存储库被命名为 `google-auth-library-[name of language]`。

**Facebook**：在 [Facebook 开发者](https://developers.facebook.com) 中搜索 "library" 或 "sdk"。

**常规 OAuth 2.0**：指向十几种语言库的链接页面由 IETF OAuth 工作组在以下位置进行维护：[OAuth 代码](https://oauth.net/code/)。请注意，其中一些库可用来实现 OAuth 兼容服务。作为外接程序开发人员，你所感兴趣的库就是此页上称为*客户端*的库，因为 Web 服务器是 OAuth 兼容服务的客户端。

## <a name="middleman-services"></a>中间人服务

加载项可以使用中间人服务（如 [OAuth.io](https://oauth.io) 或 [Auth0](https://auth0.com)）执行授权。中间人服务可以提供热门在线服务的访问令牌，和/或简化加载项社交登录的启用过程。通过极少量的代码，加载项就可以使用客户端脚本或服务器端代码，连接到中间人服务，然后中间人服务会向加载项发送所需的任何在线服务令牌。所有授权实现代码都位于中间人服务中。 

我们建议外接程序中用于身份验证/授权的 UI 使用对话框 API 打开登录页面。 有关详细信息，请参阅[在身份验证流中使用对话框 API](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow)。 以这种方式打开 Office 对话框时，对话框具有全新和单独的浏览器实例，以及父页面的实例中的 JavaScript 引擎（如外接程序的任务窗格或 FunctionFile）。 一个标记以及可转换为字符串的其他信息被传递回使用名为 `messageParent` 的 API 的父页面。 然后父页面可以使用标记对资源进行经过授权的调用。 由于此体系结构，用户必须谨慎地使用中间人服务提供的 API。 服务通常会提供 API 集，其中代码创建某种上下文对象，该对象获取标记并使用该标记对资源进行后续调用。 该服务通常具有单个 API 方法，该方法进行初始调用并创建上下文对象**。 此类对象无法完全字符串化，因此无法从 Office 对话框传递到父页面。 通常，中间人服务在较低抽象级别提供第二个 API 集，例如 REST API。 第二个集将具有从该服务获取标记的 API，以及获取对资源的授权访问权限时将标记传递到服务的其他 API。 需要在此较低抽象级别使用 API，以便在 Office 对话框中获取标记并使用 `messageParent` 将其传递到父页面。 

## <a name="what-is-cors"></a>什么是 CORS？

CORS 的全称是 [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS)，即“跨源资源共享”。若要了解如何在加载项内使用 CORS，请参阅[解决 Office 加载项中的同源策略限制](addressing-same-origin-policy-limitations.md)。
