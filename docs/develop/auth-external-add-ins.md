---
title: 在 Office 加载项中授权外部服务
description: 获得对非 Microsoft 数据的授权，如 Google、Facebook、LinkedIn、SalesForce 和使用 OAuth 2.0、授权代码和隐式流的 GitHub。
ms.date: 08/07/2019
localization_priority: Normal
ms.openlocfilehash: fd180e11106e7e1e2f20f539746535c4310ad81e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093740"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>在 Office 加载项中授权外部服务

Popular online services, including Microsoft 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in.

> [!NOTE]
> 本文的其余部分涉及的是访问非 Microsoft 服务。 若要了解如何访问 Microsoft Graph (包括 Microsoft 365) ，请参阅[使用 Sso 访问 Microsoft graph](overview-authn-authz.md#access-to-microsoft-graph-with-sso) ，并[访问 microsoft graph 而不使用 sso](overview-authn-authz.md#access-to-microsoft-graph-without-sso)。

The industry standard framework for enabling web application access to an online service is **OAuth 2.0**. In most situations, you don't need to know the details of how the framework works to use it in your add-in. Many libraries are available that simplify the details for you.

OAuth 的基本概念是，应用程序本身可以是一个[安全主体](/windows/security/identity-protection/access-control/security-principals)，就像一个用户或组，拥有其自己的标识和权限集。 在最典型的应用场景中，当用户在需要联机服务的 Office 加载项中进行操作时，加载项会向服务发送请求，请求为用户帐户提供一组特定权限。 然后，该服务会提示用户向加载项授予这些权限。 授予权限之后，该服务会向外接程序发送一个小的编码*访问令牌*。 外接程序可以通过在其向服务 API 发送的所有请求中包含令牌来使用该服务。 但外接程序只能在用户授予它的权限范围内进行操作。 令牌还会在某个指定时间后过期。

几种称为*流*或*授权类型*的 OAuth 模式专为不同方案而设计。 以下两种模式最常实现：

- **隐式流**：加载项与在线服务的通信是通过客户端 JavaScript 实现。 此流常用于单页应用程序 (SPA)。
- **授权代码流**：外接程序的 Web 应用程序和联机服务之间的通信是*服务器到服务器*。 因此，它是通过服务器端代码实现。

OAuth 流旨在保护应用程序的标识和授权。 授权代码流中提供了需要一直隐藏的*客户端密码*。 由于没有服务器端后端的应用程序（如 SPA）无法保护密码，因此建议在 SPA 中使用隐式流。

应熟悉隐式流和授权代码流的利与弊。 若要详细了解这两个流，请参阅[授权代码流](https://tools.ietf.org/html/rfc6749#section-1.3.1)和[隐式流](https://tools.ietf.org/html/rfc6749#section-1.3.2)。

> [!NOTE]
> You also have the option of using a middleman service to perform authorization and pass the access token to your add-in. For details about this scenario, see the **Middleman services** section later in this article.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>在 Office 外接程序中使用隐式流

若要确定在线服务是否支持隐式流，最好是查阅服务文档。

有关支持隐式流的库的信息，请参阅本文后面的**库**部分。

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>在 Office 加载项中使用授权代码流

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For more information about some of these libraries, see the **Libraries** section later in this article.

## <a name="libraries"></a>库

Libraries are available for many languages and platforms, for both the Implicit flow and the Authorization Code flow. Some libraries are general purpose, while others are for specific online services.

**Google**: Search [GitHub.com/Google](https://github.com/google) for "auth" or the name of your language. Most of the relevant repos are named `google-auth-library-[name of language]`.

**Facebook**：在 [Facebook 开发者](https://developers.facebook.com) 中搜索 "library" 或 "sdk"。

**General OAuth 2.0**: A page of links to libraries for over a dozen languages is maintained by the IETF OAuth Working Group at: [OAuth Code](https://oauth.net/code/). Note that some of these libraries are for implementing an OAuth compliant service. The libraries of interest to you as a an add-in developer are called *client* libraries on this page because your web server is a client of the OAuth compliant service.

## <a name="middleman-services"></a>中间人服务

Your add-in can use a middleman service such as [OAuth.io](https://oauth.io) or [Auth0](https://auth0.com) to perform authorization. A middleman service may either provide access tokens for popular online services or simplify the process of enabling social login for your add-in, or both. With very little code, your add-in can use either client-side script or server-side code to connect to the middleman service and it will send your add-in any required tokens for the online service. All of the authorization implementation code is in the middleman service. 

我们建议外接程序中用于身份验证/授权的 UI 使用对话框 API 打开登录页面。 有关详细信息，请参阅[在身份验证流中使用对话框 API](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow)。 以这种方式打开 Office 对话框时，对话框具有全新和单独的浏览器实例，以及父页面的实例中的 JavaScript 引擎（如外接程序的任务窗格或 FunctionFile）。 一个标记以及可转换为字符串的其他信息被传递回使用名为 `messageParent` 的 API 的父页面。 然后父页面可以使用标记对资源进行经过授权的调用。 由于此体系结构，用户必须谨慎地使用中间人服务提供的 API。 服务通常会提供 API 集，其中代码创建某种上下文对象，该对象获取标记并使用该标记对资源进行后续调用。 该服务通常具有单个 API 方法，该方法进行初始调用并创建上下文对象**。 此类对象无法完全字符串化，因此无法从 Office 对话框传递到父页面。 通常，中间人服务在较低抽象级别提供第二个 API 集，例如 REST API。 第二个集将具有从该服务获取标记的 API，以及获取对资源的授权访问权限时将标记传递到服务的其他 API。 需要在此较低抽象级别使用 API，以便在 Office 对话框中获取标记并使用 `messageParent` 将其传递到父页面。 

## <a name="what-is-cors"></a>什么是 CORS？

CORS stands for [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). For information about how to use CORS inside add-ins, see [Addressing same-origin policy limitations in Office Add-ins](addressing-same-origin-policy-limitations.md).
