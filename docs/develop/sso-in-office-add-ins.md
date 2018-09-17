---
title: 为 Office 加载项启用单一登录
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 534ac41e7518756a2aa5b4408ce7adb0f434e27d
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945339"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>为 Office 加载项启用单一登录（预览）

用户可以使用自己的个人 Microsoft 帐户/工作或学校 (Office 365) 帐户，登录 Office（在线、移动和桌面平台）。 你可以利用这一优势并使用单一登录（SSO）授权用户访问你的加载项，无需用户再次登录。


![显示加载项登录过程的图像](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 预览版支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets?view=office-js)。
> 若要使用 SSO，您必须从加载项启动 HTML 页的 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 中加载 beta 版的 Office JavaScript 库。
> 如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。 若要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

对于用户来说，这只涉及一次登录，因而使得运行加载项的体验更流畅。 对于开发人员来说，这意味着你的加载项不必使用加密的密码维护自己的用户表。

### <a name="how-it-works-at-runtime"></a>运行时的工作方式

以下关系图显示了 SSO 流程的工作方式。

![SSO 过程关系图](../images/sso-overview-diagram.png)

1. 在加载项中， JavaScript 调用新的 Office.js API [getAccessTokenAsync](#sso-api-reference)。 这会指示 Office 主机应用程序获取对加载项的访问令牌。 请参阅[示例访问令牌](#example-access-token)。
2. 如果用户未登录，Office 主机应用会打开弹出窗口，以供用户登录。
3. 如果当前用户是首次使用加载项，则会看到同意提示。
4. Office 主机应用程序从当前用户的 Azure AD v2.0 终结点请求获取**加载项令牌**。
5. Azure AD 将加载项令牌发送给 Office 主机应用程序。
6. Office 主机应用程序在 `getAccessTokenAsync` 调用返回的结果对象中，将**加载项令牌**发送给加载项。
7. 加载项中的 JavaScript 可以分析令牌并提取它所需的信息，如用户的电子邮件地址。 
8. 可选地，加载项可以向其服务器端发送 HTTP 请求以获取关于用户的更多数据；如用户的首选项。 或者，访问令牌本身可以发送到服务器端进行分析和验证。 

## <a name="develop-an-sso-add-in"></a>开发 SSO 加载项

此部分介绍了创建启用 SSO 的 Office 加载项所需完成的任务。 其中介绍的这些任务与语言和框架无关。 有关详细演练的示例，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>创建服务应用程序

在 Azure v2.0 端点的注册门户注册加载项：https://apps.dev.microsoft.com。 这是一个 5 – 10 分钟过程，包括以下任务：

* 获取外接程序的客户端 ID 和机密。
* 指定加载项访问 AAD v. 2.0 端点（以及可选的 Microsoft Graph）所需的权限。 总是需要“个人资料”权限。
* 向外接程序授予 Office 主机应用程序信任。
* 将 Office 主机应用程序预授权给具有 *access_as_user* 默认权限的外接程序。

有关此过程的更多详细信息，请参阅[注册使用 SSO 和 Azure AD v2.0 端点的 Office 加载项](register-sso-add-in-aad-v2.md)。

### <a name="configure-the-add-in"></a>配置外接程序

向外接程序清单添加新标记：

* **WebApplicationInfo** —下列元素的父元素。
* **ID** - 加载项的客户端 ID 这是你在注册加载项的过程中获得的应用程序 ID。 请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项](register-sso-add-in-aad-v2.md)。
* **Resource** —加载项 URL。
* **Scopes** —一个或多个 **Scope**元素的父元素。
* **Scope** —指定加载项访问 AAD 所需的权限。 权限总是需要的，如果加载项不能访问 Microsoft Graph，它可能是唯一需要的权限。`profile` 如果是这样，你也需要用于所需 Microsoft Graph 权限的 **作用域** 元素；例如：`User.Read`、`Mail.Read`。 你在代码中用于访问 Microsoft Graph 的库可能需要额外的权限。 例如，用于 .NET 的 Microsoft 身份验证库（MSAL）需要 `offline_access` 权限。 有关更多信息，请参阅 [从 Office 加载项授权给 Microsoft Graph](authorize-to-microsoft-graph.md)。

对于除 Outlook 之外的 Office 主机，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 部分的末尾。对 Outlook，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 部分的末尾。

下面的示例展示了标记：

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a>添加客户端代码

将 JavaScript 添加到外接程序，以执行以下操作：

* 调用 [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)。

* 分析访问令牌或将其传递给加载项的服务器端代码。 

下面是调用 `getAccessTokenAsync` 的简单例子。 

> [!NOTE]
> 这个例子明确地只处理一种错误。 有关更详细的错误处理示例，请参阅 [Office-Add-in-ASPNET-SSO 中的 Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) 和 [Office-Add-in-NodeJS-SSO 中的 program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)。 并参阅[单一登录（SSO）错误消息故障排除](troubleshoot-sso-in-office-add-ins.md)。
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

下面是将加载项令牌传递给服务器端的一个简单示例。 令牌包含在向服务器端发回请求时的 `Authorization` 标头中。 这个例子设想发送 JSON 数据，所以使用了 `POST` 方法，但是当不写入服务器时 `GET` 足以发送访问令牌。

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a>何时调用方法

如果在没有用户登录到 Office 时无法使用加载项，那么`getAccessTokenAsync` *当加载项启动时*你应该调用。

如果加载项具有某些不需要登录用户的功能，那么`getAccessTokenAsync` *当用户采取需要登录用户的操作时*你可以调用。 `getAccessTokenAsync` 由于 Office 缓存访问令牌，并且将重用它，直到它过期，而没有再次调用 AAD v.，因此没有显著的性能下降。 在 调用 `getAccessTokenAsync` 时重新调用 AAD v.2.0 端点。 因此，可以将 `getAccessTokenAsync` 调用添加到所有在需要令牌时启动操作的函数和处理程序。

### <a name="add-server-side-code"></a>添加服务器端代码

大多数情况下，如果加载项没有将访问令牌传递到服务器端并在其中使用它，那么获取访问令牌的意义就不大。 加载项可以执行的一些服务器端任务：

* 创建一个或多个 Web API 方法，使用从令牌中提取的有关用户的信息；例如，在托管数据库中查找用户首选项的方法。 （请参阅下面的 **使用 SSO 令牌作为标识** 。）根据你的语言和框架，可能会有库提供，因而简化你必须编写的代码。
* 获取 Microsoft Graph 数据。 服务器端代码应执行以下操作：

    * 验证访问令牌（请参阅下面的 **验证访问令牌** ）。
    * 通过调用 Azure AD v2.0 端点来启动＂代表＂流程，此端点包含访问令牌、一些关于用户的元数据以及加载项凭据（ID 和机密）。 在这种情况下，访问令牌称为引导令牌。
    * 缓存从代表流程返回的新访问令牌。
    * 使用新令牌从 Microsoft Graph 获取数据。

 有关获得用户 Microsoft Graph 数据的授权访问的更多详细信息，请参阅[在 Office 加载项中授权给 Microsoft Graph](authorize-to-microsoft-graph.md)。

#### <a name="validate-the-access-token"></a>验证访问令牌

Web API 收到访问令牌后，必须在使用该令牌前对其进行验证。 该令牌是 JSON Web 令牌 (JWT)，这意味着验证方式与大多数标准 OAuth 流中的令牌验证方式类似。 有许多可用于处理 JWT 验证的库，而它们的基本内容为：

- 检查令牌的格式是否正确
- 检查令牌是否由预期的颁发机构颁发
- 检查令牌是否是针对 Web API

验证令牌时，请牢记以下准则：

- 有效的 SSO 令牌是由 Azure 颁发机构 `https://login.microsoftonline.com` 的。 令牌中的 `iss` 声明应以此值开头。
- 令牌的 `aud` 参数将被设置为加载项注册的应用程序 ID。
- 令牌的 `scp` 参数将被设置为 `access_as_user`。

#### <a name="using-the-sso-token-as-an-identity"></a>将 SSO 令牌用作标识

如果加载项需要验证用户标识，则 SSO 令牌包含的信息可用于创建此标识。 令牌中的以下声明与标识相关。

- `name` —用户的显示名称。
- `preferred_username` —用户的电子邮件地址。
- `oid` —表示 Azure Active Directory 中的用户 ID 的 GUID。
- `tid` —表示 Azure Active Directory 中的用户组织 ID 的 GUID。

因为 `name` 和 `preferred_username`值可以更改，我们建议使用 `oid`和 `tid` 值将标识与后端的授权服务关联。

例如，你的服务可以将这些值组合在一起，并设置为类似 `{oid-value}@{tid-value}` 的格式，然后将其存储为内部用户数据库中的用户记录值。 然后，在后续的请求中，可以使用同一值检索此用户，并可基于现有访问控制机制确定对特定资源的访问。

### <a name="example-access-token"></a>示例访问令牌

以下是访问令牌的典型解码有效负载。 有关这些属性的信息，请参阅 [Azure Active Directory v2.0 令牌引用](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)。


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",         
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",         
    iat: 1521143967,         
    nbf: 1521143967,         
    exp: 1521147867,         
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",         
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",         
    azpacr: "0",         
    e_exp: 262800,         
    name: "Mila Nikolova",         
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",         
    preferred_username: "milan@contoso.com",         
    scp: "access_as_user",         
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",         
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",         
    uti: "MICAQyhrH02ov54bCtIDAA",         
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a>在 Outlook 加载项中使用 SSO

在 Outlook 加载项中使用 SSO 与在 Excel、PowerPoint 或Word 加载项中使用，存在一些小而重要的差异。 一定要阅读[使用 Outlook 加载项中的单一登录令牌对用户进行身份验证](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)和[方案：在 Outlook 加载项中对你的服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。

## <a name="sso-api-reference"></a>SSO API 引用

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Office 身份验证命名空间 `Office.context.auth` 提供了方法 `getAccessTokenAsync`，使 Office 主机可以获取加载项 Web 应用程序的访问令牌。 这样间接地使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户二次登录。

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

调用 Azure Active Directory V 2.0 端点以获取令牌来访问加载项的 Web 应用程序。 这样加载项便能识别用户。 通过[“代表”OAuth 流](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)，服务器端代码可以使用此令牌访问加载项 Web 应用程序的 Microsoft Graph。

> [!NOTE]
> 在 Outlook 中，如果加载项加载于 Outlook.com 或 Gmail 邮箱，此 API 将不受支持。

<table><tr><td>主机</td><td>Excel、Outlook、PowerPoint、Word</td></tr>

 <tr><td>要求集</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a>参数

`options` - 可选。 接受 `AuthOptions` 对象 （请参阅下面）以定义登录行为。

`callback` - 可选。 接受一个回调方法，可为用户的 ID 分析令牌或使用“代表”流中的令牌获取 Microsoft Graph 访问权限。 |||UNTRANSLATED_CONTENT_START|||If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.|||UNTRANSLATED_CONTENT_END||| 2.0 格式的访问令牌。

接口提供了 Office 从 AAD v.`AuthOptions` 2.0 通过 `getAccessTokenAsync` 方法获取加载项访问令牌时的用户体验选项。

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



