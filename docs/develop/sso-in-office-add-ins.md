---
title: 为 Office 加载项启用单一登录
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: 1a75f7d619d2375a2f7fcb07f6afb7e0d6261ead
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579903"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a>为 Office 加载项启用单一登录（预览）

用户使用个人 Microsoft 账户或工作/学校（Office 365）账户登录到 Office （在线、移动设备和桌面平台） 。可以趁机使用单一登录 (SSO) 来授权用户使用加载项而无需用户二次登陆。

![显示加载项登录过程的图像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a>预览状态

单一登录 API 只支持预览。可供开发人员实验；但是，不应用于生产加载项。此外, 使用 SSO 的加载项不被 [AppSource](https://appsource.microsoft.com) 接受。

并非所有 Office 应用都支持 SSO 预览。在 Word、 Excel、 Outlook 和 PowerPoint 中可用。目前关于哪里支持单一登录 API 的详细信息，请参阅 [IdentityAPI 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js) 。

### <a name="requirements-and-best-practices"></a>要求和最佳做法

若要使用 SSO，必须从加载项启动 HTML 页的 `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` 中加载 beta 版的 Office JavaScript 库。

若正在使用 **Outlook** 加载项，须为 Office 365 租户启用新式身份验证。有关如何执行此操作的信息，请参阅 [Exchange Online：如何为租户启用新式身份验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx) 。

应该 *不* 依赖于 SSO 作为加载项的唯一的身份验证方法。当加载项在某种错误情况回退时，应执行一个备用的身份验证系统。可以使用用户表和身份验证系统，或利用一个社交登录服务商。欲知如何使用 Office 加载项执行此操作的详细信息，请参阅 [在 Office 加载项中授权外部服务](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins) 。对 *Outlook* 建议一个回退系统。欲知详情，请参阅 [方案：在 Outlook 加载项中实现服务单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) 。

### <a name="how-sso-works-at-runtime"></a>运行时 SSO 工作方式

以下关系图显示了 SSO 流程的工作方式。

![SSO 过程关系图](../images/sso-overview-diagram.png)

1. 在外接程序 JavaScript 调用新的 Office.js API [getAccessTokenAsync](#sso-api-reference)。这会告诉 Office 主机应用程序获取访问令牌到外接程序。请参阅 [示例访问令牌](#example-access-token)。
2. 如果用户未登录，Office 主机应用会打开弹出窗口，以供用户登录。
3. 如果当前用户是首次使用加载项，则会看到同意提示。
4. Office 主机应用程序从当前用户的 Azure AD v2.0 端点请求获取**加载项令牌**。
5. Azure AD 将加载项令牌发送给 Office 主机应用程序。
6. 作为 `getAccessTokenAsync` 调用返回的结果对象的一部分，Office 主机应用程序将**加载项令牌**发送给加载项。
7. 加载项中的 JavaScript 可以分析令牌并提取它所需的信息，如用户的电子邮件地址。 
8. 加载项可以向服务器端发送 HTTP 请求来获取更多用户数据；比如用户偏好。此外，访问令牌本身可被发送到服务器端进行分析和检验。 

## <a name="develop-an-sso-add-in"></a>开发 SSO 加载项

本节介绍创建使用 SSO Office 的加载项时的任务。这些任务再次以语言-框架-不可知论证的方式描述。详细步骤，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>创建服务应用程序

在 Azure v2.0 端点的注册门户注册加载项：https://apps.dev.microsoft.com。此过程用时 5-10 分钟，包括以下任务：

* 为加载项获取客户 ID 和机密。
* 加载项需要 AAD v. 2.0 终点 （可选 Microsoft Graph） 许可。始终需要"配置文件"权限。
* 向加载项授予 Office 主机应用信任。
* 将 Office 主机应用程序预授权给具有 *access_as_user* 默认权限的加载项。

有关此过程的更多详细信息，请参阅[注册使用 SSO 和 Azure AD v2.0 端点的 Office 加载项](register-sso-add-in-aad-v2.md)。

### <a name="configure-the-add-in"></a>配置加载项

向加载项清单添加新标记：

* **WebApplicationInfo** - 下列元素的母元素。
* **Id** - 客户加载项 ID ,这是注册加载项的申请 ID。参阅 [注册使用 Azure AD v2.0 终点的加载项](register-sso-add-in-aad-v2.md) 。
* **Resource** - 加载项 URL。
* **Scopes** - 一个或多个 **Scope** 元素的母元素.
* **范围** 指加载项需要 AAD 的许可。如果加载项不访问 Microsoft Graph的话， `profile` 就是始终并唯一需要的权限。如果是，可能还需要 **范围** 元素以获取所需的 Microsoft Graph 许可; 例如， `User.Read`， `Mail.Read`。用于访问 Microsoft Graph 的代码库可能需要其他权限。例如，适用于.NET 的 Microsoft 身份验证库 (MSAL) 需要 `offline_access` 权限。详细信息，参阅 [ Office 加载项授权 Microsoft Graph](authorize-to-microsoft-graph.md)。

对于除 Outlook 之外的 Office 主机，在 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 文末添加标记。对于 Outlook，在 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 文末添加标记。

以下是标记的示例：

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

将 JavaScript 添加到加载项，以执行以下操作：

* 调用 [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)。

* 分析访问令牌或将其传递给加载项的服务器端代码。 

下面是调用 `getAccessTokenAsync` 的简单例子。 

> [!NOTE]
> 此例只明显处理一种错误类型。更精细的错误处理的示例，请参阅 [Office-加载项-SPNET- SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) 和 [program.js in Office-加载项-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)。请参阅 [诊断错误消息的单一登录 (SSO)](troubleshoot-sso-in-office-add-ins.md)。
 

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

下面是将加载项令牌传递到服务器端的简单示例。需求被送回服务器端时，令牌标题为 `Authorization` 。本示例预期发送 JSON 数据，因此使用 `POST` 方法，但 `GET` 就足够在不写入服务器时发送访问令牌。

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

如果在没有用户登录到 Office 时无法使用加载项，应调用 `getAccessTokenAsync` *当加载项启动时* 。

如果加载项不需要登录某些功能，然后调用`getAccessTokenAsync`*如果用户执行需要登录的用户的操作* 。没有冗余的调用不会显著的性能下降`getAccessTokenAsync`因为 Office 缓存访问令牌并将反复使用，直到过期，而不会每次`getAccessTokenAsync`被调用时都尝试调用 AAD v. 2.0 终结点。这样可将添加调用`getAccessTokenAsync`到所有需要令牌的功能和处理程序来启动操作。

### <a name="add-server-side-code"></a>添加服务器端代码

如果加载项没有将访问令牌传递到服务器端并在服务器端使用它，多数情况下获取访问令牌意义极小。加载项可以在服务器端做如下事情：

* 多创建一个或多个 Web API 方法，该方法从令牌中提取客户信息；比如，从托管数据库查看用户偏好（参阅如下 **使用 SSO 令牌作为身份** ）。根据语言和框架，库可用来简化代码编写。
* 获取 Microsoft Graph 数据。服务器端代码应执行以下操作：

    * 验证访问令牌（请参阅如下 **验证访问令牌** ）。
    * 开始“代表”流调用 Azure AD v2.0 结点，包括访问令牌、 用户元数据，以及加载项身份 （其 ID 和机密）。在此上下文中，访问令牌被称为 bootstrap 令牌。
    * 缓存从代表流程返回的新访问令牌。
    * 使用新令牌从 Microsoft Graph 获取数据。

 有关获得对用户 Microsoft Graph 数据的授权访问的更多详细信息，请参阅[在 Office 加载项中授权给 Microsoft Graph](authorize-to-microsoft-graph.md)。

#### <a name="validate-the-access-token"></a>验证访问令牌

一旦 Web API 接收访问令牌，必须在使用它之前进行验证。该标记是 JSON Web 令牌 (JWT)，这意味着验证与多数标准的 OAuth 流中的令牌验证一致。有许多库可用来处理 JWT 验证，但基础库包括：

- 检查令牌的格式是否正确
- 检查令牌是否由预期的颁发机构颁发
- 检查令牌是否是针对 Web API

验证令牌时，请记住以下准则：

- 有效的 SSO 令牌于 Azure 权威颁发， `https://login.microsoftonline.com`。 `iss` 声明令牌的值应由此开始。
- 令牌的 `aud` 参数将用来设置加载项注册的应用程序 ID。
- 将令牌的 `scp` 参数设置为 `access_as_user`。

#### <a name="using-the-sso-token-as-an-identity"></a>将 SSO 令牌用作标识

如果加载项需要验证用户的身份，SSO 令牌包含可以用于建立标识的信息。令牌中与标识相关的有如下几点。

- `name` - 用户的显示名称。
- `preferred_username` - 用户的电子邮件地址。
- `oid` - 表示 Azure Active Directory 中的用户 ID 的 GUID。
- `tid` - 表示 Azure Active Directory 中的用户组织 ID 的 GUID。

因为 `name` 和 `preferred_username`值可以更改，我们建议使用 `oid`和 `tid` 值将标识与后端的授权服务关联。

例如，服务器无法格式化这些值如 `{oid-value}@{tid-value}`，那么将这些值存储为内部用户数据库的记录。后续若有请求时，可以通过检索相同的值来找到用户，并根据现有访问机制决定获取特定资源的访问权限。

### <a name="example-access-token"></a>访问令牌样本

下面是典型的解码有效负载的访问令牌。欲知属性的信息，请参阅 [Azure Active Directory v2.0 令牌参照](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens) 。


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

在 Outlook 加载项中使用与在 Excel、 PowerPoint、或 Word 加载项中使用 SSO 之间有一些细微、但很重要的差别。请务必阅读 [使用 Outlook 加载项中的单一登录令牌对用户进行身份验证](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)和[方案：在 Outlook 加载项中对你的服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。

## <a name="sso-api-reference"></a>SSO API 参照

### <a name="getaccesstokenasync"></a>getAccessTokenAsync

Office Auth 命名空间， `Office.context.auth`，提供了 Office 主机获取登陆令牌以完成加载项网申的方法 `getAccessTokenAsync` 。并且非直接地允许加载项访问已登录的用户的 Microsoft Graph 数据而无需用户再次登录。

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

该方法调用了 Azure Active Directory V 2.0 终结点来获取加载项网络申请的登陆令牌。使加载项可以识别用户。服务器端代码可以使用此令牌，用["代表" OAuth 流](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) 接入 Microsoft Graph。

> [!NOTE]
> 在 Outlook 中，如果加载项加载于 Outlook.com 或 Gmail 邮箱，此 API 不受支持。

<table><tr><td>Hosts</td><td>Excel、Outlook、PowerPoint、Word</td></tr>

 <tr><td>[要求集](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)</td></tr></table>

#### <a name="parameters"></a>参数

`options` - 可选。 接受 `AuthOptions` 对象 （请参阅下方）以定义登录行为。

`callback` - 可选。接受回调方法分析令牌，用用户 ID 或用“代表”流进入 Microsoft Graph。如果 AsyncResult  为"成功"，则 是原始 AAD v. 2.0-被格式化的登陆令牌。[ ](https://docs.microsoft.com/javascript/api/office/office.asyncresult) `.status` `AsyncResult.value`

当 Office 得到 AAD v. 2.0 通过 `getAccessTokenAsync` 方法获取的加载项登陆令牌， `AuthOptions` 交互就有了不同的用户体验。

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



