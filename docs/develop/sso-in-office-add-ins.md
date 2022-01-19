---
title: 为 Office 加载项启用单一登录
description: 了解如何使用常用的 Microsoft 个人、工作或教育帐户来为 Office 加载项启用单一登录。
ms.date: 09/03/2021
ms.localizationpriority: high
ms.openlocfilehash: c371372bc954496ccbce12f65191c76e01ce0bd2
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074264"
---
# <a name="enable-single-sign-on-for-office-add-ins"></a>为 Office 加载项启用单一登录

用户可以使用自己的个人 Microsoft 帐户/Microsoft 365 教育版或工作帐户，登录 Office（在线、移动和桌面平台）。可以利用这一点，使用单点登录 (SSO) 将用户授权给加载项，让用户无需登录第二次。

![显示加载项登录过程的图像。](../images/sso-for-office-addins.png)

## <a name="requirements-and-best-practices"></a>要求和最佳做法

如果使用的是 **Outlook** 加载项，请务必为 Microsoft 365 租赁启用新式验证。 若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。

*不应* 依赖 SSO 作为加载项的唯一身份验证方法。 应实现备用身份验证系统，在某些错误情况下，加载项可以返回到该系统。 可以使用包含用户表和身份验证的系统，也可以利用其中某个社交登录提供者。 有关如何使用 Office 插件进行此操作的详细信息，请参见[授权 Office 加载项中的外部服务](auth-external-add-ins.md)。 对于 *Outlook*，建议使用回退系统。 有关详细信息，请参阅[应用场景：在 Outlook 外接程序中对服务实现单一登录](../outlook/implement-sso-in-outlook-add-in.md)。 有关使用 Azure Active Directory 作为回退系统的示例，请参阅 [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) 和 [Office 加载项 ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)。

## <a name="how-sso-works-at-runtime"></a>运行时 SSO 的工作方式

以下关系图显示了 SSO 流程的工作方式。

![显示 SSO 流程的关系图。](../images/sso-overview-diagram.png)

1. 在加载项中，JavaScript 调用新的 Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_)。 该操作告诉 Office 客户端应用程序获取加载项的访问令牌。 请参阅[示例访问令牌](#example-access-token)。
2. 如果用户未登录，Office 客户端应用程序会打开弹出窗口，以供用户登录。
3. 如果当前用户是首次使用加载项，则会看到同意提示。
4. Office 客户端应用程序从当前用户的 Azure AD v2.0 终结点请求获取 **加载项令牌**。
5. Azure AD 将加载项令牌发送给 Office 客户端应用程序。
6. Office 客户端应用程序在 `getAccessToken` 调用返回的结果对象中，将“**加载项令牌**”发送给加载项。
7. 加载项中的 JavaScript 可以解析令牌并提取所需信息，如用户的电子邮件地址。
8. （可选）加载项可以向其服务器端发送 HTTP 请求以获取关于用户的更多数据，如用户的偏好。 此外，访问令牌本身也可发送到服务器端以进行解析和验证。

## <a name="develop-an-sso-add-in"></a>开发 SSO 加载项

此部分介绍了创建启用 SSO 的 Office 加载项所需完成的任务。 其中介绍的这些任务与语言和框架无关。 有关详细演练的示例，请参阅：

- [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
- [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)

> [!NOTE]
> 可使用 Yeoman 生成器创建启用了 SSO 的  Node.js Office 加载项。 Yeoman 生成器简化了启用了 SSO 的加载项创建流程，能够自动执行在 Azure 内配置所需的步骤，并生成加载项使用 SSO 所需的代码。 有关详细信息，请参阅“[单一登录（SSO）快速入门](../quickstarts/sso-quickstart.md)”。

### <a name="create-the-service-application"></a>创建服务应用程序

在 Azure v2.0 终结点的注册门户注册外接程序。该流程用时 5-10 分钟，包括以下任务。

- 获取加载项的客户端 ID 和机密。
- 指定加载项访问 AAD v 所需的权限。 2.0 端点（可选 Microsoft Graph）。 始终需要“配置文件”和“openid”权限。
- 授予 Office 客户端应用程序信任加载项。
- 将 Office 客户端应用程序预授权给具有 *access_as_user* 默认权限的加载项。

有关此过程的详细信息，请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项](register-sso-add-in-aad-v2.md)。

### <a name="configure-the-add-in"></a>配置加载项

向加载项清单添加新标记。

- **WebApplicationInfo** - 下列元素的父元素。
- **ID** - 加载项的客户端 ID。这是在注册加载项时获得的应用程序 ID。 请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项](register-sso-add-in-aad-v2.md)。
- **Resource** - 加载项 URL。 这是在 AAD 中注册加载项时使用的相同 URI（包括 `api:` 协议）。 这个 URI 的域名部分必须与加载项的清单 `<Resources>` 中的 URL 中使用的域名（包括任何子域名）相匹配，并且 URI 必须以`<Id>`中的客户端 ID 结束。
- **Scopes** - 一个或多个“**Scope**”元素的父元素。
- **Scope** - 指定加载项访问 AAD 所需的权限。 如果加载项不访问 Microsoft Graph，则始终需要`profile` 和 `openID` 权限，并且可能是唯一需要的权限。 如果可以访问，则还需要“**Scope**”元素来获取所需的 Microsoft Graph 权限（如 `User.Read``Mail.Read`）。 在代码中用于访问 Microsoft Graph 的库可能需要其他权限。 例如，用于 .NET 的 Microsoft 身份验证库 (MSAL) 需要 `offline_access` 权限。 有关详细信息，请参阅[向 Office 加载项中的 Microsoft Graph 授权](authorize-to-microsoft-graph.md)。

对于除 Outlook 之外的 Office 应用程序，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 部分的末尾。对于 Outlook，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 部分的末尾。

下面是一个标注示例。

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>openid</Scope>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

> [!NOTE]
> 如果不符合SSO清单中的格式要求，将导致AppSource拒绝加载项，直到它符合所需格式。

### <a name="add-client-side-code"></a>添加客户端代码

将 JavaScript 添加到加载项，以执行以下操作：

- 调用 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_)。

- 解析访问令牌或将其传递到加载项的服务器端代码。

下面是调用 `getAccessToken` 的简单示例。

> [!NOTE]
> 此示例只显式处理一种错误。 有关更详细的错误处理的示例，请参阅 [Office 加载项 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) 和 [Office 加载项 ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)。

```js
async function getGraphData() {
    try {
        let bootstrapToken = await OfficeRuntime.auth.getAccessToken();

        // The /api/DoSomething controller will make the token exchange and use the
        // access token it gets back to make the call to MS Graph.
        getData("/api/DoSomething", bootstrapToken);
    }
    catch (exception) {
        if (exception.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // Microsoft 365 Education or work account, or a Microsoft account.
        } else {
            // Handle error
        }
    }
}
```

下面是一个将加载项令牌传递到服务器端的简单示例。 将请求发送回服务器端时，令牌作为 `Authorization` 标头包含在内。 此示例设想发送 JSON 数据，因此它使用 `POST` 方法，但使用 `GET` 就足以在未写入服务器时发送访问令牌。

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + bootstrapToken
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

如果因当前没有用户登录 Office 而无法使用加载项，则应 *在加载项启动时* 调用 `getAccessToken`，并在 `getAccessToken` 的 `options` 参数中传递 `allowSignInPrompt: true`。 例如：`OfficeRuntime.auth.getAccessToken( { allowSignInPrompt: true });`

如果加载项具有一些无需用户登录的功能，那么 *当用户执行需要用户登录的操作时*，请调用 `getAccessToken`。 `getAccessToken` 的冗余调用不会导致性能严重下降，因为 Office 缓存并重用启动没有过期的令牌，无需每次调用 AAD v。 `getAccessToken` 都重新调用 AAD V 2.0 端点。 因此，可以将 `getAccessToken` 调用添加到所有在需要令牌时启动操作的函数和处理程序。

### <a name="add-server-side-code"></a>添加服务器端代码

大多数情况下，如果加载项没有将访问令牌传递到服务器端并在其中使用它，那么获取访问令牌的意义就不大。

- 创建一种或多种 Web API 方法（例如，一种在托管数据库中查找用户首选项的方法），使用有关从令牌中提取的用户的信息。 （请参阅下文“**使用 SSO 令牌作为标识**”。）可以使用一些库简化需要编写的代码，具体视语言和框架而定。
- 获取 Microsoft Graph 数据。 服务器端代码应执行以下操作：

  - 通过调用 Azure AD v2.0 终结点启动“代表”流，该终结点包括访问令牌、关于用户的一些元数据以及加载项的凭据（其 ID 和机密）。在此上下文中，访问令牌称为启动令牌。
  - 使用新的令牌从 Microsoft Graph 获取数据。
  - 或者，在启动流之前，验证访问令牌（请参阅下文 **验证访问令牌**）。
  - 或者，在代表流完成后，缓存从流返回的新访问令牌，以便在对 Microsoft Graph 的其他调用中重复使用它，直到过期为止。

 如需深入了解如何获得对用户的 Microsoft Graph 数据的授权访问，请参阅[向 Office 加载项中的 Microsoft Graph 授权](authorize-to-microsoft-graph.md)。

#### <a name="validate-the-access-token"></a>验证访问令牌

Web API 收到访问令牌后，可以在使用该令牌前对其进行验证。 该令牌是 JSON Web 令牌 (JWT)，这意味着验证方式与大多数标准 OAuth 流中的令牌验证方式类似。 有许多可用于处理 JWT 验证的库，而它们的基本内容为：

- 检查令牌的格式是否正确
- 检查令牌是否由预期的颁发机构颁发
- 检查令牌是否是针对 Web API

验证令牌时，请牢记以下准则。

- 有效的 SSO 令牌是由 Azure 颁发机构 `https://login.microsoftonline.com` 的。 令牌中的 `iss` 声明应以此值开头。
- 令牌的 `aud` 参数将被设置为加载项注册的应用程序 ID。
- 令牌的 `scp` 参数将被设置为 `access_as_user`。

#### <a name="using-the-sso-token-as-an-identity"></a>将 SSO 令牌用作标识

如果加载项需要验证用户标识，则 SSO 令牌包含的信息可用于创建此标识。令牌中的以下声明与标识相关。

- `name` - 用户的显示名称。
- `preferred_username` - 用户的电子邮件地址。
- `oid` - 表示 Azure Active Directory 中的用户 ID 的 GUID。
- `tid` - 表示 Azure Active Directory 中的用户组织 ID 的 GUID。

由于 `name` 和 `preferred_username` 值可以更改，因此建议使用 `oid` 和 `tid` 值将标识与后端的授权服务关联。

例如，你的服务可以将这些值组合在一起，并设置为类似 `{oid-value}@{tid-value}` 的格式，然后将其存储为内部用户数据库中的用户记录值。 然后，在后续的请求中，可以使用同一值检索此用户，并可基于现有访问控制机制确定对特定资源的访问。

### <a name="example-access-token"></a>示例访问令牌

以下是访问令牌的典型解码有效负载。 有关属性的详细信息，请参阅 [Azure Active Directory v2.0 令牌参考](/azure/active-directory/develop/active-directory-v2-tokens)。

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

## <a name="using-sso-with-an-outlook-add-in"></a>将 SSO 与 Outlook 加载项一起使用

在 Outlook 加载项中使用 SSO 与在 Excel、PowerPoint 或 Word 加载项中使用 SSO 存在一些细微但却重要的差别。 请务必阅读[使用 Outlook 加载项的单一登录对用户进行身份验证](../outlook/authenticate-a-user-with-an-sso-token.md)和[：在 Outlook 加载项中为服务实现单一登录](../outlook/implement-sso-in-outlook-add-in.md)。

## <a name="sso-api-reference"></a>SSO API 参考

### <a name="getaccesstoken"></a>getAccessToken

OfficeRuntime [Auth](/javascript/api/office-runtime/officeruntime.auth) 命令空间 `OfficeRuntime.Auth` 提供了方法 `getAccessToken`，它使 Office 应用程序能够获取加载项的 Web 应用程序的访问令牌。 这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。

```typescript
getAccessToken(options?: AuthOptions: (result: AsyncResult<string>) => void): void;
```

该方法调用 Azure Active Directory V 2.0 端点以获取令牌来访问加载项的 Web 应用程序。 这样可以使加载项识别用户。 通过[“代表”OAuth 流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)，服务器端代码可以使用此令牌访问加载项 Web 应用程序的 Microsoft Graph。

> [!NOTE]
> 在 Outlook 中，如果加载项加载到 Outlook.com 或 Gmail 邮箱中，则此 API 不受支持。

|Hosts|Excel、Outlook、PowerPoint、Word|
|---|---|
|[要求集](specify-office-hosts-and-api-requirements.md)|[IdentityAPI](../reference/requirement-sets/identity-api-requirement-sets.md)|

#### <a name="parameters"></a>参数

`options` - 可选。 接受 [AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) 对象（参见下文）以定义登录行为。

`callback` - 可选。 接受可以解析用户 ID 的令牌或使用“代表”流中的令牌来访问 Microsoft Graph 的回调方法。 如果 [AsyncResult](/javascript/api/office/office.asyncresult) `.status`为“成功”，则 `AsyncResult.value` 是原始 AAD v。 2.0 格式的访问令牌。

当 Office 从 AAD v 获取加载项的访问令牌时，[AuthOptions](/javascript/api/office-runtime/officeruntime.authoptions) 接口会提供用户体验选项。 2.0 使用 `getAccessToken` 方法。
