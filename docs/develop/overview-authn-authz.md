---
title: Office 加载项中的身份验证和授权概述
description: 了解 Office 加载项中的身份验证和授权工作原理。
ms.date: 01/25/2022
ms.localizationpriority: high
---

# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a>Office 加载项中的身份验证和授权概述

默认情况下，Office 加载项允许匿名访问，但你可以要求用户登录以后再使用Microsoft 帐户、Microsoft 365 教育版或工作帐户或其他公共帐户的外接程序。 此任务被称为“用户身份验证”，因为它让加载项能够知道用户的身份。

你的加载项还能从用户处获得对其以下数据的访问许可：Microsoft Graph 数据（例如其 Microsoft 365 个人资料、OneDrive 文件和 SharePoint 数据），或者 Google、Facebook、领英、SalesForce 和 GitHub 等其他外部源中的数据。 此任务被称为“加载项（或应用）授权”，因为要获得授权的是 *加载项*，而不是用户。

## <a name="key-resources-for-authentication-and-authorization"></a>用于身份验证和授权的关键资源

本文档介绍如何构建和配置 Office 加载项，以成功实现身份验证和授权。 但是，所述的许多概念和安全技术超出了本文档的范围。 例如，此处未介绍常规安全概念，如 OAuth 流、令牌缓存或标识管理。 本文档也没有记录任何特定于 Microsoft Azure 或 Microsoft 标识平台的内容。 如果需要这些方面的信息，建议参考以下资源。

- [Microsoft 标识平台](/azure/active-directory/develop)
- [Microsoft 标识平台开发人员的支持和帮助选项](/azure/active-directory/develop/developer-support-help-options)
- Microsoft 标识平台上的 [OAuth 2.0 和 OpenID Connect 协议](/azure/active-directory/develop/active-directory-v2-protocols)

## <a name="sso-scenarios"></a>SSO 方案

用户使用单一登录(SSO)非常方便，因为他们只需登录一次即可使用 Office。 他们无需单独登录加载项。 所有版本的 Office 均不支持 SSO，因此仍需通过[使用 Microsoft 标识平台](#authenticate-with-the-microsoft-identity-platform)采用另一种登录方法。 有关支持的 Office 版本的详细信息，请参阅 [标识 API 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)

### <a name="get-the-users-identity-through-sso"></a>通过 SSO 获取用户标识

通常，外接程序仅需要用户的标识。 例如，你可能只想对加载项进行个性化设置，并在任务窗格中显示用户的名称。 或者，你可能希望唯一 ID 将用户与数据库中的数据关联。 只需从 Office 获取用户的访问令牌即可实现此目的。

若要通过 SSO 获取用户标识，请调用 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) 方法。 此方法返回一个访问令牌，该令牌也是一个标识令牌，其中包含对当前已登录用户唯一的多个声明，包括 `preferred_username`、`name`、`sub`和`oid`。 有关这些属性的详细信息，请参阅 [Microsoft 标识平台 ID 令牌](/azure/active-directory/develop/id-tokens)。 有关 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) 返回的令牌示例，请参阅[示例访问令牌](sso-in-office-add-ins.md#example-access-token)。

如果用户未登录，Office 将打开一个对话框，并使用 Microsoft 标识平台请求用户登录。 然后，此方法将返回访问令牌，或在用户无法登录时弹出错误。

在需要存储用户数据的方案中，请参阅 [Microsoft 标识平台 ID 令牌](/azure/active-directory/develop/id-tokens)，了解如何从令牌获取值以唯一标识用户。 使用该值在你维护的用户表或用户数据库中查找用户。 使用数据库来用户用户首选项或用户帐户状态等用户相关信息。 由于你在使用 SSO，因此你的用户不单独登录到你的加载项，你无需存储用户的密码。

在开始使用 SSO 实现用户身份验证之前，请确保完全了解文章[为 Office 加载项启用单一登录](sso-in-office-add-ins.md)。

### <a name="access-your-web-apis-through-sso"></a>通过 SSO 访问 Web API

如果加载项具有需要授权用户的服务器端 API，请调用 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) 方法以获取访问令牌。 访问令牌提供对你自己 Web 服务器的访问权限（通过 [Microsoft Azure 应用注册](register-sso-add-in-aad-v2.md)配置）。在 Web 服务器上调用 API 时，还会传递访问令牌来授权用户。

以下代码演示如何构造对外接程序 Web 服务器 API 的 HTTPS GET 请求，以获取数据。 代码在客户端运行，例如在任务窗格中运行。 它首先通过调用 `getAccessToken` 获取访问令牌。 然后，它使用服务器 API 的正确授权标头和 URL 构造 AJAX 调用。

```javascript
function getOneDriveFileNames() {

    let accessToken = await Office.auth.getAccessToken();

    $.ajax({
        url: "/api/data",
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
        .done(function (result) {
            //... work with data from the result...
        });
}
```

下面的代码显示了上一个代码示例中 REST 调用的示例 /api/data 处理程序。 ASP.NET Web 为在服务器上运行的代码。 属性 `[Authorize]` 将要求从客户端传递有效的访问令牌，否则它将向客户端返回错误。

```csharp
    [Authorize]
    // GET api/data
    public async Task<HttpResponseMessage> Get()
    {
        //... obtain and return data to the client-side code...
    }
```

### <a name="access-microsoft-graph-through-sso"></a>通过 SSO 访问 Microsoft Graph

在某些情况下，不仅需要用户标识，还需要代表用户访问 [Microsoft Graph](/graph)资源。 例如，可能需要发送电子邮件，或代表用户在 Teams 中创建聊天。 这些操作等可通过 Microsoft Graph 来完成。 需要按照以下步骤操作：

1. 通过 SSO 调用 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) 获取当前用户的访问令牌。 如果用户未登录，Office 将打开一个对话框，并使用 Microsoft 标识平台请求用户登录。 用户登录后或者在用户已登录时，该方法会返回一个访问令牌。
1. 将访问令牌传递到服务器端代码。
1. 在服务器端，使用 [OAuth 2.0 代理流](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow) 来交换访问令牌，以获得包含必要的委托用户标识和调用 Microsoft Graph 权限的新访问令牌。 

> [!NOTE]
> 为获得最佳安全性以避免泄露访问令牌，请始终在服务器端执行代表流。 从服务器（而不是客户端）调用Microsoft Graph API。 不要将访问令牌返回到客户端代码。

在开始执行 SSO 以访问外接程序中的Microsoft Graph 之前，请确保完全熟悉以下文章。

- [为 Office 加载项启用单一登录](sso-in-office-add-ins.md)
- [使用 SSO 对 Microsoft Graph 授权](authorize-to-microsoft-graph.md)

还应该阅读以下至少一篇文章，这些文章将指导你构建 Office 外接程序，以使用 SSO 和访问 Microsoft Graph。 即使不执行这些步骤，也包含有关实现 SSO 和代表流的有用信息。

- [创建使用单一登录的 ASP.NET Office 外接程序](create-sso-office-add-ins-aspnet.md)，该加载项将指导你完成 [Office 外接程序 ASP.NET SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)的示例。
- [创建使用单一登录的 Node.js Office 外接程序](create-sso-office-add-ins-nodejs.md)，该加载项将指导你完成 [Office 外接程序 NodeJS SSO ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)的示例。

## <a name="non-sso-scenarios"></a>非 SSO 方案

在某些情况下，你可能不希望使用 SSO。 例如，可能需要使用不同的标识提供程序进行身份验证，而不是 Microsoft 标识平台。 此外，并非所有方案都支持 SSO。 例如，较旧版本的 Office 不支持 SSO。 在这种情况下，需要回退到外接程序的备用身份验证系统。

### <a name="authenticate-with-the-microsoft-identity-platform"></a>向 Microsoft 标识平台进行身份验证。

加载项可以使用作为身份验证提供程序的 [Microsoft 标识平台登录用户](/azure/active-directory/develop)。 登录用户后，可以使用 Microsoft 标识平台向 [Microsoft Graph](/graph) 或其他 Microsoft 托管服务授权加载项。 当 SSO 在 Office 不可用时，使用此方法作为备用登录方法。 此外，在某些方案中，即使在 SSO 可用的情况下，用户也需要单独登录加载项；例如，当你希望他们可以选择使用与当前登录到 Office 的 ID 不同的 ID 登录到该加载项时。

需要注意的是，Microsoft 标识平台不允许其登录页在 iframe 中打开。 当 Office 加载项在 *Office 网页版* 中运行时，任务窗格是一个 iFrame。 这意味着你需要使用通过 Office 对话框 API 打开的对话框打开登录页。 这会影响你使用身份验证帮助程序库的方式。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

有关使用 Microsoft 标识平台实现身份验证的信息，请参阅 [Microsoft 标识平台(v2.0)概述](/azure/active-directory/develop/v2-overview)。 该文档包含许多教程和指南，以及指向相关示例和库的链接。 正如[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)中所述，你可能需要调整示例中的代码以在 Office 对话框中运行。

### <a name="access-to-microsoft-graph-without-sso"></a>在不使用 SSO 的情况下访问 Microsoft Graph

可以通过从 Microsoft 标识平台获取Microsoft Graph 的访问令牌，以获取对加载项 Microsoft Graph 数据的授权。 无需通过 Office 依赖 SSO 即可执行此操作（如果 SSO 失败或不受支持）。 有关详细信息，请参阅 [在没有 SSO 的情况下访问 Microsoft Graph](authorize-to-microsoft-graph-without-sso.md)，其中包含更多详细信息和示例链接。

### <a name="access-to-non-microsoft-data-sources"></a>访问非 Microsoft 数据源

借助 Google、Facebook、领英、SalesForce 和 GitHub 等热门在线服务，开发人员可授权用户访问自己在其他应用中的帐户。 这样，便可在 Office 加载项中添加这些服务。 要概述了解加载项可执行此操作的方法，请参阅[在 Office 加载项中授权外部服务](auth-external-add-ins.md)。

> [!IMPORTANT]
> 在开始编码之前，请查明数据源是否允许在 iframe 中打开其登录页。 当 Office 加载项在 *Office 网页版* 中运行时，任务窗格是一个 iFrame。 如果数据源不允许在 iframe 中打开其登录页，则需要在使用 Office 对话框 API 打开的对话框中打开登录页。 有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。

## <a name="see-also"></a>另请参阅

- [Microsoft 标识平台文档](/azure/active-directory/develop/)
- [Microsoft 标识平台访问令牌](/azure/active-directory/develop/access-tokens)
- Microsoft 标识平台上的 [OAuth 2.0 和 OpenID Connect 协议](/azure/active-directory/develop/active-directory-v2-protocols)
- [Microsoft 标识平台和 OAuth 2.0 代表流](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
- [JSON Web 令牌 (JWT) ](https://en.wikipedia.org/wiki/JSON_Web_Token)
- [JSON Web 令牌查看器](https://jwt.ms/)
