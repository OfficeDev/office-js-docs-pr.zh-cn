# <a name="enable-single-sign-on-for-office-add-ins"></a>为 Office 加载项启用单一登录

用户可以使用自己的个人 Microsoft 帐户/工作或学校 (Office 365) 帐户，登录 Office（在线、移动和桌面平台）。 可以利用这一点，使用 SSO 执行以下操作（用户无需再次登录）：

* 授权用户登录加载项。
* 授权加载项访问 [Microsoft Graph](https://developer.microsoft.com/graph/docs)。

![显示加载项登录过程的图像](../../images/OfficeHostTitleBarLogin.png)

>**注意：**目前，Word、Excel 和 PowerPoint 支持单一登录 API。 若要详细了解目前单一登录 API 的受支持情况，请参阅 [Identity API 要求集](../../reference/requirement-sets/identity-api-requirement-sets.md)。
> Outlook 的单一登录当前处于预览阶段。 如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。 若要了解如何这样做，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

对于用户，这样可以方便他们顺畅地运行加载项，因为只需登录一次。 对于开发者，这意味着加载项可以验证用户身份，并使用用户已提供给 Office 应用程序的凭据，通过 Microsoft Graph 获取对用户数据的访问权限。

## <a name="sso-add-in-architecture"></a>SSO 外接程序体系结构

除了托管 Web 应用程序的页面和 JavaScript 之外，外接程序还必须以相同的[完全限定的域名](https://msdn.microsoft.com/en-us/library/windows/desktop/ms682135.aspx#_dns_fully_qualified_domain_name_fqdn__gly)托管一个或多个 Web API，这些 API 可获取 Microsoft Graph 的访问令牌，并向它发出请求。

外接程序清单包含标记，用于指定外接程序在 Azure Active Directory (Azure AD) v2.0 终结点中的注册方式，并指定外接程序需要的 Microsoft Graph 的任何权限。

### <a name="how-it-works-at-runtime"></a>运行时的工作方式

以下关系图显示了 SSO 流程的工作方式。
<!-- Minor fixes to the text in the diagram - change V2 to v2.0, and change "(e.g. Word, Excel, etc.)" to "(for example, Word, Excel)". -->
![SSO 过程关系图](../../images/SSOOverviewDiagram.png)

1. 在加载项中，JavaScript 调用新的 Office.js API `getAccessTokenAsync`。 这会指示 Office 主机应用程序获取对加载项的访问令牌。 （以下称为**加载项令牌**。）
1. 如果用户未登录，Office 主机应用程序会打开弹出窗口，以供用户登录。
1.  如果当前用户是首次使用加载项，则会看到同意提示。
1. Office 主机应用程序从当前用户的 Azure AD v2.0 终结点请求获取**加载项令牌**。
1. Azure AD 将加载项令牌发送给 Office 主机应用程序。
1. Office 主机应用程序在 `getAccessTokenAsync` 调用返回的结果对象中，将**加载项令牌**发送给加载项。
1. 加载项中的 JavaScript 向 Web API（与加载项托管在同一完全限定的域中）发出 HTTP 请求，并添加**加载项令牌**作为授权证明。  
1. 服务器端代码验证传入的**加载项令牌**。
1. 服务器端代码使用“代表”流（在 [OAuth2 令牌交换](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)和 [Web API Azure 应用场景的守护程序或服务器应用程序](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)中定义），获取对 Microsoft Graph 的访问令牌（以下称为 **MSG 令牌**），以交换加载项令牌。
1. Azure AD 将 **MSG 令牌**（如果加载项请求获取 *offline_access* 权限，则同时返回刷新令牌）返回给加载项。
1. 服务器端代码缓存 **MSG 令牌**。
1. 服务器端代码向 Microsoft Graph 发出请求，并添加 **MSG 令牌**。
1. Microsoft Graph 将数据返回给加载项，从而将数据传递到加载项 UI。
1. 如果 MSG 令牌过期，服务器端代码可以使用刷新令牌获取新的 **MSG 令牌**。

## <a name="develop-an-sso-add-in"></a>开发 SSO 加载项

此部分介绍了创建启用 SSO 的 Office 加载项所需完成的任务。 其中介绍的这些任务与语言和框架无关。 有关详细演练的示例，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](../../docs/develop/create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](../../docs/develop/create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a>创建服务应用程序

在 Azure v2.0 终结点的注册门户注册外接程序：https://apps.dev.microsoft.com。该流程用时 5-10 分钟，包括以下任务：

* 获取外接程序的客户端 ID 和机密。
* 指定外接程序访问 Microsoft Graph 所需权限。
* 向外接程序授予 Office 主机应用程序信任。
* 将 Office 主机应用程序预授权给具有 *access_as_user* 默认权限的外接程序。

### <a name="configure-the-add-in"></a>配置外接程序

向外接程序清单添加新标记：

* **WebApplicationInfo** - 下列元素的父元素。
* **Id** - 外接程序的客户端 ID。
* **Resource** - 加载项 URL。
* **Scopes** - 一个或多个 **Scope** 元素的父元素。
* **Scope** - 指定加载项访问 Microsoft Graph 所需的权限。 例如，`User.Read`、`Mail.Read` 或 `offline_access`。 有关详细信息，请参阅 [Microsoft Graph 权限](https://developer.microsoft.com/en-us/graph/docs/concepts/permissions_reference)

对于除 Outlook 之外的 Office 主机，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 部分的末尾。对 Outlook，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 部分的末尾。

### <a name="add-client-side-code"></a>添加客户端代码

将 JavaScript 添加到外接程序，以执行以下操作：

* 调用 `Office.context.auth.getAccessTokenAsync(myTokenHandler)`。
* 创建将外接程序令牌传递给外接程序服务器端代码的处理程序。例如：

```js
function mytokenHandler(asyncResult) {
    // Passes asyncResult.value (which has the add-in access token)
    // to the add-in’s web API as an Authorization header.
}
```

### <a name="when-to-call-the-method"></a>何时调用方法

如果因没有用户登录 Office 且 Office 没有对加载项的访问令牌而无法使用加载项，应*在加载项启动时*调用 `getAccessTokenAsync`。

如果加载项的一些功能不需要访问 Microsoft Graph 或用户登录，那么*当用户执行的操作需要访问 Microsoft Graph 或最起码需要用户登录时*，调用 `getAccessTokenAsync`。 `getAccessTokenAsync` 的冗余调用不会导致性能严重下降，因为 Office 缓存并重用访问没有过期的令牌，无需每次调用 `getAccessTokenAsync` 都重新调用 AAD V 2.0 终结点。 因此，可以将 `getAccessTokenAsync` 调用添加到所有在需要令牌时启动操作的函数和处理程序。

### <a name="add-server-side-code"></a>添加服务器端代码

创建可获取 Microsoft Graph 数据的一个或多个 Web API 方法。 可以使用一些库简化需要编写的代码，具体视语言和框架而定。 服务器端代码应执行以下操作：

* 验证从之前创建的令牌处理程序收到的加载项令牌。
* 通过调用 Azure AD v2.0 终结点启动“代表”流，该终结点包括外接程序访问令牌、关于用户的一些元数据以及外接程序的凭据（其 ID 和机密）。
* 缓存返回的 MSG 令牌。
* 使用 MSG 令牌从 Microsoft Graph 获取数据。
