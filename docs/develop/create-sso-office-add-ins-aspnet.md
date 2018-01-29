# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>创建使用单一登录的 ASP.NET Office 加载项

如果用户已登录 Office，加载项可以使用相同的凭据，这样用户无需重新登录，即可访问多个应用程序。 有关概述，请参阅[在 Office 加载项中启用 SSO](../develop/sso-in-office-add-ins.md)。

本文将引导你完成在使用 ASP.NET、OWIN 和适用于 .NET 的 Microsoft 验证库 (MSAL) 生成的外接程序中启用单一登录 (SSO) 的过程。

> **注意：**有关基于 Node.js 的外接程序的类似文章，请参阅[创建使用单一登录的 Node.js Office 外接程序](../develop/create-sso-office-add-ins-nodejs.md)。

## <a name="prerequisites"></a>先决条件

* Visual Studio 2017 Preview 最新可用版本。

>**注意：**最新版 Visual Studio 2017 Preview 与 SSO 需要的加载项清单标记暂不兼容。 下面的过程中详细介绍了如何解决此问题。

* Office 2016，版本 1708，内部版本 8424.nnnn 或更高版本（Office 365 订阅版本，有时称为“即点即用”）。可能需要成为 Office 预览体验成员才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/en-us/office-insider?tab=tab-1)。

## <a name="set-up-the-starter-project"></a>设置初学者项目

1. 在 [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso) 处克隆或下载存储库。

1. 打开 **Before** 文件夹，并打开 Visual Studio 中的 .sln 文件。这是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。

    > 注意：同一存储库中还有已完成版本的样本加载项。 这与完成本文中过程生成的加载项基本一样，不同之处在于，已完成项目中的代码注释对本文来说是多余的。 若要使用已完成版本，只需打开 *.sln 文件，并按照本文中的说明操作即可，但要跳过**编写客户端代码**和**编写服务器端代码**。

1. 项目打开后，在 Visual Studio 中生成它，这将安装 packages.config 文件中列出的包。 这可能需要几秒钟到几分钟的时间才能完成，具体视计算机本地包缓存中的包数量而定。

    > **重要说明！** Web API 项目根中的 packages.config 指定了 `1.1.1-alpha0393` 版 MSAL 库 Microsoft.Identity.Client。 首次按 F5 后，应验证是否已安装此版本（或更高版本）：在“工具”****菜单中，依次转到“Nuget 包管理器”**** > “管理解决方案的 Nuget 包”**** > “已安装”****。 滚动到“Microsoft.Identity.Client”****，查看已安装的版本。 如果低于 `1.1.1-alpha0393`（或“已安装”****列表中没有此版本），请依次转到“Nuget 包管理器”**** > “包管理器控制台”****。 在控制台中，运行命令 `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json`。

1. 该项目完全生成后，请按 F5。PowerPoint 将打开，“主页”****功能区上会有一个“SSO ASP.NET”****组。

1. 按此组中的“显示加载项”****按钮，即可在任务窗格中看到此加载项的 UI。 任务窗格中的按钮尚未相互关联。
2. 停止 Visual Studio 中的调试器。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>向 Azure AD v2.0 终结点注册加载项

1. 转到 [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)。

1. 使用管理员凭据登录 Office 365 租户。例如，MyName@contoso.onmicrosoft.com

1. 单击“添加应用”****。

1. 出现提示时，使用“Office-Add-in-ASPNET-SSO”作为应用名称，然后按“创建应用程序”****。

1. 当应用程序的配置页打开时，复制并保存**应用程序 ID**。 在后面的过程中，将会用到它。

    > **注意**：当其他应用程序（如 PowerPoint、Word、Excel 等 Office 主机应用程序）请求获取对此应用程序的访问权限时，此 ID 的值为“audience”。 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

1. 在“应用程序机密”****部分中，按“生成新密码”****。 此时，系统会打开弹出对话框，并显示新密码（亦称为“应用程序密码”）。 *立即复制此密码，并将它与应用程序 ID 一起保存。* 在后面的过程中，将需要用到它。 然后，关闭此对话框。

1. 在“平台”****部分中，单击“添加平台”****。

1. 在打开的对话框中，选择“Web API”****。

1. “应用程序 ID URI”****已生成，格式为“api://{App ID GUID}”。在双正斜线和 GUID 之间插入字符串“localhost:44355/”。完整 ID 应为 `api://localhost:44355/{App ID GUID}`。（位于“应用程序 ID URI”****下方的“作用域”****名称的域部分将自动更改以匹配。应显示 `api://localhost:44355/{App ID GUID}/access_as_user`。）

1. 在“预授权应用程序”****部分中，确定要授权给加载项 Web 应用程序的应用程序。 下面每个 ID 都需要进行预授权。 每次输入一个 ID，都会看到新的空文本框。 （仅输入 GUID）。
 * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
 * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
 * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. 打开每个“应用程序 ID”****旁边的“作用域”****下拉列表，并选中 `api://localhost:44355/{App ID GUID}/access_as_user` 对应的框。

1. 在“平台”****部分顶部附近，再次单击“添加平台”****并选择“Web”****。

1. 在“平台”****下的新“Web”****部分中，输入以下内容作为“重定向 URL”****：`https://localhost:44355`。

    > 注意：在撰写本文时，“Web API”****平台有时会从“平台”****部分消失，特别是在添加“Web”****平台以及“保存注册页面”**后，如果刷新页面就会出现上述情况。为了确保“Web API”****平台仍然是注册的一部分，请单击页面底部附近的“编辑应用程序清单”****按钮。应该会看到清单的 **identifierUris** 属性中的 `api://localhost:44355/{App ID GUID}` 字符串。还有一个 **oauth2Permissions** 属性，它的 **value** 子属性的值为 `access_as_user`。

1. 向下滚动到“Microsoft Graph 权限”****部分，“委派的权限”****小节。使用“添加”****按钮打开“选择权限”****对话框。

1. 在对话框中，选中以下权限对应的框（默认情况下，某些权限可能已处于选中状态）： 加载项本身真正需要的只是第一项权限，但服务器端代码使用的 MSAL 库需要有 `offline_access` 和 `openid`。 Office 主机必须有 `profile` 权限，才能获取对加载项 Web 应用程序的令牌。
 * Files.Read.All
 * offline_access
 * openid
 * profile

1. 单击对话框底部的“确定”****。

1. 单击注册页底部的“保存”****。

## <a name="grant-admin-consent-to-the-add-in"></a>向加载项授予管理员许可

> **注意：**仅在开发加载项时，才需要执行此过程。 将生产加载项部署到 Office 应用商店或加载项目录时，用户会独自信任它，或者管理员会在组织安装时授予许可。

1. 如果加载项未在 Visual Studio 中运行，请按 **F5** 运行它。 必须在 IIS 中运行，才能顺利完成此过程。

1. 在以下字符串中，将占位符“{application_ID}”替换为注册加载项时复制的应用程序 ID：`https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. 将生成的 URL 粘贴到浏览器地址栏，并转到此 URL。

1. 看到提示时，使用管理员凭据登录 Office 365 租户。

1. 然后系统提示你授予外接程序访问 Microsoft Graph 数据的权限。单击“接受”****。

1. 然后，将浏览器窗口/选项卡重定向到注册外接程序时指定的**重定向 URL**；因此，外接程序的主页将在浏览器中打开。

2. 在浏览器的地址栏中，你将看到一个带有 GUID 值的“租户”查询参数。 这是 Office 365 租户的 ID。 复制并保存此值。 在后面的步骤中，将会用到它。

3. 关闭此窗口/选项卡。

1. 停止 Visual Studio 中的调试器。

## <a name="configure-the-add-in"></a>配置外接程序

1. 在下面的字符串中，将占位符“{tenant_ID}”替换为之前获得的 Office 365 租户 ID。如果出于任何原因，你以前没有获得 ID，请使用[查找 Office 365 租户 ID](https://support.office.com/zh-cn/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) 中的一种方法来获取 ID。

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

1. 在 Visual Studio 中打开 web.config。你需要为 **appSettings** 部分中的某些键分配值。

1. 将在步骤 1 中构造的字符串用作名为“ida:Issuer”的键的值。请确保此值中没有空格。

1. 将下面的值分配给相应的键：

|键|值|
|:-----|:-----|
|ida:ClientID|注册外接程序时获取的应用程序 ID。|
|ida:Audience|注册外接程序时获取的应用程序 ID。|
|ida:Password|注册外接程序时获取的密码。|


下面的示例展示了更改后的四个键。 *请注意，ClientID 和 Audience 是一样的*。 也可以将一个键用于这两种用途，但如果由于它们不是一直都相同而将它们区分开来，web.config 标记的可重用性将会更高。 此外，将键区分开来也可以强化以下概念：相对于 Office 主机，加载项是 OAuth 资源；相对于 Microsoft Graph，加载项是 OAuth 客户端。

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    ```

> **注意：****appSettings** 部分中的其他设置保持不变。

1. 保存并关闭文件。

1. 在外接程序项目中，打开外接程序清单文件“Office-Add-in-ASPNET-SSO.xml”。

1. 滚动到文件底部。

1. 结束 `</VersionOverrides>` 标记的正上方有以下标记：

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}<Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. 将标记中的*两处*占位符“{application_GUID here}”均替换成在注册加载项时复制的应用程序 ID。 由于 ID 并不包含“{}”，因此请勿添加它们。 这与在 web.config 中对 ClientID 和 Audience 使用的 ID 相同。

    > **注意**：
    >* **Resource** 值是向注册的加载项添加 Web API 平台时设置的**应用程序 ID URI**。
    >* 如果通过 Office 应用商店销售该外接程序，则 **Scopes** 部分仅用于生成同意对话框。

1. 在 Visual Studio 中打开“错误列表”****的“警告”****选项卡。 如果存在关于 `<WebApplicationInfo>` 不是 `<VersionOverrides>` 的有效子级的警告，则该 Visual Studio 2017 Preview 版本无法识别 SSO 标记。 作为解决方法，请对 Word、Excel 或 PowerPoint 外接程序执行以下操作。 （如果使用的是 Outlook 外接程序，请参阅下面的解决方法。）

   - **Word、Excel 和 Powerpoint 的解决方法**

   > 1. 在结束 `</VersionOverrides>` 标记正上方的清单中，注释掉 `<WebApplicationInfo>` 部分。

   > 2. 按 F5 启动调试会话。此操作会在下列文件夹（相比 Visual Studio，在“文件资源管理器”****中访问此文件夹更方便）中创建清单副本：`Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

   > 3. 在清单副本中，删除 `<WebApplicationInfo>` 部分周围的注释语法。

   > 4. 保存此清单副本。

   > 5. 现在，必须防止 Visual Studio 在下次按 F5 时重写此清单副本。 右键单击“解决方案资源管理器”****顶部的解决方案节点（而不是任何项目节点）。

   > 6. 从关联菜单中选择“属性”****，随后“解决方案属性页”****对话框便会打开。

   > 7. 展开“配置属性”****，并选择“配置”****。

   > 8. 在 **Office-Add-in-ASPNET-SSO** 项目（*不是* **Office-Add-in-ASPNET-SSO-WebAPI** 项目）行中取消选择“生成”****和“部署”****。

   > 9. 按“确定”****关闭对话框。

   - **Outlook 的解决方法**

   > 1. 在开发计算机上找到现有的 `MailAppVersionOverridesV1_1.xsd`。 它应位于 `./Xml/Schemas/{lcid}` 下的 Visual Studio 安装目录中。 例如，在英语（美国）的系统上进行 VS 2017 32 位的典型安装时，完整路径为 `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`。

   > 2. 将现有文件重命名为 `MailAppVersionOverridesV1_1.old`。

   > 3. 将此修改后的文件版本复制到文件夹中：[修改后的 MailAppVersionOverrides 架构](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. 在 Visual Studio 中保存并关闭该主清单文件。

## <a name="code-the-client-side"></a>编写客户端代码

1. 打开 **Scripts** 文件夹中的 Home.js 文件。其中已存在一些代码：
    * 针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。
    * 在任务窗格底部，显示从 Microsoft Graph（或错误消息）返回的数据的 `showResult` 方法。

1. 在针对 `Office.initialize` 的分配下面，添加下面的代码。关于此代码，请注意以下几点：

    * `getAccessTokenAsync` 是 Office.js 中的新 API，它使外接程序能够让 Office 主机应用程序（Excel、PowerPoint、Word 等）请求外接程序的访问令牌（针对登录 Office 的用户）。反过来，Office 主机应用程序会向 Azure AD 2 终结点请求令牌。由于在注册外界程序时将 Office 主机预授权给外接程序，因此 Azure AD 将发送此令牌。
    * 如果没有用户登录 Office，则 Office 主机将提示用户登录。
    * options 参数将 `forceConsent` 设置为 false，因此将不会提示用户同意为 Office 主机提供访问外接程序的权限。

    ```js
    function getOneDriveFiles() {
        getDataWithToken({ forceConsent: false });
    }

    function getDataWithToken(options) {
        Office.context.auth.getAccessTokenAsync(options,
            function (result) {
                if (result.status === "succeeded") {
                    TODO1: Use the access token to get Microsoft Graph data.
                }
                else {
                    console.log("Code: " + result.error.code);
                    console.log("Message: " + result.error.message);
                    console.log("name: " + result.error.name);
                    document.getElementById("getGraphAccessTokenButton").disabled = true;
                }
            });
    }
    ```

1. 用以下行替换 TODO1。可以在后续步骤中创建 `getData` 方法和服务器端“/api/values”路由。相对 URL 用于终结点，因为它必须与外接程序托管在相同的域中。

    ```js
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. 在 `getOneDriveFiles` 方法下面添加以下内容。此实用程序方法调用指定的 Web API 终结点，并向其传递与 Office 主机应用程序用于获取外接程序访问权限的令牌相同的访问令牌。在服务器端，此访问令牌将用于“代表”流，以获取 Microsoft Graph 的访问令牌。

    ```js
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            TODO2: Handle errors and the case where Microsoft Graph
                   requires additional form of authentication.
        });
    }
    ```

1. 将 TODO2 替换为以下代码行。 关于此代码，请注意以下几点：

    * 如果是因为 Microsoft Graph 需要其他形式的身份验证而失败，`exceptionMessage` 是包含“capolids”的 JSON 字符串。 在这种情况下，Office 主机需要获取新令牌。  
    * 异常消息指示 AAD，应提示用户进行所有相应形式的身份验证，因此它必须被传递到 Office 主机，而反过来当请求获取新令牌时，它就会传递给 AAD。
    * `authChallenge` 选项是将此字符串传递到 Office 主机的方法。
    * 如果错误不是要求进行其他身份验证，那么它就会被记录到控制台中。

    ```js
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    if (exceptionMessage.indexOf("capolids") !== -1) {
        getDataWithToken({ authChallenge: exceptionMessage });
    } else {
        console.log(result.error);
    }
    ```

1. 保存并关闭文件。

## <a name="code-the-server-side"></a>编写服务器端代码

### <a name="configure-the-owin-middleware"></a>配置 OWIN 中间件

1. 在项目的根目录中打开 Startup.cs 文件。

1. 将关键字 `partial` 添加到 Startup 类（如果其中尚不存在该关键字）的声明。具体应如下所示：

    `public partial class Startup`

1. 将以下行添加至 `Configuration` 方法的正文。在后续步骤中创建 `ConfigureAuth` 方法。

    `ConfigureAuth(app);`

1. 保存并关闭文件。

1. 右键单击“App_Start”****文件夹，并依次选择“添加”>“类”****。

1. 在“添加新项”****对话框中，命名文件“Startup.Auth.cs”****，再单击“添加”****。

1. 将新文件中的命名空间名称缩短为 `Office_Add_in_ASPNET_SSO_WebAPI`。

1. 确保下列所有 `using` 语句都位于文件的顶部。

    ```
    using Owin;
    using System.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. 将关键字 `partial` 添加到 `Startup` 类（如果其中尚不存在该关键字）的声明。具体应如下所示：

    `public partial class Startup`

1. 将下列方法添加到 `Startup` 类。该方法指定 OWIN 中间件如何验证从客户端 Home.js 文件的 `getData` 方法传递给它的访问令牌。每次调用使用 `[Authorize]` 属性修饰的 Web API 终结点时都会触发授权过程。

    ```
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. 将 TODO3 替换为以下代码行。 注意：

    * 代码指示 OWIN 以确保在来自 Office 主机（并通过客户端调用 `getData` 进行传递）的访问令牌中指定的受众和令牌颁发者必须与 web.config 中指定的值相匹配。
    * 将 `SaveSigninToken` 设置为 `true` 将导致 OWIN 从 Office 主机保存原始令牌。外接程序需要它来获取具有“代表”流的 Microsoft Graph 的访问令牌。
    * OWIN 中间件不验证作用域。应包括 `access_as_user` 的访问令牌的作用域在控制器中进行验证。

    ```
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. 将 TODO4 替换为以下代码行。 注意：

    * 调用方法 `UseOAuthBearerAuthentication` 而不是更常见的 `UseWindowsAzureActiveDirectoryBearerAuthentication`，因为后者与 Azure AD V2 终结点不兼容。
    * 传递给该方法的发现 URL 是 OWIN 中间件获得用于获取所需密钥说明的位置，以验证从 Office 主机接收到的访问令牌上的签名。

    ```
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
            {
                AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
            });
    ```

1. 保存并关闭文件。

### <a name="create-the-apivalues-controller"></a>创建 /api/values 控制器

1. 打开文件 **Controllers\ValueController.cs**。

2. 请确保下列 `using` 语句位于文件顶部。

    ```
    using Microsoft.Identity.Client;
    using System;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web;
    using System.Web.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. 在声明 `ValuesController` 的行的正上方，添加属性 `[Authorize]`。 这样可确保，每当调用控制器方法时，加载项都会运行上一过程中配置的授权过程。 只有拥有对加载项的有效访问令牌，调用方才能调用控制器方法。

4. 将下列方法添加到 `ValuesController`：

    ```
    // GET api/values
    public async Task<IEnumerable<string>> Get()
    {
        // TODO5: Validate the scopes of the access token.
    }
    ```

5. 将 TODO5 替换为以下代码行，以验证令牌中指定的作用域是否包括 `access_as_user`。

    ```
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO6: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO7: Get the access token for Microsoft Graph.
        // TODO8: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO9: Remove excess information from the data and send the data to the client.
    }
    return new string[] { "Error", "Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user." };
    ```

> **注意：**只可使用 `access_as_user` 作用域授权 API 为 Office 加载项处理代表流。服务中的其他 API 应有自己的作用域要求。 这就限制了使用 Office 获得的令牌可以访问的内容。

6. 将 TODO6 替换为以下代码。注意：
    * 此代码将从 Office 主机收到的原始访问令牌转换为，将传递给另一个方法的 `UserAssertion` 对象。
    * 外接程序不再扮演 Office 主机和用户需要访问的资源（或受众）的角色。现在它本身就是一个需要访问 Microsoft Graph 的客户端。`ConfidentialClientApplication` 是 MSAL“客户端上下文”对象。
    * `ConfidentialClientApplication` 构造函数的第三个参数是在“代表”流中实际不使用的重定向 URL，但使用正确的 URL 是一个很好的做法。第四和第五个参数可用于定义持久性存储，该存储使得外接程序能在不同的会话之间重用未过期的令牌。此示例不实现任何持久性存储。
    * MSAL 要求 `openid`、`offline_access` 作用域能够发挥作用，但如果代码过多地发出请求，则会抛出错误。 如果代码请求获取 `profile`，也会抛出错误，这真正仅适用于 Office 主机应用程序获取对加载项 Web 应用程序的令牌时。 因此，只会显式请求 `Files.Read.All`。

    ```
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. 将 TODO7 替换为以下代码行。 注意：

    * `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` 方法将首先查找内存中的 MSAL 缓存，获取匹配的访问令牌。仅当不存在任何令牌时，该方法才会通过 Azure AD V2 终结点启动“代表”流。
    * 如果 MS Graph 资源要求进行多重身份验证，但用户尚未提供，AAD 就会抛出包含 Claims 属性的异常。
    * Claims 属性值必须传递到客户端，然后才会传递到 Office 主机，并被添加到新令牌请求中。 AAD 将提示用户进行所有相应形式的身份验证。
    * 任何不属于类型 `MsalUiRequiredException` 的异常都是有意不捕获的，这样才能传播到客户端。

    ```
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalUiRequiredException e)
    {        
        if (String.IsNullOrEmpty(e.Claims))
        {
            throw e;
        }
        else
        {
            throw new HttpException(e.Claims);
        }   
    }
    ```

8. 将 TODO8 替换为以下代码行。 注意：

    * `GraphApiHelper` 和 `ODataHelper` 类在 **Helpers** 文件夹的文件中定义。`OneDriveItem` 类在 **Models** 文件夹的一个文件中定义。 这些类的详细讨论内容与授权或 SSO 无关，因此不在本文的讨论范围内。
    * 通过向 Microsoft Graph 请求仅获取实际需要的数据，可以提升性能，因此代码使用 ` $select` 查询参数来指定仅需要 name 属性，并使用 `$top` 参数来指定仅需要前 3 个文件夹或文件名。

    ```
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    var getFilesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    ```

9. 将 TODO9 替换为以下代码行。 请注意，尽管上述代码仅需要 OneDrive 项的 *name* 属性，但 Microsoft Graph 始终包括 OneDrive 项的 *eTag* 属性。 为减少发送到客户端的有效负载，下面的代码仅使用项目名称重建结果。

    ```
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in getFilesResult)
    {
      itemNames.Add(item.Name);
    }                    
    return itemNames;
    ```

## <a name="run-the-add-in"></a>运行加载项

1. 请确保 OneDrive 中有一些文件，以便可以验证结果。

1. 在 Visual Studio 中，按 F5。PowerPoint 将打开，“主页”****功能区上会有一个“SSO ASP.NET”****组。

1. 按此组中的“显示外接程序”****按钮，在任务窗格中查看此外接程序的 UI。

1. 按“从 OneDrive 获取我的文件”****按钮。 如果尚未登录 Office，便会看到登录提示。
    > **注意：**如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用程序现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已在 PowerPoint 中更改过。 在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。 为了防止发生这种情况，请务必先*关闭其他所有 Office 应用程序*，然后再按“从 OneDrive 获取我的文件”****。

1. 登录后，OneDrive 上的文件和文件夹列表将会显示在此按钮下方。 这可能需要超过 15 秒才能完成，特别是首次运行时。
