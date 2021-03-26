---
title: 创建使用单一登录的 ASP.NET Office 加载项
description: 有关如何使用 (后端创建 (或) Office 外接程序以使用单一登录 ASP.NET SSO (的分步指南) 。
ms.date: 03/11/2021
localization_priority: Normal
ms.openlocfilehash: e92bac3be81254a4c15f5e071602edbe788692ac
ms.sourcegitcommit: 5ad32261f80e7ab371aba032d9024ad1275c23f9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2021
ms.locfileid: "51221372"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>创建使用单一登录的 ASP.NET Office 加载项

如果用户已登录 Office，加载项可以使用相同的凭据，这样用户无需重新登录，即可访问多个应用程序。 有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。
本文将引导你完成在内置加载项 (SSO) 启用单一登录 ASP.NET。

> [!NOTE]
> 有关与此类似的 Node.js 加载项文章，请参阅[创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)。

## <a name="prerequisites"></a>先决条件

* Visual Studio 2019 或更高版本。

* [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* 在 Microsoft 365 订阅中，OneDrive for Business 上至少存储了一些文件和文件夹。

* 一个 Microsoft Azure 订阅。 此加载项需要 Azure Active Directory (AD)。 Azure AD 为应用程序提供了用于进行身份验证和授权的标识服务。 你还可在 [Microsoft Azure](https://account.windowsazure.com/SignUp) 获得试用订阅。

## <a name="set-up-the-starter-project"></a>设置初学者项目

在 [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso) 处克隆或下载存储库。

> [!NOTE]
> 示例项目有两个版本：
>
> * **Before** 文件夹是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。本文后续章节将引导你完成此过程。
> * 如果完成了本文中的过程，该示例的 **已完成** 版本会与所生成的加载项类似，只不过完成的项目具有对本文文本冗余的代码注释。 若要使用已完成的版本，请按照本文中的说明进行操作即可，但需要将“Before”替换为“Complete”，并跳过 **编写客户端代码** 和 **编写服务器端代码** 部分。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>向 Azure AD v2.0 终结点注册加载项。

1. 导航到“Azure 门户 - 应用注册”[](https://go.microsoft.com/fwlink/?linkid=2083908)页面以注册你的应用。

1. 使用管理员 ***凭据*** 登录 Microsoft 365 租赁。 例如，MyName@contoso.onmicrosoft.com。

1. 选择“新注册”。 在“注册应用”页上，按如下方式设置值。

    * 将“名称”设置为“`Office-Add-in-ASPNET-SSO`”。
    * 将“**受支持的帐户类型**”设置为“**任何组织目录中的帐户和个人 Microsoft 帐户(任何 Azure AD 目录 - 多租户)**”（例如，Skype、Xbox）。 （如果希望加载项仅可供注册该加载项的租户中的用户使用，则可以选择“**仅限此组织目录中的帐户...**”，但需要执行一些额外的设置步骤。 请参阅下面的 **单租户设置**。）
    * 在“**重定向 URI**”部分，确保在下拉列表中选择“**Web**”，然后将 URI 设置为 ` https://localhost:44355/AzureADAuth/Authorize`。
    * 选择“**注册**”。

1. 在 **"Office-Add-in-ASPNET-SSO"** 页上，复制并保存 Application **(client) ID** 和 **Directory (tenant) ID 的值**。 你将在后面的过程中使用它们。

    > [!NOTE]
    > 当其他应用程序（如 PowerPoint、Word、Excel) 等 Office 客户端应用程序 (）寻求对该应用程序的授权访问权限时，此应用程序客户端) ID 是"受众"值。 **(** 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

1. 在“**管理**”下，选择“**证书和密码**”。 选择“**新客户端密码**”按钮。 输入“**描述**”的值，然后选择适当的“**到期**”选项，并选择“**添加**”。 在继续操作前，*立即复制客户端密码值并使用应用程序 ID 保存它*，因为在后面的过程中，将需要用到它。

1. 在“**管理**”下，选择“**公开 API**”。 选择“**设置**”链接以在窗体“api://$App ID GUID$”中生成应用 ID URI，其中 $App ID GUID$ 是 **应用程序（客户端）ID**。 在 `//` 后面和 GUID 前面插入 `localhost:44355/`（请注意结尾附加的正斜杠“/”）。 整个 ID 的格式应为 `api://localhost:44355/$App ID GUID$`；例如 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

1. 在对话框中选择“**保存**”。

1. 选择“**添加一个作用域**”按钮。 在打开的面板中，输入 `access_as_user` 作为 **作用域** 名称。

1. 将“谁能同意?”设置为“管理员和用户”。

1. 使用适用于作用域的值填写用于配置管理员和用户同意提示的字段，这些值使 Office 客户端应用程序能够使用与当前用户相同的权限使用外接程序的 Web API。 `access_as_user` 建议：

    * **管理员显示名称：Office** 可以充当用户。
    * **管理员许可描述**：使 Office 能够借助与当前用户相同的权限调用加载项的 Web API。
    * **用户同意显示名称：Office** 可以充当您。
    * **用户同意描述**：允许 Office 使用你拥有的相同权限调用外接程序的 Web API。

1. 确保将“**状态**”设置为“**已启用**”。

1. 选择“**添加作用域**”。

    > [!NOTE]
    > 显示在文本字段正下方的 **作用域** 名称的域部分应自动与你先前设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。

1. 在“授权客户端应用程序”部分中，确定要授权给加载项 Web 应用程序的应用程序。 下面每个 ID 都需要进行预授权。

    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）
    * `08e18876-6177-487e-b8b5-cf950c1e598c`（Office 网页版）
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）

    对于每个 ID，执行以下步骤：

    a. 选择“**添加客户端应用程序**”按钮，然后在打开的面板中，将“客户端 ID”设置为相应的 GUID 并勾选 `api://localhost:44355/$App ID GUID$/access_as_user` 框。

    b. 选择“添加应用程序”。

1. 在“**管理**”下，选择“**API 权限**”，然后选择“**添加权限**”。 在打开的面板上，选择 **Microsoft Graph**，然后选择“委派权限”。

1. 使用“选择权限”搜索框来搜索加载项需要的权限。 选择以下选项。 外接程序本身确实只需要第一项;但 `profile` Office 应用程序需要权限才能获取外接程序 Web 应用程序的令牌。 （该加载项实际上仅需要 Files.Read.All 和 profile。 但必须请求其他两个，因为 MSAL.NET 库需要它们。）

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > `User.Read` 权限可能已默认列出。 根据最佳做法，最好不要请求授予不需要的权限，因此，如果加载项实际上不需要此权限，我们建议取消选中此权限对应的框。

1. 选择所显示的每个权限的复选框。 选择加载项需要的权限后，选择面板底部的“**添加权限**”按钮。

1. 在同一页面上，选择“**为[租户名称]授予管理员许可**”按钮，然后在显示的确认中选择“**接受**”。

    > [!NOTE]
    > 选择“**为[租户名称]授予管理员许可** 后，可能会看到一条横幅消息，要求你在几分钟后再次尝试，以便能够构建许可提示。 如果是这样，你可以开始下一部分，但不要忘记回到门户并 **_按此按钮_**！

## <a name="configure-the-solution"></a>配置解决方案

1. 在 **Before** 文件夹的根部，打开 **Visual Studio** 中的解决方案 (.sln) 文件。 右键单击“**解决方案资源管理器**”最上面的节点（即“解决方案”节点，而非任何项目节点），然后选择“**设置启动项目**”。

1. 在“**通用属性**”下，选择“**启动项目**”，然后选择“**多个启动项目**”。 确保两个项目的“**操作**”均设置为“**启动**”，并且以“...WebAPI”结尾的项目排在前面。 关闭该对话框。

1. 返回到解决方案 **资源管理器**， (不要右键) **Office-Add-in-ASPNET-SSO-WebAPI** 项目。 随后将打开“**属性**”窗格。 确保“**已启用 SSL**”为“**True**”。 验证“**SSL URL**”是否为 `http://localhost:44355/`。

1. 在“Web.config”中，使用先前复制的值。 将“**ida:ClientID**”和“**ida:Audience**”均设置为“**应用程序(客户端) ID**”，并将“**ida:Password**”设置为客户端密码。 此外，将 **ida：Domain** 设置为 (末尾没有正斜杠 `http://localhost:44355` "/") 。 

    > [!NOTE]
    > 当其他应用程序（如 PowerPoint、Word、Excel) 等 Office 客户端应用程序 (）寻求对该应用程序的授权访问权限时，Application (客户端) **ID** 是"受众"值。 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

1. 如果在注册该加载项时，“**受支持的帐户类型**”未选择“仅限此组织目录中的帐户”，请保存并关闭 web.config。否则，请保存，但将其保持打开状态。

1. 仍在 **"** 解决方案资源管理器"中，选择 **"Office-Add-in-ASPNET-SSO"** 项目，打开外接程序清单文件"Office-Add-in-ASPNET-SSO.xml"，然后滚动到文件底部。 在结尾的 `</VersionOverrides>` 标记的正上方有以下标记：

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. 将标记中的 *两处* 占位符“$application_GUID here$”均替换为在注册加载项时复制的应用程序 ID。 由于 ID 并不包含“$”符号，因此请勿添加它们。 这与在 web.config 中对 ClientID 和 Audience 所使用的 ID 相同。

  > [!NOTE]
  > **资源** 值是注册加载项时设置的 **应用程序 ID URI**。 仅在通过 AppSource 销售加载项时，才使用 **作用域** 部分生成许可对话框。

1. 保存并关闭此文件。

### <a name="setup-for-single-tenant"></a>单租户设置

如果在注册该加载项时，“**受支持的帐户类型**”选择了“仅限此组织目录中的帐户”，则需要执行以下额外的设置步骤：

1. 返回 Azure 门户，并打开加载项注册界面的“**概述**”边栏选项卡。 复制“**目录(租户) ID**”。

1. 在 web.config 中，将“**ida:Authority**”的值中的“common”替换为上一步复制的 GUID。 完成后，值应如下所示：`<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`。

1. 保存并关闭 web.config。

## <a name="code-the-client-side"></a>编写客户端代码

1. 打开 **Scripts** 文件夹中的 HomeES6.js 文件。 其中已存在一些代码：

    * 有一些填充代码用于向全局窗口对象分配 Office.Promise 对象，以便在 Office 为 UI 使用 Internet Explorer 时可运行该加载项。 （有关详细信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。）
    * 针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。
    * `showResult` 方法，用于在任务窗格底部显示从 Microsoft Graph 返回的数据（或错误消息）。
    * `logErrors` 方法，用于记录最终用户不应看到的控制台错误。
    * 一些代码实现了加载项在 SSO 不受支持或有错误的情况下使用的回退授权系统。

1. 在针对 `Office.initialize` 的分配下面，添加下面的代码。 关于此代码，请注意以下几点：

    * 加载项中的错误处理有时会自动尝试使用一组不同的选项，重新获取访问令牌。 计数器变量 `retryGetAccessToken` 用于确保用户不会重复循环失败的尝试来获取令牌。
    * `getGraphData` 函数通过 ES6 `async` 关键字进行定义。 使用 ES6 语法可以使 Office 加载项中的 SSO API 更易于使用。 此文件是该解决方案中唯一会使用 Internet Explorer 不支持的语法的文件。 我们在文件名中放入“ES6”作为提醒用途。 该解决方案使用 tsc 转译器将此文件转译为 ES5，以便在 Office 为 UI 使用 Internet Explorer 时可运行加载项。 （请查看项目根目录中的 tsconfig.json 文件。）

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. 在 `getGraphData` 函数下方，添加下列函数。 请注意，你将在稍后的步骤中创建 `handleClientSideErrors` 函数。

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graph and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```

1. 将 `TODO 1` 替换为以下代码。 关于此代码，请注意以下几点：

    * `getAccessToken` 告知 Office 从 Azure AD 获取启动令牌并返回给加载项。
    * `allowSignInPrompt` 在用户尚未登录 Office 的情况下告知 Office 提示用户进行登录。
    * `allowConsentPrompt` 指示 Office 提示用户同意允许外接程序访问用户的 AAD 配置文件（如果尚未授予同意）。  (生成的提示 *不允许* 用户同意任何 Microsoft Graph 范围。) 
    * `forMSGraphAccess` 告知 Office 该加载项打算使用启动令牌来换取 Microsoft Graph 的访问令牌（而不是仅将启动令牌用作用户 ID 令牌）。 通过设置此选项，如果用户的租户管理员尚未向加载项授予许可，则 Office 有机会取消获取启动令牌的过程（并返回错误代码 13012）。 加载项的客户端代码可以通过分支到回退授权系统来响应 13012。 如果未使用 且管理员未授予同意，将返回启动令牌，但尝试与代表流交换它将导致 `forMSGraphAccess` 错误。 因此，通过 `forMSGraphAccess` 选项可以快速将加载项分支到回退系统。
    * 你将在稍后的步骤中创建 `getData` 函数。
    * `/api/values` 参数是服务器端控制器的 URL，它将进行令牌交换并使用它返回的访问令牌来对 Microsoft Graph 执行调用。

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. 在 `getGraphData` 函数下方，添加以下内容。 关于此代码，请注意以下几点：

    * SSO 和回退授权系统均会使用它。
    * `relativeUrl` 参数是服务器端控制器。
    * `accessToken` 参数可以是启动令牌或完全访问令牌。
    * `writeFileNamesToOfficeDocument` 已是项目的一部分。
    * 你将在稍后的步骤中创建 `handleServerSideErrors` 函数。

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a>处理客户端错误

1. 在 `getData` 函数下方，添加下列函数。 请注意，`error.code` 是一个数字，通常处于 13xxx 范围内。

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. 将 `TODO 2` 替换为下面的代码。 有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the web.
        showResult(["Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. 将 `TODO 3` 替换为下面的代码。 对于所有其他错误，加载项会分支到回退授权系统。 有关这些错误的详细信息，请参阅 Office 加载项中的 [SSO 疑难解答](troubleshoot-sso-in-office-add-ins.md)。在此外接程序中，回退系统将打开一个对话框，要求用户登录，即使用户已登录。

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a>处理服务器端错误

1. 在 `handleClientSideErrors` 函数下方，添加下列函数。

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. 将 `TODO 4` 替换为以下代码。 关于此代码，请注意，ASP.NET 错误类是在有类似于 MFA 的功能之前创建的。 服务器端逻辑处理针对第二种身份验证因素的请求时有一个副作用，即发送到客户端的服务器端错误有 **Message** 属性，但没有 **ExceptionMessage** 属性。 但是，所有其他错误都有 **ExceptionMessage** 属性，因此客户端代码必须分析这两者的响应。 一个或另一个变量将是未定义的。

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. 将 `TODO 5` 替换为以下代码。 Microsoft Graph 要求进行其他形式的身份验证时，将发送错误 AADSTS50076。 其中包括 **Message.Claims** 属性中的附加要求的相关信息。 为处理这种情况，该代码会再次尝试获取启动令牌，但这一次还包括请求额外的因素作为 `authChallenge` 选项的值，这会告诉 Azure AD 提示用户输入所有必需的身份验证形式。

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. 将 `TODO 6` 替换为以下代码。

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. 将 `TODO 7` 替换为以下代码。 请注意，在极少数情况下，启动令牌在由 Office 验证时未过期，但是会在发动到 Azure AD 进行交换时过期。 Azure AD 将以错误 AADSTS500133 做出响应。 发生这种情况时，代码会回调 SSO API（但不超过一次）。 这次，Office 将返回新的未过期的启动令牌。

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. 将 `TODO 8` 替换为以下代码。

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. 保存文件。

## <a name="code-the-server-side"></a>编写服务器端代码

### <a name="configure-the-owin-middleware"></a>配置 OWIN 中间件

1. 在 **Office-Add-in-ASPNET-SSO-WebAPI** 项目的根目录中打开 Startup.cs 文件，并将以下方法添加到 **Startup** 类。 请注意，你将在稍后的步骤中创建 `ConfigureAuth` 方法。

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. 保存并关闭此文件。

1. 右键单击“App_Start”文件夹，并依次选择“添加”>“类”。

1. 在“添加新项”对话框中，命名文件“Startup.Auth.cs”，再单击“添加”。

1. 将新文件中的命名空间名称缩短为 `Office_Add_in_ASPNET_SSO_WebAPI`。

1. 确保下列所有 `using` 语句都位于文件的顶部。

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. 将关键字 `partial` 添加到 `Startup` 类（如果其中尚不存在该关键字）的声明。具体应如下所示：

    `public partial class Startup`

1. 将下列方法添加到 `Startup` 类。该方法指定 OWIN 中间件如何验证从客户端 Home.js 文件的 `getData` 方法传递给它的访问令牌。每次调用使用 `[Authorize]` 属性修饰的 Web API 终结点时都会触发授权过程。

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. 将 `TODO 1` 替换为以下代码。 关于此代码，请注意以下几点：

    * 该代码指示 OWIN 确保在来自 Office 应用程序的启动令牌中指定的访问群体必须与 web.config。
    * Microsoft 帐户具有不同于任何组织租户 GUID 的颁发者 GUID，因此为了支持这两种类型的帐户，我们不会验证颁发者。
    * 设置为 将导致 OWIN 从 Office 应用程序保存原始 `SaveSigninToken` `true` 启动令牌。 加载项需要该令牌来获取具有代理流的 Microsoft Graph 访问令牌。
    * OWIN 中间件不验证作用域。 启动令牌作用域应包括 `access_as_user`，在控制器中加以验证。

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. 将 `TODO 2` 替换为以下代码。 关于此代码，请注意以下几点：

    * 调用的是方法 `UseOAuthBearerAuthentication`，而不是更常见的 `UseWindowsAzureActiveDirectoryBearerAuthentication`，因为后者与 Azure AD V2 终结点不兼容。
    * 传递到方法的 URL 是 OWIN 中间件获取获取密钥的说明，以验证从 Office 应用程序收到的启动令牌上的签名。 此 URL 的 Authority 区段来自 web.config。它可能是“common”字符串，而对于单租户加载项，则是一个 GUID。

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. 保存并关闭此文件。

### <a name="create-the-apivalues-controller"></a>创建 /api/values 控制器

1. 打开文件 **Controllers\ValueController.cs**。 SSO 系统成功获得启动令牌后，将使用此控制器。 此控制器不用作回退授权系统的一部分。 该系统使用的是已为你创建的 AzureADAuthController。

1. 请确保下列 `using` 语句位于文件顶部。

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. 在声明 `ValuesController` 的代码行的正上方，添加属性 `[Authorize]`。这可确保只要调用控制器方法时，加载项就会运行在上一过程中配置的授权过程。只有拥有对加载项的有效访问令牌，调用方才能调用控制器的方法。

1. 将下列方法添加到 `ValuesController`。 请注意，返回值是 `Task<HttpResponseMessage>`（而不是 `Task<IEnumerable<string>>`），这对于 `GET api/values` 方法而言更为常见。 由于 OAuth 授权逻辑必须在控制器中，而不是 ASP.NET 筛选器中，所以这是一种副作用。 该逻辑中的一些错误条件要求将 HTTP 响应对象发送到加载项的客户端。

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. 将 `TODO1` 替换为以下代码，以验证令牌中指定的作用域是否包括 `access_as_user`。 请注意，`SendErrorToClient` 方法的第二个参数是 **Exception** 对象。 在此示例中，代码传递 `null`，因为添加 **Exception** 对象会阻止在生成的 HTTP Response 中添加 **Message** 属性。


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. 将 `TODO 2` 替换为以下代码，以便整合在使用代理流来获取 Microsoft Graph 的令牌时所需的所有信息。 关于此代码，请注意以下几点：

    * 您的外接程序不再扮演 Office 应用程序或用户 (访问) 访问群体的角色。 现在它本身就是一个需要访问 Microsoft Graph 的客户端。 是 MSAL“客户端上下文”对象。
    * 从 MSAL.NET 3.x.x 开始，`bootstrapContext` 仅仅是启动令牌本身。
    * Authority 来自 web.config。它可能是“common”字符串，而对于单租户加载项，则是一个 GUID。
    * MSAL 要求 `openid`、`offline_access` 作用域能够发挥作用，但如果代码过多地发出请求，则会抛出错误。 如果代码请求 ，也会引发错误，这仅在 Office 客户端应用程序获取加载项 Web 应用程序的令牌时 `profile` 真正使用。 因此，只会显式请求获取 `Files.Read.All`。

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri(ConfigurationManager.AppSettings["ida:Domain"])
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. 将 `TODO 3` 替换为下面的代码。 关于此代码，请注意以下几点：

    * `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` 方法将首先查找内存中的 MSAL 缓存，获取匹配的访问令牌。 仅当不存在任何令牌时，该方法才会通过 Azure AD V2 终结点启动代理流。
    * 任何不属于类型 `MsalServiceException` 的异常都是有意不捕获的，这样才能作为 `500 Server Error` 消息传播到客户端。

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. 将 `TODO 3a` 替换为下面的代码。 关于此代码，请注意以下几点：

    * 如果 Microsoft Graph 资源要求进行多重身份验证，但用户尚未提供，则 Azure AD 会返回“400 错误请求”以及错误 `AADSTS50076` 和 **Claims** 属性。 MSAL 抛出包含此信息的 **MsalUiRequiredException**（继承自 **MsalServiceException**）。
    * **必须将 Claims** 属性值传递到客户端，客户端应传递到 Office 应用程序，然后它将包括在请求新的启动令牌中。 Azure AD 会提示用户进行所有必需形式的身份验证。
    * 由于创建异常 HTTP Response 的 API 并不知道 **Claims** 属性，因此它们不会在 Response 对象中添加这个属性。 必须手动创建消息来添加它。 不过，自定义 **Message** 属性会阻止创建 **ExceptionMessage** 属性，因此向客户端发送错误 ID `AADSTS50076` 的唯一方法是，将它添加到自定义 **Message** 中。 客户端中的 JavaScript 需要发现响应是否包含 **Message** 或 **ExceptionMessage**，这样才能了解要读取的内容。
    * 自定义消息被格式化为 JSON，以便客户端 JavaScript 能够使用已知的 JavaScript `JSON` 对象方法分析它。

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. 将 `TODO 3b` 替换为下面的代码。 关于此代码，请注意以下几点：

    * 如果 Azure AD 调用包含至少一个作用域（权限）未获得用户和租户管理员的许可（或许可被撤消），则 Azure AD 将返回“400 错误请求”和错误 `AADSTS65001`。 MSAL 抛出包含此信息的 **MsalUiRequiredException**。
    * 如果 Azure AD 调用包含至少一个 Azure AD 无法识别的作用域，则 AAD 将返回“400 错误请求”和错误 `AADSTS70011`。 MSAL 抛出包含此信息的 **MsalUiRequiredException**。
    * 其中包含完整说明，因为 70011 也会在其他情况下返回，只有在它表示存在无效范围时，才需要在此加载项中处理它。
    * **MsalUiRequiredException** 对象传递给 `SendErrorToClient`。这样可确保 HTTP 响应中有包含错误消息的 **ExceptionMessage** 属性。

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. 将 `TODO 3c` 替换为以下代码，以处理所有其他 **MsalServiceException**。 正如前文所述，

    ```csharp
    else
    {
        throw e;
    }
    ```

1. 将 `TODO 4` 替换为以下代码。 事先为你创建的 `GraphApiHelper.GetOneDriveFileNames` 方法将向 Microsoft Graph 请求数据并包含访问令牌。

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. 保存并关闭文件。

## <a name="run-the-solution"></a>运行解决方案

1. 打开 Visual Studio 解决方案文件。
1. 在“**生成**”菜单上，选择“**清理解决方案**”。 完成后，再次打开“**生成**”菜单，并选择“**生成解决方案**”。
1. 在“**解决方案资源管理器**”中，选择“**Office-Add-in-ASPNET-SSO**”项目节点（而不是顶部的解决方案节点，也不是名称以“WebAPI”结尾的项目）。
1. 在“**属性**”窗格中，打开“**启动文档**”下拉列表，然后选择三个选项之一（“Excel”、“Word”或“PowerPoint”）。

    ![选择所需的 Office 客户端应用程序：Excel、PowerPoint 或 Word](../images/SelectHost.JPG)

1. 按 F5。
1. 在 Office 应用程序的“**主页**”功能区上，选择“**SSO ASP.NET**”组中的“**显示加载项**”以打开任务窗格加载项。
1. 单击“**获取 OneDrive 文件名**”按钮。 如果使用 Microsoft 365 教育版或工作帐户或 Microsoft 帐户登录 Office，并且 SSO 按预期工作，OneDrive for Business 中的前 10 个文件和文件夹名称将显示在任务窗格中。 如果你未登录，或者处于不支持 SSO 的情形中，或者 SSO 出于任何原因无法正常工作，则系统将提示你登录。 登录后，将显示文件和文件夹名称。

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>转到暂存和生产时更新外接程序

与所有的 Office Web 外接程序一样，当您准备好移动到暂存服务器或生产服务器时，必须使用新域更新清单 `localhost:44355` 中的域。 同样，您必须更新域的 web.config 文件。

由于该域出现在 AAD 注册中，因此您需要更新该注册以使用新域，以在它出现 `localhost:44355` 的位置进行更改。
