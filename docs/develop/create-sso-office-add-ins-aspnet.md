---
title: 创建使用单一登录的 ASP.NET Office 加载项
description: 有关如何使用 ASP.NET 后端创建 (或转换) Office 加载项以使用单一登录 (SSO) 的分步指南。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 980ae1b9c36dfdf7fcf84ad4fb1ba9088687cf7a
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889378"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a>创建使用单一登录的 ASP.NET Office 加载项

如果用户已登录 Office，加载项可以使用相同的凭据，这样用户无需重新登录，即可访问多个应用程序。 有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。
本文介绍如何在使用 ASP.NET 生成的加载项中启用单一登录 (SSO) 。

## <a name="prerequisites"></a>先决条件

- Visual Studio 2019 或更高版本。

- 配置 Visual Studio 时的 **Office/SharePoint 开发** 工作负荷。

- [Office 开发人员工具](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

- Microsoft 365 订阅中至少存储在OneDrive for Business上的几个文件和文件夹。

- 具有活动订阅的 Azure 帐户 - [免费创建帐户](https://azure.microsoft.com/free/?WT.mc_id=A261C142F)。

## <a name="set-up-the-starter-project"></a>设置初学者项目

在 [Office Add-in ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) 处克隆或下载存储库。

> [!NOTE]
> 示例有两个版本。
>
> - **Before** 文件夹是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。本文后续章节将引导你完成此过程。
> - 如果完成了本文中的过程，该示例的 **已完成** 版本会与所生成的加载项类似，只不过完成的项目具有对本文文本冗余的代码注释。 若要使用已完成的版本，请按照本文中的说明进行操作即可，但需要将“Before”替换为“Complete”，并跳过 **编写客户端代码** 和 **编写服务器端代码** 部分。

## <a name="register-the-add-in-through-an-app-registration"></a>通过应用注册注册加载项

首先，完成[快速入门中的步骤：向Microsoft 标识平台注册](/azure/active-directory/develop/quickstart-register-app)应用程序以注册加载项。

对应用注册使用以下设置。

- 名称：`Office-Add-in-ASPNET-SSO`
- 支持的帐户类型： **任何组织目录中的帐户 (任何 Azure AD 目录 - 多租户) 和个人 Microsoft 帐户 (，例如 Skype、Xbox)**

    > [!NOTE]
    >  如果希望加载项仅供你正在注册的租户中的用户使用，则 **只能在此组织目录中选择帐户...** 请参阅本文后面 **的单租户安装程序** 。

- 平台： **Web**
- 重定向 URI： **https://localhost:44355/AzureADAuth/Authorize**
- 客户端机密： `*********` (Web 应用程序在请求令牌时使用客户端机密来证明其标识。 *记录此值以便在后续步骤中使用 - 它只显示一次。*) 

### <a name="expose-a-web-api"></a>公开 Web API

1. 在创建的应用注册中，选择 **“公开 API >添加范围**。
   如果尚未配置应用程序 **ID URI** ，系统会提示你设置该 URI。

    应用 ID URI 充当将在 API 代码中引用的范围的前缀，并且它必须全局唯一。 使用窗体 `api://localhost:44355/[application-id-guid]`;例如 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

1. 在 **“添加范围** ”窗格中指定作用域的属性。

    |字段          |值  |
    |---------------|---------|
    |**范围名称** | `access_as_user`|
    |**谁可以同意** | **管理员和用户**|
    |**管理员许可显示名称** | Office 可以充当用户。|
    |**管理员许可说明** | 使 Office 能够使用与当前用户相同的权限调用外接程序的 Web API。|
    |**用户同意显示名称** | Office 可以充当你。|
    |**用户同意说明** | 使 Office 能够使用你拥有的相同权限调用外接程序的 Web API。|

1. 将 **状态** 设置为 **“已启用**”，然后选择 **“添加范围**”。

    > [!NOTE]
    > 文本字段正下方显示的名称的 **\<Scope\>** 域部分应自动匹配之前设置的应用程序 ID URI，`/access_as_user`并追加到末尾;例如。 `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`

1. 在 **“授权客户端应用程序** ”部分中，输入以下 ID 以预授权所有 Microsoft Office 应用程序终结点。

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (所有 Microsoft Office 应用程序终结点) 

    > [!NOTE]
    > 该 `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID 在以下所有平台上预授权 Office。 或者，如果出于任何原因想要拒绝某些平台上的 Office 授权，则可以输入以下 ID 的适当子集。 只需保留要从中隐瞒授权的平台的 ID 即可。 这些平台上外接程序的用户将无法调用 Web API，但外接程序中的其他功能仍将有效。
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5`（Office 网页版）
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）

1. 选择 **添加客户端应用程序**。 在打开的面板中，将客户端 ID 设置为相应的 GUID，然后选中该框 `api://localhost:44355/[application-id-guid]/access_as_user`。

1. 选择“添加应用程序”。

### <a name="configure-microsoft-graph-permissions"></a>配置 Microsoft Graph 权限

1. 选择 **API 权限>添加权限> Microsoft Graph**。

1. 选择“**委托的权限**”。 Microsoft Graph 公开了许多权限，最常用的权限显示在列表顶部。

1. 在 **“选择权限**”下，选择以下权限。

    |权限     |说明  |
    |---------------|-------------|
    |Files.Read.All |读取用户可以访问的所有文件。 |
    |个人资料        |查看用户的基本配置文件。 Office 应用程序需要获取加载项 Web 应用程序的令牌。 |

    > [!NOTE]
    > `User.Read` 权限可能已默认列出。 根据最佳做法，最好不要请求授予不需要的权限，因此，如果加载项实际上不需要此权限，我们建议取消选中此权限对应的框。

1. 选择 **“添加权限** ”以完成此过程。

每当你配置权限时，都会在登录时要求应用的用户同意允许应用代表他们访问资源 API。 作为管理员，你还可以代表所有用户授予同意，这样就不会提示他们这样做。

1. 在同一页面上，选择“**为[租户名称]授予管理员许可**”按钮，然后在显示的确认中选择“**接受**”。

    > [!NOTE]
    > 选择“**为[租户名称]授予管理员许可** 后，可能会看到一条横幅消息，要求你在几分钟后再次尝试，以便能够构建许可提示。 如果是这样，则可以在下一部分开始工作， **_但不要忘记回到门户并按下此按钮_**！

## <a name="configure-the-solution"></a>配置解决方案

1. 在 **Before** 文件夹的根部，打开 **Visual Studio** 中的解决方案 (.sln) 文件。 右键单击“**解决方案资源管理器**”最上面的节点（即“解决方案”节点，而非任何项目节点），然后选择“**设置启动项目**”。

1. 在“**通用属性**”下，选择“**启动项目**”，然后选择“**多个启动项目**”。 确保两个项目的“**操作**”均设置为“**启动**”，并且以“...WebAPI”结尾的项目排在前面。 关闭该对话框。

1. 返回 **解决方案资源管理器**，选择 (不要右键单击 **Office-Add-in-ASPNET-SSO-WebAPI** 项目) 。 随后将打开“**属性**”窗格。 确保“**已启用 SSL**”为“**True**”。 验证“**SSL URL**”是否为 `http://localhost:44355/`。

1. 在“Web.config”中，使用先前复制的值。 将“**ida:ClientID**”和“**ida:Audience**”均设置为“**应用程序(客户端) ID**”，并将“**ida:Password**”设置为客户端密码。 此外，将 **ida：Domain** 设置为 `http://localhost:44355` (末尾没有正斜杠“/”，) 。

    > [!NOTE]
    > 当其他应用程序 **（例如 Office 客户** 端应用程序 (（例如，PowerPoint、Word、Excel) ）寻求对应用程序的授权访问时，应用程序 (客户端) ID 是“受众”值。 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

1. 如果在注册该加载项时，“**受支持的帐户类型**”未选择“仅限此组织目录中的帐户”，请保存并关闭 web.config。否则，请保存，但将其保持打开状态。

1. 在 **解决方案资源管理器** 中，选择 **Office-Add-in-ASPNET-SSO** 项目，打开加载项清单文件“Office-Add-in-ASPNET-SSO.xml”，然后滚动到文件底部。 在结束 `</VersionOverrides>` 标记的正上方，你会发现以下标记。

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. 将标记中的 *两处* 占位符“$application_GUID here$”均替换为在注册加载项时复制的应用程序 ID。 由于 ID 并不包含“$”符号，因此请勿添加它们。 这与在 web.config 中对 ClientID 和 Audience 所使用的 ID 相同。

    > [!NOTE]
    > 该 **\<Resource\>** 值是注册加载项时设置的应用程序 **ID URI** 。 仅当加载项通过 AppSource 出售时，该 **\<Scopes\>** 部分仅用于生成同意对话框。

1. 保存并关闭此文件。

### <a name="setup-for-single-tenant"></a>单租户设置

如果在注册外接程序时为 **支持的帐户类型** 选择了“仅此组织目录中的帐户”，则需要执行这些额外的设置步骤。

1. 返回 Azure 门户，并打开加载项注册界面的“**概述**”边栏选项卡。 复制“**目录(租户) ID**”。

1. 在 web.config 中，将“**ida:Authority**”的值中的“common”替换为上一步复制的 GUID。 完成后，值应如下所示：`<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`。

1. 保存并关闭 web.config。

## <a name="code-the-client-side"></a>编写客户端代码

1. 打开 **Scripts** 文件夹中的 HomeES6.js 文件。 其中已有一些代码。

    - 有一些填充代码用于向全局窗口对象分配 Office.Promise 对象，以便在 Office 为 UI 使用 Internet Explorer 时可运行该加载项。 （有关详细信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。）
    - 针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。
    - `showResult` 方法，用于在任务窗格底部显示从 Microsoft Graph 返回的数据（或错误消息）。
    - `logErrors` 方法，用于记录最终用户不应看到的控制台错误。
    - 在不支持 SSO 或发生错误的情况下实现外接程序将使用的回退授权系统的代码。

1. 在针对 `Office.initialize` 的分配下面，添加下面的代码。 关于此代码，请注意以下几点：

    - 加载项中的错误处理有时会自动尝试使用一组不同的选项，重新获取访问令牌。 计数器变量 `retryGetAccessToken` 用于确保用户不会重复循环失败的尝试来获取令牌。
    - `getGraphData` 函数通过 ES6 `async` 关键字进行定义。 使用 ES6 语法可以使 Office 加载项中的 SSO API 更易于使用。 此文件是该解决方案中唯一会使用 Internet Explorer 不支持的语法的文件。 我们在文件名中放入“ES6”作为提醒用途。 该解决方案使用 tsc 转译器将此文件转译为 ES5，以便在 Office 为 UI 使用 Internet Explorer 时可运行加载项。 （请查看项目根目录中的 tsconfig.json 文件。）

    ```javascript
    let retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. 在 `getGraphData` 函数下方，添加下列函数。 请注意，你将在稍后的步骤中创建 `handleClientSideErrors` 函数。

    > [!NOTE]
    > 为了区分本文中使用的两个访问令牌，从 getAccessToken () 返回的令牌称为启动令牌。 稍后会通过代表流交换该令牌，以获取有权访问 Microsoft Graph 的新令牌。

    ```javascript
    async function getDataWithToken(options) {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for a new access token to Microsoft Graph and then get the data
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

1. 替换 `TODO 1` 为以下代码以从 Office 主机获取访问令牌。 **options** 参数包含从上一 **个 getGraphData ()** 函数传递的以下设置。

    - `allowSignInPrompt` 设置为 true。 这会告知 Office 在用户尚未登录 Office 时提示用户登录。
    - `allowConsentPrompt` 设置为 true。 这会指示 Office 在尚未授予许可的情况下，提示用户同意允许外接程序访问用户的Microsoft Azure Active Directory配置文件。  (生成的提示 *不允许* 用户同意任何 Microsoft Graph 范围。) 
    - `forMSGraphAccess` 设置为 true。 如果用户或管理员尚未授予对加载项的 Graph 范围的同意，则此操作会通知 Office 返回 (代码 13012) 的错误。 若要访问 Microsoft Graph，外接程序必须通过代表流交换新访问令牌的访问令牌。 设置 `forMSGraphAccess` 为 true 有助于避免 **getAccessToken ()** 成功，但随后 Microsoft Graph 代表流失败的情况。 加载项的客户端代码可以通过分支到回退授权系统来响应 13012。

    另请注意以下代码：

    - 你将在稍后的步骤中创建 `getData` 函数。
    - 该 `/api/values` 参数是服务器端控制器的 URL，该控制器将使用代表流交换令牌以调用 Microsoft Graph 的新访问令牌。

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken(options);

    getData("/api/values", bootstrapToken);
    ```

1. 在 `getGraphData` 函数下方，添加以下内容。 关于此代码，请注意以下几点：

    - SSO 和回退授权系统均会使用它。
    - `relativeUrl` 参数是服务器端控制器。
    - `accessToken` 参数可以是启动令牌或完全访问令牌。
    - `writeFileNamesToOfficeDocument` 已是项目的一部分。
    - 你将在稍后的步骤中创建 `handleServerSideErrors` 函数。

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
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to sign in, press the Get OneDrive File Names button again."]);
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

1. 将 `TODO 3` 替换为下面的代码。 对于所有其他错误，加载项会分支到回退授权系统。 有关这些错误的详细信息，请参阅 [Office 加载项中的 SSO 疑难解答](troubleshoot-sso-in-office-add-ins.md)。在此加载项中，回退系统将打开一个对话框，要求用户登录，即使用户已经登录。

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
    const message = JSON.parse(result.responseText).Message;
    const exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. 将 `TODO 5` 替换为以下代码。 Microsoft Graph 要求进行其他形式的身份验证时，将发送错误 AADSTS50076。 其中包括 **Message.Claims** 属性中的附加要求的相关信息。 为处理这种情况，该代码会再次尝试获取启动令牌，但这一次还包括请求额外的因素作为 `authChallenge` 选项的值，这会告诉 Azure AD 提示用户输入所有必需的身份验证形式。

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            const claims = JSON.parse(message).Claims;
            const claimsAsString = JSON.stringify(claims);
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

    - 代码指示 OWIN 确保来自 Office 应用程序的启动令牌中指定的受众必须与web.config中指定的值匹配。
    - Microsoft 帐户具有不同于任何组织租户 GUID 的颁发者 GUID，因此为了支持这两种类型的帐户，我们不验证颁发者。
    - 导致 `true` OWIN 从 Office 应用程序保存原始启动令牌的设置`SaveSigninToken`。 加载项需要该令牌来获取具有代理流的 Microsoft Graph 访问令牌。
    - OWIN 中间件不验证作用域。 启动令牌作用域应包括 `access_as_user`，在控制器中加以验证。

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. 将 `TODO 2` 替换为以下代码。 关于此代码，请注意以下几点：

    - 调用的是方法 `UseOAuthBearerAuthentication`，而不是更常见的 `UseWindowsAzureActiveDirectoryBearerAuthentication`，因为后者与 Azure AD V2 终结点不兼容。
    - 传递给方法的 URL 是 OWIN 中间件获取获取密钥的说明的位置，它需要验证从 Office 应用程序收到的启动令牌上的签名。 此 URL 的 Authority 区段来自 web.config。它可能是“common”字符串，而对于单租户加载项，则是一个 GUID。

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

1. 将下列方法添加到 `ValuesController`。 请注意，返回值是 `Task<HttpResponseMessage>`（而不是 `Task<IEnumerable<string>>`），这对于 `GET api/values` 方法而言更为常见。 这是一个副作用，即 OAuth 授权逻辑必须位于控制器中，而不是在 ASP.NET 筛选器中。 该逻辑中的一些错误条件要求将 HTTP 响应对象发送到加载项的客户端。

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //         token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get a new access token for Microsoft Graph.

        // TODO 4: Use the new access token to call Microsoft Graph.
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

    - 加载项不再扮演资源 (或受众) Office 应用程序和用户需要访问的角色。 现在它本身就是一个需要访问 Microsoft Graph 的客户端。 是 MSAL“客户端上下文”对象。
    - 从 MSAL.NET 3.x.x 开始，`bootstrapContext` 仅仅是启动令牌本身。
    - Authority 来自 web.config。它可能是“common”字符串，而对于单租户加载项，则是一个 GUID。
    - 如果代码请求 `profile`，MSAL 将引发错误，而此错误仅在 Office 客户端应用程序获取加载项 Web 应用程序的令牌时才使用。 因此，只会显式请求获取 `Files.Read.All`。

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

    - `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` 方法将首先查找内存中的 MSAL 缓存，获取匹配的访问令牌。 仅当不存在任何令牌时，该方法才会通过 Azure AD V2 终结点启动代理流。
    - 任何不属于类型 `MsalServiceException` 的异常都是有意不捕获的，这样才能作为 `500 Server Error` 消息传播到客户端。

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

    - 如果 Microsoft Graph 资源要求进行多重身份验证，但用户尚未提供，则 Azure AD 会返回“400 错误请求”以及错误 `AADSTS50076` 和 **Claims** 属性。 MSAL 抛出包含此信息的 **MsalUiRequiredException**（继承自 **MsalServiceException**）。
    - 声明 **属性值** 必须传递给客户端，客户端应将其传递给 Office 应用程序，然后在请求新的启动令牌时包含该属性值。 Azure AD 会提示用户进行所有必需形式的身份验证。
    - 由于创建异常 HTTP Response 的 API 并不知道 **Claims** 属性，因此它们不会在 Response 对象中添加这个属性。 必须手动创建消息来添加它。 不过，自定义 **Message** 属性会阻止创建 **ExceptionMessage** 属性，因此向客户端发送错误 ID `AADSTS50076` 的唯一方法是，将它添加到自定义 **Message** 中。 客户端中的 JavaScript 需要发现响应是否包含 **Message** 或 **ExceptionMessage**，这样才能了解要读取的内容。
    - 自定义消息被格式化为 JSON，以便客户端 JavaScript 能够使用已知的 JavaScript `JSON` 对象方法分析它。

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. 将 `TODO 3b` 替换为下面的代码。 关于此代码，请注意以下几点：

    - 如果 Azure AD 调用包含至少一个作用域（权限）未获得用户和租户管理员的许可（或许可被撤消），则 Azure AD 将返回“400 错误请求”和错误 `AADSTS65001`。 MSAL 抛出包含此信息的 **MsalUiRequiredException**。
    - 如果 Azure AD 调用包含至少一个 Azure AD 无法识别的作用域，则 AAD 将返回“400 错误请求”和错误 `AADSTS70011`。 MSAL 抛出包含此信息的 **MsalUiRequiredException**。
    - 其中包含完整说明，因为 70011 也会在其他情况下返回，只有在它表示存在无效范围时，才需要在此加载项中处理它。
    - **MsalUiRequiredException** 对象传递给 `SendErrorToClient`。这样可确保 HTTP 响应中有包含错误消息的 **ExceptionMessage** 属性。

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. 将 `TODO 3c` 替换为以下代码，以处理所有其他 **MsalServiceException**。

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

    ![选择所需的 Office 客户端应用程序：Excel、PowerPoint 或 Word。](../images/SelectHost.JPG)

1. 按 F5。
1. 在 Office 应用程序的“**主页**”功能区上，选择“**SSO ASP.NET**”组中的“**显示加载项**”以打开任务窗格加载项。
1. 单击“**获取 OneDrive 文件名**”按钮。 如果使用Microsoft 365 教育版或工作帐户或 Microsoft 帐户登录到 Office，并且 SSO 按预期工作，则OneDrive for Business中的前 10 个文件和文件夹名称将显示在任务窗格中。 如果未登录，或者处于不支持 SSO 的方案中，或者 SSO 出于任何原因无法正常工作，系统会提示你登录。 登录后，将显示文件和文件夹名称。

### <a name="testing-the-fallback-path"></a>测试回退路径

若要测试回退授权路径，请使用以下步骤强制 SSO 路径失败。

1. 将以下代码添加到HomeES6.js文件中方法的顶部 `getDataWithToken` 。

    ```javascript
    function MockSSOError(code) {
        this.code = code;
    }
    ```

1. 然后，将以下行添加到同一方法中的块顶部 `try` ，就在调用 `getAccessToken`的正上方。

    ```javascript
    throw new MockSSOError("13003");
    ```

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a>转到过渡和生产时更新加载项

与所有 Office Web 加载项一样，当你准备好移动到过渡服务器或生产服务器时，必须使用新域更新 `localhost:44355` 清单中的域。 同样，必须更新web.config文件中的域。

由于域显示在 AAD 注册中，因此需要更新该注册以使用新域来代替 `localhost:44355` 它出现的位置。
