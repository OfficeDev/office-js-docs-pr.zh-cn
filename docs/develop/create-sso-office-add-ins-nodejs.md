---
title: 创建使用单一登录的 Node.js Office 加载项
description: 了解如何创建使用 Office 单一登录的基于 Node.js 的 Office 加载项
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: d9abb65351a0c3d4a26f06462f2a425c6a104a4a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59148866"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>创建使用单一登录的 Node.js Office 加载项

用户可以登录 Office，Office Web 加载项能够利用此登录进程，授权用户访问加载项和 Microsoft Graph，而无需要求用户再登录一次。有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。

本文将逐步介绍如何在使用 Node.js 和 Express 生成的加载项中启用单一登录 (SSO) 。 有关与此类似的 ASP.NET 加载项文章，请参阅[创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)。

> [!NOTE]
> 作为完成本文中所述步骤的替代方法，可使用 Yeoman 生成器创建启用 SSO 的 Node.js Office 加载项。 Yeoman 生成器简化了启用了 SSO 的加载项创建流程，能够自动执行在 Azure 内配置所需的步骤，并生成加载项使用 SSO 所需的代码。 有关详细信息，请参阅“[单一登录（SSO）快速入门](../quickstarts/sso-quickstart.md)”。

## <a name="prerequisites"></a>先决条件

* [Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）

* [Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）

* TypeScript，版本 3.6.2 或更高版本

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* 一个代码编辑器。 建议使用 Visual Studio Code。

* 至少存储在你的 OneDrive for Business 订阅中的Microsoft 365文件夹。

* 一个 Microsoft Azure 订阅。 此加载项需要 Azure Active Directory (AD)。 Azure AD 为应用程序提供了用于进行身份验证和授权的标识服务。 你还可在 [Microsoft Azure](https://account.windowsazure.com/SignUp) 获得试用订阅。

## <a name="set-up-the-starter-project"></a>设置初学者项目

1. 克隆或下载 [Office 外接程序 NodeJS SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) 中的存储库。

    > [!NOTE]
    > 示例有三个版本：
    >
    > * Begin 文件夹是初学者项目。 未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。 本文后续章节将引导你完成此过程。
    > * 如果完成了本文中的过程，该示例的 **已完成** 版本会与所生成的加载项类似，只不过完成的项目具有对本文文本冗余的代码注释。 若要使用已完成的版本，请按照本文中的说明操作，但将"Begin"替换为"Completed"，并跳过编写客户端代码和编写 **服务器端** 代码部分。
    > * **SSOAutoSetup** 版本是一个完整示例，可自动执行大多数步骤以在 Azure AD 中注册加载项并对其进行配置。 如果想要快速查看使用 SSO 的加载项，请使用此版本。 按照文件夹自述文件中的步骤操作即可。 我们建议你在某些时候完成本文中的手动注册和设置步骤，以更好地了解 Azure AD 与加载项之间的关系。

1. 打开 Begin 文件夹中 **的命令** 提示符。

1. 在该控制台中输入 `npm install` 以安装 package.json 文件中列出明细的所有依赖项。

1. 运行命令 `npm run install-dev-certs`。 为安装证书的提示选择“**是**”。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>向 Azure AD v2.0 终结点注册加载项。

1. 导航到“Azure 门户 - 应用注册”[](https://go.microsoft.com/fwlink/?linkid=2083908)页面以注册你的应用。

1. 使用管理员 ***凭据*** 登录您的Microsoft 365租户。 例如，MyName@contoso.onmicrosoft.com。

1. 选择“新注册”。 在“注册应用”页上，按如下方式设置值。

    * 将“名称”设置为“`Office-Add-in-NodeJS-SSO`”。
    * 将“**受支持的帐户类型**”设置为“**任何组织目录中的帐户和个人 Microsoft 帐户**”（例如，Skype、Xbox、Outlook.com）。
    * 将应用程序类型设置为 **Web，** 然后将" **重定向 URI"** 设置为 `https://localhost:44355/dialog.html` 。
    * 选择“**注册**”。

1. 在 **Office-Add-in-NodeJS-SSO** 页面上，复制并保存“**应用程序（客户端）ID**”和“**目录（租户）ID**”的值。 你将在后面的过程中使用它们。

    > [!NOTE]
    > 当其他应用程序（如 Office 客户端应用程序 (例如 PowerPoint、Word、Excel) ）寻求应用程序的授权访问权限时，此应用程序客户端) ID 是"受众"值。 **(** 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

1. 选择“**管理**”下的“**身份验证**”。 在 **隐式授予** 部分中，为访问令牌和 ID令牌 **启用复选框**。 该示例具有一个回退授权系统，当 SSO 不可用时，将调用此系统。 该系统使用隐式流。

1. 在窗体顶部，选择“**保存**”。

1. 选择“管理”下的“证书和密码”。 选择“新客户端密码”按钮。 输入“描述”的值，然后选择“到期”的适当选项，并选择“添加”。 在继续操作前，*立即复制客户端机密码值并使用应用程序 ID 保存它*，因为在后面的过程中，将需要用到它。

1. 在“管理”下选择“公开 API”。 选择" **设置"** 链接。 这将以"api：//$App ID GUID$"的形式生成应用程序 ID URI，其中 $App ID GUID$ 是应用程序 (客户端) **ID。**

1. 在生成的 ID 中， (注意双正斜杠和 GUID) 末尾附加的正斜杠 `localhost:44355/` "/"。 完成后，整个 ID 应格式为 `api://localhost:44355/$App ID GUID$` ;例如 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7` 。

1. 选择“**添加一个作用域**”按钮。 在打开的面板中，输入 `access_as_user` 作为 **作用域** 名称。

1. 将“谁能同意?”设置为“管理员和用户”。

1. 使用适用于范围的值填写用于配置管理员和用户同意提示的字段，使 Office 客户端应用程序能够使用与当前用户相同的权限使用外接程序的 Web API。 `access_as_user` 建议：

    - **管理员显示名称：Office** 可以充当用户。
    - **管理员许可描述**：使 Office 能够借助与当前用户相同的权限调用加载项的 Web API。
    - **用户同意显示名称：Office** 可以充当你。
    - **用户同意描述**：Office以您具有的相同权限调用外接程序的 Web API。

1. 确保将“**状态**”设置为“**已启用**”。

1. 选择“**添加作用域**”。

    > [!NOTE]
    > 显示在文本字段正下方的 **作用域** 名称的域部分应自动与你先前设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。

1. 在“授权客户端应用程序”部分中，确定要授权给加载项 Web 应用程序的应用程序。 下面每个 ID 都需要进行预授权。

    - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    - `57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）
    - `08e18876-6177-487e-b8b5-cf950c1e598c`（Office 网页版）
    - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）

    对于每个 ID，请执行以下步骤。

    a. 选择“**添加客户端应用程序**”按钮，然后在打开的面板中，将“客户端 ID”设置为相应的 GUID 并勾选 `api://localhost:44355/$App ID GUID$/access_as_user` 框。

    b. 选择“添加应用程序”。

1. 选择“管理”下的“API 权限”，然后选择“添加权限”。 在打开的面板上，选择 **Microsoft Graph**，然后选择“委派权限”。

1. 使用“选择权限”搜索框来搜索加载项需要的权限。 选择以下选项。 外接程序本身确实只需要第一项;但 `profile` 应用程序需要权限Office才能获取外接程序 Web 应用程序的令牌。

    * Files.Read.All
    * profile

    > [!NOTE]
    > `User.Read` 权限可能已默认列出。 根据最佳做法，最好不要请求授予不需要的权限，因此，如果加载项实际上不需要此权限，我们建议取消选中此权限对应的框。

1. 选择所显示的每个权限的复选框。 选择加载项需要的权限后，选择面板底部的“**添加权限**”按钮。

1. 在同一页面上，选择“**为[租户名称]授予管理员许可**”按钮，然后为显示的确认选择“**是**”。

## <a name="configure-the-add-in"></a>配置加载项

1. 在代码编辑器中打开克隆项目中的 `\Begin` 文件夹。

1. 打开 `.ENV` 文件，并使用先前复制的值。 将 **CLIENT_ID** 设置为 **应用程序（客户端）ID**，并将 **CLIENT_SECRET** 设置为客户端密码。 该值 **不** 能用引号引起来。 完成后，文件应当类似于以下示例：

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. 打开 `\public\javascripts\fallbackAuthDialog.js` 文件。 在 `msalConfig` 声明中，将占位符 $application_GUID here$ 替换为在注册加载项时复制的应用程序 ID。 该值应该用引号引起来。

1. 打开加载项清单文件“manifest\manifest_local.xml”，然后滚动到该文件的底部。 在结束 `</VersionOverrides>` 标记的正上方，你将找到以下标记。

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

1. 将标记中的 *两处* 占位符“$application_GUID here$”均替换为在注册加载项时复制的应用程序 ID。 由于 ID 并不包含“$”符号，因此请勿包含它们。 这是用于 中"用户"和"CLIENT_ID访问群体"的相同 ID。ENV 文件。

   > [!NOTE]
   > **资源** 值是注册加载项时设置的 **应用程序 ID URI**。 仅在通过 AppSource 销售加载项时，才使用 **作用域** 部分生成许可对话框。

## <a name="code-the-client-side"></a>编写客户端代码

### <a name="create-the-sso-logic"></a>创建 SSO 逻辑

1. 在代码编辑器中，打开文件 `public\javascripts\ssoAuthES6.js`。 它已经具有确保即使在 Internet Explorer 11 中也支持 Promise 的代码，并且具有 `Office.onReady` 调用，可将处理程序分配给加载项的唯一按钮。

   > [!NOTE]
   > 顾名思义，ssoAuthES6.js 使用 JavaScript ES6 语法，因为使用 `async` 和 `await` 可以最好地显示 SSO API 本质的简单性。 启动 localhost 服务器时，此文件将转换为 ES5 语法，以便在 Internet Explorer 11 中运行该示例。

1. 在 Office.onReady 方法下方添加以下代码。

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exchange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         Office.auth.getAccessToken call.

        }
    }
    ```

1. 将 `TODO 1` 替换为以下代码。 关于此代码，请注意以下几点：

    - `Office.auth.getAccessToken` 指示 Office 从 Azure AD 获取引导令牌。 引导令牌类似于 ID令 牌，但是它具有值为 `access-as-user` 的 `scp`（作用域）属性。 Web 应用程序可将此类令牌与 Microsoft Graph 的访问令牌进行交换。
    - 将选项设置为 true 意味着如果当前没有用户登录 `allowSignInPrompt` Office，Office将打开弹出窗口登录提示。
    - 将选项设置为 true 意味着如果用户未同意允许外接程序访问用户的 AAD 配置文件，Office将打开同意 `allowConsentPrompt` 提示。  (提示仅允许用户同意用户的 AAD 配置文件，而不是 Microsoft Graph范围。) 
    - 将 选项设置为 true 可Office指示加载项打算使用启动令牌获取 Microsoft Graph 的访问令牌，而不只是将其用作 `forMSGraphAccess` ID 令牌。 如果租户管理员未向加载项授予对 Microsoft Graph 的访问许可，则 `Office.auth.getAccessToken` 将返回错误 **13012**。 该加载项可通过回退到备用的授权系统来做出响应，这是必需的，因为 Office 可以提示仅同意访问用户的 Azure AD 配置文件，而不是任何 Microsoft Graph 作用域。 回退授权系统要求用户重新登录，并且可以提示用户同意 Microsoft  Graph作用域。 因此，`forMSGraphAccess` 选项可确保加载项不会进行令牌交换，交换会因缺乏许可而失败。 （由于先前步骤中已授予管理员许可，此加载项不会发生此情况。 但这里包含了一个选项来说明最佳实践。）

    ```javascript
    let bootstrapToken = await Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true }); 
    ```

1. 将 `TODO 2` 替换为下面的代码。 将在后续步骤中创建 `getGraphToken` 方法。

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. 将 `TODO 3` 替换为以下代码。 关于此代码，请注意以下几点： 

    - 如果Microsoft 365租户已配置为需要多重身份验证，则 将包括包含有关其他必需 `exchangeResponse` `claims` 因素的信息的属性。 在这种情况下，应该再次调用 `Office.auth.getAccessToken`，并将 `authChallenge` 选项设置为 claims 属性的值。 这就指示 AAD 提示用户进行所有必需形式的身份验证。

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await Office.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. 将 `TODO 4` 替换为以下代码。 关于此代码，请注意以下几点： 

    - 将在后续步骤中创建 `handleAADErrors` 方法。 Azure AD 错误作为 HTTP 代码 200 响应返回给客户端。 它们不会引发错误，因此不会触发 `getGraphData` 方法的 `catch` 块。
    - 将在后续步骤中创建 `makeGraphApiCall` 方法。 它将对 MS Graph 终结点进行 AJAX 调用。 在该调用的 `.fail` 回调中捕获到错误，而不是在 `getGraphData` 方法的 `catch` 块中。

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. 将 `TODO 5` 替换为以下内容：

    - 来自 `getAccessToken` 调用的错误将具有 `code` 属性，其错误号通常处于 13xxx 范围内。 将在后续步骤中创建 `handleClientSideErrors` 方法。
    - `showMessage` 方法在任务窗格上显示文本。

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. 在 `getGraphData` 方法下方，添加下列函数。 请注意，这是一个服务器端 Express 路由，用于将启动令牌与 Azure AD 交换为 `/auth` Microsoft Graph。

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. 在 `getGraphToken` 方法下方，添加下列函数。 请注意，`error.code` 是一个数字，通常处于 13xxx 范围内。

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```

1. 将 `TODO 6` 替换为下面的代码。 有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to sign in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // Office.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the web.
        showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The Office.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. 将 `TODO 7` 替换为下面的代码。 有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。函数 `dialogFallback` 用于调用备用授权系统。 在此加载项中，回退系统将打开一个对话框，它要求用户登录（即使用户已登录），并使用 msal.js 和隐式流来获取 Microsoft Graph 访问令牌。

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. 在 `handleClientSideErrors` 函数下方，添加下列函数。 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. 在极少数情况下，Office 缓存的引导令牌在 Office 验证时未过期，但是会在到达 Azure AD 进行交换时过期。 Azure AD 将以错误 **AADSTS500133** 做出响应。 在这种情况下，加载项应仅以递归方式调用 `getGraphData`。 由于缓存的引导令牌现在已过期，Office 将从 Azure AD 获取一个新令牌。 因此， `TODO 8` 将 替换为以下内容：


    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
    {
        getGraphData();
    }
    ```

1. 若要确保加载项不会进入 `getGraphData` 调用的无限循环，该加载项应跟踪调用 `getGraphData` 的次数，并确保不会多次对它进行递归式调用。 因此，应在 `handleAADErrors` 和 `getGraphData` 函数的全局范围内创建计数器变量。 全局变量的理想位置就在 `Office.onReady` 方法调用的正下方。

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. 在 `handleAADErrors` 方法中更改 `if` 结构，以使其：

    - 在调用 `getGraphData` 之前递增计数器。
    - 执行测试以确保尚未对 `getGraphData` 进行第二次调用。

    因此，`if` 结构的最终版本应如下所示：

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. 将 `TODO 9` 替换为以下内容：

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. 保存并关闭此文件。

### <a name="get-the-data-and-add-it-to-the-office-document"></a>获取数据并将其添加到 Office 文档

1. 在 `public\javascripts` 文件夹中，创建名为 `data.js` 的新文件。

1. 将以下函数添加到文件中。 这是 `getGraphData` 函数在获得 Microsoft Graph 访问令牌后调用的函数。 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. 将 `TODO 10` 替换为以下代码。 关于此代码，请注意以下几点：

    - 此对象是 `$.ajax` 方法的参数。
    - `/getuserdata` 是你在后续步骤中创建的加载项服务器上的 Express 路由。 它将调用 Microsoft Graph 终结点，并在其调用中包含访问令牌。 

    ```javascript
    {
        type: "GET",
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. 将 `TODO11` 替换为以下代码。 关于此代码，请注意以下几点：

    - `writeFileNamesToOfficeDocument` 会将来自 Graph 的数据插入到 Office 文档中。 它在 `public\javascripts\document.js` 文件中定义。
    - 如果 `writeFileNamesToOfficeDocument` 返回错误，它将以“无法将文件名添加到文档中”开头。

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () {
        showMessage("Your data has been added to the document.");
    })
    .catch(function (error) {
        showMessage(error);
    });
    ```

1. 保存并关闭此文件。

## <a name="code-the-server-side"></a>编写服务器端代码

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a>创建身份验证路由器和令牌交换逻辑

1. 打开文件 `routes\authRoute.js`，然后在 `require` 语句正下方和 `module.exports` 语句上方添加以下路由函数。 请注意，`router.get` 的 URL 参数是“/”。 由于此路由是在负责处理 URL“/auth”的所有 HTTP 请求的路由器中定义的，因此该路由可有效处理“/auth”的所有请求。 先前创建的客户端 `getGraphToken` 函数将调用此路由。  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exchange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. 将 `TODO 12` 替换为下面的代码。

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. 将 `TODO 13` 替换为下面的代码。 关于此代码，请注意以下几点：

    - 这是一个长 `else` 块的开头，但是结尾 `}` 尚未结束，因为你将向其添加更多代码。
    - `authorization` 字符串是“持有者”，后跟引导令牌，因此 `else` 块的第一行将令牌分配给 `jwt`。 （“JWT”代表“JSON Web 令牌”。）
    - 两个 `process.env.*` 值是你配置加载项时分配的常量。
    - `requested_token_use` 窗体参数设置为“on_behalf_of”。 它告知 Azure AD 加载项正在使用“代理流”请求 Microsoft Graph 访问令牌。 通过验证分配给 `assertion` 窗体参数的引导令牌是否具有设置为 `access-as-user` 的 `scp` 属性，Azure 将对此做出响应。
    - `scope` 窗体参数设置为“Files.Read.All”，这是加载项唯一需要的 Microsoft Graph 作用域。

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. 将 `TODO 14` 替换为以下代码，它将完成 `else` 块。 关于此代码，请注意以下几点：

    - 常量 `tenant` 设置为“通用”，因为你在 Azure AD 中注册加载项时已将其配置为多租户；特别是当你将“**支持的帐户类型**”设置为“**任何组织目录中的帐户和个人 Microsoft 帐户（例如，Skype、Xbox、Outlook.com）**”时。 如果改为选择仅支持注册加载项的同Microsoft 365租户中的帐户，则此代码将设置为租户 `tenant` 的 GUID。 
    - 如果 POST 请求没有错误，那么 Azure AD 的响应将转换为 JSON 并发送到客户端。 此 JSON 对象具有 `access_token` 属性，Azure AD 已为其分配 Microsoft Graph 访问令牌。

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: formurlencoded(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();

            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. 保存并关闭此文件。

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a>创建将从 Microsoft Graph 获取数据的路由

1. 打开项目根目录中的 `app.js` 文件。 在“/dialog.html”路由的正下方，添加以下路由。 此路由由你在前面步骤中创建的 `makeGraphApiCall` 函数调用。

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. 将 `TODO 15` 替换为以下代码。 关于此代码，请注意以下几点：

    - 此路由的调用方 `makeGraphApiCall` 将 Microsoft Graph 访问令牌作为名为“access_token”的标头添加到 HTTP 请求中。
    - `getGraphData` 函数在 `msgraph-helper.js` 文件中定义。 （此函数与在 `ssoAuthES6.js` 文件中定义的客户端 `getGraphData` 函数不同。）
    - `queryParamsSegment` 的最后一个参数是硬编码值。 如果你在生产加载项中重复使用此代码，并且 `queryParamsSegment` 的任何部分均来自用户输入，请确保它已被清理，以便它不能用于响应标头注入攻击。
    - 通过仅指定所需的属性（“名称”）以及仅前 10 个文件夹或文件名，该代码可最大限度地减少来自 Microsoft Graph 的数据量。

    ```javascript
    const graphToken = req.get('access_token');
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. 将 `TODO 16` 替换为以下代码。 关于此代码，请注意以下几点：

    - 如果 Microsoft Graph 返回错误（例如无效或过期的令牌），则返回的对象中将有一个 code 属性设置为 HTTP 状态（例如 401）。 代码会将错误转发给客户端。 它将在 `makeGraphApiCall` 的 `.fail` 回调中被捕获。
    - Microsoft Graph 数据包含该加载项不需要的 OData 元数据和 eTag，因此代码将构造一个新数组，其中仅包含要发送到客户端的文件名。

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. 保存并关闭此文件。

## <a name="run-the-project"></a>运行项目

1. 请确保 OneDrive 中有一些文件，以便可以验证结果。

1. 在 `\Begin` 文件夹的根目录中打开命令提示符。

1. 运行命令 `npm start`。

1. 需要将加载项旁加载到 Office 应用程序（Excel、Word 或 PowerPoint），以便对其进行测试。 说明取决于你的平台。 在[旁加载 Office 加载项以供测试](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)中有指向说明的链接。

1. 在 Office 应用程序的“**主页**”功能区上，选择“**SSO Node.js**”组中的“**显示加载项**”按钮以打开任务窗格加载项。

1. 单击“**获取 OneDrive 文件名**”按钮。 如果使用 Microsoft 365 教育版 或工作帐户或 Microsoft 帐户登录 Office 并且 SSO 正常工作，OneDrive for Business 中的前 10 个文件和文件夹名称将插入到文档中。  (首次登录可能需要 15 秒。) 如果您未登录，或者您位于不支持 SSO 的方案中，或者 SSO 因任何原因无法工作，系统将提示您登录。 登录后，将显示文件和文件夹名称。

> [!NOTE]
> 如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已更改过，也不例外。 在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。 为了防止发生这种情况，请务必先 *关闭其他所有 Office 应用程序*，然后再按“**获取 OneDrive 文件名**”。
