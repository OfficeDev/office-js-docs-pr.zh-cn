# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>创建使用单一登录的 Node.js Office 加载项

用户可以登录到 Office，并且你的 Office Web 外接程序可以利用此登录过程来授权用户访问你的外接程序和 Microsoft Graph，而无需要求用户再次登录。有关概述，请参阅[在 Office 外接程序中启用 SSO](../../docs/develop/sso-in-office-add-ins.md)。

本文将引导你完成在使用 Node.js 和 Express 构建的外接程序中启用单一登录 (SSO) 的过程。 

> **注意：**有关基于 ASP.NET 的外接程序的类似文章，请参阅[创建使用单一登录的 ASP.NET Office 外接程序](../../docs/develop/create-sso-office-add-ins-aspnet.md)。

## <a name="prerequisites"></a>先决条件

* [节点和 npm](https://nodejs.org/en/)，版本 6.9.4 或更高版本。
* [Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）。
* TypeScript 版本 2.2.2 或更高版本。
* Office 2016，版本 1708，内部版本 8424.nnnn 或更高版本（Office 365 订阅版本，有时称为“即点即用”）。可能需要成为 Office 预览体验成员才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/en-us/office-insider?tab=tab-1)。

## <a name="set-up-the-starter-project"></a>设置初学者项目

1. 在 [Office 外接程序 NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso) 中克隆或下载存储库。 


    > **注意：**该示例有两个版本。 
    > 
    > * **Before** 文件夹是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。本文后续章节将引导你完成此过程。 
    > * 如果完成了本文中的过程，该示例的**已完成**版本会与所生成的外接程序类似，只不过完成的项目具有对本文文本冗余的代码注释。若要使用已完成的版本，请按照本文中的说明进行操作即可，但需要将“Before”替换为“Completed”，并跳过**编写客户端代码**和**编写服务器端代码**部分。

1. 在 **Before** 文件夹中打开 Git bash 控制台。

2. 在该控制台中输入 `npm install` 以安装 package.json 文件中列出明细的所有依赖项。

3. 在控制台中输入 `npm run build ` 以生成该项目。 
     > 注意：你可能会看到一些生成错误，指出声明了某些变量但未使用。 请忽略这些错误。 这是“旧”版样本缺少稍后将添加的一些代码带来的负面影响。

## <a name="register-the-add-in-with-azure-ad-v2-endpoint"></a>向 Azure AD V2 终结点注册加载项

1. 转到 [https://apps.dev.microsoft.com](https://apps.dev.microsoft.com)。 

1. 使用管理员凭据登录 Office 365 租户。例如，MyName@contoso.onmicrosoft.com

1. 单击“添加应用”****。

1. 出现提示时，使用“Office-Add-in-NodeJS-SSO”作为应用名称，然后按“创建应用程序”****。

1. 当应用的配置页面打开时，复制**应用程序 ID** 并保存。你将在后续过程中使用它。 

    > 注意：当其他应用程序（例如 PowerPoint、Word、Excel 等 Office 主机应用程序）寻求对应用程序的授权访问权限时，此 ID 是“受众”值。当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

1. 在“应用程序机密”****部分中，按“生成新密码”****。将打开弹出对话框，并显示一个新密码（也称为“应用机密”）。*立即复制密码并使用应用程序 ID 进行保存。*你将在后续过程中需要它。然后关闭此对话框。

1. 在“平台”****部分中，单击“添加平台”****。 

1. 在打开的对话框中，选择“Web API”****。

1. “应用程序 ID URI”****已生成，格式为“api://{App ID GUID}”。在双正斜线和 GUID 之间插入字符串“localhost:3000”。完整 ID 应为 `api://localhost:3000/{App ID GUID}`。（位于“应用程序 ID URI”****下方的“作用域”****名称的域部分将自动更改以匹配。应显示 `api://localhost:3000/{App ID GUID}/access_as_user`。）

1. 这一步和下一步将授予 Office 主机应用程序对加载项的访问权限。 在“预授权应用程序”****部分中，确定要授权给加载项 Web 应用程序的应用程序。 下面每个 ID 都需要进行预授权。 每次输入一个 ID，都会看到新的空文本框。 （仅输入 GUID）。

 * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
 * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
 * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online) 

1. 打开每个“应用程序 ID”****旁边的“作用域”****下拉列表，并选中 `api://localhost:44355/{App ID GUID}/access_as_user` 对应的框。

1. 在“平台”****部分顶部附近，再次单击“添加平台”****并选择“Web”****。

1. 在“平台”****下的新“Web”****部分中，输入以下内容作为“重定向 URL”****：`https://localhost:3000`。 

    > 注意：在撰写本文时，“Web API”****平台有时会从“平台”****部分消失，特别是在添加“Web”****平台以及“保存注册页面”**后，如果刷新页面就会出现上述情况。为了确保“Web API”****平台仍然是注册的一部分，请单击页面底部附近的“编辑应用程序清单”****按钮。应该会看到清单的 **identifierUris** 属性中的 `api://localhost:3000/{App ID GUID}` 字符串。还有一个 **oauth2Permissions** 属性，它的 **value** 子属性的值为 `access_as_user`。

1. 向下滚动到“Microsoft Graph 权限”****部分，“委派的权限”****小节。使用“添加”****按钮打开“选择权限”****对话框。

1. 在对话框中，选中以下权限对应的框： 
    * Files.Read.All
    * 配置文件

1. 单击对话框底部的“确定”****。

1. 单击注册页面底部的“保存”****。

## <a name="grant-admin-consent-to-the-add-in"></a>向加载项授予管理员许可

> **注意：**仅在开发加载项时，才需要执行此过程。 将生产加载项部署到 Office 应用商店或加载项目录时，用户会在安装时独自信任它。

1. 在以下字符串中，将占位符“{application_ID}”替换为注册加载项时复制的应用程序 ID。

    `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. 将生成的 URL 粘贴到浏览器地址栏并导航到该 URL。

1. 出现提示时，使用管理员凭据登录 Office 365 租户。

1. 然后系统提示你授予外接程序访问 Microsoft Graph 数据的权限。单击“接受”****。 

1. 然后，浏览器窗口/选项卡会重定向到注册加载项时指定的**重定向 URL**；因此，如果加载项正在运行，那么浏览器中会打开加载项的主页。 如果加载项未在运行，将会看到错误消息，指明找不到或打不开 localhost:3000 处的资源。 *不过，尝试进行重定向便意味着管理员许可过程成功完成*。 所以，无论是打开了主页，还是看到错误，都可以继续执行下一步。

2. 在浏览器的地址栏中，你将看到一个带有 GUID 值的“租户”查询参数。这是 Office 365 租户的 ID。复制并保存此值。你将在后续步骤中使用它。

3. 关闭该窗口/选项卡。

## <a name="configure-the-add-in"></a>配置外接程序

1. 在代码编辑器中打开 src\server.ts 文件。顶部附近存在对 `AuthModule` 类的构造函数的调用。该构造函数中存在一些需要为其分配值的字符串参数。

2. 对于 `client_id` 属性，将占位符 `{client GUID}` 替换为注册外接程序时保存的应用程序 ID。完成后，应该有一个括在单引号中的 GUID。而不应存在任何“{}”字符。

3. 对于 `client_secret` 属性，将占位符 `{client secret}` 替换为注册外接程序时保存的应用程序机密。

4. 对于 `audience` 属性，将占位符 `{audience GUID}` 替换为注册外接程序时保存的应用程序 ID。（即分配给 `client_id` 属性的同一值）。
  
3. 在分配给 `issuer` 属性的字符串中，你将看到占位符 *{O365 tenant GUID}*。将此替换为在最后一个过程结束时保存的 Office 365 租户 ID。如果出于任何原因，你以前没有获得 ID，请使用[查找 Office 365 租户 ID](https://support.office.com/en-us/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b)中的一种方法来获取 ID。完成后，`issuer` 属性值应如下所示：

    `https://login.microsoftonline.com/12345678-1234-1234-1234-123456789012/v2.0`

1. 保持 `AuthModule` 构造函数中的其他参数不变。 保存并关闭文件。

1. 在项目的根目录中，打开外接程序清单文件“Office-Add-in-NodeJS-SSO.xml”。

1. 滚动到文件底部。

1. 在结束 `</VersionOverrides>` 标记的正上方，你会发现以下标记：

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:3000/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. 将标记中的*两处*占位符“{application_GUID here}”均替换成在注册加载项时复制的应用程序 ID。 （由于 ID 并不包含“{}”，因此请勿添加它们。）这与在 web.config 中对 ClientID 和 Audience 使用的 ID 相同。

    >注意： 
    >
    >* **Resource** 值是将 Web API 平台添加到外接程序注册时设置的“应用程序 ID URI”****。
    >* 如果通过 Office 应用商店销售该外接程序，则 **Scopes** 部分仅用于生成同意对话框。

1. 保存并关闭文件。

## <a name="code-the-client-side"></a>编写客户端代码

1. 打开 **public** 文件夹中的 program.js 文件。其中已存在一些代码：

    * 针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。
    * 在任务窗格底部，显示从 Microsoft Graph（或错误消息）返回的数据的 `showResult` 方法。

1. 在针对 `Office.initialize` 的分配下面，添加下面的代码。关于此代码，请注意以下几点： 

    * 首次尝试使用“代表”流时，便会调用 `getDataWithoutAuthChallenge` 函数。 假设只需要单一身份验证。 将在稍后的步骤中添加代码，以处理需要多重身份验证的情况。
    * `getAccessTokenAsync` 是 Office.js 中新增的 API，支持加载项向 Office 主机应用程序（Excel、PowerPoint、Word 等）请求获取对加载项的访问令牌（对于已登录 Office 的用户）。 反过来，Office 主机应用程序会向 Azure AD 2.0 终结点请求获取令牌。 由于已在注册加载项时将 Office 主机预授权给加载项，因此 Azure AD 将会发送令牌。 
     * 如果没有用户登录 Office，则 Office 主机将提示用户登录。 
     * options 参数将 `forceConsent` 设置为 false，因此将不会提示用户同意为 Office 主机提供访问外接程序的权限。

    ```js
    function getOneDriveItems() {
        getDataWithoutAuthChallenge();
    }   
    
    function getDataWithoutAuthChallenge() {       
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    // TODO1: Use the access token to get Microsoft Graph data.
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

1. 将 TODO1 替换为以下代码行。 可以在稍后的步骤中创建 `getData` 方法和服务器端“/api/onedriveitems”路由。 相对 URL 适用于终结点，因为它必须与加载项托管在同一域中。

    ```
    accessToken = result.value;
    getData("/api/onedriveitems", accessToken);
    ```

1. 在 `getOneDriveFiles` 方法下面添加以下内容。此实用程序方法调用指定的 Web API 终结点，并向其传递与 Office 主机应用程序用于获取外接程序访问权限的令牌相同的访问令牌。在服务器端，此访问令牌将用于“代表”流，以获取 Microsoft Graph 的访问令牌。 

    ```
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET",
        })
        .done(function (result) {
            TODO2: Display data and handle demand for multi-factor authentication.
        })
        .fail(function (result) {
            console.log(result.error);
       });
    }
    ```

1. 用以下代码替换 TODO2。关于此代码，请注意以下几点：
    * 如果 Microsoft Graph 目标要求有其他身份验证因素，就不会生成数据。 而是生成声明 JSON，指示 AAD 必须提示用户提供什么其他身份验证因素。 在这种情况下，客户端必须启动新登录，将此 Claims 字符串传递到 AAD，这样 AAD 才能提供所需的提示。
    * 如果生成声明 JSON，将包含字符串“capolids”。
    * 将在稍后的步骤中创建 `getDataUsingAuthChallenge` 函数。

    ```
    if (result[0].indexOf('capolids') !== -1) {                
        result[0] = JSON.parse(result[0])
        getDataUsingAuthChallenge(result[0]);
    } else {  
        showResult(result);
    }
    ```

1. 在文件中 `getData` 函数的正下方，添加以下函数。 关于此函数，请注意以下几点：
    * 在 AAD 要求提供其他身份验证因素时，使用此函数。 
    * 此函数会触发第二次登录，以提示用户提供其他身份验证因素。 
    * `authChallenge` 选项包含的字符串可指示 AAD 应提示提供什么身份验证因素。 请求获取对加载项的加载项令牌时，Office 主机会将此字符串传递给 AAD。

    ```
    function getDataUsingAuthChallenge(authChallengeString) {       
        Office.context.auth.getAccessTokenAsync({authChallenge: authChallengeString},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/onedriveitems", accessToken);
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

1. 保存并关闭文件。

## <a name="code-the-server-side"></a>编写服务器端代码

有两个需要修改的服务器端文件。 
- src\auth.js 提供授权 helper 函数。它已具有在各种授权流中使用的泛型成员。我们需要为其添加可实现“代表”流的函数。
- src\server.js文件具有运行服务器和 Express 中间件所需的基本成员。我们需要为其添加服务于主页和 Web API 的函数，以获取 Microsoft Graph 数据。

### <a name="create-a-method-to-exchange-tokens"></a>创建交换令牌的方法

1. 打开 \src\auth.ts 文件。将下面的方法添加到 `AuthModule` 类。关于此代码，请注意以下几点：
    * jwt 参数是该应用程序的访问令牌。在“代表”流中，它与 AAD 进行交换，以获取资源的访问令牌。
    * scopes 参数具有默认值，但在此示例中，它将被调用代码覆盖。
    * resource 参数是可选的。它不应在 STS 是 AAD V2 终结点时使用。后者从作用域推断资源，如果在 HTTP 请求中发送资源，则它将返回错误。 
    

    ```
    private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        try {
            // TODO3: Construct the parameters that will be sent in the body of the 
            //        HTTP Request to the STS that starts the "on behalf of" flow.
            // TODO4: Send the request to the STS.
            // TODO5: Process the response and persist the access token to resource.
        }
        catch (exception) {
            throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                        + JSON.stringify(exception), 
                                        exception);
        }
    }
    ```

2. 将 TODO3 替换为以下代码行。 关于此代码，请注意以下几点：
    * 支持“代表”流的 STS 需要 HTTP 请求正文中的某些属性/值对。此代码构造一个可成为请求正文的对象。 
    * 仅当资源传递给方法时，才会将资源属性添加到正文。

    ```
    const v2Params = {
            client_id: this.clientId,
            client_secret: this.clientSecret,
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            assertion: jwt,
            requested_token_use: 'on_behalf_of',
            scope: scopes.join(' ')
        };
        let finalParams = {};
        if (resource) {
            // In JavaScript we could just add the resource property to the v2Params
            // object, but that won't compile in TypeScript.
            let v1Params  = { resource: resource };  
            for(var key in v2Params) { v1Params[key] = v2Params[key]; }
            finalParams = v1Params;
        } else {
            finalParams = v2Params;
        } 
    ```

3. 将 TODO4 替换为以下代码行，可将 HTTP 请求发送到 STS 的令牌终结点。

    ```
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. 将 TODO5 替换为以下代码行。 请注意，此代码除了返回对资源的访问令牌之外，还将保留它及其到期时间。 调用代码通过重用资源未到期的访问令牌来避免对 STS 的不必要的调用。 下一部分将介绍如何执行此操作。

    ```
    if (res.status !== 200) {
        TODO6: Handle failure and the case where AAD asks for additional
               authentication factors.
    }
    const json = await res.json();
    // Persist the token and it's expiration time.
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

5. 将 TODO6 替换为以下代码行。 关于此代码，请注意以下几点：

    * 一些 Azure Active Directory 配置要求用户，必须提供其他身份验证因素，才能访问一些 Microsoft Graph 目标（例如 OneDrive），即使用户仅使用密码就能登录 Office。 在这种情况下，AAD 将发送包含 `Claims` 属性的响应。 
    * 此 `Claims` 值需要传递回客户端，它应为用户启动第二次登录，并在 AAD 调用中添加 `Claims` 值。 AAD 将提示用户提供其他身份验证因素。
    * 作为预防措施，此代码清除在用户仅使用密码登录时所获取的任何访问令牌缓存。  

    ```
    const exception = await res.json();
    // Check if AAD is the STS.
    if (this.stsDomain === 'https://login.microsoftonline.com') {
        if (JSON.stringify(exception.claims)) {                       
            ServerStorage.clear();
            return JSON.stringify(exception.claims);    
        } else {                    
            throw exception;
        }
    }
    else {                    
        throw exception;
    }
    ```

5. 保存但不关闭文件。

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>使用“代表”流创建一个获取资源访问权限的方法

1. 还是在 src/auth.ts 中，将下面的方法添加到 `AuthModule` 类。关于此代码，请注意以下几点：
    * 上面关于 `exchangeForToken` 方法参数的注释也适用于此方法的参数。
    * 此方法先检查永久性存储中是否有尚未过期且下一分钟也不会过期的资源访问令牌。 仅在需要时，它才会调用在上一部分中创建的 `exchangeForToken` 方法。

    ```
    async acquireTokenOnBehalfOf(jwt: string, scopes: string[] = ['openid'], resource?: string) {
        const resourceTokenExpirationTime = ServerStorage.retrieve('ResourceTokenExpiresAt');
        if (moment().add(1, 'minute').diff(resourceTokenExpirationTime) < 1 ) {
            return ServerStorage.retrieve('ResourceToken');
        } else if (resource) {
            return this.exchangeForToken(jwt, scopes, resource);
        } else {
            return this.exchangeForToken(jwt, scopes);
        }
    } 
    ```

2. 保存并关闭文件。

### <a name="create-the-endpoints-that-will-serve-the-add-ins-home-page-and-data"></a>创建服务于外接程序主页和数据的终结点

1. 打开 src\server.ts 文件。 

2. 将以下方法添加到文件底部。此方法将为外接程序的主页提供服务。外接程序清单指定主页 URL。

    ```
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. 将以下方法添加到文件底部。此方法将处理 `onedriveitems` API 的任何请求。
    ```
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Send to the client only the data that it actually needs.
    })); 
    ```

4. 将 TODO7 替换为以下代码行，可验证从 Office 主机应用程序收到的访问令牌。 `verifyJWT` 方法在 src\auth.ts 文件中进行定义。 它始终验证受众和颁发者。 此可选参数可用于指定是否还要它验证访问令牌中的作用域是否为 `access_as_user`。 这是用户和 Office 主机通过“代表”流获取对 Microsoft Graph 的访问令牌时，唯一需要拥有的对加载项的权限。 

    ```
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

> **注意：**只可使用 `access_as_user` 作用域授权 API 为 Office 加载项处理代表流。服务中的其他 API 应有自己的作用域要求。 这就限制了使用 Office 获得的令牌可以访问的内容。

5. 将 TODO8 替换为以下代码。 关于此代码，请注意以下几点：

    * `acquireTokenOnBehalfOf` 调用中不包括 resource 参数，因为 `AuthModule` 对象 (`auth`) 是使用不支持 resource 属性的 AAD V2.0 终结点进行构造。
    * 调用的第二个参数指定了加载项获取 OneDrive 上用户文件和文件夹列表时所需的权限。 （之所以不需要 `profile` 权限是因为，只有当 Office 主机获取对加载项的访问令牌时，才需要此权限，用此令牌交换对 Microsoft Graph 的访问令牌时并不需要。）
    * 如果响应是包含“capolids”的字符串，表明这是来自 AAD 的声明消息，要求进行多重身份验证。 此消息会被传递给客户端，再用它来启动第二次登录。 此字符串指示 AAD，应提示用户提供什么其他身份验证因素。

    ```
    let graphToken = null;
    const tokenAcquisitionResponse = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    if (tokenAcquisitionResponse.includes('capolids')) {
        const claims: string[] = [];
        claims.push(tokenAcquisitionResponse);
        return res.json(claims);
    } else {
        // The response is the token to Microsoft Graph itself. Rename it so remaining code
        // is self-documenting.
        graphToken = tokenAcquisitionResponse;
    }
    ```

6. 将 TODO9 替换为以下代码行。 关于此代码，请注意以下几点：

    * MSGraphHelper 类在 Src\msgraph helper.ts 中定义。 
    * 通过指定只需要 name 属性和前 3 项，可以最大限度地减少必须返回的数据。

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. 将 TODO10 替换为以下代码行。 请注意，Microsoft Graph 会为每一项都返回一些 OData 元数据和 **eTag** 属性，即使 `name` 是所请求的唯一属性，也是如此。 该代码仅向客户端发送项目名称。

    ```
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. 保存并关闭文件。

## <a name="deploy-the-add-in"></a>部署外接程序

现在，你需要让 Office 知道在哪里可以找到该外接程序。

1. 创建网络共享，或[将文件夹共享到网络](https://technet.microsoft.com/en-us/library/cc770880.aspx)。

2. 将 Office-Add-in-NodeJS-SSO.xml 清单文件从项目根目录复制到共享文件夹。

3. 启动 PowerPoint 并打开文档。

4. 选择“文件”****选项卡，然后选择“选项”****。

5. 选择**信任中心**，然后选择**信任中心设置**按钮。

6. 选择“受信任的外接程序目录”****。

7. 在“目录 URL”****字段中，输入包含 Office-Add-in-NodeJS-SSO.xml 的文件夹共享的网络路径，然后选择“添加目录”****。

8. 选中“显示在菜单中”****复选框，然后选择“确定”****。

9. 随后会出现一条消息，告知你下次启动 Microsoft Office 时将应用你的设置。关闭 PowerPoint。

## <a name="build-and-run-the-project"></a>生成和运行项目

根据是否使用 Visual Studio Code，有两种生成和运行项目的方法。对于这两种方法，当更改代码时，该项目将生成和自动生成并重新运行。

1. 如果使用的不是 Visual Studio Code： 
 1. 打开节点终端，然后导航到该项目的根文件夹。
 2. 在终端中，输入 **npm run build**。 
 3. 打开第二个节点终端，然后导航到该项目的根文件夹。
 4. 在终端中，输入 **npm run start**。

2. 如果使用的是 VS Code：
 1. 通过 VS Code 打开项目。
 2. 按 CTRL-SHIFT-B 生成项目。
 3. 按 F5 键在调试会话中运行该项目。


## <a name="add-the-add-in-to-an-office-document"></a>将外接程序添加到 Office 文档

1. 重启 PowerPoint 并打开或创建演示文稿。 

2. 在 PowerPoint 中的“开发工具”****选项卡上，选择“我的外接程序”****。

3. 选择“共享文件夹”****选项卡。

4. 选择“SSO NodeJS 示例”****，然后选择“确定”****。

5. “主页”****功能区上有一个名为“**SSO NodeJS**”的新组，包含标记为“显示外接程序”****的按钮和一个图标。 

## <a name="test-the-add-in"></a>测试加载项

1. 请确保 OneDrive 中有一些文件，以便可以验证结果。

2. 单击“显示加载项”****按钮，打开此加载项。

2. 此加载项打开，并显示欢迎页。 单击“从 OneDrive 获取我的文件”****按钮。

2. 如果你已登录 Office，则 OneDrive 上的文件和文件夹列表将显示在该按钮的下方。首次操作需要的时间可能会超过 15 秒。

3. 如过没有登录 Office，弹出窗口将打开并提示进行登录。完成登录后，文件和文件夹的列表将在几秒钟后显示。*请勿再次按下此按钮。*
> **注意：**如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用程序现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已在 PowerPoint 中更改过。 在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。 为了防止发生这种情况，请务必先*关闭其他所有 Office 应用程序*，然后再按“从 OneDrive 获取我的文件”****。
