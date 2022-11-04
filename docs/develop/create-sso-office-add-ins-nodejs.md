---
title: 创建使用单一登录的 Node.js Office 加载项
description: 了解如何创建使用 Office 单一登录的基于 Node.js 的加载项。
ms.date: 10/06/2022
ms.localizationpriority: medium
ms.openlocfilehash: 35128da43b3f27a58df5e188a5001bfa8aba4a4c
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/28/2022
ms.locfileid: "68841691"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>创建使用单一登录的 Node.js Office 加载项

Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).

本文将指导你完成在加载项中启用单一登录 (SSO) 的过程。 创建的示例外接程序包含两个部分：在 Microsoft Excel 中加载的任务窗格，以及处理任务窗格对 Microsoft Graph 的调用的中间层服务器。 中间层服务器使用 Node.js 和 Express 构建， `/getuserfilenames`并公开单个 REST API ，该 API 返回用户的 OneDrive 文件夹中的前 10 个文件名的列表。 任务窗格使用 `getAccessToken()` 方法获取已登录用户到中间层服务器的访问令牌。 中间层服务器使用代表流 (OBO) 将访问令牌交换为有权访问 Microsoft Graph 的新令牌。 可以扩展此模式以访问任何 Microsoft Graph 数据。 任务窗格始终调用中间层 REST API， (需要 Microsoft Graph 服务时) 传递访问令牌。 中间层使用通过 OBO 获取的令牌调用 Microsoft Graph 服务并将结果返回到任务窗格。

本文使用使用 Node.js 和 Express 的加载项。 有关与此类似的 ASP.NET 加载项文章，请参阅[创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）

- [Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）

- 代码编辑器 - 建议Visual Studio Code

- Microsoft 365 订阅中至少存储在OneDrive for Business上的一些文件和文件夹

- 支持 [IdentityAPI 1.3 要求集](/javascript/api/requirement-sets/common/identity-api-requirement-sets) 的 Microsoft 365 内部版本。 可以获取[免费的开发人员沙盒](https://developer.microsoft.com/microsoft-365/dev-program#Subscription)，该沙盒提供可续订的 90 天Microsoft 365 E5开发人员订阅。 开发人员沙盒包含 Microsoft Azure 订阅，可在本文后面的步骤中使用该订阅进行应用注册。 如果需要，可以使用单独的 Microsoft Azure 订阅进行应用注册。 在 [Microsoft Azure](https://account.windowsazure.com/SignUp) 获取试用版订阅。

## <a name="set-up-the-starter-project"></a>设置初学者项目

1. 克隆或下载 [Office 外接程序 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) 中的存储库。

   > [!NOTE]
   > 示例项目有两个版本：
   >
   > - **Begin** 文件夹是初学者项目。 未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。 本文后续章节将引导你完成此过程。
   > - **Complete** 文件夹包含相同的示例，已完成本文中的所有编码步骤。 若要使用已完成的版本，只需按照本文中的说明进行操作，但将“Begin”替换为“Complete”，并跳过编写 **客户端代码** 和 **编写中间层服务器端** 代码部分。

1. 在 **Begin** 文件夹中打开命令提示符。

1. 在该控制台中输入 `npm install` 以安装 package.json 文件中列出明细的所有依赖项。

1. 运行命令 `npm run install-dev-certs`。 为安装证书的提示选择“**是**”。

将以下值用于后续应用注册步骤的占位符。

| 占位符           | 值                                 |
|-----------------------|---------------------------------------|
| `<add-in-name>`       | **Office-Add-in-NodeJS-SSO**          |
| `<redirect-platform>` | **单页应用程序 (SPA)**     |
| `<redirect-uri>`      | `https://localhost:44355/dialog.html` |

[!INCLUDE [register-sso-add-in-aad-v2-include](../includes/register-sso-add-in-aad-v2-include.md)]

## <a name="configure-the-add-in"></a>配置加载项

1. 在代码编辑器中打开克隆项目中的 `\Begin` 文件夹。

1. 打开 文件， `.ENV` 并使用之前从 **Office-Add-in-NodeJS-SSO** 应用注册复制的值。 按如下所示设置值：

   | 名称              | 值                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | 应用程序注册概述页中的 **应用程序 (客户端) ID**。 |
   | **CLIENT_SECRET** | 从 **“证书&机密**”页保存的 **客户端** 密码。       |
   | **DIRECTORY_ID**  | 应用注册概述页中的 **目录 (租户) ID**。   |

   该值 **不** 能用引号引起来。 完成后，文件应当类似于以下示例：

   ```javascript
   CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
   CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
   DIRECTORY_ID=478aa78e-20ba-4c0d-9ffe-c4f62e5de3d5
   NODE_ENV=development
   SERVER_SOURCE=https://localhost:44355

1. Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file. Just above the `</VersionOverrides>` end tag, you'll find the following markup.

   ```xml
   <WebApplicationInfo>
     <Id>$app-id-guid$</Id>
     <Resource>api://localhost:44355/$app-id-guid$</Resource>
     <Scopes>
         <Scope>Files.Read</Scope>
         <Scope>profile</Scope>
         <Scope>openid</Scope>
     </Scopes>
   </WebApplicationInfo>
   ```

1. 将标记 _中两个位置的_ 占位符“$app-id-guid$”替换为创建 **Office-Add-in-NodeJS-SSO** 应用注册时复制 **的应用程序 ID**。 “$”符号不是 ID 的一部分，因此请勿包含它们。 此 ID 与在 中用于CLIENT_ID的 ID 相同。ENV 文件。

   > [!NOTE]
   > 该值 **\<Resource\>** 是注册加载项时设置 **的应用程序 ID URI** 。 如果加载项是通过 AppSource 销售的，则 **\<Scopes\>** 节仅用于生成同意对话框。

1. 打开 `\public\javascripts\fallback-msal\authConfig.js`文件。 将占位符“$app-id-guid$”替换为之前创建的 **Office-Add-in-NodeJS-SSO** 应用注册中保存的应用程序 ID。

1. 保存对文件所做的更改。

## <a name="code-the-client-side"></a>编写客户端代码

### <a name="create-client-request-and-response-handler"></a>创建客户端请求和响应处理程序

1. 在代码编辑器中，打开文件 `public\javascripts\ssoAuthES6.js`。 它已经具有确保即使在 Internet Explorer 11 中也支持 Promise 的代码，并且具有 `Office.onReady` 调用，可将处理程序分配给加载项的唯一按钮。

   > [!NOTE]
   > 顾名思义，ssoAuthES6.js 使用 JavaScript ES6 语法，因为使用 `async` 和 `await` 可以最好地显示 SSO API 本质的简单性。 当 localhost 服务器启动时，此文件将转译为 ES5 语法，以便该示例支持 Internet Explorer 11。

    示例代码的一个关键部分是客户端请求。 客户端请求是一个 对象，用于跟踪有关在中间层服务器上调用 REST API 的请求的信息。 这是必要的，因为需要通过以下方案跟踪或更新客户端请求状态：

    - SSO 失败，需要回退身份验证。 访问令牌是通过弹出对话框中的 MSAL 获取的。 目标是在这种情况下不会失败，并正常回退到备用身份验证方法。

    客户端请求对象跟踪以下数据：

    - `authSSO` - 如果使用 SSO，则为 true，否则为 false。
    - `verb` - REST API 谓词，例如 GET 和 POST。
    - `accessToken`- ASP.NET Core服务器的访问令牌。
    - `url`- 在 ASP.NET Core服务器上调用的 REST API 的 URL。
    - `callbackRESTApiHandler` - 用于传递 REST API 调用结果的函数。
    - `callbackFunction` - 在准备就绪时将客户端请求传递到 的函数。

1. 若要初始化客户端请求对象，请在 `createRequest` 函数中将 替换为 `TODO 1` 以下代码。

    ```javascript
    const clientRequest = {
      authSSO: authSSO,
      verb: verb,
      accessToken: null,
      url: url,
      callbackRESTApiHandler: restApiCallback,
        callbackFunction: callbackFunction,
    };
    ```

1. 将 `TODO 2` 替换为下面的代码。 关于此代码，请注意以下几点：

    - 它会检查是否正在使用 SSO。 对于 SSO，获取访问令牌的方法不同于回退身份验证。
    - 如果 SSO 返回访问令牌，它将调用 `callbackfunction` 函数。 对于回退身份验证，它会调用 `dialogFallback`，最终将在用户通过 MSAL 登录后调用回调函数。

    ```javascript
    // Get access token.

    if (authSSO) {
    try {
      // Get access token from Office SSO.
      clientRequest.accessToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      callbackFunction(clientRequest);
    } catch (error) {
      // handle the SSO error which will inform us if we need to switch to fallback auth.
      let fallbackRequired = handleSSOErrors(error);
      if (fallbackRequired) switchToFallbackAuth(clientRequest);
    }
   } else {
     // Use fallback auth to get access token.
     dialogFallback(clientRequest);
   }
    ```

1. 在 `getFileNameList` 函数中，将 `TODO 3` 替换为下列代码。 关于此代码，请注意以下几点：

    - 当用户在任务窗格上选择“**获取 OneDrive 文件名”** 按钮时，将调用 函数`getFileNameList`。
    - 它会创建一个客户端请求来跟踪有关调用的信息，例如 REST API 的 URL。
    - 当 REST API 返回结果时，它会传递给 `handleGetFileNameResponse` 函数。 此回调作为参数传递给 `createRequest` ，并在 中 `clientRequest.callbackRESTApiHandler`跟踪。
    - 代码使用客户端请求调用 `callWebServer` 以执行后续步骤并调用 REST API。

    ```javascript
    createRequest(
      "GET",
      "/getuserfilenames",
      handleGetFileNameResponse,
      async (clientRequest) => {
        await callWebServer(clientRequest);
      }
    );
    ```

1. 在 `handleGetFileNameResponse` 函数中，将 `TODO 4` 替换为下列代码。 关于此代码，请注意以下几点：

    - 代码传递响应 (，其中包含) 文件名 `writeFileNamesToOfficeDocument` 写入文档的文件名列表。
    - 代码检查错误。 如果写入文件名，则显示成功消息，否则会显示错误。

    ```javascript
    if (response !== null) {
      try {
        await writeFileNamesToOfficeDocument(response);
        showMessage("Your OneDrive filenames are added to the document.");
      } catch (error) {
        // The error from writeFileNamesToOfficeDocument will begin
        // "Unable to add filenames to document."
        showMessage(error);
      }
    } else
    showMessage("A null response was returned to handleGetFileNameResponse.");
    ```

1. 在 `handleSSOErrors` 函数中，将 `TODO 5` 替换为下列代码。 有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。

    ```javascript
    let fallbackRequired = false;

   switch (err.code) {
     case 13001:
       // No one is signed into Office. If the add-in cannot be effectively used when no one
       // is logged into Office, then the first call of getAccessToken should pass the
       // `allowSignInPrompt: true` option. Since this sample does that, you should not see
       // this error.
       showMessage(
         "No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again."
       );
       break;
     case 13002:
       // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
       // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
       showMessage(
         "You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."
       );
       break;
     case 13006:
       // Only seen in Office on the web.
       showMessage(
         "Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."
       );
       break;
     case 13008:
       // Only seen in Office on the web.
       showMessage(
        "Office is still working on the last operation. When it completes, try this operation again."
       );
       break;
     case 13010:
       // Only seen in Office on the web.
       showMessage(
         "Follow the instructions to change your browser's zone configuration."
       );
       break;
    ```

1. 将 `TODO 6` 替换为下面的代码。 有关这些错误的详细信息，请参阅 [Office 外接程序中的 SSO 疑难解答](troubleshoot-sso-in-office-add-ins.md)。对于无法处理的任何错误， `true` 将返回到调用方。 这表示调用方应切换到使用 MSAL 作为回退身份验证。

    ```javascript
     default:
      // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
      // to non-SSO sign-in.
      fallbackRequired = true;
      break;
    }
    return fallbackRequired;
    ```

### <a name="call-the-rest-api-on-the-middle-tier-server"></a>在中间层服务器上调用 REST API

1. 在 `callWebServer` 函数中，将 `TODO 7` 替换为下列代码。 关于此代码，请注意以下几点：

    - 实际的 AJAX 调用将由 函数进行 `ajaxCallToRESTApi` 。
    - 如果中间层服务器返回指示当前令牌已过期的错误，则此函数将尝试获取新的访问令牌。
    - 如果无法成功完成 AJAX 调用， `switchToFallbackAuth` 将调用 以使用 MSAL 身份验证而不是 Office SSO。

    ```javascript
    try {
    const data = await $.ajax({
      type: clientRequest.verb,
      url: clientRequest.url,
      headers: { Authorization: "Bearer " + clientRequest.accessToken },
      cache: false,
    });
    clientRequest.callbackRESTApiHandler(data);

    } catch (error) {
     // TODO 8: Check for expired SSO token and refresh if needed.

    // TODO 9: Check for Microsoft Graph and other errors.

    }
    ```

1. 将 `TODO 8` 替换为下面的代码。 关于此代码，请注意以下几点：

    - 当服务器标识过期的令牌时，它将返回类型为“TokenExpiredError”的错误。
    - 尝试...catch 将调用 Office.auth.getAccessToken 以获取具有新过期的刷新令牌。
    - 代码将尝试再次调用服务器 API。

    ```javascript
    // Check for expired SSO token. Refresh and retry the call if it expired.
    if (
      error.responseJSON &&
      authSSO === true &&
      error.responseJSON.type === "TokenExpiredError"
    ) {
      try {
        const accessToken = await Office.auth.getAccessToken({
          allowSignInPrompt: true,
          allowConsentPrompt: true,
          forMSGraphAccess: true,
        });
        const data = await $.ajax({
          type: clientRequest.verb,
          url: clientRequest.url,
          headers: { Authorization: "Bearer " + accessToken },
          cache: false,
        });
        clientRequest.callbackRESTApiHandler(data);
      } catch (error) {
        showMessage(error.responseText);
        switchToFallbackAuth(clientRequest);
        return;
      }
    }
    ```

1. 将 `TODO 9` 替换为下面的代码。 关于此代码，请注意以下几点：

    - 对于 **Microsoft Graph** 错误，请在任务窗格上显示消息。
    - 对于所有其他消息，请在任务窗格上显示消息。

    ```javascript
    // Check for a Microsoft Graph API call error. which is returned as bad request (403)
    if (error.status === 403) {
      if (error.responseJSON && error.responseJSON.type === "Microsoft Graph") {
        showMessage(error.responseJSON.errorDetails);
      } else {
        showMessage(error);
      }
      return;
    }

    // For all other error scenarios, display the message and use fallback auth.
    showMessage("Unknown error from web server: " + JSON.stringify(error));
    if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
    ```

回退身份验证使用 MSAL 库来登录用户。 加载项本身是 SPA，使用 SPA 应用注册来访问中间层服务器。

1. 在 `switchToFallbackAuth` 函数中，将 `TODO 10` 替换为下列代码。 关于此代码，请注意以下几点：

    - 它将全局 `authSSO` 设置为 false，并创建使用 MSAL 进行身份验证的新客户端请求。新请求具有对中间层服务器的 MSAL 访问令牌。
    - 创建请求后，它会调用 `callWebServer` 以继续尝试成功调用中间层服务器。

    ```javascript
    // Guard against accidental call to this function when fallback is already in use.

    if (authSSO === false) return;

    showMessage("Switching from SSO to fallback auth.");
    authSSO = false;
    // Create a new request for fallback auth.
    createRequest(
      clientRequest.verb,
      clientRequest.url,
      clientRequest.callbackRESTApiHandler,
      async (fallbackRequest) => {
        // Hand off to call using fallback auth.
        await callWebServer(fallbackRequest);
      }
    );
    ```

## <a name="code-the-middle-tier-server"></a>对中间层服务器进行编码

中间层服务器提供 REST API 供客户端调用。 例如，REST API `/getuserfilenames` 从用户的 OneDrive 文件夹中获取文件名列表。 每个 REST API 调用都需要客户端提供访问令牌，以确保正确的客户端正在访问其数据。 访问令牌通过代表流 (OBO) 交换 Microsoft Graph 令牌。 新的 Microsoft Graph 令牌由 MSAL 库缓存，用于后续 API 调用。 它永远不会发送到中间层服务器之外。 有关详细信息，请参阅 [中间层访问令牌请求](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>创建路由并实现代理流

1. 打开 文件 `routes\getFilesRoute.js` ，并将 替换为 `TODO 11` 以下代码。 关于此代码，请注意以下几点：

    - 它调用 `authHelper.validateJwt`。 这可确保访问令牌有效且未被篡改。
    - 有关详细信息，请参阅 [验证令牌](/azure/active-directory/develop/access-tokens#validating-tokens)。

    ```javascript
    router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 12: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
    );
    ```

1. 将 `TODO 12` 替换为下面的代码。 关于此代码，请注意以下几点：

    - 它仅请求所需的最小范围，例如 `files.read`。
    - 它使用 MSAL `authHelper` 在调用 `acquireTokenOnBehalfOf`中执行 OBO 流。

    ```javascript
    try {
      const authHeader = req.headers.authorization;
      let oboRequest = {
        oboAssertion: authHeader.split(" ")[1],
        scopes: ["files.read"],
      };

      // The Scope claim tells you what permissions the client application has in the service.
      // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      const accessAsUserScope = tokenScopes.find(
        (scope) => scope === "access_as_user"
      );
      if (!accessAsUserScope) {
        res.status(401).send({ type: "Missing access_as_user" });
        return;
      }
      const cca = authHelper.getConfidentialClientApplication();
      const response = await cca.acquireTokenOnBehalfOf(oboRequest);
      // TODO 13: Call Microsoft Graph to get list of filenames.
    } catch (err) {
      // TODO 14: Handle any errors.
    }
    ```

1. 将 `TODO 13` 替换为下面的代码。 关于此代码，请注意以下几点：

    - 它构造 Microsoft 图形 API 调用的 URL，然后通过 `getGraphData` 函数进行调用。
    - 它通过发送 HTTP 500 响应以及详细信息来返回错误。
    - 成功后，它会将包含文件名列表的 JSON 返回到客户端。

    ```javascript
    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 10 folder or file names.
    const rootUrl = "/me/drive/root/children";

    // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
    // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
    // sanitized so that it cannot be used in a Response header injection attack.
    const params = "?$select=name&$top=10";

    const graphData = await getGraphData(
      response.accessToken,
      rootUrl,
      params
    );

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // there will be a code property in the returned object set to a HTTP status (e.g. 401).
    // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
    if (graphData.code) {
      res
        .status(403)
        .send({
          type: "Microsoft Graph",
          errorDetails:
            "An error occurred while calling the Microsoft Graph API.\n" +
            graphData,
        });
    } else {
      // MS Graph data includes OData metadata and eTags that we don't need.
      // Send only what is actually needed to the client: the item names.
      const itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
      }

      res.status(200).send(itemNames);
    }
    ```

1. 将 `TODO 14` 替换为下面的代码。 此代码专门检查令牌是否过期，因为客户端可以请求新令牌并再次调用。

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(401).send({ type: "TokenExpiredError", errorDetails: err });
   } else {
     res.status(403).send({ type: "Unknown", errorDetails: err });
   }
   ```

该示例必须处理通过 MSAL 的回退身份验证和通过 Office 的 SSO 身份验证。 该示例将首先尝试 SSO，文件 `authSSO` 顶部的布尔值将跟踪示例是否使用 SSO 或已切换到回退身份验证。

## <a name="run-the-project"></a>运行项目

1. 请确保 OneDrive 中有一些文件，以便可以验证结果。

1. 在 `\Begin` 文件夹的根目录中打开命令提示符。

1. 运行 命令 `npm install` 以安装所有包依赖项。

1. 运行 命令 `npm start` 以启动中间层服务器。

1. 需要将加载项旁加载到 Office 应用程序（Excel、Word 或 PowerPoint），以便对其进行测试。 说明取决于你的平台。 在[旁加载 Office 加载项以供测试](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)中有指向说明的链接。

1. 在 Office 应用程序的“**主页**”功能区上，选择“**SSO Node.js**”组中的“**显示加载项**”按钮以打开任务窗格加载项。

1. 单击“**获取 OneDrive 文件名**”按钮。 如果你使用Microsoft 365 教育版或工作帐户或 Microsoft 帐户登录 Office，并且 SSO 按预期工作，则OneDrive for Business中的前 10 个文件和文件夹名称将插入到文档中。  (第一次可能需要 15 秒的时间。) 如果未登录，或者你处于不支持 SSO 的方案中，或者 SSO 由于任何原因无法正常工作，系统会提示你登录。 登录后，将显示文件和文件夹名称。

> [!NOTE]
> 如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已更改过，也不例外。 在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。 为了防止发生这种情况，请务必先 _关闭其他所有 Office 应用程序_，然后再按“**获取 OneDrive 文件名**”。

## <a name="security-notes"></a>安全说明

- 中的`/getuserfilenames``getFilesroute.js`路由使用文本字符串来编写对 Microsoft Graph 的调用。 如果更改调用以便字符串的任何部分来自用户输入，请清理输入，使其不能用于响应标头注入攻击。

- 在 `app.js` 以下内容中，脚本的安全策略已到位。 你可能希望根据加载项的安全需求指定其他限制。

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

始终遵循[Microsoft 标识平台文档中](/azure/active-directory/develop/)的安全最佳做法。
