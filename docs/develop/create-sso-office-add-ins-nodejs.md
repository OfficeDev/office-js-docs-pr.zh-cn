---
title: 创建使用单一登录的 Node.js Office 加载项
description: 了解如何创建使用 Office 单一登录的基于 Node.js 的加载项。
ms.date: 07/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6f71630f2694db9c53ba6d2e3e6d07f54ab91cb8
ms.sourcegitcommit: c62d087c27422db51f99ed7b14216c1acfda7fba
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/08/2022
ms.locfileid: "66689402"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on"></a>创建使用单一登录的 Node.js Office 加载项

用户可以登录 Office，Office Web 加载项能够利用此登录进程，授权用户访问加载项和 Microsoft Graph，而无需要求用户再登录一次。有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。

本文介绍如何在加载项中启用单一登录 (SSO) 。 创建的示例加载项包含两个部分：在 Microsoft Excel 中加载的任务窗格，以及用于处理任务窗格对 Microsoft Graph 的调用的中间层服务器。 中间层服务器使用 Node.js 和 Express 生成，并公开单个 REST API， `/getuserfilenames`该 API 返回用户 OneDrive 文件夹中前 10 个文件名的列表。 任务窗格使用该 `getAccessToken()` 方法获取已登录用户到中间层服务器的访问令牌。 中间层服务器使用代理流 (OBO) 来交换访问 Microsoft Graph 的新令牌。 可以扩展此模式以访问任何 Microsoft Graph 数据。 任务窗格始终调用中间层 REST API (在需要 Microsoft Graph 服务时传递访问令牌) 。 中间层使用通过 OBO 获取的令牌调用 Microsoft Graph 服务并将结果返回到任务窗格。

本文适用于使用 Node.js 和 Express 的加载项。 有关与此类似的 ASP.NET 加载项文章，请参阅[创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）

- [Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）

- 代码编辑器 - 建议Visual Studio Code

- Microsoft 365 订阅中OneDrive for Business上存储的至少几个文件和文件夹

- 支持 [IdentityAPI 1.3 要求集](/javascript/api/requirement-sets/common/identity-api-requirement-sets) 的 Microsoft 365 内部版本。 可以获取[免费开发人员沙盒，该沙盒](https://developer.microsoft.com/microsoft-365/dev-program#Subscription)提供可续订的 90 天Microsoft 365 E5开发人员订阅。 开发人员沙盒包含 Microsoft Azure 订阅，可在本文后面的步骤中用于应用注册。 如果愿意，可以使用单独的 Microsoft Azure 订阅进行应用注册。 获取 [Microsoft Azure](https://account.windowsazure.com/SignUp) 的试用订阅。

## <a name="set-up-the-starter-project"></a>设置初学者项目

1. 克隆或下载 [Office 外接程序 NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) 中的存储库。

   > [!NOTE]
   > 示例项目有两个版本：
   >
   > - **Begin** 文件夹是一个初学者项目。 未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。 本文后续章节将引导你完成此过程。
   > - **完整** 文件夹包含与本文中完成的所有编码步骤相同的示例。 若要使用已完成的版本，只需按照本文中的说明操作，但将“Begin”替换为“完成”，并跳过“ **代码客户端** ”和 **“编码中间层服务器** 端”部分。

1. 在 **Begin** 文件夹中打开命令提示符。

1. 在该控制台中输入 `npm install` 以安装 package.json 文件中列出明细的所有依赖项。

1. 运行命令 `npm run install-dev-certs`。 为安装证书的提示选择“**是**”。

## <a name="register-the-add-in-with-microsoft-identity-platform"></a>使用Microsoft 标识平台注册加载项

需要在 Azure 中创建表示中间层服务器的应用注册。 这可以启用身份验证支持，以便可以向 JavaScript 中的客户端代码颁发适当的访问令牌。 此注册支持客户端中的 SSO，以及使用 Microsoft 身份验证库 (MSAL) 进行回退身份验证。

1. 若要注册应用，请导航到[Azure 门户 - 应用注册](https://go.microsoft.com/fwlink/?linkid=2083908)页注册应用。

1. 使用 **_管理员_** 凭据登录到 Microsoft 365 租户。 例如，MyName@contoso.onmicrosoft.com。

1. 选择“新注册”。 在“注册应用”页上，按如下方式设置值。

   - 将“名称”设置为“`Office-Add-in-NodeJS-SSO`”。
   - 将 **支持的帐户类型** 设置为 **任何组织目录中的帐户 (任何 Azure AD 目录 - 多租户) 和个人 Microsoft 帐户 (例如 Skype、Xbox) 。**
   - 在 **“重定向 URI”** 部分中，使用重定向 URI 值`https://localhost:44355/dialog.html`将平台设置 **为单页应用程序 (SPA)**。
   - 选择 **“注册”**。

   > [!NOTE]
   > 仅当客户端使用 MSAL 进行回退身份验证时，才使用 SPA 应用程序类型。

1. 在 **Office-Add-in-NodeJS-SSO** 页面上，复制并保存“**应用程序（客户端）ID**”和“**目录（租户）ID**”的值。 你将在后面的过程中使用它们。

   > [!NOTE]
   > 当其他应用程序 **（例如 Office 客户** 端应用程序 (（例如，PowerPoint、Word、Excel) ）寻求对应用程序的授权访问时，此应用程序 (客户端) ID 是“受众”值。 它也是应用程序在寻求对 Microsoft Graph 的授权访问权限时的“客户端 ID”。

1. 在最左侧栏中，在 **“管理**”下选择 **“身份验证**”。 在 **“隐式授予”和“混合流** ”部分中，选择 **访问令牌** 和 **ID** 令牌的复选框。 当 SSO 不可用时，该示例使用 Microsoft 身份验证库 (MSAL) 进行回退身份验证。

1. 选择“**保存**”。

1. 在 **“管理”** 下，选择 **“证书”&机密** ，然后选择 **“新建客户端机密**”。 输入“**描述**”的值，然后选择适当的“**到期**”选项，并选择“**添加**”。

   Web 应用程序在请求令牌时使用客户端机密 **值** 来证明其标识。 _记录此值以便在后续步骤中使用 - 它只显示一次。_

1. 在最左侧栏中，选择“**管理**”下 **的“公开 API**”。 选择 **“设置** ”链接。 这将以“api：//$App ID GUID$”的形式生成应用程序 ID URI，其中$App ID GUID$ 是 **应用程序 (客户端) ID**。

1. 在生成的 ID 中，插入 `localhost:44355/` (记下追加到两个正斜杠和 GUID 之间的结束) 的正斜杠“/”。 完成后，整个 ID 应具有窗体 `api://localhost:44355/$App ID GUID$`，例如 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。 然后选择“**保存**”。

1. 选择“添加一个作用域”按钮。 在打开的面板中，输入 `access_as_user` 作为“作用域名称”。

1. 将“谁能同意?”设置为“管理员和用户”。

1. 填写用于配置管理员和用户同意提示的字段，其中包含适合 `access_as_user` 作用域的值，使 Office 客户端应用程序能够使用与当前用户具有相同权限的外接程序的 Web API。 建议：

   - **管理员许可显示名称**：Office 可以充当用户。
   - **管理员许可描述**：使 Office 能够借助与当前用户相同的权限调用加载项的 Web API。
   - **用户同意显示名称**：Office 可以充当你。
   - **用户同意说明**：允许 Office 使用与你拥有的权限相同的权限调用外接程序的 Web API。

1. 确保将“状态”设置为“已启用”。

1. 选择“添加作用域”。

   > [!NOTE]
   > 显示在文本字段正下方的 **作用域** 名称的域部分应自动与你先前设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。

1. 在 **“授权客户端应用程序**”部分中，选择 **“添加客户端应用程序**”按钮，然后在打开的面板中，将客户端 ID 设置为`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e``api://localhost:44355/$app-id-guid$/access_as_user`“**授权范围**”复选框。

1. 选择“添加应用程序”。

   > [!NOTE]
   > 该 `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID 预授权所有 Microsoft Office 应用程序终结点。 如果要在 Windows 和 Mac 上的 Office 上支持 MICROSOFT 帐户 (MSA) ，则还需要此功能。 或者，如果出于任何原因想要拒绝某些平台上的 Office 授权，则可以输入以下 ID 的适当子集。 只需保留要从中隐瞒授权的平台的 ID 即可。 这些平台上外接程序的用户将无法调用 Web API，但外接程序中的其他功能仍将有效。
   >
   > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
   > - `93d53678-613d-4013-afc1-62e9e444a0a5`（Office 网页版）
   > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）

1. 在最左侧栏中，选择“**管理”** 下 **的 API 权限**，然后选择 **“添加权限**”。 在打开的面板上，选择 **Microsoft Graph**，然后选择“委派权限”。

1. 使用“选择权限”搜索框来搜索加载项需要的权限。 选择以下选项。 加载项本身仅需要第一个加载项; `profile` 但 Office 应用程序需要这些权限和 `openid` 权限才能获取具有用户标识的访问令牌才能访问中间层服务器。

   - **Files.Read**
   - **个人资料**
   - **openid**

   > [!NOTE]
   > `User.Read` 权限可能已默认列出。 最好不要请求不需要的权限，因此，如果外接程序实际上不需要，建议取消选中此权限的框。

1. 选择所显示的每个权限的复选框。 选择加载项需要的权限后，选择面板底部的“**添加权限**”按钮。

1. 在同一页面上，选择“**为[租户名称]授予管理员许可**”按钮，然后为显示的确认选择“**是**”。

## <a name="configure-the-add-in"></a>配置加载项

1. 在代码编辑器中打开克隆项目中的 `\Begin` 文件夹。

1. 打开该 `.ENV` 文件并使用之前从 **Office-Add-in-NodeJS-SSO** 应用注册复制的值。 按如下所示设置值：

   | 名称              | 值                                                            |
   | ----------------- | ---------------------------------------------------------------- |
   | **CLIENT_ID**     | 应用程序 (应用注册概述页中 **的客户端) ID**。 |
   | **CLIENT_SECRET** | 从 **“证书&机密**”页保存 **的客户端机密**。       |
   | **DIRECTORY_ID**  | 从应用注册概述页 **(租户) ID 的目录**。   |

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

1. 将标记 _中两个位置的_ 占位符“$app-id-guid$”替换为创建 **Office-Add-in-NodeJS-SSO** 应用注册时复制的应用程序 **ID**。 “$”符号不是 ID 的一部分，因此请不要包含它们。 这是用于CLIENT_ID的相同 ID。ENV 文件。

   > [!NOTE]
   > 该 **\<Resource\>** 值是注册加载项时设置的应用程序 **ID URI** 。 仅当加载项通过 AppSource 出售时，该 **\<Scopes\>** 部分仅用于生成同意对话框。

1. 打开 `\public\javascripts\fallback-msal\authConfig.js`文件。 将占位符“$app-id-guid$”替换为之前创建的 **Office-Add-in-NodeJS-SSO** 应用注册中保存的应用程序 ID。

1. 保存对文件所做的更改。

## <a name="code-the-client-side"></a>编写客户端代码

### <a name="create-client-request-and-response-handler"></a>创建客户端请求和响应处理程序

1. 在代码编辑器中，打开文件 `public\javascripts\ssoAuthES6.js`。 它已经具有确保即使在 Internet Explorer 11 中也支持 Promise 的代码，并且具有 `Office.onReady` 调用，可将处理程序分配给加载项的唯一按钮。

   > [!NOTE]
   > 顾名思义，ssoAuthES6.js 使用 JavaScript ES6 语法，因为使用 `async` 和 `await` 可以最好地显示 SSO API 本质的简单性。 启动 localhost 服务器时，此文件将转译为 ES5 语法，以便该示例将支持 Internet Explorer 11。

    示例代码的一个关键部分是客户端请求。 客户端请求是一个对象，用于跟踪有关在中间层服务器上调用 REST API 的请求的信息。 这是必要的，因为需要通过以下方案跟踪或更新客户端请求状态：

    - SSO 重试 REST API 调用失败的位置，因为它需要额外的同意。 使用更新的身份验证选项进行示例代码调 `getAccessToken` 用，获取所需的用户同意，然后再次调用 REST API。 目标是在 REST API 需要额外同意的情况下不失败。
    - SSO 失败，需要回退身份验证。 访问令牌通过 MSAL 在弹出对话框中获取。 目标是在此方案中不失败，并正常地回退到替代身份验证方法。

    客户端请求对象跟踪以下数据：

    - `authOptions` - SSO [的身份验证配置参数](/javascript/api/office/office.authoptions)。
    - `authSSO` - 如果使用 SSO，则为 true，否则为 false。
    - `accessToken` - 中间层服务器的访问令牌。 对于 SSO，获取此令牌的方法不同于回退身份验证。
    - `url` - 要在中间层服务器上调用的 REST API 的 URL。
    - `callbackHandler` - 传递 REST API 调用结果的函数。
    - `callbackFunction` - 准备就绪时要将客户端请求传递到的函数。

1. 若要初始化客户端请求对象，请在函数中 `createRequest` 替换 `TODO 1` 为以下代码。

   ```javascript
   const clientRequest = {
     authOptions: {
       allowSignInPrompt: true,
       allowConsentPrompt: true,
       forMSGraphAccess: true,
     },
     authSSO: authSSO,
     accessToken: null,
     url: url,
     callbackRESTApiHandler: restApiCallback,
     callbackFunction: callbackFunction,
   };
   ```

1. 将 `TODO 2` 替换为下面的代码。 关于此代码，请注意以下几点：

   - 它检查是否正在使用 SSO。 对于 SSO，获取访问令牌的方法不同于回退身份验证。
   - 如果 SSO 返回访问令牌，它将调用该 `callbackfunction` 函数。 对于它调用 `dialogFallback`的回退身份验证，最终会在用户通过 MSAL 登录后调用回调函数。

   ```javascript
   // Get access token.

   if (authSSO) {
     try {
       // Get access token from Office SSO.
       clientRequest.accessToken = await getAccessTokenFromSSO(
         clientRequest.authOptions
       );
       callbackFunction(clientRequest);
     } catch {
       // Use fallback authentication if SSO failed to get access token.
       switchToFallbackAuth(clientRequest);
     }
   } else {
     // Use fallback authentication to get access token.
     dialogFallback(clientRequest);
   }
   ```

1. 在 `getFileNameList` 函数中，将 `TODO 3` 替换为下列代码。 关于此代码，请注意以下几点：

   - `getFileNameList`当用户选择任务窗格上的 **“获取 OneDrive 文件名**”按钮时，将调用该函数。
   - 它会创建一个客户端请求来跟踪有关调用的信息，例如 REST API 的 URL。
   - 当 REST API 返回结果时，它将传递给函 `handleGetFileNameResponse` 数。 此回调作为参数传递给`createRequest`并跟踪。`clientRequest.callbackRESTApiHandler`
   - 使用客户端请求执行后续步骤并调用 REST API 的代码调 `callWebServer` 用。

   ```javascript
   createRequest(
     "/getuserfilenames",
     handleGetFileNameResponse,
     async (clientRequest) => {
       await callWebServer(clientRequest);
     }
   );
   ```

1. 在 `handleGetFileNameResponse` 函数中，将 `TODO 4` 替换为下列代码。 关于此代码，请注意以下几点：

   - 代码将响应传递 (其中包含要将 `writeFileNamesToOfficeDocument` 文件名写入文档的文件名) 列表。
   - 代码检查错误。 如果写入文件名，则显示成功消息，否则显示错误。

   ```javascript
   if (response != null) {
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

### <a name="get-the-sso-access-token"></a>获取 SSO 访问令牌

1. 在 `getAccessTokenFromSSO` 函数中，将 `TODO 5` 替换为下列代码。 关于此代码，请注意以下几点：

   - 它调用 `Office.auth.getAccessToken` 从 Office 获取访问令牌。
   - 如果发生错误，它将调用 `handleSSOErrors` 函数。 如果无法处理该错误，它将向调用方引发错误。 这是指示调用方切换到回退身份验证。

   ```javascript
   try {
     // The access token returned from getAccessToken only has permissions to your middle-tier server APIs,
     // and it contains the identity claims of the signed-in user.

     const accessToken = await Office.auth.getAccessToken(authOptions);
     return accessToken;
   } catch (error) {
     let fallbackRequired = handleSSOErrors(error);
     if (fallbackRequired) throw error; // Rethrow the error and caller will switch to fallback auth.
     return null; // Returning a null token indicates no need for fallback (an explanation about the error condition was shown by handleSSOErrors).
   }
   ```

1. 在 `handleSSOErrors` 函数中，将 `TODO 6` 替换为下列代码。 有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。

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

1. 将 `TODO 7` 替换为下面的代码。 有关这些错误的详细信息，请参阅 [Office 加载项中的 SSO 疑难解答](troubleshoot-sso-in-office-add-ins.md)。对于无法处理的任何错误， `true` 将返回给调用方。 这表示调用方应切换到使用 MSAL 作为回退身份验证。

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

1. 在 `callWebServer` 函数中，将 `TODO 8` 替换为下列代码。 关于此代码，请注意以下几点：

   - 实际的 AJAX 调用将由函 `ajaxCallToRESTApi` 数进行。
   - 如果中间层服务器返回一个错误，指示当前令牌已过期，则此函数将尝试获取新的访问令牌。
   - 如果无法成功完成 AJAX 调用， `switchToFallbackAuth` 将调用它来使用 MSAL 身份验证而不是 Office SSO。

   ```javascript
   try {
     await ajaxCallToRESTApi(clientRequest);
   } catch (error) {
     if (error.statusText === "Internal Server Error") {
       const retryCall = handleWebServerErrors(error, clientRequest);
       if (retryCall && clientRequest.authSSO) {
         try {
           clientRequest.accessToken = await getAccessTokenFromSSO(
             clientRequest.authOptions
           );
           await ajaxCallToRESTApi(clientRequest);
         } catch {
           // If still an error go to fallback.
           switchToFallbackAuth(clientRequest);
           return;
         }
       }
     } else {
       console.log(JSON.stringify(error)); // Log any errors.
       showMessage(error.responseText);
     }
   }
   ```

1. 在 `ajaxCallToRESTApi` 函数中，将 `TODO 9` 替换为下列代码。 关于此代码，请注意以下几点：

   - 该函数显式重新引发调用方要处理的任何错误。

   ```javascript
   try {
     await $.ajax({
       type: "GET",
       url: clientRequest.url,
       headers: { Authorization: "Bearer " + clientRequest.accessToken },
       cache: false,
       success: function (data) {
         result = data;
         // Send result to the callback handler.
         clientRequest.callbackRESTApiHandler(result);
       },
     });
   } catch (error) {
     // This function explicitly requires the caller to handle any errors
     throw error;
   }
   ```

1. 在 `handleWebServerErrors` 函数中，将 `TODO 10` 替换为下列代码。 关于此代码，请注意以下几点：

   - 该错误由中间层服务器返回，该服务器指示错误的类型，并使此处更易于处理。
   - 对于 **Microsoft Graph** 错误，请在任务窗格上显示消息。
   - 对于 **AADSTS500133** 错误，返回 true，以便调用方知道令牌已过期，应获取新的令牌。
   - 对于所有其他消息，请在任务窗格中显示消息。

   ```javascript
   let retryCall = false;
   // Our middle-tier server returns a type to help handle the known cases.
   switch (err.responseJSON.type) {
     case "Microsoft Graph":
       // An error occurred when the middle-tier server called Microsoft Graph.
       showMessage(
         "Error from Microsoft Graph: " +
           JSON.stringify(err.responseJSON.errorDetails)
       );
       retryCall = false;
       break;
     case "Missing access_as_user":
       // The access_as_user scope was missing.
       showMessage("Error: Access token is missing the access_as_user scope.");
       retryCall = false;
       break;
     case "AADSTS500133": // expired token
       // On rare occasions the access token could expire after it was sent to the middle-tier server.
       // Microsoft identity platform will respond with
       // "The provided value for the 'assertion' is not valid. The assertion has expired."
       // Return true to indicate to caller they should refresh the token.
       retryCall = true;
       break;
     default:
       showMessage(
         "Unknown error from web server: " +
           JSON.stringify(err.responseJSON.errorDetails)
       );
       retryCall = false;
       if (clientRequest.authSSO) switchToFallbackAuth(clientRequest);
   }
   return retryCall;
   ```

回退身份验证将使用 MSAL 库登录用户。 外接程序本身是 SPA，使用 SPA 应用注册来访问中间层服务器。

1. 在 `switchToFallbackAuth` 函数中，将 `TODO 11` 替换为下列代码。 关于此代码，请注意以下几点：

   - 它将全局 `authSSO` 设置为 false，并创建使用 MSAL 进行身份验证的新客户端请求。新请求具有对中间层服务器的 MSAL 访问令牌。
   - 创建请求后，它会调用 `callWebServer` 继续尝试成功调用中间层服务器。

   ```javascript
   showMessage("Switching from SSO to fallback auth.");
   authSSO = false;
   // Create a new request for fallback auth.
   createRequest(
     clientRequest.url,
     clientRequest.callbackRESTApiHandler,
     async (fallbackRequest) => {
       // Hand off to call using fallback auth.
       await callWebServer(fallbackRequest);
     }
   );
   ```

## <a name="code-the-middle-tier-server"></a>对中间层服务器进行编码

中间层服务器提供 REST API 供客户端调用。 例如，REST API `/getuserfilenames` 从用户的 OneDrive 文件夹中获取文件名列表。 每个 REST API 调用都需要客户端的访问令牌，以确保正确的客户端正在访问其数据。 访问令牌通过 OBO)  (代表流交换 Microsoft Graph 令牌。 新的 Microsoft Graph 令牌由 MSAL 库缓存，用于后续 API 调用。 它永远不会在中间层服务器外部发送。 有关详细信息，请参阅 [中间层访问令牌请求](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#middle-tier-access-token-request)

### <a name="create-the-route-and-implement-on-behalf-of-flow"></a>创建路由并实现代理流

1. 打开文件 `routes\getFilesRoute.js` 并替换 `TODO 12` 为以下代码。 关于此代码，请注意以下几点：

   - 它调用 `authHelper.validateJwt`。 这可确保访问令牌有效且未被篡改。
   - 有关详细信息，请参阅 [验证令牌](/azure/active-directory/develop/access-tokens#validating-tokens)。

   ```javascript
   router.get(
     "/getuserfilenames",
     authHelper.validateJwt,
     async function (req, res) {
       // TODO 13: Exchange the access token for a Microsoft Graph token
       //          by using the OBO flow.
     }
   );
   ```

1. 将 `TODO 13` 替换为下面的代码。 关于此代码，请注意以下几点：

   - 它只请求所需的最小范围，例如 `files.read`。
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
     // TODO 14: Call Microsoft Graph to get list of filenames.
   } catch (err) {
     // TODO 15: Handle any errors.
   }
   ```

1. 将 `TODO 14` 替换为下面的代码。 关于此代码，请注意以下几点：

   - 它构造 Microsoft 图形 API 调用的 URL，然后通过函`getGraphData`数进行调用。
   - 它通过发送 HTTP 500 响应以及详细信息来返回错误。
   - 成功后，它将包含文件名列表的 JSON 返回到客户端。

   ```javascript
   // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
   // and only the top 10 folder or file names.
   const rootUrl = "/me/drive/root/children";

   // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
   // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
   // sanitized so that it cannot be used in a Response header injection attack.
   const params = "?$select=name&$top=10";

   const graphData = await getGraphData(response.accessToken, rootUrl, params);

   // If Microsoft Graph returns an error, such as invalid or expired token,
   // there will be a code property in the returned object set to a HTTP status (e.g. 401).
   // Return it to the client. On client side it will get handled in the fail callback of `makeWebServerApiCall`.
   if (graphData.code) {
     res.status(500).send({ type: "Microsoft Graph", errorDetails: graphData });
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

1. 将 `TODO 15` 替换为以下代码。 此代码专门检查令牌是否已过期，因为客户端可以请求新令牌并再次调用。

   ```javascript
   // On rare occasions the SSO access token is unexpired when Office validates it,
   // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
   // with "The provided value for the 'assertion' is not valid. The assertion has expired."
   // Construct an error message to return to the client so it can refresh the SSO token.
   if (err.errorMessage.indexOf("AADSTS500133") !== -1) {
     res.status(500).send({ type: "AADSTS500133", errorDetails: err });
   } else {
     res.status(500).send({ type: "Unknown", errorDetails: err });
   }
   ```

该示例必须通过 MSAL 和 SSO 身份验证通过 Office 处理回退身份验证。 如果示例使用 SSO 或已切换到回退身份验证，则该示例将首先尝试 SSO， `authSSO` 并且文件顶部的布尔值会跟踪。

## <a name="run-the-project"></a>运行项目

1. 请确保 OneDrive 中有一些文件，以便可以验证结果。

1. 在 `\Begin` 文件夹的根目录中打开命令提示符。

1. 运行该命令 `npm install` 以安装所有包依赖项。

1. 运行命令 `npm start` 以启动中间层服务器。

1. 需要将加载项旁加载到 Office 应用程序（Excel、Word 或 PowerPoint），以便对其进行测试。 说明取决于你的平台。 在[旁加载 Office 加载项以供测试](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)中有指向说明的链接。

1. 在 Office 应用程序的“**主页**”功能区上，选择“**SSO Node.js**”组中的“**显示加载项**”按钮以打开任务窗格加载项。

1. 单击“**获取 OneDrive 文件名**”按钮。 如果使用Microsoft 365 教育版或工作帐户或 Microsoft 帐户登录到 Office，并且 SSO 按预期工作，则OneDrive for Business中的前 10 个文件和文件夹名称将插入到文档中。  (第一次登录可能需要多达 15 秒的时间。) 如果未登录，或者你处于不支持 SSO 或 SSO 因任何原因不起作用的方案中，系统会提示你登录。 登录后，将显示文件和文件夹名称。

> [!NOTE]
> 如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已更改过，也不例外。 在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。 为了防止发生这种情况，请务必先 _关闭其他所有 Office 应用程序_，然后再按“**获取 OneDrive 文件名**”。

## <a name="security-notes"></a>安全说明

* 其中的`/getuserfilenames``getFilesroute.js`路由使用文本字符串来撰写对 Microsoft Graph 的调用。 如果更改调用以使字符串的任何部分来自用户输入，请清理输入，使其不能用于响应标头注入攻击。

* 以下 `app.js` 内容安全策略适用于脚本。 可能需要根据加载项安全需求指定其他限制。

    `"Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,`

始终遵循[Microsoft 标识平台文档](/azure/active-directory/develop/)中的安全最佳做法。
