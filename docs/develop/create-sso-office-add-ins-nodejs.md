---
title: 创建使用单一登录的 Node.js Office 加载项
description: 2018 年 1 月23 日
ms.openlocfilehash: bb77d037140f8c56ca05f3817fb2b9d0271297ae
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876611"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a>创建使用单一登录的 Node.js Office 加载项（预览）

用户可以登录 Office，Office Web 加载项能够利用此登录进程，授权用户访问加载项和 Microsoft Graph，而无需要求用户再登录一次。有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。

本文将逐步介绍如何在使用 Node.js 和 Express 生成的加载项中启用单一登录 (SSO) 。 

> [!NOTE]
> 有关与此类似的 ASP.NET 加载项文章，请参阅[创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)。

## <a name="prerequisites"></a>先决条件

* [节点和 npm](https://nodejs.org/en/) 版本 6.9.4 或更高版本

* [Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）

* TypeScript 版本 2.2.2 或更高版本

* Office 2016 版本 1708（生成号 8424.nnnn）或更高版本（Office 365 订阅版本，有时亦称为“即点即用”）

  可能必须成为 Office 预览体验成员，才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。

## <a name="set-up-the-starter-project"></a>创建起始项目

1. 克隆或下载 [Office 加载项 NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso) 中的存储库。 

    > [!NOTE]
    > 示例项目有三个版本：  
    > * **Before** 文件夹是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。本文后续章节将引导你完成此过程。 
    > * 如果完成了本文中的过程，该示例的**已完成**版本会与所生成的外接程序类似，只不过完成的项目具有对本文文本冗余的代码注释。若要使用已完成的版本，请按照本文中的说明进行操作即可，但需要将“Before”替换为“Completed”，并跳过**编写客户端代码**和**编写服务器端代码**部分。
    > * **完整多租户**版本是支持多租户的完整示例。 如果要支持带 SSO 的来自不同域的 Microsoft 帐户，请浏览此示例。

2. 在 **Before** 文件夹中打开 Git bash 控制台。

3. 在该控制台中输入 `npm install` 以安装 package.json 文件中列出明细的所有依赖项。

4. 在控制台中输入 `npm run build `，以生成项目。 

    > [!NOTE]
    > 可能会看到一些生成错误，提示某些变量已声明但未使用。请忽略这些错误。之所以会看到这些错误是因为，示例项目的“之前”版本缺少某代码，将在后续步骤中添加。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>向 Azure AD v2.0 端点注册加载项

以下说明以通用方式书写，以便可以在多个地方使用。 对于本文而言，请执行以下操作：
- 将占位符 **$ADD-IN-NAME$** 替换为 `“Office-Add-in-NodeJS-SSO`。
- 将占位符 **$FQDN-WITHOUT-PROTOCOL$** 替换为 `localhost:3000`。
- 在**选择权限**对话框中指定权限时，选中以下权限框。 只有第一个是加载 项本身真正需要的；但 Office 主机需要 `profile` 权限来为加载项 Web 应用程序获取令牌。
    * Files.Read.All
    * profile

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]


## <a name="grant-administrator-consent-to-the-add-in"></a>向加载项授予管理员许可

[!INCLUDE[](../includes/grant-admin-consent-to-an-add-in-include.md)]

## <a name="configure-the-add-in"></a>配置加载项

1. 在代码编辑器中打开 src\server.ts 文件。顶部附近存在对 `AuthModule` 类的构造函数的调用。该构造函数中存在一些需要为其分配值的字符串参数。

2. 对于 `client_id` 属性，将占位符 `{client GUID}` 替换为注册加载项时保存的应用程序 ID。 完成后，单引号中应该只有一个 GUID。 不应出现任何 "{}" 字符。

3. 对于 `client_secret` 属性，将占位符 `{client secret}` 替换为注册加载项时保存的应用程序机密。

4. 对于 `audience` 属性，将占位符 `{audience GUID}` 替换为注册外接程序时保存的应用程序 ID。（即分配给 `client_id` 属性的同一值）。
  
3. 在分配给 `issuer` 属性的字符串中，你会看到占位符 *{O365 tenant GUID}*。 将其替换为 Office 365 租约 ID。 使用[找到你的 Office 365 租户 ID](https://docs.microsoft.com/onedrive/find-your-office-365-tenant-id) 中的一种方法获得它。 完成后，`issuer` 属性值应该如下所示：

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

    > [!NOTE]
    > * **Resource** 值是向注册的加载项添加 Web API 平台时设置的**应用 ID URI**。
    > * 仅在通过 AppSource 销售加载项时，才使用 **Scopes** 部分生成许可对话框。

1. 保存并关闭文件。

## <a name="code-the-client-side"></a>编写客户端代码

1. 打开 **public** 文件夹中的 program.js 文件。其中已存在一些代码：

    * 针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。
    * 方法，用于在任务窗格底部显示从 Microsoft Graph 返回的数据（或错误消息）。`showResult`
    * 方法，用于记录最终用户不应看到的控制台错误。`logErrors`

11. 在 `Office.initialize` 赋值语句的下方，添加下列代码。关于此代码，请注意以下几点：

    * 加载项中的错误处理有时会自动尝试使用一组不同的选项，重新获取访问令牌。 计数器变量 `timesGetOneDriveFilesHasRun` 以及标志变量 `triedWithoutForceConsent` 和 `timesMSGraphErrorReceived` 用于确保用户不会重复循环失败的尝试来获取令牌。 
    * 虽然 `getDataWithToken` 方法是在下一步中创建，但请注意，它会将 `forceConsent` 选项设置为 `false`。有关详细信息，请参阅下一步。

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;
    var timesMSGraphErrorReceived = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. 在 `getOneDriveFiles` 方法下方，添加下列代码。关于此代码，请注意以下几点：

    * [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) 是 Office.js 中的新 API，可便于加载项要求 Office 主机应用（Excel、PowerPoint、Word 等）提供加载项访问令牌（对于已登录 Office 的用户）。反过来，Office 主机应用会向 Azure AD 2.0 端点请求获取令牌。由于已在注册加载项时将 Office 主机预授权给加载项，因此 Azure AD 会发送该令牌。
    * 如果用户未登录 Office，Office 主机会提示用户登录。
    * options 参数将 `forceConsent` 设置为 `false`，因此用户不会在每次使用加载项时都看到提示，要求其许可向 Office 主机授予对加载项的访问权限。 用户首次运行加载项时，`getAccessTokenAsync` 调用会失败，但在后续步骤中添加的错误处理逻辑会自动重新调用（`forceConsent` 选项设置为 `true`），并提示用户许可，但仅限首次运行。
    * 方法将在后续步骤中创建。`handleClientSideErrors`

    ```javascript
    function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                TODO1: Use the access token to get Microsoft Graph data.
            }
            else {
                handleClientSideErrors(result);
            }
        });
    }
    ```

1. 用以下行替换 TODO1。可以在后续步骤中创建 `getData` 方法和服务器端“/api/values”路由。相对 URL 用于终结点，因为它必须与外接程序托管在相同的域中。

    ```javascript
    accessToken = result.value;
    getData("/api/values", accessToken);
    ```

1. 在 `getOneDriveFiles` 方法下方，添加下列代码。关于此代码，请注意以下几点：

    * 此方法调用指定 Web API 终结点，并向它传递访问令牌，这也是 Office 主机应用用于获取对加载项的访问权限的令牌。在服务器端，此访问令牌将用于“代表”流，以获取对 Microsoft Graph 的访问令牌。
    * 方法将在后续步骤中创建。`handleServerSideErrors`

    ```javascript
    function getData(relativeUrl, accessToken) {
        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        }); 
    }
    ```

### <a name="create-the-error-handling-methods"></a>创建错误处理方法

1. 在 `getData` 方法下方，添加下列方法。 当 Office 主机无法获取对加载项 Web 服务的访问令牌时，此方法便会处理加载项客户端中的错误。 这些错误通过错误代码进行报告，因此下面的方法使用 `switch` 语句区分它们。

    ```javascript
    function handleClientSideErrors(result) {

        switch (result.error.code) {
    
            // TODO2: Handle the case where user is not logged in, or the user cancelled, without responding, a
            //        prompt to provide a 2nd authentication factor. 
    
            // TODO3: Handle the case where the user's sign-in or consent was aborted.
    
            // TODO4: Handle the case where the user is logged in with an account that is neither work or school, 
            //        nor Micrososoft Account.
    
            // TODO5: Handle an unspecified error from the Office host.
    
            // TODO6: Handle the case where the Office host cannot get an access token to the add-ins 
            //        web service/application.
    
            // TODO7: Handle the case where the user tiggered an operation that calls `getAccessTokenAsync` 
            //        before a previous call of it completed.
    
            // TODO8: Handle the case where the add-in does not support forcing consent.
    
            // TODO9: Log all other client errors.
        }
    }
    ```

1. 将 `TODO2` 替换为以下代码。 如果用户未登录或用户取消（未响应）提供辅助身份验证因素的提示，错误 13001 发生。 无论属于上述哪种情况，代码都会重新运行 `getDataWithToken` 方法，并设置强制登录提示选项。

    ```javascript
    case 13001:
        getDataWithToken({ forceAddAccount: true });
        break;
    ```

1. 将 `TODO3` 替换为以下代码。 如果用户登录或许可被中止，错误 13002 发生。 建议用户重试一次，但只能重试一次。

    ```javascript
    case 13002:
        if (timesGetOneDriveFilesHasRun < 2) {
            showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
        } else {
            logError(result);
        }          
        break; 
    ```

1. 将 `TODO4` 替换为以下代码。 如果用户用于登录的帐户既不是工作帐户或学校帐户，也不是 Microsoft 帐户，错误 13003 发生。 建议用户注销，再使用受支持的帐户类型重新登录。

    ```javascript
    case 13003: 
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
        break;   
    ```

    > [!NOTE]
    > 此方法不处理错误 13004 和 13005，因为它们只在开发期间出现。 无法通过运行时代码进行修复，并且向最终用户报告这两个错误也没有意义。

1. 将 `TODO5` 替换为下列代码。如果 Office 主机中出现可能表明主机处于不稳定状态的未指定错误，就会发生错误 13006。建议用户重启 Office。

    ```javascript
    case 13006:
        showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
        break;        
    ```

1. 将 `TODO6` 替换为以下代码。 如果 Office 主机与 AAD 之间的交互出现问题，导致主机无法获得对加载项 Web 服务/应用的访问令牌，错误 13007 发生。 这可能由于暂时网络问题所致。 建议用户稍后重试。

    ```javascript
    case 13007:
        showResult(['That operation cannot be done at this time. Please try again later.']);
        break;      
    ```

1. 将 `TODO7` 替换为下列代码。如果用户触发的操作未等到上一次调用完成就调用了 `getAccessTokenAsync`，就会发生错误 13008。

    ```javascript
    case 13008:
        showResult(['Please try that operation again after the current operation has finished.']);
        break;
    ```      

1. 将 `TODO8` 替换为以下代码。 如果加载项不支持强制许可，但调用 `getAccessTokenAsync` 时 `forceConsent` 选项设置为 `true`，错误 13009 发生。 通常情况下，如果发生这种情况，代码应自动重新运行 `getAccessTokenAsync`，同时将许可选项设置为 `false`。 不过，在某些情况下，调用将 `forceConsent` 设置为 `true` 的方法本身就是在自动响应调用将选项设置为 `false` 的方法时出现的错误。 此时，不得重试代码，而是应建议用户注销并重新登录。

    ```javascript
    case 13009:
        if (triedWithoutForceConsent) {
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
        } else {
            getDataWithToken({ forceConsent: false });
        }
        break;
    ```      
    
1. 将 `TODO9` 替换为以下代码。

    ```javascript
    default:
        logError(result);
        break;
    ```  

1. 在 `handleClientSideErrors` 方法下方，添加下列方法。此方法可处理加载项 Web 服务中发生的以下错误：无法执行代表流，或无法从 Microsoft Graph 获取数据。

    ```javascript
    function handleServerSideErrors(result) {
    
        // TODO10: Handle the case where AAD asks for an additional form of authentication.

        // TODO11: Handle the case where consent has not been granted, or has been revoked.

        // TODO12: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow

        // TODO13: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO14: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO15: Log all other server errors.
    }
    ```

1. 将 `TODO10` 替换为下列代码。关于此代码，请注意以下几点：

    * 一些 Azure Active Directory 配置要求用户，必须提供其他一个或多个身份验证因素，才能访问一些 Microsoft Graph 目标（例如 OneDrive），即使用户仅使用密码就能登录 Office，也不例外。在这种情况下，AAD 将发送包含错误 50076 的响应（具有 `Claims` 属性）。 
    * Office 主机应获取新令牌（使用 **Claims** 值作为 `authChallenge` 选项）。 这就指示 AAD 提示用户进行所有必需形式的身份验证。 

    ```javascript
    if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 50076){
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }
    ```

1. *在上一步添加的代码的最后一个右大括号正下方*，将 `TODO11` 替换为下列代码。关于此代码，请注意以下几点：

    * 错误 65001 表示未许可授予（或已撤消）一个或多个对 Microsoft Graph 的访问权限。 
    * 加载项应获取新令牌（`forceConsent` 选项设置为 `true`）。

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 65001){
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
        // getDataWithToken({ forceConsent: true });
    }
    ```

1. *在上一步添加的代码的最后一个右大括号正下方*，将 `TODO12` 替换为下列代码。关于此代码，请注意以下几点：

    * 错误 70011 表示已请求获取的范围（权限）无效。 加载项应报告此错误。
    * 代码使用 AAD 错误号记录其他任何错误。

    ```javascript
    else if (result.responseJSON.error.innerError
            && result.responseJSON.error.innerError.error_codes
            && result.responseJSON.error.innerError.error_codes[0] === 70011){
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }
    ```

1. *在上一步添加的代码的最后一个右大括号正下方*，将 `TODO13` 替换为下列代码。关于此代码，请注意以下几点：

    * 如果 `access_as_user` 范围（权限）不在访问令牌中，此令牌由加载项客户端发送到 AAD 以便在代表流中使用，那么在后续步骤中创建的服务器端代码将发送以 `... expected access_as_user` 结尾的消息。
    * 加载项应报告此错误。

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1){
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }
    ```

1. *在上一步添加的代码的最后一个右大括号正下方*，将 `TODO14` 替换为下列代码。关于此代码，请注意以下几点：

    * 不太可能将到期或无效令牌发送到 Microsoft Graph，但如果这种情况确实发生，在后续步骤中创建的服务器端代码将以字符串 `Microsoft Graph error` 结尾。
    * 在这种情况下，加载项应重置 `timesGetOneDriveFilesHasRun` 计数器和 `timesGetOneDriveFilesHasRun` 标志变量，再重新调用按钮处理程序方法，以从头开始执行整个身份验证流程。 但它只能执行此操作一次。 如果再次发生，它应只记录此错误。
    * 如果连续两次出现这种情况，代码会记录此错误。

    ```javascript
    else if (result.responseJSON.error.name
            && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }        
    }
    ```

1. *在上一步添加的代码的最后一个右大括号正下方*，将 `TODO15` 替换为以下代码。

    ```javascript
    else {
        logError(result);
    }
    ```

## <a name="code-the-server-side"></a>编写服务器端代码

有两个需要修改的服务器端文件。 
- src\auth.js 提供授权 helper 函数。它已具有在各种授权流中使用的泛型成员。我们需要为其添加可实现“代表”流的函数。
- src\server.js文件具有运行服务器和 Express 中间件所需的基本成员。我们需要为其添加服务于主页和 Web API 的函数，以获取 Microsoft Graph 数据。

### <a name="create-a-method-to-exchange-tokens"></a>创建交换令牌的方法

1. 打开 \src\auth.ts 文件。将下面的方法添加到 `AuthModule` 类。关于此代码，请注意以下几点：

    * 参数是对应用的访问令牌。在“代表”流中，它与 AAD 进行交换，以获取对资源的访问令牌。`jwt`
    * 虽然 scopes 参数有默认值，但在此示例中，它将被调用代码覆盖。
    * resource 是可选参数。不得在 STS 是 AAD V 2.0 终结点时使用它。V 2.0 终结点通过范围推断资源。如果在 HTTP 请求中发送资源，它会返回错误。 
    * 信息块中抛出异常*不会*导致立即向客户端发送“500 内部服务器错误”。`catch` server.js 文件中的调用代码会捕获此异常，并将它变成发送到客户端的错误消息。

        ```javascript
        private async exchangeForToken(jwt: string, scopes: string[] = ['openid'], resource?: string) {
            try {
                // TODO3: Construct the parameters that will be sent in the body of the 
                //        HTTP Request to the STS that starts the "on behalf of" flow.
                // TODO4: Send the request to the STS.
                // TODO5: Catch errors from the STS and relay them to the client.
                // TODO6: Process the response and persist the access token to resource.
            }
            catch (exception) {
                throw new UnauthorizedError('Unable to obtain an access token to the resource' 
                                            + JSON.stringify(exception), 
                                            exception);
            }
        }
        ```

2. 将 `TODO3` 替换为以下代码。关于此代码，请注意以下几点：
    * 支持“代表”流的 STS 需要 HTTP 请求正文中的某些属性/值对。此代码构造一个可成为请求正文的对象。 
    * 仅当资源传递到方法时，才将 resource 属性添加到正文。

        ```javascript
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

3. 将 `TODO4` 替换为以下代码，用于将 HTTP 请求发送到 STS 的令牌终结点。

    ```javascript
    const res = await fetch(`${this.stsDomain}/${this.tenant}/${this.tokenURLsegment}`, {
        method: 'POST',
        body: form(finalParams),
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    }); 
    ```

4. 将 `TODO5` 替换为以下代码。 请注意，抛出异常*不会*导致立即向客户端发送“500 内部服务器错误”。 server.js 文件中的调用代码会捕获此异常，并将它变成发送到客户端的错误消息。

    ```javascript
     if (res.status !== 200) {
        const exception = await res.json();
        throw exception;                
    } 
    ```

5. 将 `TODO6` 替换为以下代码。请注意，代码会返回并保留对资源的访问令牌及其到期时间。调用代码可以重用对资源的未到期访问令牌，避免了对 STS 执行不必要的调用。下一部分将介绍如何执行此操作。

    ```javascript  
    const json = await res.json();
    const resourceToken = json['access_token'];
    ServerStorage.persist('ResourceToken', resourceToken);
    const expiresIn = json['expires_in'];  // seconds until token expires.
    const resourceTokenExpiresAt = moment().add(expiresIn, 'seconds');
    ServerStorage.persist('ResourceTokenExpiresAt', resourceTokenExpiresAt);
    return resourceToken; 
    ```

6. 保存但不关闭文件。

### <a name="create-a-method-to-get-access-to-the-resource-using-the-on-behalf-of-flow"></a>使用“代表”流创建一个获取资源访问权限的方法

1. 还是在 src/auth.ts 中，将下面的方法添加到 `AuthModule` 类。关于此代码，请注意以下几点：

    * 上面关于 `exchangeForToken` 方法参数的注释也适用于此方法的参数。
    * 方法先检查对资源（尚未到期且不会在下一分钟到期）的访问令牌是否有永久性存储。仅在需要的情况下，它才会调用在上一部分中创建的 `exchangeForToken` 方法。

    ```javascript
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

    ```javascript
    app.get('/index.html', handler(async (req, res) => {
        return res.sendfile('index.html');
    })); 
    ```

3. 将以下方法添加到文件底部。此方法将处理任何 `onedriveitems` API 请求。
    ```javascript
    app.get('/api/onedriveitems', handler(async (req, res) => {
        // TODO7: Initialize the AuthModule object and validate the access token 
        //        that the client-side received from the Office host.
        // TODO8: Get a token to Microsoft Graph from either persistent storage 
        //        or the "on behalf of" flow.
        // TODO9: Use the token to get data from Microsoft Graph.
        // TODO10: Relay any errors from Microsoft Graph to the client.
        // TODO11: Send to the client only the data that it actually needs.
    })); 
    ```

4. 将 `TODO7` 替换为以下代码，用于验证从 Office 主机应用收到的访问令牌。`verifyJWT` 方法是在 src\auth.ts 文件中定义。它始终验证受众和颁发者。使用可选参数是为了指定，还希望它验证访问令牌中的范围是否为 `access_as_user`。这是用户和 Office 主机唯一需要的对加载项的权限，以便通过“代表”流获取对 Microsoft Graph 的访问令牌。 

    ```javascript
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    ```

    > [!NOTE]
    > 只能使用 `access_as_user` 范围授权 API 为 Office 加载项处理代表流。服务中的其他 API 应有自己的范围要求。这就限制了使用 Office 获得的令牌可以访问的内容。

5. 将 `TODO8` 替换为以下代码。关于此代码，请注意以下几点：

    * 调用中不包括 resource 参数，因为 `AuthModule` 对象 (`auth`) 是使用不支持 resource 属性的 AAD V2.0 终结点进行构造。`acquireTokenOnBehalfOf`
    * 调用的第二个参数指定了加载项获取 OneDrive 上用户文件和文件夹列表时所需的权限。 （之所以不需要 `profile` 权限是因为，只有当 Office 主机获取对加载项的访问令牌时，才需要此权限，用此令牌交换对 Microsoft Graph 的访问令牌时并不需要。）

    ```javascript
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);
    ```

6. 将 `TODO9` 替换为以下代码行。关于此代码，请注意以下几点：

    * MSGraphHelper 类是在 Src\msgraph helper.ts 中定义。 
    * 通过指定只需要 name 属性和前 3 项，可以最大限度地减少必须返回的数据。

    `const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");`

7. 将 `TODO10` 替换为以下代码。 请注意，此代码处理 Microsoft Graph 返回的“401 未授权”错误，此错误表示令牌到期或无效。 由于令牌暂留逻辑应该会阻止，因此这种情况不太可能会发生。 （请参阅上面的**使用“代表”流创建方法以获取对资源的访问权限**部分。）如果这种情况确实发生，此代码会将错误中继到客户端，并在错误名称中显示“Microsoft Graph 错误”。 （请参阅在之前步骤中在 program.js 文件内创建的 `handleClientSideErrors` 方法。）在后续步骤中添加到 ODataHelper.js 文件的代码有助于处理 Microsoft Graph 返回的错误。

    ```javascript
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    ```


1. 将 `TODO11` 替换为以下代码。请注意，Microsoft Graph 对每项返回某 OData 元数据和 **eTag** 属性，即使 `name` 是所请求的唯一属性，也不例外。代码仅向客户端发送项名称。

    ```javascript
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    ```

8. 保存并关闭文件。

### <a name="add-response-handling-to-the-odatahelper"></a>向 ODataHelper 添加响应处理

1. 打开文件 src\odata-helper.ts。 文件几乎已完成。 缺少的是，请求“结束”事件处理程序的回调主体。 将 `TODO` 替换为以下代码。 关于此代码，请注意以下几点：

    * OData 终结点返回的响应可能是错误（如 401）。如果终结点需要访问令牌，但令牌无效或到期，就会生成 401 错误。 不过，错误消息仍是*消息*，而不是 `https.get` 调用中的错误，因此不会触发 `https.get` 末尾的 `on('error', reject)` 代码行。 所以，代码区分成功 (200) 消息和错误消息，并向调用方发送 JSON 对象，其中包含请求获取的 OData 或错误消息。

    ```javascript
    var error;
    if (response.statusCode === 200) {
        // TODO1: Return the data to the caller and resolve the Promise.
    } else {
       // TODO2: Return an error object to the caller and resolve the Promise.
    }
    ```

1.  将 `TODO1` 替换为下列代码。请注意，此代码假设数据是以 JSON 形式返回。

    ```javascript
    let parsedBody = JSON.parse(body);
    resolve(parsedBody);
    ```

1.  将 `TODO2` 替换为下列代码。关于此代码，请注意以下几点：

    * OData 源返回的错误响应将始终包含 statusCode，通常是 statusMessage。 一些 OData 源还向主体添加错误属性，以提供更多信息，如内部或更具体的代码和消息。
    * Promise 对象已解析，未被拒绝。 Web 服务在服务器间调用 OData 终结点时，`https.get` 运行。 但这种调用出现的上下文是，客户端在 Web 服务中调用 Web API。 如果此“内部”请求被拒绝，客户端向 Web 服务发送的“外部”请求永不会完成。 此外，如果 `http.get` 的调用方需要将 OData 终结点返回的错误中继到客户端，必须解析具有自定义 `Error` 对象的请求。

    ```javascript
    error = new Error();
    error.code = response.statusCode;
    error.message = response.statusMessage;
    
    // The error body sometimes includes an empty space
    // before the first character, remove it or it causes an error.
    body = body.trim();
    error.bodyCode = JSON.parse(body).error.code;
    error.bodyMessage = JSON.parse(body).error.message;
    resolve(error);
    ```

1. 保存并关闭文件。

## <a name="deploy-the-add-in"></a>部署外接程序

现在，你需要让 Office 知道在哪里可以找到该外接程序。

1. 创建网络共享，或[将文件夹共享到网络](https://docs.microsoft.com/previous-versions/windows/it-pro/windows-server-2008-R2-and-2008/cc770880(v=ws.11))。

2. 将 Office-Add-in-NodeJS-SSO.xml 清单文件从项目根目录复制到共享文件夹。

3. 启动 PowerPoint 并打开文档。

4. 选择“文件”**** 选项卡，然后选择“选项”****。

5. 选择**信任中心**，然后选择**信任中心设置**按钮。

6. 选择“受信任的外接程序目录”****。

7. 在“目录 URL”**** 字段中，输入包含 Office-Add-in-NodeJS-SSO.xml 的文件夹共享的网络路径，然后选择“添加目录”****。

8. 选中“显示在菜单中”**** 复选框，然后选择“确定”****。

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

2. 在 PowerPoint 中的“开发工具”**** 选项卡上，选择“我的外接程序”****。

3. 选择“共享文件夹”**** 选项卡。

4. 选择“SSO NodeJS 示例”****，然后选择“确定”****。

5. “主页”**** 功能区上有一个名为“**SSO NodeJS**”的新组，包含标记为“显示外接程序”**** 的按钮和一个图标。 

## <a name="test-the-add-in"></a>测试加载项

1. 请确保 OneDrive 中有一些文件，以便可以验证结果。

2. 单击“显示加载项”**** 按钮，打开此加载项。

2. 此时，加载项打开并显示欢迎页。单击“从 OneDrive 获取我的文件”**** 按钮。

2. 如果你已登录 Office，则 OneDrive 上的文件和文件夹列表将显示在该按钮的下方。首次操作需要的时间可能会超过 15 秒。

3. 如过没有登录 Office，弹出窗口将打开并提示进行登录。完成登录后，文件和文件夹的列表将在几秒钟后显示。*请勿再次按下此按钮。*

> [!NOTE]
> 如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已在 PowerPoint 中更改过，也不例外。 在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。 为了防止发生这种情况，请务必先*关闭其他所有 Office 应用程序*，然后再按“从 OneDrive 获取我的文件”****。
