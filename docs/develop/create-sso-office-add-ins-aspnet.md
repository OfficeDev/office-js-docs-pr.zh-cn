---
title: 创建使用单一登录的 ASP.NET Office 加载项
description: null
ms.date: 01/23/2018
---

# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on-preview"></a>创建使用单一登录的 ASP.NET Office 加载项（预览）

如果用户已登录 Office，加载项可以使用相同的凭据，这样用户无需重新登录，即可访问多个应用。有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。

本文将引导你完成在使用 ASP.NET、OWIN 和适用于 .NET 的 Microsoft 验证库 (MSAL) 生成的外接程序中启用单一登录 (SSO) 的过程。

> [!NOTE]
> 有关与此类似的 Node.js 加载项文章，请参阅[创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)。

## <a name="prerequisites"></a>先决条件

* 最新版 Visual Studio 2017 Preview。

* Office 2016，版本 1708，内部版本 8424.nnnn 或更高版本（Office 365 订阅版本，有时称为“即点即用”）。可能需要成为 Office 预览体验成员才能获取此版本。有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/zh-cn/office-insider?tab=tab-1)。

## <a name="set-up-the-starter-project"></a>设置初学者项目

1. 在 [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso) 处克隆或下载存储库。

1. 打开 **Before** 文件夹，并打开 Visual Studio 中的 .sln 文件。这是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。

    > [!NOTE]
    > 在同一存储库中，还有此示例的已完成版本。这就像是完成本文中的过程后生成的加载项，不同之处在于已完成的项目有代码注释，但这对本文文本来说是多余的。若要使用已完成版本，只需打开 `sln` 文件，再按照本文中的说明操作即可，但要跳过**编写客户端代码**和**编写服务器端代码**这两部分。

1. 项目打开后，在 Visual Studio 中执行生成，这会安装 packages.config 文件中列出的包。此过程的完成耗时可能需要几秒到几分钟不等，具体视计算机本地包缓存中的包数量而定。

    > [!NOTE]
    > 将看到有关 Identity 命名空间的错误消息。 这是由于将在下一步中修复的配置问题间接造成。 重要的是，包已安装。

1. 目前，SSO 所需的 MSAL 库 (Microsoft.Identity.Client) 版本（`1.1.1-alpha0393` 版本）没有在标准 NuGet 目录中列出，因此也没有在 package.config 中列出，必须单独进行安装。 

   > 1. 在“工具”菜单上，依次转到“NuGet 包管理器” > “包管理器控制台”。 

   > 2. 在控制台中，运行以下命令。 即使 Internet 连接速度很快，也可能需要一分钟或更长时间才能完成。 完成后，应该会在控制台输出末尾处附近看到**已成功安装“Microsoft.Identity.Client 1.1.1-alpha0393”...**。

   >    `Install-Package Microsoft.Identity.Client -Version 1.1.1-alpha0393 -Source https://www.myget.org/F/aad-clients-nightly/api/v3/index.json`

   > 3. 在“解决方案资源管理器”中，右键单击“引用”。验证是否列出了“Microsoft.Identity.Client”。如果没有列出或它的条目上有警告图标，请先删除此条目，再使用“Visual Studio 添加引用”向导，添加对 **... \[Begin | Complete]\packages\Microsoft.Identity.Client.1.1.1-alpha0393\lib\net45\Microsoft.Identity.Client.dll** 处程序集的引用。

1. 重新生成此项目。

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a>向 Azure AD v2.0 终结点注册加载项

1. 转到 [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)。

1. 使用管理员凭据登录 Office 365 租户。例如，MyName@contoso.onmicrosoft.com

1. 单击“添加应用”。

1. 当出现提示时，为应用命名“Office-Add-in-ASPNET-SSO”，再按“创建应用”。

1. 当应用的配置页面打开时，复制并保存“应用 ID”。将在后续过程中用到它。

    > [!NOTE]
    > 如果其他应用（如 PowerPoint、Word、Excel 等 Office 主机应用）寻求对应用的授权访问权限，此 ID 是“受众”值。反过来，如果它寻求对 Microsoft Graph 的授权访问权限，此 ID 同时也是应用的“客户端 ID”。

1. 在“应用机密”部分中，按“生成新密码”。此时，弹出式对话框打开，并显示新密码（亦称为“应用密码”）。*立即复制密码，并将它与应用 ID 一起保存。*将需要在后续过程中用到它。然后，关闭对话框。

1. 在“平台”部分中，单击“添加平台”。

1. 在随即打开的对话框中，选择“Web API”。

1. 此时，生成了“api://{应用 ID GUID}”形式的“应用 ID URI”。在双斜杠和 GUID 之间插入字符串“localhost:44355/”。整个 ID 应为 `api://localhost:44355/{App ID GUID}`。 

    > [!NOTE]
    > “应用 ID URI”正下方的“范围”名称的域部分会自动更改为与之匹配。 它应为 `api://localhost:44355/{App ID GUID}/access_as_user`。

1. 在“预授权应用”部分中，确定要授权给加载项 Web 应用的应用。 下面每个 ID 都需要进行预授权。 每次输入一个 ID，都会看到新的空文本框。 （仅输入 GUID）。
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. 打开每个“应用程序 ID”旁边的“作用域”下拉列表，并选中 `api://localhost:44355/{App ID GUID}/access_as_user` 对应的框。

1. 在“平台”部分顶部附近，再次单击“添加平台”并选择“Web”。

1. 在“平台”下的新“Web”部分中，输入下列内容作为“重定向 URL”：`https://localhost:44355`。

    > [!NOTE]
    > 截至本文撰写之时，“Web API”平台有时会从“平台”部分中消失，特别是在添加“Web”平台和*保存注册页面*后刷新页面时。为了确保仍可以在注册期间选择“Web API”平台，请单击页面底部附近的“编辑应用清单”按钮。应该会看到清单的 **identifierUris** 属性中的 `api://localhost:44355/{App ID GUID}` 字符串。还有 **oauth2Permissions** 属性，它的 **value** 子属性的值为 `access_as_user`。

1. 向下滚动到“Microsoft Graph 权限”部分的“委派的权限”子部分。使用“添加”按钮，打开“选择权限”对话框。

1. 在对话框中，选中以下权限对应的框。 加载项本身真正需要的只是第一项权限，但服务器端代码使用的 MSAL 库需要有 `offline_access` 和 `openid`。 Office 主机必须有 `profile` 权限，才能获取对加载项 Web 应用程序的令牌。
    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > `User.Read` 权限可能已默认列出。根据最佳做法，最好不要请求授予不需要的权限，因此建议取消选中此权限对应的框。

1. 单击对话框底部的“确定”。

1. 单击注册页底部的“保存”。

## <a name="grant-admin-consent-to-the-add-in"></a>向加载项授予管理员许可

> [!NOTE]
> 仅在开发加载项时，才需要执行此过程。将生产加载项部署到 AppSource 或加载项目录时，用户需要单独信任它，否则管理员会在安装时授予组织许可。

1. 如果加载项未在 Visual Studio 中运行，请按 **F5** 运行它。必须在 IIS 中运行它，才能顺利完成此过程。

1. 在以下字符串中，将占位符“{application_ID}”替换为注册加载项时复制的应用 ID：`https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. 将生成的 URL 粘贴到浏览器地址栏，并转到此 URL。

1. 看到提示时，使用管理员凭据登录 Office 365 租户。

1. 然后系统提示你授予外接程序访问 Microsoft Graph 数据的权限。单击“接受”。

1. 然后，将浏览器窗口/选项卡重定向到注册外接程序时指定的**重定向 URL**；因此，外接程序的主页将在浏览器中打开。

2. 浏览器地址栏中将显示带 GUID 值的“tenant”查询参数。这是 Office 365 租赁 ID。请复制并保存此值。将在后续步骤中用到它。

3. 关闭窗口/选项卡。

1. 停止 Visual Studio 中的调试器。

## <a name="configure-the-add-in"></a>配置外接程序

1. 在下面的字符串中，将占位符“{tenant_ID}”替换为之前获得的 Office 365 租户 ID。如果出于任何原因，你以前没有获得 ID，请使用[查找 Office 365 租户 ID](https://support.office.com/zh-cn/article/Find-your-Office-365-tenant-ID-6891b561-a52d-4ade-9f39-b492285e2c9b) 中的一种方法来获取 ID。

    `https://login.microsoftonline.com/{tenant_ID}/v2.0`

2. 在 Visual Studio 中打开 web.config。你需要为 **appSettings** 部分中的某些键分配值。

3. 将在步骤 1 中构造的字符串用作名为“ida:Issuer”的键的值。请确保此值中没有空格。

4. 将下面的值分配给相应的键：

    |键|值|
    |:-----|:-----|
    |ida:ClientID|注册外接程序时获取的应用程序 ID。|
    |ida:Audience|注册外接程序时获取的应用程序 ID。|
    |ida:Password|注册加载项时获取的密码。|

   下面的示例展示了四个键的更改后效果。*请注意，ClientID 和 Audience 是相同的*。也可以将一个键同时用于这两种用途，但如果继续单独使用两个键，web.config 标记的可重用性将更高，因为它们并非始终相同。此外，单独使用两个键也可以强化以下概念：相对于 Office 主机而言，加载项是 OAuth 资源；相对于 Microsoft Graph 而言，它同时又是 OAuth 客户端。

    ```xml
    <add key=”ida:ClientID" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Audience" value="12345678-1234-1234-1234-123456789012" />
    <add key="ida:Password" value="rFfv17ezsoGw5XUc0CDBHiU" />
    <add key="ida:Issuer" value="https://login.microsoftonline.com/aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee/v2.0" />
    
    ```

   > [!NOTE]
   > **appSettings** 部分中的其他设置保持不变。

1. 保存并关闭文件。

1. 在外接程序项目中，打开外接程序清单文件“Office-Add-in-ASPNET-SSO.xml”。

1. 滚动到文件底部。

1. 结束 `</VersionOverrides>` 标记的正上方有以下标记：

    ```xml
    <WebApplicationInfo>
      <Id>{application_GUID here}</Id>
      <Resource>api://localhost:44355/{application_GUID here}</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. 将标记中的*两处*占位符“{此为 application_GUID }”均替换为在注册加载项时复制的应用 ID。（由于“{}”不属于 ID，因此请勿添加。）这与在 web.config 中对 ClientID 和 Audience 使用的 ID 相同。

    > [!NOTE]
    > * **Resource** 值是向注册的加载项添加 Web API 平台时设置的**应用 ID URI**。
    > * 仅在通过 AppSource 销售加载项时，才使用 **Scopes** 部分生成许可对话框。

1. 在 Visual Studio 中，打开“错误列表”的“警告”选项卡。 如果存在关于 `<WebApplicationInfo>` 不是 `<VersionOverrides>` 的有效子级的警告，则该 Visual Studio 2017 Preview 版本无法识别 SSO 标记。 作为解决方法，请对 Word、Excel 或 PowerPoint 外接程序执行以下操作。 （如果使用的是 Outlook 外接程序，请参阅下面的解决方法。）

   - **Word、Excel 和 Powerpoint 的解决方法**

        1. 在结束 `</VersionOverrides>` 标记正上方的清单中，注释掉 `<WebApplicationInfo>` 部分。

        2. 按 F5 启动调试会话。此操作会在下列文件夹（相比 Visual Studio，在“文件资源管理器”中访问此文件夹更方便）中创建清单副本：`Office-Add-in-ASP.NET-SSO\Complete\Office-Add-in-ASPNET-SSO\bin\Debug\OfficeAppManifests`

        3. 在清单副本中，删除 `<WebApplicationInfo>` 部分周围的注释语法。

        4. 保存此清单副本。

        5. 现在，必须阻止 Visual Studio 在用户下次按 F5 时重写此清单副本。右键单击“解决方案资源管理器”顶部的解决方案节点（而不是任何项目节点）。

        6. 选择上下文菜单中的“属性”，随后“解决方案属性页”对话框打开。

        7. 展开“配置属性”，并选择“配置”。

        8. 在 **Office-Add-in-ASPNET-SSO** 项目（*不是* **Office-Add-in-ASPNET-SSO-WebAPI** 项目）行中取消选择“生成”和“部署”。

        9. 按“确定”关闭对话框。

   - **Outlook 的解决方法**

        1. 在开发计算机上找到现有的 `MailAppVersionOverridesV1_1.xsd`。 它应位于 `./Xml/Schemas/{lcid}` 下的 Visual Studio 安装目录中。 例如，在英语（美国）的系统上进行 VS 2017 32 位的典型安装时，完整路径为 `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Xml\Schemas\1033`。

        2. 将现有文件重命名为 `MailAppVersionOverridesV1_1.old`。

        3. 将此修改后的文件版本复制到文件夹中：[修改后的 MailAppVersionOverrides 架构](https://github.com/OfficeDev/outlook-add-in-attachments-demo/blob/sso-conversion/manifest-schema-fix/MailAppVersionOverridesV1_1.xsd)

1. 在 Visual Studio 中保存并关闭该主清单文件。

## <a name="code-the-client-side"></a>编写客户端代码

1. 打开 **Scripts** 文件夹中的 Home.js 文件。其中已存在一些代码：
    * 针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。
    * `showResult` 方法，用于在任务窗格底部显示从 Microsoft Graph 返回的数据（或错误消息）。
    * `logErrors` 方法，用于记录最终用户不应看到的控制台错误。

1. 在向 `Office.initialize` 分配函数下方，添加下列代码。关于此代码，请注意以下几点：

    * 加载项中的错误处理有时会自动尝试使用一组不同的选项，重新获取访问令牌。 计数器变量 `timesGetOneDriveFilesHasRun` 和标志变量 `triedWithoutForceConsent` 用于确保用户不会重复循环失败的尝试来获取令牌。 
    * 虽然 `getDataWithToken` 方法是在下一步中创建，但请注意，它会将 `forceConsent` 选项设置为 `false`。有关详细信息，请参阅下一步。

    ```javascript
    var timesGetOneDriveFilesHasRun = 0;
    var triedWithoutForceConsent = false;

    function getOneDriveFiles() {
        timesGetOneDriveFilesHasRun++;
        triedWithoutForceConsent = true;
        getDataWithToken({ forceConsent: false });
    }   
    ```

1. 在 `getOneDriveFiles` 方法下方，添加下列代码。关于此代码，请注意以下几点：

    * `getAccessTokenAsync` 是 Office.js 中的新 API，可便于加载项要求 Office 主机应用（Excel、PowerPoint、Word 等）提供加载项访问令牌（对于已登录 Office 的用户）。反过来，Office 主机应用会向 Azure AD 2.0 终结点请求获取令牌。由于已在注册加载项时将 Office 主机预授权给加载项，因此 Azure AD 会发送访问令牌。
    * 如果用户未登录 Office，Office 主机会提示用户登录。
    * options 参数将 `forceConsent` 设置为 `false`，因此用户不会在每次使用加载项时都看到提示，要求其许可向 Office 主机授予对加载项的访问权限。 用户首次运行加载项时，`getAccessTokenAsync` 调用会失败，但在后续步骤中添加的错误处理逻辑会自动重新调用（`forceConsent` 选项设置为 `true`），并提示用户许可，但仅限首次运行。
    * `handleClientSideErrors` 方法将在后续步骤中创建。

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
    * `handleServerSideErrors` 方法将在后续步骤中创建。

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
    
        // TODO10: Parse the JSON response.

        // TODO11: Handle the case where AAD asks for an additional form of authentication.

        // TODO12: Handle the case where consent has not been granted, or has been revoked.

        // TODO13: Handle the case where an invalid scope (permission) was used in the on-behalf-of flow.

        // TODO14: Handle the case where the token that the add-in's client-side sends to it's 
        //         server-side is not valid because it is missing `access_as_user` scope (permission).

        // TODO15: Handle the case where the token sent to Microsoft Graph in the request for 
        //         data is expired or invalid.

        // TODO16: Log all other server errors.
    }
    ```

1. 将 `TODO10` 替换为以下代码。 请注意，对于加载项 Web 服务传递给加载项客户端的大多数 `4xx` 错误，响应中都有 **ExceptionMessage** 属性，其中包含 AADSTS（Azure Active Directory 安全令牌服务）错误号和其他数据。 不过，如果 AAD 向加载项 Web 服务发送消息，请求提供其他身份验证因素，那么消息包含特殊的 **Claims** 属性，用于指定（使用代码编号）需要的其他因素。 由于创建并向客户端发送 HTTP Response 的 ASP.NET API 并不知道此 **Claims** 属性，因此它们不会在 Response 对象中添加这个属性。 将在后续步骤中创建的服务器端代码负责处理这个问题，具体方法是手动向 Response 对象添加 **Claims** 值。 因为此值位于 **Message** 属性中，所以代码也需要解析相应属性。

    ```javascript
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    var message = JSON.parse(result.responseText).Message;
    }
    ```

1. 将 `TODO11` 替换为以下代码。关于此代码，请注意以下几点：

    * 如果 Microsoft Graph 要求进行其他形式的身份验证，错误 50076 发生。
    * Office 主机应获取新令牌（使用 **Claims** 值作为 `authChallenge` 选项）。 这就指示 AAD 提示用户进行所有必需形式的身份验证。 

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
        }
    }    
    ```

1. 将 `TODO12` 替换为以下代码。关于此代码，请注意以下几点：

    * 错误 65001 表示未许可授予（或已撤消）一个或多个对 Microsoft Graph 的访问权限。 
    * 加载项应获取新令牌（`forceConsent` 选项设置为 `true`）。

    ```javascript
    if (exceptionMessage.indexOf('AADSTS65001') !== -1) {
        showResult(['Please grant consent to this add-in to access your Microsoft Graph data.']);        
        /*
            THE FORCE CONSENT OPTION IS NOT AVAILABLE IN DURING PREVIEW. WHEN SSO FOR
            OFFICE ADD-INS IS RELEASED, REMOVE THE showResult LINE ABOVE AND UNCOMMENT
            THE FOLLOWING LINE.
        */
       // getDataWithToken({ forceConsent: true });
    }    
    ```

1. 将 `TODO13` 替换为下列代码。关于此代码，请注意以下几点：

    * 错误 70011 有多重含义。对于此加载项而言，最重要的含义是已请求获取的范围（权限）无效。因此，代码会检查是否有完整错误说明，而不仅仅是数字。
    * 加载项应报告此错误。

    ```javascript
     else if (exceptionMessage.indexOf("AADSTS70011: The provided value for the input parameter 'scope' is not valid.") !== -1) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }    
    ```

1. 将 `TODO14` 替换为以下代码。关于此代码，请注意以下几点：

    * 如果 `access_as_user` 范围（权限）不在访问令牌中，此令牌由加载项客户端发送到 AAD 以便在代表流中使用，那么在后续步骤中创建的服务器端代码将发送消息 `Missing access_as_user`。
    * 加载项应报告此错误。

    ```javascript
    else if (exceptionMessage.indexOf('Missing access_as_user.') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }    
    ```

1. 将 `TODO15` 替换为下列代码。关于此代码，请注意以下几点：

    * 要在服务器端代码中使用的标识库（Microsoft 身份验证库 (MSAL)）应确保没有向 Microsoft Graph 发送任何到期或无效令牌；但如果这种情况确实发生，从 Microsoft Graph 返回到加载项 Web 服务的错误包含代码 `InvalidAuthenticationToken`。在后续步骤中创建的服务器端代码会将此消息中继到加载项客户端。
    * 在这种情况下，加载项应重置计数器和标志变量，再重新调用按钮处理程序方法，以从头开始执行整个身份验证流程。

    ```javascript
    // If the token sent to MS Graph is expired or invalid, start the whole process over.
    else if (result.code === 'InvalidAuthenticationToken') {
        timesGetOneDriveFilesHasRun = 0;
        triedWithoutForceConsent = false;
        getOneDriveFiles();
    }    
    ```

1. 将 `TODO16` 替换为以下代码。

    ```javascript
    else {
        logError(result);
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

1. 右键单击“App_Start”文件夹，并依次选择“添加”>“类”。

1. 在“添加新项”对话框中，命名文件“Startup.Auth.cs”，再单击“添加”。

1. 将新文件中的命名空间名称缩短为 `Office_Add_in_ASPNET_SSO_WebAPI`。

1. 确保下列所有 `using` 语句都位于文件的顶部。

    ```csharp
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

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO3: Configure the validation settings
        // TODO4: Specify the type of authorization and the discovery endpoint
        // of the secure token service.
    }
    ```

1. 将 TODO3 替换为以下代码行。 关于此代码，请注意以下几点：

    * 代码指示 OWIN 以确保在来自 Office 主机（并通过客户端调用 `getData` 进行传递）的访问令牌中指定的受众和令牌颁发者必须与 web.config 中指定的值相匹配。
    * 将 `SaveSigninToken` 设置为 `true` 将导致 OWIN 从 Office 主机保存原始令牌。外接程序需要它来获取具有“代表”流的 Microsoft Graph 的访问令牌。
    * OWIN 中间件不验证范围。应包括 `access_as_user` 的访问令牌范围是在控制器中进行验证。

    ```csharp
    var tvps = new TokenValidationParameters
        {
            ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
            ValidIssuer = ConfigurationManager.AppSettings["ida:Issuer"],
            SaveSigninToken = true
        };
    ```

1. 将 TODO4 替换为下列代码。关于此代码，请注意以下几点：

    * 调用的是方法 `UseOAuthBearerAuthentication`，而不是更常见的 `UseWindowsAzureActiveDirectoryBearerAuthentication`，因为后者与 Azure AD V2 终结点不兼容。
    * 传递给该方法的发现 URL 是 OWIN 中间件获得用于获取所需密钥说明的位置，以验证从 Office 主机接收到的访问令牌上的签名。

    ```csharp
    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
        {
            AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider("https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"))
        });
    ```

1. 保存并关闭文件。

### <a name="create-the-apivalues-controller"></a>创建 /api/values 控制器

1. 打开文件 **Controllers\ValueController.cs**。

2. 请确保下列 `using` 语句位于文件顶部。

    ```csharp
    using Microsoft.Identity.Client;
    using System.IdentityModel.Tokens;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    using Office_Add_in_ASPNET_SSO_WebAPI.Models;
    ```

3. 在声明 `ValuesController` 的代码行的正上方，添加属性 `[Authorize]`。这可确保只要调用控制器方法时，加载项就会运行在上一过程中配置的授权过程。只有拥有对加载项的有效访问令牌，调用方才能调用控制器的方法。

    > [!NOTE]
    > 生产 ASP.NET MVC Web API 服务应在一个或多个自定义 [FilterAttribute](https://msdn.microsoft.com/zh-cn/library/system.web.http.filters(v=vs.108).aspx) 类中有代表流的自定义逻辑。 此说明性示例将逻辑放入主控制器中，以便能够轻松跟进授权和数据提取逻辑的整个流。 这也可以让示例与 [Azure 示例](https://github.com/Azure-Samples/)中的授权示例模式保持一致。    

4. 将下列方法添加到 `ValuesController`。 请注意，返回值是 `Task<HttpResponseMessage>`（而不是 `Task<IEnumerable<string>>`），这对于 `GET api/values` 方法而言更为常见。 这是将自定义授权逻辑放入控制器中造成不良影响：此逻辑中的一些错误条件要求将 HTTP Response 对象发送到加载项客户端。 

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO1: Validate the scopes of the access token.
    }
    ```

5. 将 `TODO1` 替换为以下代码行，以验证令牌中指定的范围是否包括 `access_as_user`。

    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (addinScopes.Contains("access_as_user"))
    {
        // TODO2: Assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.
        // TODO3: Get the access token for Microsoft Graph.
        // TODO4: Get the names of files and folders in OneDrive by using the Microsoft Graph API.
        // TODO5: Remove excess information from the data and send the data to the client.
    }
    return SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    ```

    > [!NOTE]
    > 只能使用 `access_as_user` 范围授权 API 为 Office 加载项处理代表流。服务中的其他 API 应有自己的范围要求。这就限制了使用 Office 获得的令牌可以访问的内容。

6. 将 `TODO2` 替换为以下代码。关于此代码，请注意以下几点：
    * 它将从 Office 主机收到的原始访问令牌转换为，传递给另一个方法的 `UserAssertion` 对象。
    * 外接程序不再扮演 Office 主机和用户需要访问的资源（或受众）的角色。现在它本身就是一个需要访问 Microsoft Graph 的客户端。`ConfidentialClientApplication` 是 MSAL“客户端上下文”对象。
    * `ConfidentialClientApplication` 构造函数的第三个参数是在“代表”流中实际不使用的重定向 URL，但使用正确的 URL 是一个很好的做法。第四和第五个参数可用于定义持久性存储，该存储使得外接程序能在不同的会话之间重用未过期的令牌。此示例不实现任何持久性存储。
    * MSAL 要求 `openid`、`offline_access` 作用域能够发挥作用，但如果代码过多地发出请求，则会抛出错误。 如果代码请求获取 `profile`，也会抛出错误，这真正仅适用于 Office 主机应用程序获取对加载项 Web 应用程序的令牌时。 因此，只会显式请求获取 `Files.Read.All`。

    ```csharp
    var bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext as BootstrapContext;
    UserAssertion userAssertion = new UserAssertion(bootstrapContext.Token);
    ClientCredential clientCred = new ClientCredential(ConfigurationManager.AppSettings["ida:Password"]);
    ConfidentialClientApplication cca =
                    new ConfidentialClientApplication(ConfigurationManager.AppSettings["ida:ClientID"],
                                                      "https://localhost:44355", clientCred, null, null);
    string[] graphScopes = { "Files.Read.All" };
    ```

7. 将 `TODO3` 替换为以下代码。关于此代码，请注意以下几点：

    * `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` 方法将首先查找内存中的 MSAL 缓存，获取匹配的访问令牌。仅当不存在任何令牌时，该方法才会通过 Azure AD V2 终结点启动“代表”流。
    * 如果 MS Graph 资源要求进行多重身份验证，但用户尚未提供，AAD 就会抛出包含 Claims 属性的异常。
    * 必须将 Claims 属性值传递到客户端，接着客户端会将它传递到 Office 主机，然后主机会将它添加到新令牌请求中。AAD 将提示用户进行所有必需形式的身份验证。
    * 任何不属于类型 `MsalServiceException` 的异常都是有意不捕获的，这样才能作为 `500 Server Error` 消息传播到客户端。

    ```csharp
    AuthenticationResult result = null;
    try
    {
        result = await cca.AcquireTokenOnBehalfOfAsync(graphScopes, userAssertion, "https://login.microsoftonline.com/common/oauth2/v2.0");
    }
    catch (MsalServiceException e)
    {        
        // TODO3a: Handle request for multi-factor authentication.
        // TODO3b: Handle lack of consent.
        // TODO3c: Handle invalid scope (permission).
        // TODO3d: Handle all other MsalServiceExceptions.
    }
    ```

8. 将 `TODO3a` 替换为下列代码。关于此代码，请注意以下几点：

    * 如果 MS Graph 资源要求进行多重身份验证，但用户尚未提供，AAD 就会返回包含错误 AADSTS50076 和 **Claims** 属性的“400 错误请求”。MSAL 会抛出包含此信息的 **MsalUiRequiredException**（继承自 **MsalServiceException**）。 
    * 必须将 **Claims** 属性值传递到客户端，接着客户端应将它传递到 Office 主机，然后主机会将它添加到新令牌请求中。AAD 会提示用户进行所有必需形式的身份验证。
    * 由于创建异常 HTTP Response 的 API 并不知道 **Claims** 属性，因此它们不会在 Response 对象中添加这个属性。 必须手动创建消息来添加它。 不过，自定义 **Message** 属性会阻止创建 **ExceptionMessage** 属性，因此向客户端发送错误 ID `AADSTS50076` 的唯一方法是，将它添加到自定义 **Message** 中。 客户端中的 JavaScript 需要发现响应是否包含 **Message** 或 **ExceptionMessage**，这样才能了解要读取的内容。
    * 自定义消息被格式化为 JSON，以便客户端 JavaScript 能够使用已知的 `JSON` 对象方法分析它。
    * `SendErrorToClient` 方法将在后续步骤中创建。 它的第二个参数是 **Exception** 对象。 在此示例中，代码传递 `null`，因为添加 **Exception** 对象会阻止在生成的 HTTP Response 中添加 **Message** 属性。

    ```csharp
    if (e.Message.StartsWith("AADSTS50076")) {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

9. 将 `TODO3b` 和 `TODO3c` 替换为下列代码。关于此代码，请注意以下几点：

    * 如果 AAD 调用包含至少一个范围（权限）未获用户和租户管理员的许可（或许可被撤消）， AAD 返回“400 错误请求”和错误 `AADSTS65001`。 MSAL 抛出包含此信息的 **MsalUiRequiredException**。 客户端应通过选项 `{ forceConsent: true }` 重新调用 `getAccessTokenAsync`。
    *  如果 AAD 调用包含至少一个 AAD 无法识别的范围，AAD 返回“400 错误请求”和错误 `AADSTS70011`。 MSAL 抛出包含此信息的 **MsalUiRequiredException**。 客户端应通知用户。
    *  包含完整说明，因为 70011 也会在其他情况下返回，只有在表示存在无效范围时，才需要在此加载项中处理它。 
    *  **MsalUiRequiredException** 对象传递给 `SendErrorToClient`。这样可确保 HTTP 响应中有包含错误消息的 **ExceptionMessage** 属性。
    *  由于没有自定义消息，因此会对第三个参数传递 `null`。

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001"))
    || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

10. 将 `TODO3d` 替换为以下代码。 请注意，代码会重新抛出异常，而不是在包含 **HttpStatusCode.Forbidden** (401) 的自定义 HTTP Response 内中继它。 结果就是，ASP.NET 发送自己的 HTTP Response，其中包含“500 服务器错误”状态。

    ```csharp
    else
    {
        throw e;
    }  
    ```

11. 将 `TODO4` 替换为以下代码。关于此代码，请注意以下几点：

    * `GraphApiHelper` 和 `ODataHelper` 类在 **Helpers** 文件夹的文件中定义。`OneDriveItem` 类在 **Models** 文件夹的一个文件中定义。 这些类的详细讨论内容与授权或 SSO 无关，因此不在本文的讨论范围内。
    * 通过只请求 Microsoft Graph 提供实际所需数据，可以提升性能，因此代码使用 ` $select` 查询参数来指定仅需要 name 属性，并使用 `$top` 参数来指定仅需要前 3 个文件夹或文件名。
    * 如果发送到 Microsoft Graph 的令牌无效，Microsoft Graph 会发送“401 未授权”错误和“InvalidAuthenticationToken”代码。 然后，ASP.NET 抛出 **RuntimeBinderException**。 这也是当令牌到期时发生的情况，尽管 MSAL 应阻止这种情况发生。 

    ```csharp
    var fullOneDriveItemsUrl = GraphApiHelper.GetOneDriveItemNamesUrl("?$select=name&$top=3");
    IEnumerable<OneDriveItem> filesResult;
    try
    {
        filesResult = await ODataHelper.GetItems<OneDriveItem>(fullOneDriveItemsUrl, result.AccessToken);
    }
    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException e)
    {
        return SendErrorToClient(HttpStatusCode.Unauthorized, e, null);                    
    }
    ```

12. 将 `TODO5` 替换为以下代码。关于此代码，请注意以下几点： 

    * 尽管上述代码仅请求获取 OneDrive 项的 *name* 属性，但 Microsoft Graph 始终包括 OneDrive 项的 *eTag* 属性。为减少发送到客户端的有效负载，下面的代码重新构造了仅包含项名称的结果。
    * 包含三个 OneDrive 文件和文件夹的列表作为“200 OK”HTTP Response 发送到客户端。

    ```csharp
    List<string> itemNames = new List<string>();
    foreach (OneDriveItem item in filesResult)
    {
        itemNames.Add(item.Name);
    }

    var requestMessage = new HttpRequestMessage();
    requestMessage.SetConfiguration(new HttpConfiguration());
    var response = requestMessage.CreateResponse<List<string>>(HttpStatusCode.OK, itemNames); 
    return response;
    ```

13. 在 Get 方法下方，添加下列方法。 关于此代码，请注意以下几点：  

    * 此方法将服务器端异常信息中继到客户端。 
    * 如果将原始异常传递到此方法，那么 HttpError 构造函数会在 **ExceptionMessage** 属性中添加来自 Exception 对象的信息。  
    * 如果对异常传递了 `null`，那么 HttpError 构造函数会在 **Message** 属性中添加 message 参数，且 **ExceptionMessage** 属性不存在。

    ```csharp
    private HttpResponseMessage SendErrorToClient(HttpStatusCode statusCode, Exception e, string message)
    {
        HttpError error;
        if (e != null)
        {
            error = new HttpError(e, true);
        }
        else
        {
            error = new HttpError(message);
        }
        var requestMessage = new HttpRequestMessage();
        var errorMessage = requestMessage.CreateErrorResponse(statusCode, error);
        return errorMessage;
    }        
    ```

## <a name="run-the-add-in"></a>运行加载项

1. 请确保 OneDrive 中有一些文件，以便可以验证结果。

1. 在 Visual Studio 中，按 F5。PowerPoint 将打开，“主页”功能区上会有一个“SSO ASP.NET”组。

1. 按此组中的“显示加载项”按钮，在任务窗格中查看此加载项的 UI。

1. 按“从 OneDrive 获取我的文件”按钮。如果尚未登录 Office，便会看到登录提示。
    
    > [!NOTE]
    > 如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已在 PowerPoint 中更改过，也不例外。 在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。 为了防止发生这种情况，请务必先*关闭其他所有 Office 应用*，再按“从 OneDrive 获取我的文件”。

1. 登录后，便会在按钮下方看到 OneDrive 文件和文件夹列表。此过程可能需要超过 15 秒才能完成，特别是首次使用时。
