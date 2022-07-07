---
title: 使用 SSO 获取已登录用户的标识
description: 调用 getAccessToken API 以获取 ID 令牌，其中包含名称、电子邮件和有关已登录用户的其他信息。
ms.date: 02/16/2022
localization_priority: Normal
ms.openlocfilehash: 5416c469a15d7eda9333f511c61e2cff1a901018
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660065"
---
# <a name="use-sso-to-get-the-identity-of-the-signed-in-user"></a>使用 SSO 获取已登录用户的标识

`getAccessToken`使用 API 获取一个访问令牌，其中包含登录到 Office 的当前用户的标识。 访问令牌也是一个 ID 令牌，因为它包含有关已登录用户的标识声明，例如其名称和电子邮件。 在调用自己的 Web 服务时，还可以使用 ID 令牌来标识用户。 若要调用 `getAccessToken` ，必须将 Office 加载项配置为将 SSO 与 Office 配合使用。

在本文中，你将创建一个获取 ID 令牌的 Office 加载项，并在任务窗格中显示用户的姓名、电子邮件和唯一 ID。

> [!NOTE]
> 使用 Office 和 API 的 `getAccessToken` SSO 在所有方案中都不起作用。 始终实现回退对话框，以便在 SSO 不可用时登录用户。 有关详细信息，请参阅 [Office 对话框 API 进行身份验证和授权](auth-with-office-dialog-api.md)。

## <a name="create-an-app-registration"></a>创建应用注册

若要将 SSO 与 Office 配合使用，需要在Azure 门户中创建应用注册，以便Microsoft 标识平台可以为 Office 加载项及其用户提供身份验证和授权服务。

1. 若要注册应用，请转到[Azure 门户 - 应用注册](https://go.microsoft.com/fwlink/?linkid=2083908)页。

1. 使用 **_管理员_** 凭据登录到 Microsoft 365 租户。 例如，MyName@contoso.onmicrosoft.com。

1. 选择“新注册”。 在“注册应用”页上，按如下方式设置值。

   - 将“名称”设置为“`Office-Add-in-SSO`”。
   - 将“**受支持的帐户类型**”设置为“**任何组织目录中的帐户和个人 Microsoft 帐户**”（例如，Skype、Xbox、Outlook.com）。
   - 将应用程序类型设置为 **Web** ，然后将 **重定向 URI** 设置为 `https://localhost:[port]/dialog.html`。 替换 `[port]` 为 Web 应用程序的正确端口号。 如果使用 yo office 创建了加载项，则端口号通常为 3000，并且在 package.json 文件中找到。 如果使用 Visual Studio 2019 创建了加载项，则会在 Web 项目的 **SSL URL** 属性中找到该端口。
   - 选择 **“注册”**。

1. 在 **Office-Add-in-SSO 页上** ，复制并保存应用程序 **(客户端) ID** 和 **目录 (租户) ID 的值**。 你将在后面的过程中使用它们。

   > [!NOTE]
   > 当其他应用程序 **（例如 Office 客户** 端应用程序 (（例如，PowerPoint、Word、Excel) ）寻求对应用程序的授权访问时，此应用程序 (客户端) ID 是“受众”值。 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

1. 选择“**管理**”下的“**身份验证**”。 在 **“隐式授予”** 部分中，启用 **Access 令牌** 和 **ID 令牌** 的复选框。

1. 在窗体顶部，选择“保存”。

1. 在“管理”下选择“公开 API”。 选择 **“设置** ”链接。 这将在窗体 `api://[app-id-guid]`中生成应用程序 ID URI，其中 `[app-id-guid]` 应用程序 **(客户端) ID**。

1. 在生成的 ID 中，插入 `localhost:[port]/` (记下追加到两个正斜杠和 GUID 之间的结束) 的正斜杠“/”。 替换 `[port]` 为 Web 应用程序的正确端口号。 如果使用 yo office 创建了加载项，则端口号通常为 3000，并且在 package.json 文件中找到。 如果使用 Visual Studio 2019 创建了加载项，则会在 Web 项目的 **SSL URL** 属性中找到该端口。
   完成后，整个 ID 应具有窗体 `api://localhost:[port]/[app-id-guid]`，例如 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

1. 选择“添加一个作用域”按钮。 在打开的面板中 **\<Scope\>**，输入`access_as_user`为名称。

1. 将“谁能同意?”设置为“管理员和用户”。

1. 填写用于配置管理员和用户同意提示的字段，其中包含适合 `access_as_user` 作用域的值，使 Office 客户端应用程序能够使用与当前用户具有相同权限的外接程序的 Web API。 建议：

   - **管理员许可显示名称**：Office 可以充当用户。
   - **管理员许可描述**：使 Office 能够借助与当前用户相同的权限调用加载项的 Web API。
   - **用户同意显示名称**：Office 可以充当你。
   - **用户同意说明**：允许 Office 使用与你拥有的权限相同的权限调用外接程序的 Web API。

1. 确保将“**状态**”设置为“**已启用**”。

1. 选择“**添加作用域**”。

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

1. 选择 **“添加客户端应用程序** ”按钮，然后在打开的面板中，将应用程序 (客户端) ID 设置 `[app-id-guid]` 为 `api://localhost:44355/[app-id-guid]/access_as_user`复选框。

1. 选择“添加应用程序”。

1. 选择“管理”下的“API 权限”，然后选择“添加权限”。 在打开的面板上，选择 **Microsoft Graph**，然后选择“委派权限”。

1. 使用“选择权限”搜索框来搜索加载项需要的权限。 搜索并选择 **配置文件** 权限。 `profile` Office 应用程序需要该权限才能获取加载项 Web 应用程序的令牌。

   - profile

   > [!NOTE]
   > `User.Read` 权限可能已默认列出。 根据最佳做法，最好不要请求授予不需要的权限，因此，如果加载项实际上不需要此权限，我们建议取消选中此权限对应的框。

1. 选择窗格下方选择“**添加权限**”。

1. 在同一页上，选择“ **授予管理员同意 \<tenant-name\>** ”按钮，然后为显示的确认选择 **“是** ”。

## <a name="create-the-office-add-in"></a>创建 Office 加载项

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. 启动 Visual Studio 2019 并选择 **创建新项目**。
1. 搜索并选择 **Excel Web 外接** 程序项目模板。 然后选择“**下一步**”。 注意：SSO 适用于任何 Office 应用程序，但本文适用于 Excel。
1. 输入项目名称，如 **sso-display-user-info** ，然后选择 **“创建**”。 可以将其他字段保留为默认值。
1. 在 **“选择外接程序类型** ”对话框中，选择 **“向 Excel 添加新功能**”，然后选择 **“完成**”。

项目已创建，将在解决方案中包含两个项目。

- **sso-display-user-info**：包含用于将外接程序旁加载到 Excel 的清单和详细信息。
- **sso-display-user-infoWeb**：托管加载项网页的 ASP.NET 项目。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

请确保已 [设置开发环境](../overview/set-up-your-dev-environment.md)。

1. 输入以下命令创建项目。

   ```command line
   yo office --projectType taskpane --name 'sso-display-user-info' --host excel --js true
   ```

该项目是在名为 **sso-display-user-info** 的新文件夹中创建的。

---

## <a name="configure-the-manifest"></a>配置清单

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. 在 **解决方案资源管理器** 打开 **sso-display-user-info > sso-display-user-infoManifest > sso-display-user-info.xml**

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. 在 Visual Studio 代码中，打开 **manifest.xml** 文件。

---

1. 清单底部附近是一个结束 `</Resources>` 元素。 在元素下方 `</Resources>` ，但在结束 `</VersionOverrides>` 元素之前插入以下 XML。 对于 Outlook 以外的 Office 应用程序，请将标记添加到该部分的 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 末尾。 对 Outlook，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 部分的末尾。

   ```xml
   <WebApplicationInfo>
       <Id>[application-id]</Id>
       <Resource>api://localhost:[port]/[application-id]</Resource>
       <Scopes>
           <Scope>openid</Scope>
           <Scope>user.read</Scope>
           <Scope>profile</Scope>
       </Scopes>
   </WebApplicationInfo>
   ```

1. 替换 `[port]` 为项目的正确端口号。 如果使用 yo office 创建了加载项，则端口号通常为 3000，并且在 package.json 文件中找到。 如果使用 Visual Studio 2019 创建了加载项，则会在 Web 项目的 **SSL URL** 属性中找到该端口。
1. 将这两个占位符替换 `[application-id]` 为应用注册中的实际应用程序 ID。
1. 保存文件。

插入的 XML 包含以下元素和信息。

- **\<WebApplicationInfo\>** - 以下元素的父级。
- **\<Id\>** - 外接程序的客户端 ID 这是注册加载项时获取的应用程序 ID。 请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项](register-sso-add-in-aad-v2.md)。
- **\<Resource\>** - 加载项的 URL。 这是在 AAD 中注册加载项时使用的相同 URI（包括 `api:` 协议）。 此 URI 的域部分必须与加载项清单部分的 URL 中 **\<Resources\>** 使用的域（包括任何子域）匹配，URI 必须以客户端 ID **\<Id\>** 结尾。
- **\<Scopes\>** - 一个或多个 **\<Scope\>** 元素的父元素。
- **\<Scope\>** - 指定加载项需要 AAD 的权限。 如果加载项不访问 Microsoft Graph，则始终需要`profile` 和 `openID` 权限，并且可能是唯一需要的权限。 如果确实如此，则还需要 **\<Scope\>** 具有所需的 Microsoft Graph 权限的元素;例如， `User.Read``Mail.Read`. 在代码中用于访问 Microsoft Graph 的库可能需要其他权限。 例如，用于 .NET 的 Microsoft 身份验证库 (MSAL) 需要 `offline_access` 权限。 有关详细信息，请参阅[向 Office 加载项中的 Microsoft Graph 授权](authorize-to-microsoft-graph.md)。

## <a name="add-the-jwt-decode-package"></a>添加 jwt-decode 包

可以调用 `getAccessToken` API 从 Office 获取 ID 令牌。 首先，可以添加 jwt-decode 包，以便更轻松地解码和查看 ID 令牌。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. 打开 Visual Studio 解决方案。
1. 在菜单上， **选择“工具> NuGet 包管理器>包管理器控制台**。
1. 在 **包管理器控制台** 中输入以下命令。

   `Install-Package jwt-decode -Projectname sso-display-user-infoWeb`

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. 从终端/控制台窗口转到加载项项目的根文件夹。
1. 输入以下命令

   `npm install jwt-decode`

---

## <a name="add-ui-to-the-task-pane"></a>将 UI 添加到任务窗格

我们需要修改任务窗格，以便它可以显示从 ID 令牌获取的用户信息。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. 打开Home.html文件。
1. 将以下脚本标记添加到 `<head>` 页面部分。 这将包括我们之前添加的 jwt-decode 包。

   ```html
   <script src="Scripts/jwt-decode-2.2.0.js" type="text/javascript"></script>
   ```

1. 将该 `<body>` 部分替换为以下 HTML。

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. 打开 **src/taskpane/taskpane.html** 文件。
1. 将该 `<body>` 部分替换为以下 HTML。

   ```html
   <body>
     <h1>Welcome</h1>
     <p>
       Sign in to Office, then choose the <b>Get ID Token</b> button to see your
       ID token information.
     </p>
     <button id="getIDToken">Get ID Token</button>
     <div>
       <span id="userInfo"></span>
     </div>
   </body>
   ```

---

## <a name="call-the-getaccesstoken-api"></a>调用 getAccessToken API

最后一步是通过调用 `getAccessToken`获取 ID 令牌。

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. 打开 **Home.js** 文件。
1. 使用以下代码替换文件的全部内容。

   ```javascript
   (function () {
     "use strict";

     // The initialize function must be run each time a new page is loaded.
     Office.initialize = function (reason) {
       $(document).ready(function () {
         $("#getIDToken").click(getIDToken);
       });
     };

     async function getIDToken() {
       try {
         let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
           allowSignInPrompt: true,
         });
         let userToken = jwt_decode(userTokenEncoded);
         document.getElementById("userInfo").innerHTML =
           "name: " +
           userToken.name +
           "<br>email: " +
           userToken.preferred_username +
           "<br>id: " +
           userToken.oid;
         console.log(userToken);
       } catch (error) {
         document.getElementById("userInfo").innerHTML =
           "An error occurred. <br>Name: " +
           error.name +
           "<br>Code: " +
           error.code +
           "<br>Message: " +
           error.message;
         console.log(error);
       }
     }
   })();
   ```

1. 保存文件。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

1. 打开 **src/taskpane/taskpane.js** 文件。
1. 使用以下代码替换文件的全部内容。

   ```javascript
   import jwt_decode from "jwt-decode";

   Office.onReady((info) => {
     if (info.host === Office.HostType.Excel) {
       document.getElementById("getIDToken").onclick = getIDToken;
     }
   });

   async function getIDToken() {
     try {
       let userTokenEncoded = await OfficeRuntime.auth.getAccessToken({
         allowSignInPrompt: true,
       });
       let userToken = jwt_decode(userTokenEncoded);
       document.getElementById("userInfo").innerHTML =
         "name: " +
         userToken.name +
         "<br>email: " +
         userToken.preferred_username +
         "<br>id: " +
         userToken.oid;
       console.log(userToken);
     } catch (error) {
       document.getElementById("userInfo").innerHTML =
         "An error occurred. <br>Name: " +
         error.name +
         "<br>Code: " +
         error.code +
         "<br>Message: " +
         error.message;
       console.log(error);
     }
   }
   ```

1. 保存文件。

---

## <a name="run-the-add-in"></a>运行加载项

# <a name="visual-studio-2019"></a>[Visual Studio 2019](#tab/vs2019)

1. 选择 **“调试>开始调试**，或按 **F5**。

# <a name="yo-office"></a>[yo office](#tab/yooffice)

从命令行运行 `npm start` 。

---

1. Excel 启动时，使用用于创建应用注册的同一租户帐户登录 Office。
1. 在 **“主页** ”功能区上，选择 **“显示任务窗格** ”以打开加载项。
1. 在加载项的任务窗格中，选择 **“获取 ID 令牌**”。

外接程序将显示已登录帐户的名称、电子邮件和 ID。

> [!NOTE]
> 如果遇到任何错误，请查看本文中有关应用注册的注册步骤。 设置应用注册时缺少详细信息是使用 SSO 时出现问题的常见原因。 如果仍然无法成功运行加载项，请参阅对 [单一登录 (SSO) 的错误消息进行故障排除 ](troubleshoot-sso-in-office-add-ins.md)。

## <a name="see-also"></a>另请参阅

[使用声明可靠地标识用户 (使用者和对象 ID) ](/azure/active-directory/develop/id-tokens#using-claims-to-reliably-identify-a-user-subject-and-object-id)

