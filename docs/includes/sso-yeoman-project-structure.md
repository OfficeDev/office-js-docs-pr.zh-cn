### <a name="configuration"></a>配置

以下文件指定外接程序的配置设置。

- 项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。

- **./.项目** 根目录中的 ENV 文件定义外接程序项目使用的常量。

### <a name="task-pane"></a>任务窗格 

以下文件定义加载项的任务窗格 UI 和功能。

- **./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。

- **./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。

- 在 JavaScript 项目中 **，./src/taskpane/taskpane.js** 文件包含初始化外接程序的代码。 在 TypeScript 项目中 **，./src/taskpane/taskpane.ts** 文件包含初始化外接程序的代码，还包含使用 Office JavaScript API 库将数据从 Microsoft Graph 添加到 Office 文档的代码。

### <a name="authentication"></a>身份验证

以下文件可加快 SSO 进程，并将数据写入Office文档。

- 在 JavaScript 项目中 **，./src/helpers/documentHelper.js** 文件包含使用 Office JavaScript API 库将数据从 Microsoft Graph 添加到 Office 文档的代码。 TypeScript 项目中没有此类文件;使用 Office JavaScript API 库将数据从 Microsoft Graph 添加到 Office 文档的代码存在于 **./src/taskpane/taskpane.ts** 中。

- **./src/helpers/fallbackauthdialog.html** 文件是无 UI 页面，用于加载用于回退身份验证策略的 JavaScript。

- **./src/helpers/fallbackauthdialog.js** 文件包含用于回退身份验证策略的 JavaScript，该策略使用 msal.js。

- **./src/helpers/fallbackauthhelper.js** 文件包含任务窗格 JavaScript，在不支持 SSO 身份验证的情况下调用回退身份验证策略。

- **./src/helpers/ssoauthhelper.js** 文件包含调用 SSO API、`getAccessToken` 的 JavaScript ，接收引导令牌，针对 Microsoft Graph 访问令牌启动引导令牌交换，同时调用 Microsoft Graph 以获得数据。