### <a name="configuration"></a>配置

以下文件指定外接程序的配置设置。

- 项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。

- **./。** 项目根目录中的环境文件定义外接程序项目使用的常量。

### <a name="task-pane"></a>任务窗格 

以下文件定义加载项的任务窗格 UI 和功能。

- **./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。

- **./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。

- 在 JavaScript 项目中， **/src/taskpane/taskpane.js**文件包含用于初始化加载项的代码。 在 TypeScript 项目中， **/src/taskpane/taskpane.ts**文件包含用于初始化外接程序的代码，以及使用 Office JavaScript 库将数据从 Microsoft Graph 添加到 Office 文档的代码。

### <a name="authentication"></a>身份验证

以下文件可帮助 SSO 进程并将数据写入 Office 文档。

- 在 JavaScript 项目中， **/src/helpers/documentHelper.js**文件包含使用 Office JavaScript 库将数据从 Microsoft Graph 添加到 Office 文档的代码。 在 TypeScript 项目中没有此类文件;使用 Office JavaScript 库将数据从 Microsoft Graph 添加到 Office 文档的代码在 **/src/taskpane/taskpane.ts**中存在。

- **./Src/helpers/fallbackauthdialog.html**文件是为回退身份验证策略加载 JavaScript 的无 UI 页面。

- **/Src/helpers/fallbackauthdialog.js**文件包含使用 msal 登录用户的回退身份验证策略的 JavaScript。

- 在不支持 SSO 身份验证的情况下， **/src/helpers/fallbackauthhelper.js**文件包含用于调用回退身份验证策略的任务窗格 JavaScript。

- **./src/helpers/ssoauthhelper.js** 文件包含调用 SSO API、`getAccessToken` 的 JavaScript ，接收引导令牌，针对 Microsoft Graph 访问令牌启动引导令牌交换，同时调用 Microsoft Graph 以获得数据。