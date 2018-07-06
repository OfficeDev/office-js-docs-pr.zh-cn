
> [!NOTE]
> 仅在开发外接程序时，才需要执行此过程。将生产加载项部署到 AppSource 或外接程序目录时，用户需要单独信任它，否则管理员会在安装时授予组织许可。

[注册了外接程序](../develop/register-sso-add-in-aad-v2.md)*之后*执行此过程。

1. 在以下字符串中，将占位符“{application_ID}”替换为注册外接程序时复制的应用程序 ID：  `https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. 将生成的 URL 粘贴到浏览器地址栏，并转到此 URL。

1. 看到提示时，使用管理员凭据登录 Office 365 租户。

1. 然后系统提示你授予外接程序访问 Microsoft Graph 数据的权限。单击**接受**。

1. 浏览器窗口/选项卡将重定向到注册外接程序时指定的**重定向网址**。 如果外接程序的 Web 应用程序正在运行，则外接程序的主页将在浏览器中打开；否则，会收到 404 错误。 但浏览器尝试打开主页的行为意味着已成功授予同意。

>[!NOTE]
>如果使用的是开发人员 O365 租户，我们建议将此过程作为最佳做法执行。 但是，如果你愿意，可以在开发中侧载 SSO 外接程序并使用同意表单提示用户。 有关更多信息，请参阅 [Windows 上的侧载](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)和 [Office Online 上的侧载](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)。

