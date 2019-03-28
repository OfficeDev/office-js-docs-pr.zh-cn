
> [!NOTE]
> 仅在开发加载项时，才需要执行此过程。将生产加载项部署到 AppSource 或加载项目录时，用户需要单独信任它，否则管理员会在安装时授予组织许可。

[注册外](../develop/register-sso-add-in-aad-v2.md)接程序*后*, 请执行此过程。

1. 在以下字符串中，将占位符“{application_ID}”替换为注册加载项时复制的应用 ID：`https://login.microsoftonline.com/common/adminconsent?client_id={application_ID}&state=12345`

1. 将生成的 URL 粘贴到浏览器地址栏，并转到此 URL。

1. 看到提示时，使用管理员凭据登录 Office 365 租户。

1. 然后系统提示你授予外接程序访问 Microsoft Graph 数据的权限。单击“接受”****。

1. 然后, 将浏览器窗口/选项卡重定向到注册外接程序时指定的**重定向 URL** 。 如果加载项的 web 应用程序正在运行, 则外接程序的主页将在浏览器中打开;否则, 将收到404错误。 但是, 浏览器试图打开主页这一事实意味着已成功授予同意。

>[!NOTE]
>如果您使用的是开发人员 O365 租户, 我们建议采用此过程作为最佳实践。 但是, 如果你愿意, 可以在开发过程中旁加载 SSO 加载项, 并提示用户提供许可表单。 有关详细信息, 请参阅[旁加载 on Windows](/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)和[旁加载 on Office Online](/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)。
