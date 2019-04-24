
> [!NOTE]
> 仅在开发加载项时，才需要执行此过程。 将生产外接程序部署到 AppSource 或外接程序目录时, 用户将单独信任它, 否则管理员会在安装时同意组织。

[注册外](../develop/register-sso-add-in-aad-v2.md)接程序*后*, 请执行此过程。 (如果刚完成该过程, 并且您的浏览器中已打开 **$ADD 名称 $** page 的**API 权限**选项卡, 则可以选择 "**授予管理员同意 [租户名称]** " 按钮, 然后在 "确认" 中选择 **"是"** 。出现。 请跳过此过程的其余部分。

1. 导航到 " [Azure 门户-应用程序注册](https://go.microsoft.com/fwlink/?linkid=2083908)" 页以查看您的应用注册。

1. 使用***管理员***凭据登录 Office 365 租户。 例如，MyName@contoso.onmicrosoft.com。

1. 选择显示名称 **$ADD 名称 $** 的应用程序。

1. 在 " **$ADD 名称 $** " 页上, 选择 " **API 权限**", 然后在 "**授予许可**" 部分下, 选择 "**授予管理员同意 [租户名称]** " 按钮。 对出现的确认选择 **"是"** 。

> [!NOTE]
> 如果您使用的是开发人员 O365 租户, 我们建议采用此过程作为最佳实践。 但是, 如果你愿意, 可以在开发过程中旁加载 SSO 加载项, 并提示用户提供许可表单。 有关详细信息, 请参阅[旁加载 on Windows](/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)和[旁加载 on Office Online](/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)。
