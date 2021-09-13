
> [!NOTE]
> 仅在开发加载项时，才需要执行此过程。 将生产加载项部署到 AppSource 或应用目录后，用户将单独信任它，或者管理员将在安装时同意组织。

在注册 *加载项* 后 [执行此过程](../develop/register-sso-add-in-aad-v2.md)。  (如果您刚刚完成该过程，并且 **"$ADD-IN-NAME$"** 页面的 **"API** 权限"选项卡在浏览器中打开，您可以选择"授予 [租户名称 **]** 管理员同意"按钮，然后选择"是"进行确认。 跳过此过程的其余部分。) 

1. 导航到 [Azure 门户 - 应用注册](https://go.microsoft.com/fwlink/?linkid=2083908) 页面以查看应用注册。

1. 使用管理员 ***凭据*** 登录您的Microsoft 365租户。 例如，MyName@contoso.onmicrosoft.com。

1. Select the app with 显示名称 **$ADD-IN-NAME$**.

1. On the **$ADD-IN-NAME$** page， select **API permissions** then， under the **Grant consent** section， choose the Grant admin consent **for [tenant name]** button. 对于 **出现的** 确认，选择"是"。

> [!NOTE]
> 如果使用的是开发人员 O365 租户，建议采用此过程作为最佳做法。 但是，如果您愿意，可以旁加载开发中的 SSO 外接程序，并提示用户提供同意表单。 有关详细信息，请参阅旁[加载和Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)[旁加载Office web 版。](../testing/sideload-office-add-ins-for-testing.md)
