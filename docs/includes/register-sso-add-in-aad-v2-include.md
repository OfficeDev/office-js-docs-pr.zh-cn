

1. 导航到 [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)。

1. 使用管理员凭据登录 Office 365 租户。例如，MyName@contoso.onmicrosoft.com

1. 单击 **“添加应用”**。

1. 收到提示时，输入 **$ ADD-IN-NAME $** 作为应用名称，然后按**创建应用程序**。

1. 当应用的配置页面打开时，复制并保存 **“应用 ID”**。将在后续过程中用到它。

    > [!NOTE]
    > 如果其他应用（如 PowerPoint、Word、Excel 等 Office 主机应用）寻求对应用的授权访问权限，此 ID 是“受众”值。反过来，如果它寻求对 Microsoft Graph 的授权访问权限，此 ID 同时也是应用的“客户端 ID”。

1. 在“应用机密”**** 部分中，按“生成新密码”****。此时，弹出式对话框打开，并显示新密码（亦称为“应用密码”）。*立即复制密码，并将它与应用 ID 一起保存。* 将需要在后续过程中用到它。然后，关闭对话框。

1. 在“平台”**** 部分中，单击“添加平台”****。

1. 在打开的对话框中，选择 **“Web API”**。

1. 生成了“api://$App ID GUID$”窗体的一个 **应用程序 ID URI**。 在双正斜杠和 GUID 之间插入 **$FQDN-WITHOUT-PROTOCOL$** （在结尾处有一个正斜杠“/”）。 整个 ID 应该具有 `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$` 窗体；例如 `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

    > [!NOTE]
    > 如果你收到一条错误消息说该域已被他人拥有，但你拥有该域，请按照[快速入门：将自定义域名添加到 Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) 的步骤进行注册，然后重复此步骤。

    > [!NOTE]
    > **应用程序 ID URI** 正下方的 **范围** 名称的域部分会自动改变以与之匹配，将 `/access_as_user` 追加在结尾处；例如：`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。

1. 在 **“预授权应用程序”** 部分中，确定要授权给加载项 Web 应用程序的应用程序。 下面每个 ID 都需要进行预授权。 每次输入一个 ID，都会看到新的空文本框。 （仅输入 GUID）。
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)

1. 打开每个“应用程序 ID”**** 旁边的“作用域”**** 下拉列表，并选中 `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user` 对应的框。

1. 在“平台”**** 部分顶部附近，再次单击“添加平台”**** 并选择“Web”****。

1. 在“平台”**** 下的新“Web”**** 部分中，输入下列内容作为“重定向 URL”****：`https://$FQDN-WITHOUT-PROTOCOL$`。

1. 向下滚动到 **“Microsoft Graph 权限”** 部分的 **“委派的权限”** 子部分。使用 **“添加”** 按钮，打开 **“选择权限”** 对话框。

1. 在对话框中，选中 `profile` 框以及你的加载项所需的任何其他 AAD 和 Microsoft Graph 权限。 示例如下：

    * Files.Read.All
    * offline_access
    * openid
    * 配置文件

    > [!NOTE]
    > `User.Read`权限可能已默认列出。 最好不要请求不需要的权限，因此，如果你的加载项实际上并不需要，我们建议你取消选中此权限的复选框。

1. 单击对话框底部的 **“确定”**。

1. 单击注册页底部的 **“保存”**。
