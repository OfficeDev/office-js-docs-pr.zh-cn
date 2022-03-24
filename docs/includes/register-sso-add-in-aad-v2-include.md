## <a name="create-an-app-registration"></a>创建应用注册

在加载项 (注册) 在加载项和加载项之间建立信任Microsoft 标识平台。 信任是单向的：加载项信任Microsoft 标识平台信任，而不是信任其他方式。

1. 使用 ***admin** _ 凭据登录到 [Azure](https://portal.azure.com/) 门户，Microsoft 365租户。 例如，_*MyName@contoso.onmicrosoft.com**。
1. 在 **"管理**"下， **选择"应用注册""** > **新注册"**。 在“注册应用”页上，按如下方式设置值。

    * 将“名称”设置为“`<add-in-name>`”。
    * 将 **支持的帐户** 类型设置为任何 (目录中Azure AD **帐户 - 多租户) 和个人 Microsoft 帐户 (例如 Skype、Xbox)**。
    * 保留“重定向 URI”为空。
    * 选择“注册”。

1. 复制并保存 **Application (客户端** id) **Directory (id) 值**。 你将在后面的过程中使用它们。

    > [!NOTE]
    > 当其他应用程序（如 Office 客户端应用程序 (例如 PowerPoint、Word、Excel) ）寻求对该应用程序的授权访问权限时，此 ID 是"受众"值。 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

## <a name="add-a-client-secret"></a>添加客户端密码

有时称为 _应用程序密码_，客户端密码是一个字符串值，你的应用可以使用它来表示身份的证书。

1. 在 Azure 门户的应用 **注册中，** 选择应用程序。
1. Select **Certificates & secretsClient** >  **secretsNew** >  **client secret**.
1. 添加客户端密码的说明。
1. 选择密码的过期时间或指定自定义生存期。
    * 客户端密码生存期限制为 2 年 (24 个月) 或更少。 不能指定超过 24 个月的自定义生命周期。
    * Microsoft 建议你设置小于 12 个月的过期值。
1. 选择“**添加**”。
1. _记录密码值_ ，以用于客户端应用程序代码。 离开此页面 _后，此_ 密码值不会再显示。

## <a name="expose-a-web-api"></a>公开 Web API

1. 请确保你正在查看刚创建的应用注册。
1. 在 **"管理**"下 **，选择"公开 API**"，然后选择" **设置"** 链接。 这将打开一 **个设置应用程序 ID URI** 框，其生成的应用程序 ID 为 URI，格式为 `api://<application-id>`。 在 之前插入完全限定的域名 `<application-id>`。 整个 ID 应格式为 `api://<fully-qualified-domain-name>/<application-id>`;例如， `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

    > [!NOTE]
    > 如果收到一条错误，指出域已有所有者，但你拥有该域，请按照[快速入门： 将自定义域名添加到 Azure Active Directory](/azure/active-directory/add-custom-domain) 中的步骤进行操作来注册该域，然后重复此步骤。  (如果未使用租户中管理员的凭据登录，Microsoft 365此错误。 请参阅步骤 2 。 注销并使用管理员凭据再次登录，然后重复步骤 3 中的过程。）

## <a name="add-a-scope"></a>添加范围

1. 选择“添加一个作用域”按钮。 在打开的面板中，输入 `access_as_user` 作为“作用域名称”。

1. 将“谁能同意?”设置为“管理员和用户”。

1. `access_as_user`使用适用于范围的值填写用于配置管理员和用户同意提示的字段，使 Office 客户端应用程序能够使用与当前用户相同的权限使用外接程序的 Web API。 建议：

    * **管理员显示名称：** Office可以充当用户。
    * **管理员同意描述:** 使 Office 能够使用与当前用户相同的权限调用加载项的 web API。
    * **用户同意显示名称：** Office可以充当你。
    * **用户同意描述:** 启用 Office 以使用与你相同的权限调用加载项的 web API。

1. 确保将“状态”设置为“已启用”。

1. 选择“添加作用域”。

    > [!NOTE]
    > 显示在文本字段正下方的“作用域名称”的域部分应自动与上一步骤中设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。

1. 在“授权客户端应用程序”部分中，确定要授权给加载项 Web 应用程序的应用程序。 下面每个 ID 都需要进行预授权。
  
    * `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    * `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)
    * `93d53678-613d-4013-afc1-62e9e444a0a5`（Office 网页版）
    * `57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）
    * `08e18876-6177-487e-b8b5-cf950c1e598c`（Office 网页版）
    * `bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）

    > [!NOTE]
    > ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` 包括列出的所有其他 ID，可用于预授权所有 Office 主机终结点，以用于 Office 外接程序 SSO 流中的服务。

    对于每个 ID，请执行以下步骤。

      a. 选择 **添加客户端应用程序**。 在打开的面板中，将客户端 **ID 设置为** 相应的 GUID，并选中 的框 `api://<fully-qualified-domain-name>/<application-id>/access_as_user`。

      b. 选择“添加应用程序”。

## <a name="add-microsoft-graph-permissions"></a>添加 Microsoft Graph权限

1. 在 **"管理**"下 **，选择"身份验证**"，然后选择" **添加平台"**。

1. 在" **配置平台"** 窗格中，选择 **"Web**"，然后将" **重定向 URI"** 值设置为 `https://<fully-qualified-domain-name>`。

1. 选择“配置”****。

1. 在 **"管理**"下， **选择"API 权限**"，然后选择" **添加权限"**。 在打开的面板上，选择 **"Microsoft Graph**"，然后选择"**委派权限"**。

1. 使用“选择权限”搜索框来搜索加载项需要的权限。 示例如下。

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > `User.Read` 权限可能已默认列出。 最佳做法是仅请求所需的权限，因此，如果加载项实际上不需要此权限，建议取消选中此权限的框。

1. 在每个权限显示时，选择其复选框（请注意，选择每个权限后，它不会在列表中保持可见）。 选择外接程序所需的权限后，选择" **添加权限"** 按钮。
