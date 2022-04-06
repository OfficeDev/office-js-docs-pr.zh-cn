## <a name="create-an-app-registration"></a>创建应用注册

在加载项 (注册应用程序) 在加载项与Microsoft 标识平台之间建立信任关系。 信任是单向的：外接程序信任Microsoft 标识平台，而不是相反。

1. 使用 ***admin** _ 凭据登录 [到Microsoft 365](https://portal.azure.com/)租户的Azure 门户。 例如，_*MyName@contoso.onmicrosoft.com**。
1. 在 **"管理**"下，选择 **"** > 应用注册 **新注册**"。 在“注册应用”页上，按如下方式设置值。

    * 将“名称”设置为“`<add-in-name>`”。
    * 将 **支持的帐户类型** 设置为 **任何组织目录中的帐户 (任何Azure AD目录 - 多租户) 和个人 Microsoft 帐户 (，例如Skype、Xbox)**。
    * 保留“重定向 URI”为空。
    * 选择“注册”。

1. 复制并保存 **应用程序 (客户端) ID** 和 **目录 (租户) ID 的** 值。 你将在后面的过程中使用它们。

    > [!NOTE]
    > 当其他应用程序（例如Office客户端应用程序 (（例如，PowerPoint、Word、Excel) ）寻求对应用程序的授权访问时，此 ID 是"受众"值。 当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。

## <a name="add-a-client-secret"></a>添加客户端机密

有时称为 _应用程序密码_，客户端密码是应用可以用来代替证书来标识自身的字符串值。

1. 在Azure 门户中，**在应用注册** 中选择应用程序。
1. 选择 **"证书&机密** > **Client secretsNew** >  **客户端机密**。
1. 添加客户端机密的说明。
1. 选择机密过期或指定自定义生存期。
    * 客户端机密生存期限制为两年 (24 个月) 或更短。 不能指定超过 24 个月的自定义生存期。
    * Microsoft 建议将过期值设置为小于 12 个月。
1. 选择“**添加**”。
1. _记录要_ 在客户端应用程序代码中使用的机密值。 离开此页面后 _，不再显示_ 此机密值。

## <a name="expose-a-web-api"></a>公开 Web API

1. 请确保查看刚创建的应用注册。
1. 在 **"管理"** 下，选择 **"公开 API**"，然后选择 **"设置** "链接。 这将打开一个 **"设置应用 ID URI** "框，其中包含生成的应用程序 `api://<application-id>`ID URI。 在该域名之前 `<application-id>`插入完全限定的域名。 整个 ID 应具有窗体`api://<fully-qualified-domain-name>/<application-id>`，例如。 `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`

    > [!NOTE]
    > 如果收到一条错误，指出域已有所有者，但你拥有该域，请按照[快速入门： 将自定义域名添加到 Azure Active Directory](/azure/active-directory/add-custom-domain) 中的步骤进行操作来注册该域，然后重复此步骤。  (如果未使用Microsoft 365租户中的管理员凭据登录，也会发生此错误。 请参阅步骤 2 。 注销并使用管理员凭据再次登录，然后重复步骤 3 中的过程。）

## <a name="add-a-scope"></a>添加范围

1. 选择“添加一个作用域”按钮。 在打开的面板中，输入 `access_as_user` 作为“作用域名称”。

1. 将“谁能同意?”设置为“管理员和用户”。

1. 填写用于配置管理员和用户同意提示的字段，其中包含适合`access_as_user`作用域的值，使Office客户端应用程序能够使用与当前用户具有相同权限的外接程序的 Web API。 建议：

    * **管理员同意显示名称：** Office可以充当用户。
    * **管理员同意描述:** 使 Office 能够使用与当前用户相同的权限调用加载项的 web API。
    * **用户同意显示名称：** Office可以充当你。
    * **用户同意描述:** 启用 Office 以使用与你相同的权限调用加载项的 web API。

1. 确保将“状态”设置为“已启用”。

1. 选择“添加作用域”。

    > [!NOTE]
    > 显示在文本字段正下方的“作用域名称”的域部分应自动与上一步骤中设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。

1. 在 **"授权客户端应用程序**"部分中，输入以下 ID 以预先授权所有Microsoft Office应用程序终结点。

   - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (所有Microsoft Office应用程序终结点) 

    > [!NOTE]
    > ID `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` 在以下所有平台上预先授权Office。 或者，如果出于任何原因想要拒绝授权在某些平台上Office，则可以输入以下 ID 的适当子集。 只需保留要从中隐瞒授权的平台的 ID 即可。 这些平台上外接程序的用户将无法调用 Web API，但外接程序中的其他功能仍将有效。
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5`（Office 网页版）
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）

1. 选择 **添加客户端应用程序**。 在打开的面板中，将 **客户端 ID** 设置为相应的 GUID，然后选中该框 `api://<fully-qualified-domain-name>/<application-id>/access_as_user`。

1. 选择“添加应用程序”。

## <a name="add-microsoft-graph-permissions"></a>添加 Microsoft Graph 权限

1. 在 **"管理"** 下，选择 **"身份验证**"，然后选择 **"添加平台**"。

1. 在" **配置平台** "窗格中，选择 **"Web**"，然后将 **重定向 URI** 值设置为 `https://<fully-qualified-domain-name>`"

1. 选择“配置”****。

1. 在 **"管理"** 下，选择 **API 权** 限，然后选择 **"添加权限**"。 在打开的面板上，选择 **Microsoft Graph**，然后选择 **委托的权限**。

1. 使用“选择权限”搜索框来搜索加载项需要的权限。 示例如下。

    * Files.Read.All
    * offline_access
    * openid
    * profile

    > [!NOTE]
    > `User.Read` 权限可能已默认列出。 最好只请求所需的权限，因此，如果外接程序实际上不需要，建议取消选中此权限的框。

1. 在每个权限显示时，选择其复选框（请注意，选择每个权限后，它不会在列表中保持可见）。 选择外接程序所需的权限后，选择" **添加权限** "按钮。
