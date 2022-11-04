## <a name="register-the-add-in-with-microsoft-identity-platform"></a>使用 Microsoft 标识平台 注册加载项

需要在 Azure 中创建表示 Web 服务器的应用注册。 这将启用身份验证支持，以便可以在 JavaScript 中向客户端代码颁发适当的访问令牌。 此注册既支持客户端中的 SSO，也支持使用 Microsoft 身份验证库 (MSAL) 进行回退身份验证。


1. 使用 Microsoft 365 租户的 ***admin** _ 凭据登录到 [Azure 门户](https://portal.azure.com/)。 例如，_*MyName@contoso.onmicrosoft.com**。
1. 选择“**应用注册**”。 如果未看到图标，请在搜索栏中搜索“应用注册”。

    :::image type="content" source="../images/azure-portal-select-app-registration.png" alt-text="Azure 门户主页。":::

    将显示 **应用注册** 页。

1. 选择“新注册”。

    :::image type="content" source="../images/azure-portal-select-new-registration.png" alt-text="“应用注册”窗格中的新注册。":::

    此时会显示 **“注册应用程序”窗格** 。

1. 在 **“管理**”下，选择“**应用注册** > **”新建注册**”。 在“ **注册应用程序** ”窗格中，按如下所示设置值。

    * 将“名称”设置为“`<add-in-name>`”。
    * 将“ **支持的帐户类型** ”设置为 **“任何组织目录中的帐户 (任何 Azure AD 目录 - 多租户) 和个人 Microsoft 帐户 (，例如 Skype、Xbox)**。
    * 将 **重定向 URI** 设置为使用平台 `<redirect-platform>` ，将 URI 设置为 `<redirect-uri>`。

    :::image type="content" source="../images/azure-portal-register-an-application.png" alt-text="注册一个应用程序窗格，其中完成了名称和支持的帐户。":::

1. 选择“**注册**”。 显示一条消息，指出已创建应用程序注册。

    :::image type="content" source="../images/azure-portal-application-created-message.png" alt-text="指示已创建应用程序注册的消息。":::

1. 复制并保存 **应用程序 (客户端) ID** 和 **目录 (租户) ID** 的值。 你将在后面的过程中使用它们。

    :::image type="content" source="../images/azure-portal-copy-client-directory-ids.png" alt-text="显示客户端 ID 和目录 ID 的 Contoso 的应用注册窗格。":::

## <a name="add-a-client-secret"></a>添加客户端密码

有时称为 _应用程序密码_，客户端密码是一个字符串值，你的应用可以使用它来代替证书来标识自身。

1. 选择“ **证书&机密**”。 然后在“ **客户端机密** ”选项卡上，选择“ **新建客户端密码**”。

    :::image type="content" source="../images/azure-portal-create-new-client-secret.png" alt-text="“证书&机密”窗格。":::

    此时会显示 **“添加客户端机密** ”窗格。

1. 添加客户端密码的说明。
1. 选择机密的过期时间或指定自定义生存期。
    * 客户端机密生存期限制为两年 (24 个月) 或更短。 不能指定超过 24 个月的自定义生存期。
    * Microsoft 建议将过期值设置为小于 12 个月。

    :::image type="content" source="../images/azure-portal-client-secret-description.png" alt-text="添加客户端密码窗格，说明和过期已完成。":::

1. 选择“**添加**”。 将创建新的机密，该值将暂时显示。

> [!IMPORTANT]
> _记录要在_ 客户端应用程序代码中使用的机密值。 离开此窗格后 _，永远不会再次显示_ 此机密值。

## <a name="expose-a-web-api"></a>公开 Web API

1. 选择 **“公开 API**”。

    此时会显示 **“公开 API** ”窗格。

    :::image type="content" source="../images/azure-portal-expose-an-api.png" alt-text="应用注册的“公开 API”窗格。":::

1. 选择“ **设置** ”以生成应用程序 ID URI。

    :::image type="content" source="../images/azure-portal-set-api-uri.png" alt-text="应用注册的“公开 API”窗格中的“设置”按钮。":::

    将显示用于设置应用程序 ID URI 的部分，其中以 格式 `api://<app-id>`显示生成的应用程序 ID URI。

1. 将应用程序 ID URI 更新为 `api://localhost:44355/<app-id>`。

    :::image type="content" source="../images/azure-portal-app-id-uri-details.png" alt-text="将 localhost 端口设置为 44355 的“编辑应用 ID URI”窗格。":::

    * **应用程序 ID URI** 以格式 `api://<app-id>` 预填充应用 ID (GUID)。
    * 应用程序 ID URI 格式应为： `api://<fully-qualified-domain-name>/<app-id>`
    * 插入 介于 `fully-qualified-domain-name` `api://` 和 `<app-id>` (这是 GUID) 。 例如，`api://contoso.com/<app-id>`。
    * 如果使用 localhost，则格式应为 `api://localhost:<port>/<app-id>`。 例如，`api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。

    有关其他应用程序 ID URI 的详细信息，请参阅 [应用程序清单 identifierUris 属性](/azure/active-directory/develop/reference-app-manifest#identifieruris-attribute)。

    > [!NOTE]
    > 如果收到一条错误，指出域已有所有者，但你拥有该域，请按照[快速入门： 将自定义域名添加到 Azure Active Directory](/azure/active-directory/add-custom-domain) 中的步骤进行操作来注册该域，然后重复此步骤。  (如果未使用 Microsoft 365 租户中管理员的凭据登录，也可能会出现此错误。 请参阅步骤 2 。 注销并使用管理员凭据再次登录，然后重复步骤 3 中的过程。）

## <a name="add-a-scope"></a>添加范围

1. 选择“**添加作用域**”。

    :::image type="content" source="../images/azure-portal-add-a-scope.png" alt-text="选择“添加范围”按钮。":::

    此时会打开 **“添加范围** ”窗格。

1. 在 **“添加范围** ”窗格中，指定作用域的属性 。

    :::image type="content" source="../images/azure-portal-add-a-scope-details.png" alt-text="添加包含示例值的范围窗格。":::

    | 字段 | 说明 | 值 |
    |-------|-------------|---------|
    | **范围名称** | 范围的名称。 常见的范围命名约定是 `resource.operation.constraint`。 | 对于 SSO，必须将其设置为 `access_as_user`。 |
    | **谁可以同意** |  确定是否需要管理员同意，或者用户是否可以在未经管理员批准的情况下同意。 | 为了学习 SSO 和示例，建议将其设置为 **管理员和用户**。 <br><br>对于更高特权的权限，请选择“ **仅管理员** ”。|
    | **管理员同意显示名称** | 仅对管理员可见的范围用途的简短说明。 | `Read-only access to user files and profiles.` |
    | **管理员同意说明** | 由范围授予的权限的更详细说明，仅供管理员查看。 | `Allow Office to have read-only access to all user files and profiles. Office can call the app's web APIs as the current user.` |
    | **用户同意显示名称** | 范围用途的简短说明。 仅当将 **“谁可以同意** ”设置为 **“管理员和用户**”时，才会向用户显示。 | `Read-only access to your files and profile.` |
    | **用户同意说明** | 范围授予的权限的更详细说明。 仅当将 **“谁可以同意** ”设置为 **“管理员和用户**”时，才会向用户显示。 | `Allow Office to have read-only access to your files and user profile.` |

1. 将 **“状态** ”设置为 **“已启用**”，然后选择“ **添加范围**”。

    :::image type="content" source="../images/azure-portal-enable-state-add-scope-button.png" alt-text="将状态设置为“启用”，然后选择“添加范围”按钮。":::

    定义的新范围将显示在窗格中。

    :::image type="content" source="../images/azure-portal-scope-added-successfully.png" alt-text="“公开 API”窗格上显示的新范围。":::

    > [!NOTE]
    > 显示在文本字段正下方的“作用域名称”的域部分应自动与上一步骤中设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。

1. 选择 **“添加客户端应用程序**”

    :::image type="content" source="../images/azure-portal-add-a-client-application.png" alt-text="选择“添加客户端应用程序”。":::

    此时会显示 **“添加客户端应用程序** ”窗格。

1. 在 **“客户端 ID”** 中输入 `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`。 此值预授权所有 Microsoft Office 应用程序终结点。

    > [!NOTE]
    > 该 `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` ID 在以下所有平台上预授权 Office。 或者，如果出于任何原因想要在某些平台上拒绝对 Office 的授权，则可以输入以下 ID 的正确子集。 只需省略要从中扣留授权的平台的 ID 即可。 这些平台上加载项的用户将无法调用 Web API，但外接程序中的其他功能仍可正常工作。
    >
    > - `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)
    > - `93d53678-613d-4013-afc1-62e9e444a0a5`（Office 网页版）
    > - `bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）

1. 在 **“授权范围”中**，选中复选框 `api://localhost:44355/<app-id>/access_as_user` 。

1. 选择“添加应用程序”。

    :::image type="content" source="../images/azure-portal-add-application.png" alt-text="“添加客户端应用程序”窗格。":::

## <a name="add-microsoft-graph-permissions"></a>添加 Microsoft Graph 权限

1. 选择 **API 权限**。

    :::image type="content" source="../images/azure-portal-api-permissions.png" alt-text="“API 权限”窗格。":::

    “ **API 权限** ”窗格随即打开。

1. 选择“**添加权限**”。

    :::image type="content" source="../images/azure-portal-add-a-permission.png" alt-text="在“API 权限”窗格中添加权限。":::

    此时会打开 **“请求 API 权限** ”窗格。

1. 选择 **Microsoft Graph**。

    :::image type="content" source="../images/azure-portal-request-api-permissions-graph.png" alt-text="带有 Microsoft Graph 按钮的“请求 API 权限”窗格。":::

1. 选择“**委托的权限**”。

    :::image type="content" source="../images/azure-portal-request-api-permissions-delegated.png" alt-text="具有委托权限按钮的“请求 API 权限”窗格。":::

1. 在 **“选择权限”** 搜索框中，搜索外接程序所需的权限。 下面是示例中使用的典型值。

    * Files.Read
    * openid
    * profile

    > [!NOTE]
    > `User.Read` 权限可能已默认列出。 最好只请求所需的权限，因此，如果加载项实际上不需要此权限，建议取消选中此权限框。

1. 选中每个权限显示的复选框。 请注意，选择每个权限后，这些权限将不会在列表中保持可见。 选择加载项所需的权限后，选择“ **添加权限**”。

    :::image type="content" source="../images/azure-portal-request-api-permissions-add-permissions.png" alt-text="“请求 API 权限”窗格，其中选择了一些权限。":::

## <a name="configure-access-token-version"></a>配置访问令牌版本

必须定义应用可接受的访问令牌版本。 此配置是在 Azure Active Directory 应用程序清单中进行的。

### <a name="define-the-access-token-version"></a>定义访问令牌版本

如果你在任何组织目录中选择了帐户类型以外的帐户类型 **(任何 Azure AD 目录 - 多租户) 和个人 Microsoft 帐户 (，例如 Skype、Xbox)**，则访问令牌版本可能会更改。 使用以下步骤确保访问令牌版本适用于 Office SSO 用法。

1. 从左窗格中选择“**管理** > **清单** ”。

    :::image type="content" source="../images/azure-portal-manifest.png" alt-text="选择“Azure 清单”。":::

    此时会显示 Azure Active Directory 应用程序清单。

1. 输入 **2** 作为 `accessTokenAcceptedVersion` 属性的值。

    :::image type="content" source="../images/azure-portal-manifest-token-version.png" alt-text="接受访问令牌版本的值。":::

1. 选择“**保存**”

    浏览器上弹出一条消息，指出清单已成功更新。

    :::image type="content" source="../images/azure-portal-manifest-updated-message.png" alt-text="清单更新的消息。":::

祝贺你！ 你已完成应用注册，以便为 Office 加载项启用 SSO。
