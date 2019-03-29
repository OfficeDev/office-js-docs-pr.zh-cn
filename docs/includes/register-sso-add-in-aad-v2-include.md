

1. <span data-ttu-id="92442-101">导航到 [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)。</span><span class="sxs-lookup"><span data-stu-id="92442-101">Navigate to [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com).</span></span>

1. <span data-ttu-id="92442-102">使用***管理员***凭据登录 Office 365 租户。</span><span class="sxs-lookup"><span data-stu-id="92442-102">Sign-in with the admin credentials to your Office 365 tenancy.</span></span> <span data-ttu-id="92442-103">例如，MyName@contoso.onmicrosoft.com</span><span class="sxs-lookup"><span data-stu-id="92442-103">For example, MyName@contoso.onmicrosoft.com</span></span>

1. <span data-ttu-id="92442-104">单击“添加应用”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="92442-104">Click **Add an app**.</span></span>

1. <span data-ttu-id="92442-105">出现提示时，输入 **$ADD-IN-NAME$** 作为应用名称，然后按“**创建应用程序**”。</span><span class="sxs-lookup"><span data-stu-id="92442-105">When prompted, use “Office-Add-in-ASPNET-SSO” as the app name, and then press Create application.</span></span>

1. <span data-ttu-id="92442-p102">当应用的配置页面打开时，复制并保存“应用 ID”\*\*\*\*。将在后续过程中用到它。</span><span class="sxs-lookup"><span data-stu-id="92442-p102">When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.</span></span>

    > [!NOTE]
    > <span data-ttu-id="92442-p103">如果其他应用（如 PowerPoint、Word、Excel 等 Office 主机应用）寻求对应用的授权访问权限，此 ID 是“受众”值。反过来，如果它寻求对 Microsoft Graph 的授权访问权限，此 ID 同时也是应用的“客户端 ID”。</span><span class="sxs-lookup"><span data-stu-id="92442-p103">This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="92442-110">在“**应用程序机密**”部分，按“**生成新密码**”。</span><span class="sxs-lookup"><span data-stu-id="92442-110">In the **Application Secrets** section, press **Generate New Password**.</span></span> <span data-ttu-id="92442-111">此时，系统会打开弹出对话框，并显示新密码（亦称为“应用程序密码”）。</span><span class="sxs-lookup"><span data-stu-id="92442-111">A popup dialog opens with a new password (also called an “app secret”) displayed.</span></span> <span data-ttu-id="92442-112">*立即复制此密码，并将它与应用程序 ID 一起保存。*</span><span class="sxs-lookup"><span data-stu-id="92442-112">*Copy the password immediately and save it with the application ID.*</span></span> <span data-ttu-id="92442-113">在后面的过程中，将需要用到它。</span><span class="sxs-lookup"><span data-stu-id="92442-113">You'll use it in a later procedure.</span></span> <span data-ttu-id="92442-114">然后，关闭此对话框。</span><span class="sxs-lookup"><span data-stu-id="92442-114">Then click OK to close the dialog box.</span></span>

1. <span data-ttu-id="92442-115">在“平台”\*\*\*\* 部分中，单击“添加平台”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="92442-115">In the **Platforms** section, click **Add Platform**.</span></span>

1. <span data-ttu-id="92442-116">在打开的对话框中，选择“Web API”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="92442-116">In the dialog that opens, select **Web API**.</span></span>

1. <span data-ttu-id="92442-117">“**应用程序 ID URI**”已生成，格式为“api://{App ID GUID}”。</span><span class="sxs-lookup"><span data-stu-id="92442-117">An **Application ID URI** has been generated of the form “api://$App ID GUID$”.</span></span> <span data-ttu-id="92442-118">在双正斜框和 GUID 之间插入 **$FQDN-WITHOUT-PROTOCOL$**（在末尾附加一个正斜框“/”）。</span><span class="sxs-lookup"><span data-stu-id="92442-118">Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="92442-119">整个 ID 的格式应为 `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`；例如 `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`。</span><span class="sxs-lookup"><span data-stu-id="92442-119">The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

    > [!NOTE]
    > <span data-ttu-id="92442-120">如果收到一条错误，指出域已有所有者，但你拥有该域，请按照[快速入门： 将自定义域名添加到 Azure Active Directory](/azure/active-directory/add-custom-domain) 中的步骤进行操作来注册该域，然后重复此步骤。</span><span class="sxs-lookup"><span data-stu-id="92442-120">If you get an error saying that the domain is already owned, but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](/azure/active-directory/add-custom-domain) to register it, and then repeat this step.</span></span> <span data-ttu-id="92442-121">（如果你未在 Office 365 租户中使用管理员凭据登录，也会出现此错误。</span><span class="sxs-lookup"><span data-stu-id="92442-121">(This error can also occur if you are not signed in with credentials of an admin in the Office 365 tenancy.</span></span> <span data-ttu-id="92442-122">请参阅步骤 2 。</span><span class="sxs-lookup"><span data-stu-id="92442-122">See step 2.</span></span> <span data-ttu-id="92442-123">注销并使用管理员凭据再次登录，然后重复步骤 3 中的过程。）</span><span class="sxs-lookup"><span data-stu-id="92442-123">Sign out and sign in again with admin credentials and repeat the process from step 3.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="92442-124">“**应用 ID URI**”正下方的“**范围**”名称的域部分会自动更改为与之匹配，并在末尾附加 `/access_as_user`；例如 `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="92442-124">The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="92442-p107">在“预授权应用”\*\*\*\* 部分中，确定要授权给加载项 Web 应用的应用。 下面每个 ID 都需要进行预授权。 每次输入一个 ID，都会看到新的空文本框。 （仅输入 GUID）。</span><span class="sxs-lookup"><span data-stu-id="92442-p107">In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application. Each of the following IDs needs to be pre-authorized. Each time you enter one, a new empty textbox appears. (Enter only the GUID.)</span></span>
    * <span data-ttu-id="92442-129">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="92442-129">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="92442-130">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="92442-130">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span></span>
    * <span data-ttu-id="92442-131">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="92442-131">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span></span>

1. <span data-ttu-id="92442-132">打开每个“应用程序 ID”\*\*\*\* 旁边的“作用域”\*\*\*\* 下拉列表，并选中 `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user` 对应的框。</span><span class="sxs-lookup"><span data-stu-id="92442-132">Open the **Scope** drop-down beside each **Application ID** and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span></span>

1. <span data-ttu-id="92442-133">在“平台”\*\*\*\* 部分顶部附近，再次单击“添加平台”\*\*\*\* 并选择“Web”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="92442-133">Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.</span></span>

1. <span data-ttu-id="92442-134">在“平台”\*\*\*\* 下的新“Web”\*\*\*\* 部分中，输入下列内容作为“重定向 URL”\*\*\*\*：`https://$FQDN-WITHOUT-PROTOCOL$`。</span><span class="sxs-lookup"><span data-stu-id="92442-134">In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span></span>

1. <span data-ttu-id="92442-p108">向下滚动到“Microsoft Graph 权限”\*\*\*\* 部分的“委派的权限”\*\*\*\* 子部分。使用“添加”\*\*\*\* 按钮，打开“选择权限”\*\*\*\* 对话框。</span><span class="sxs-lookup"><span data-stu-id="92442-p108">Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.</span></span>

1. <span data-ttu-id="92442-137">在对话框中，选中加载项所需的 `profile` 及任何其他 AAD 和 Microsoft Graph 权限对应的框。</span><span class="sxs-lookup"><span data-stu-id="92442-137">In the dialog box, check the boxes for `profile` and any other AAD and Microsoft Graph permissions that your add-in needs.</span></span> <span data-ttu-id="92442-138">示例如下：</span><span class="sxs-lookup"><span data-stu-id="92442-138">The following are examples:</span></span>

    * <span data-ttu-id="92442-139">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="92442-139">Files.Read.All</span></span>
    * <span data-ttu-id="92442-140">offline_access</span><span class="sxs-lookup"><span data-stu-id="92442-140">offline_access</span></span>
    * <span data-ttu-id="92442-141">openid</span><span class="sxs-lookup"><span data-stu-id="92442-141">openid</span></span>
    * <span data-ttu-id="92442-142">profile</span><span class="sxs-lookup"><span data-stu-id="92442-142">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="92442-143">`User.Read` 权限可能已默认列出。</span><span class="sxs-lookup"><span data-stu-id="92442-143">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="92442-144">根据最佳做法，最好不要请求授予不需要的权限，因此，如果加载项实际上不需要此权限，我们建议取消选中此权限对应的框。</span><span class="sxs-lookup"><span data-stu-id="92442-144">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission.</span></span>

1. <span data-ttu-id="92442-145">单击对话框底部的“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="92442-145">At the bottom of the dialog, click **OK**.</span></span>

1. <span data-ttu-id="92442-146">单击注册页底部的“保存”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="92442-146">At the bottom of the registration page, click **Save**.</span></span>
