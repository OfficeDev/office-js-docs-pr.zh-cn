

1. <span data-ttu-id="82eef-101">导航到 [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)。</span><span class="sxs-lookup"><span data-stu-id="82eef-101">Navigate to [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)</span></span>

1. <span data-ttu-id="82eef-102">使用***管理员***凭据登录 Office 365 租户。</span><span class="sxs-lookup"><span data-stu-id="82eef-102">Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com</span></span> <span data-ttu-id="82eef-103">例如，MyName@contoso.onmicrosoft.com</span><span class="sxs-lookup"><span data-stu-id="82eef-103">For example, MyName@contoso.onmicrosoft.com</span></span>

1. <span data-ttu-id="82eef-104">单击**添加应用**。</span><span class="sxs-lookup"><span data-stu-id="82eef-104">Click **Add an app**.</span></span>

1. <span data-ttu-id="82eef-105">收到提示时，输入 **$ADD-IN-NAME$** 作为应用名称，然后按**创建应用程序**。</span><span class="sxs-lookup"><span data-stu-id="82eef-105">When prompted, use “Office-Add-in-ASPNET-SSO” as the app name, and then press Create application.</span></span>

1. <span data-ttu-id="82eef-p102">当应用的配置页面打开时，复制并保存**应用程序 ID**。将在后续过程中用到它。</span><span class="sxs-lookup"><span data-stu-id="82eef-p102">When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.</span></span>

    > [!NOTE]
    > <span data-ttu-id="82eef-p103">如果其他应用（如 PowerPoint、Word、Excel 等 Office 主机应用）寻求对应用的授权访问权限，此 ID 是“受众”值。反过来，如果它寻求对 Microsoft Graph 的授权访问权限，此 ID 同时也是应用的“客户端 ID”。</span><span class="sxs-lookup"><span data-stu-id="82eef-p103">This ID is the “audience” value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the “client ID” of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="82eef-p104">在“应用机密”**** 部分中，按“生成新密码”****。此时，弹出式对话框打开，并显示新密码（亦称为“应用密码”）。*立即复制密码，并将它与应用 ID 一起保存。* 将需要在后续过程中用到它。然后，关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="82eef-p104">In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an “app secret”) displayed. *Copy the password immediately and save it with the application ID.* You'll need it in a later procedure. Then close the dialog.</span></span>

1. <span data-ttu-id="82eef-115">在“平台”**** 部分中，单击“添加平台”****。</span><span class="sxs-lookup"><span data-stu-id="82eef-115">In the **Platforms** section, click **Add Platform**.</span></span>

1. <span data-ttu-id="82eef-116">在打开的对话框中，选择 **Web API**。</span><span class="sxs-lookup"><span data-stu-id="82eef-116">In the dialog that opens, select **Web API**.</span></span>

1. <span data-ttu-id="82eef-117">生成了“api://$App ID GUID$”窗体的一个 **应用程序 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="82eef-117">An **Application ID URI** has been generated of the form “api://$App ID GUID$”.</span></span> <span data-ttu-id="82eef-118">在双正斜杠和 GUID 之间插入 **$FQDN-WITHOUT-PROTOCOL$** （在结尾处有一个正斜杠“/”）。</span><span class="sxs-lookup"><span data-stu-id="82eef-118">Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="82eef-119">整个 ID 应该具有 `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$` 窗体；例如 `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`。</span><span class="sxs-lookup"><span data-stu-id="82eef-119">The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

    > [!NOTE]
    > <span data-ttu-id="82eef-120">如果你收到一条错误消息说该域已被他人拥有，但你拥有该域，请按照[快速入门：将自定义域名添加到 Azure Active Directory](https://docs.microsoft.com/azure/active-directory/add-custom-domain) 的步骤进行注册，然后重复此步骤。</span><span class="sxs-lookup"><span data-stu-id="82eef-120">If you get an error saying that the domain is already owned, but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](https://docs.microsoft.com/azure/active-directory/add-custom-domain) to register it, and then repeat this step.</span></span> <span data-ttu-id="82eef-121">（如果你未使用 Office 365 租户中的管理员凭据登录，也会发生此错误。</span><span class="sxs-lookup"><span data-stu-id="82eef-121">(This error can also occur if you are not signed in with credentials of an admin in the Office 365 tenancy.</span></span> <span data-ttu-id="82eef-122">参见步骤 2。</span><span class="sxs-lookup"><span data-stu-id="82eef-122">See step 2.</span></span> <span data-ttu-id="82eef-123">注销并使用管理员凭据再次登录，从步骤 3 开始重复此过程。）</span><span class="sxs-lookup"><span data-stu-id="82eef-123">Sign out and sign in again with admin credentials and repeat the process from step 3.)</span></span>

    > [!NOTE]
    > <span data-ttu-id="82eef-124">**应用程序 ID URI** 正下方的**范围**名称的域部分会自动改变以与之匹配，将 `/access_as_user` 追加在结尾处；例如：`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="82eef-124">The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match.</span></span>

1. <span data-ttu-id="82eef-125">在**预授权应用程序**部分中，确定要授权给加载项 Web 应用程序的应用程序。</span><span class="sxs-lookup"><span data-stu-id="82eef-125">In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="82eef-126">下面每个 ID 都需要进行预授权。</span><span class="sxs-lookup"><span data-stu-id="82eef-126">Each of the following IDs needs to be pre-authorized.</span></span> <span data-ttu-id="82eef-127">每次输入一个 ID，都会看到新的空文本框。</span><span class="sxs-lookup"><span data-stu-id="82eef-127">Each time you enter one, a new empty textbox appears.</span></span> <span data-ttu-id="82eef-128">（仅输入 GUID）。</span><span class="sxs-lookup"><span data-stu-id="82eef-128">(Enter only the GUID.)</span></span>
    * <span data-ttu-id="82eef-129">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="82eef-129">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="82eef-130">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="82eef-130">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span></span>
    * <span data-ttu-id="82eef-131">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span><span class="sxs-lookup"><span data-stu-id="82eef-131">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span></span>

1. <span data-ttu-id="82eef-132">打开每个“应用程序 ID”**** 旁边的“作用域”**** 下拉列表，并选中 `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user` 对应的框。</span><span class="sxs-lookup"><span data-stu-id="82eef-132">Open the **Scope** drop-down beside each **Application ID** and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span></span>

1. <span data-ttu-id="82eef-133">在“平台”**** 部分顶部附近，再次单击“添加平台”**** 并选择“Web”****。</span><span class="sxs-lookup"><span data-stu-id="82eef-133">Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.</span></span>

1. <span data-ttu-id="82eef-134">在“平台”**** 下的新“Web”**** 部分中，输入下列内容作为“重定向 URL”****：`https://$FQDN-WITHOUT-PROTOCOL$`。</span><span class="sxs-lookup"><span data-stu-id="82eef-134">In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span></span>

1. <span data-ttu-id="82eef-p108">向下滚动到 **Microsoft Graph 权限**部分的**委派的权限**子部分。使用**添加**按钮，打开**选择权限**对话框。</span><span class="sxs-lookup"><span data-stu-id="82eef-p108">Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.</span></span>

1. <span data-ttu-id="82eef-137">在对话框中，选中 `profile` 框以及你的加载项所需的任何其他 AAD 和 Microsoft Graph 权限。</span><span class="sxs-lookup"><span data-stu-id="82eef-137">In the dialog box, check the boxes for `profile` and any other AAD and Microsoft Graph permissions that your add-in needs.</span></span> <span data-ttu-id="82eef-138">示例如下：</span><span class="sxs-lookup"><span data-stu-id="82eef-138">The following are examples:</span></span>

    * <span data-ttu-id="82eef-139">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="82eef-139">Files.Read.All</span></span>
    * <span data-ttu-id="82eef-140">offline_access</span><span class="sxs-lookup"><span data-stu-id="82eef-140">offline_access</span></span>
    * <span data-ttu-id="82eef-141">openid</span><span class="sxs-lookup"><span data-stu-id="82eef-141">openid</span></span>
    * <span data-ttu-id="82eef-142">profile</span><span class="sxs-lookup"><span data-stu-id="82eef-142">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="82eef-143">`User.Read` 权限可能已默认列出。</span><span class="sxs-lookup"><span data-stu-id="82eef-143">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="82eef-144">最好不要请求不需要的权限，因此，如果你的加载项实际上并不需要，我们建议你取消选中此权限的复选框。</span><span class="sxs-lookup"><span data-stu-id="82eef-144">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission.</span></span>

1. <span data-ttu-id="82eef-145">单击对话框底部的**确定**。</span><span class="sxs-lookup"><span data-stu-id="82eef-145">At the bottom of the dialog, click **OK**.</span></span>

1. <span data-ttu-id="82eef-146">单击注册页底部的**保存**。</span><span class="sxs-lookup"><span data-stu-id="82eef-146">At the bottom of the registration page, click **Save**.</span></span>
