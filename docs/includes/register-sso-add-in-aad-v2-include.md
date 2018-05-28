

1. <span data-ttu-id="69489-101">??? [https://apps.dev.microsoft.com/](https://apps.dev.microsoft.com)?</span><span class="sxs-lookup"><span data-stu-id="69489-101">Navigate to [site\wwwroothttps://apps.dev.microsoft.com/[nameofyourazurefunction]](https://apps.dev.microsoft.com)</span></span>

1. <span data-ttu-id="69489-p101">????????? Office 365 ??????MyName@contoso.onmicrosoft.com</span><span class="sxs-lookup"><span data-stu-id="69489-p101">Sign-in with the admin credentials to your Office 365 tenancy. For example, MyName@contoso.onmicrosoft.com</span></span>

1. <span data-ttu-id="69489-104">?? **????**?</span><span class="sxs-lookup"><span data-stu-id="69489-104">Click **Add an app**.</span></span>

1. <span data-ttu-id="69489-105">???????? **$ ADD-IN-NAME $** ?????????? **??????**?</span><span class="sxs-lookup"><span data-stu-id="69489-105">When prompted, use ?Office-Add-in-ASPNET-SSO? as the app name, and then press Create application.</span></span>

1. <span data-ttu-id="69489-p102">????????????????? **???? ID**????????????</span><span class="sxs-lookup"><span data-stu-id="69489-p102">When the configuration page for the app opens, copy the **Application Id** and save it. You'll use it in a later procedure.</span></span>

    > [!NOTE]
    > <span data-ttu-id="69489-p103">???????? PowerPoint?Word?Excel ? Office ??????????????????? ID ????????????????? Microsoft Graph ????????? ID ??????????? ID??</span><span class="sxs-lookup"><span data-stu-id="69489-p103">This ID is the ?audience? value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application. It is also the ?client ID? of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="69489-p104">???????****????????????****???????????????????????????????*????????????? ID ?????*??????????????????????</span><span class="sxs-lookup"><span data-stu-id="69489-p104">In the **Application Secrets** section, press **Generate New Password**. A popup dialog opens with a new password (also called an ?app secret?) displayed. *Copy the password immediately and save it with the application ID.* You'll need it in a later procedure. Then close the dialog.</span></span>

1. <span data-ttu-id="69489-115">?????****????????????****?</span><span class="sxs-lookup"><span data-stu-id="69489-115">In the **Platforms** section, click **Add Platform**.</span></span>

1. <span data-ttu-id="69489-116">??????????? **Web API**?</span><span class="sxs-lookup"><span data-stu-id="69489-116">In the dialog that opens, select **Web API**.</span></span>

1. <span data-ttu-id="69489-117">????api?// $ App ID GUID $?????? **???? ID URI**?</span><span class="sxs-lookup"><span data-stu-id="69489-117">An **Application ID URI** has been generated of the form ?api://$App ID GUID$?.</span></span> <span data-ttu-id="69489-118">?? **$FQDN-WITHOUT-PROTOCOL$** ????????????/????????? GUID ???</span><span class="sxs-lookup"><span data-stu-id="69489-118">Insert the **$FQDN-WITHOUT-PROTOCOL$** (with a forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="69489-119">?? ID ???????? `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`??? `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`?</span><span class="sxs-lookup"><span data-stu-id="69489-119">The entire ID should have the form `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$`; for example `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

    > [!NOTE]
    > <span data-ttu-id="69489-120">??????????????????????????????????? [??????????????Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) ?????????????</span><span class="sxs-lookup"><span data-stu-id="69489-120">If you get an error saying that the domain is already owned, but you own it, follow the procedure at [Quickstart: Add a custom domain name to Azure Active Directory](https://docs.microsoft.com/en-us/azure/active-directory/add-custom-domain) to register it, and then repeat this step.</span></span>

    > [!NOTE]
    > <span data-ttu-id="69489-121">**???? ID URI** ???? **??** ?????????????????? `/access_as_user` ??????????`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`?</span><span class="sxs-lookup"><span data-stu-id="69489-121">The domain part of the **Scope** name just below the **Application ID URI** will automatically change to match.</span></span>

1. <span data-ttu-id="69489-122">? **???????** ???????????? Web ??????????</span><span class="sxs-lookup"><span data-stu-id="69489-122">In the **Pre-authorized applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="69489-123">???? ID ?????????</span><span class="sxs-lookup"><span data-stu-id="69489-123">Each of the following IDs needs to be pre-authorized.</span></span> <span data-ttu-id="69489-124">?????? ID????????????</span><span class="sxs-lookup"><span data-stu-id="69489-124">Each time you enter one, a new empty textbox appears.</span></span> <span data-ttu-id="69489-125">???? GUID??</span><span class="sxs-lookup"><span data-stu-id="69489-125">(Enter only the GUID.)</span></span>
    * <span data-ttu-id="69489-126">`d3590ed6-52b3-4102-aeff-aad2292ab01c` ?Microsoft Office?</span><span class="sxs-lookup"><span data-stu-id="69489-126">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="69489-127">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` ?Office Online?</span><span class="sxs-lookup"><span data-stu-id="69489-127">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office Online)</span></span>
    * <span data-ttu-id="69489-128">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` ?Office Online?</span><span class="sxs-lookup"><span data-stu-id="69489-128">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office Online)</span></span>

1. <span data-ttu-id="69489-129">????????? ID?****????????****???????? `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user` ?????</span><span class="sxs-lookup"><span data-stu-id="69489-129">Open the **Scope** drop-down beside each **Application ID** and check the box for `api://$FQDN-WITHOUT-PROTOCOL$/$App ID GUID$/access_as_user`.</span></span>

1. <span data-ttu-id="69489-130">?????****?????????????????****????Web?****?</span><span class="sxs-lookup"><span data-stu-id="69489-130">Near the top of the **Platforms** section, click **Add Platform** again and select **Web**.</span></span>

1. <span data-ttu-id="69489-131">?????****????Web?****???????????????? URL?****?`https://$FQDN-WITHOUT-PROTOCOL$`?</span><span class="sxs-lookup"><span data-stu-id="69489-131">In the new **Web** section under **Platforms**, enter the following as a **Redirect URL**: `https://$FQDN-WITHOUT-PROTOCOL$`.</span></span>

1. <span data-ttu-id="69489-p107">????? **Microsoft Graph ??** ???**?????** ????? **??** ???? **????**?</span><span class="sxs-lookup"><span data-stu-id="69489-p107">Scroll down to the **Microsoft Graph Permissions** section, the **Delegated Permissions** subsection. Use the **Add** button to open a **Select Permissions** dialog.</span></span>

1. <span data-ttu-id="69489-134">???????? `profile`  ??????????????? AAD ? Microsoft Graph ???</span><span class="sxs-lookup"><span data-stu-id="69489-134">In the dialog box, check the boxes for `profile` and any other AAD and Microsoft Graph permissions that your add-in needs.</span></span> <span data-ttu-id="69489-135">?????</span><span class="sxs-lookup"><span data-stu-id="69489-135">The following are examples:</span></span>

    * <span data-ttu-id="69489-136">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="69489-136">Files.Read.All</span></span>
    * <span data-ttu-id="69489-137">offline_access</span><span class="sxs-lookup"><span data-stu-id="69489-137">offline_access</span></span>
    * <span data-ttu-id="69489-138">openid</span><span class="sxs-lookup"><span data-stu-id="69489-138">openid</span></span>
    * <span data-ttu-id="69489-139">profile</span><span class="sxs-lookup"><span data-stu-id="69489-139">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="69489-140">`User.Read`  ??????????</span><span class="sxs-lookup"><span data-stu-id="69489-140">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="69489-141">????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="69489-141">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission.</span></span>

1. <span data-ttu-id="69489-142">???????? **??**?</span><span class="sxs-lookup"><span data-stu-id="69489-142">At the bottom of the dialog, click **OK**.</span></span>

1. <span data-ttu-id="69489-143">????????? **??**?</span><span class="sxs-lookup"><span data-stu-id="69489-143">At the bottom of the registration page, click **Save**.</span></span>
