---
title: 创建使用单一登录的 ASP.NET Office 加载项
description: 如何创建 (或将 Office 后端的 ASP.NET Office) 外接程序转换为使用单一登录 (SSO) 的分步指南。
ms.date: 03/11/2021
localization_priority: Normal
ms.openlocfilehash: 36616e3388f9768c90a957ea19b47d4ec7e45de2
ms.sourcegitcommit: 4fa952f78be30d339ceda3bd957deb07056ca806
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/16/2021
ms.locfileid: "52961228"
---
# <a name="create-an-aspnet-office-add-in-that-uses-single-sign-on"></a><span data-ttu-id="ed918-103">创建使用单一登录的 ASP.NET Office 加载项</span><span class="sxs-lookup"><span data-stu-id="ed918-103">Create an ASP.NET Office Add-in that uses single sign-on</span></span>

<span data-ttu-id="ed918-104">如果用户已登录 Office，加载项可以使用相同的凭据，这样用户无需重新登录，即可访问多个应用程序。</span><span class="sxs-lookup"><span data-stu-id="ed918-104">When users are signed in to Office, your add-in can use the same credentials to permit users to access multiple applications without requiring them to sign in a second time.</span></span> <span data-ttu-id="ed918-105">有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="ed918-105">For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>
<span data-ttu-id="ed918-106">本文将引导你完成在内置加载项 (SSO) 启用单一登录 ASP.NET。</span><span class="sxs-lookup"><span data-stu-id="ed918-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with ASP.NET.</span></span>

> [!NOTE]
> <span data-ttu-id="ed918-107">有关与此类似的 Node.js 加载项文章，请参阅[创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)。</span><span class="sxs-lookup"><span data-stu-id="ed918-107">For a similar article about a Node.js-based add-in, see [Create a Node.js Office Add-in that uses single sign-on](create-sso-office-add-ins-nodejs.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ed918-108">先决条件</span><span class="sxs-lookup"><span data-stu-id="ed918-108">Prerequisites</span></span>

* <span data-ttu-id="ed918-109">Visual Studio 2019 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="ed918-109">Visual Studio 2019 or later.</span></span>

* [<span data-ttu-id="ed918-110">Office 开发人员工具</span><span class="sxs-lookup"><span data-stu-id="ed918-110">Office Developer Tools</span></span>](https://www.visualstudio.com/features/office-tools-vs.aspx)

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

* <span data-ttu-id="ed918-111">至少存储在你的订阅中OneDrive for Business一些文件和Microsoft 365文件夹。</span><span class="sxs-lookup"><span data-stu-id="ed918-111">At least a few files and folders stored on OneDrive for Business in your Microsoft 365 subscription.</span></span>

* <span data-ttu-id="ed918-112">一个 Microsoft Azure 订阅。</span><span class="sxs-lookup"><span data-stu-id="ed918-112">A Microsoft Azure subscription.</span></span> <span data-ttu-id="ed918-113">此加载项需要 Azure Active Directory (AD)。</span><span class="sxs-lookup"><span data-stu-id="ed918-113">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="ed918-114">Azure AD 为应用程序提供了用于进行身份验证和授权的标识服务。</span><span class="sxs-lookup"><span data-stu-id="ed918-114">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="ed918-115">你还可在 [Microsoft Azure](https://account.windowsazure.com/SignUp) 获得试用订阅。</span><span class="sxs-lookup"><span data-stu-id="ed918-115">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="ed918-116">设置初学者项目</span><span class="sxs-lookup"><span data-stu-id="ed918-116">Set up the starter project</span></span>

<span data-ttu-id="ed918-117">在 [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso) 处克隆或下载存储库。</span><span class="sxs-lookup"><span data-stu-id="ed918-117">Clone or download the repo at [Office Add-in ASPNET SSO](https://github.com/officedev/office-add-in-aspnet-sso).</span></span>

> [!NOTE]
> <span data-ttu-id="ed918-118">示例项目有两个版本：</span><span class="sxs-lookup"><span data-stu-id="ed918-118">There are two versions of the sample:</span></span>
>
> * <span data-ttu-id="ed918-p103">**Before** 文件夹是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。本文后续章节将引导你完成此过程。</span><span class="sxs-lookup"><span data-stu-id="ed918-p103">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
> * <span data-ttu-id="ed918-122">如果完成了本文中的过程，该示例的 **已完成** 版本会与所生成的加载项类似，只不过完成的项目具有对本文文本冗余的代码注释。</span><span class="sxs-lookup"><span data-stu-id="ed918-122">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="ed918-123">若要使用已完成的版本，请按照本文中的说明进行操作即可，但需要将“Before”替换为“Complete”，并跳过 **编写客户端代码** 和 **编写服务器端代码** 部分。</span><span class="sxs-lookup"><span data-stu-id="ed918-123">To use the completed version, just follow the instructions in this article, but replace "Before" with "Complete" and skip the sections **Code the client side** and **Code the server side**.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="ed918-124">向 Azure AD v2.0 终结点注册加载项。</span><span class="sxs-lookup"><span data-stu-id="ed918-124">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="ed918-125">导航到“Azure 门户 - 应用注册”[](https://go.microsoft.com/fwlink/?linkid=2083908)页面以注册你的应用。</span><span class="sxs-lookup"><span data-stu-id="ed918-125">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="ed918-126">使用管理员 ***凭据*** 登录到您的Microsoft 365租户。</span><span class="sxs-lookup"><span data-stu-id="ed918-126">Sign in with the ***admin*** credentials to your Microsoft 365 tenancy.</span></span> <span data-ttu-id="ed918-127">例如，MyName@contoso.onmicrosoft.com。</span><span class="sxs-lookup"><span data-stu-id="ed918-127">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="ed918-128">选择“新注册”。</span><span class="sxs-lookup"><span data-stu-id="ed918-128">Select **New registration**.</span></span> <span data-ttu-id="ed918-129">在“注册应用”页上，按如下方式设置值。</span><span class="sxs-lookup"><span data-stu-id="ed918-129">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="ed918-130">将“名称”设置为“`Office-Add-in-ASPNET-SSO`”。</span><span class="sxs-lookup"><span data-stu-id="ed918-130">Set **Name** to `Office-Add-in-ASPNET-SSO`.</span></span>
    * <span data-ttu-id="ed918-131">将“**受支持的帐户类型**”设置为“**任何组织目录中的帐户和个人 Microsoft 帐户(任何 Azure AD 目录 - 多租户)**”（例如，Skype、Xbox）。</span><span class="sxs-lookup"><span data-stu-id="ed918-131">Set **Supported account types** to **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.</span></span> <span data-ttu-id="ed918-132">（如果希望加载项仅可供注册该加载项的租户中的用户使用，则可以选择“**仅限此组织目录中的帐户...**”，但需要执行一些额外的设置步骤。</span><span class="sxs-lookup"><span data-stu-id="ed918-132">(If you want the add-in to be usable only by users in the tenancy where you are registering it, you can choose **Accounts in this organizational directory only ...** instead, but you will need to go through some additional setup steps.</span></span> <span data-ttu-id="ed918-133">请参阅下面的 **单租户设置**。）</span><span class="sxs-lookup"><span data-stu-id="ed918-133">See **Setup for single-tenant** below.)</span></span>
    * <span data-ttu-id="ed918-134">在“**重定向 URI**”部分，确保在下拉列表中选择“**Web**”，然后将 URI 设置为 ` https://localhost:44355/AzureADAuth/Authorize`。</span><span class="sxs-lookup"><span data-stu-id="ed918-134">In the **Redirect URI** section, ensure that **Web** is selected in the drop down and then set the URI to` https://localhost:44355/AzureADAuth/Authorize`.</span></span>
    * <span data-ttu-id="ed918-135">选择“**注册**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-135">Choose **Register**.</span></span>

1. <span data-ttu-id="ed918-136">在 **Office-Add-in-ASPNET-SSO** 页面上，复制并保存 **Application (client) ID** 和 **Directory (tenant) ID 的值**。</span><span class="sxs-lookup"><span data-stu-id="ed918-136">On the **Office-Add-in-ASPNET-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="ed918-137">你将在后面的过程中使用它们。</span><span class="sxs-lookup"><span data-stu-id="ed918-137">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ed918-138">当其他应用程序（如 Office 客户端应用程序 (例如 PowerPoint、Word、Excel) ）寻求应用程序的授权访问权限时，此应用程序客户端) ID 是"受众"值。 **(**</span><span class="sxs-lookup"><span data-stu-id="ed918-138">This **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="ed918-139">当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。</span><span class="sxs-lookup"><span data-stu-id="ed918-139">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="ed918-140">在“**管理**”下，选择“**证书和密码**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-140">Under **Manage**, select **Certificates & secrets**.</span></span> <span data-ttu-id="ed918-141">选择“**新客户端密码**”按钮。</span><span class="sxs-lookup"><span data-stu-id="ed918-141">Select the **New client secret** button.</span></span> <span data-ttu-id="ed918-142">输入“**描述**”的值，然后选择适当的“**到期**”选项，并选择“**添加**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-142">Enter a value for **Description**, then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="ed918-143">在继续操作前，*立即复制客户端密码值并使用应用程序 ID 保存它*，因为在后面的过程中，将需要用到它。</span><span class="sxs-lookup"><span data-stu-id="ed918-143">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="ed918-144">在“**管理**”下，选择“**公开 API**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-144">Under **Manage**, select **Expose an API**.</span></span> <span data-ttu-id="ed918-145">选择“**设置**”链接以在窗体“api://$App ID GUID$”中生成应用 ID URI，其中 $App ID GUID$ 是 **应用程序（客户端）ID**。</span><span class="sxs-lookup"><span data-stu-id="ed918-145">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="ed918-146">在 `//` 后面和 GUID 前面插入 `localhost:44355/`（请注意结尾附加的正斜杠“/”）。</span><span class="sxs-lookup"><span data-stu-id="ed918-146">Insert `localhost:44355/` (note the forward slash "/" appended to the end) after the `//` and before the GUID.</span></span> <span data-ttu-id="ed918-147">整个 ID 的格式应为 `api://localhost:44355/$App ID GUID$`；例如 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。</span><span class="sxs-lookup"><span data-stu-id="ed918-147">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span>

1. <span data-ttu-id="ed918-148">在对话框中选择“**保存**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-148">Select **Save** on the dialog.</span></span>

1. <span data-ttu-id="ed918-149">选择“**添加一个作用域**”按钮。</span><span class="sxs-lookup"><span data-stu-id="ed918-149">Select the **Add a scope** button.</span></span> <span data-ttu-id="ed918-150">在打开的面板中，输入 `access_as_user` 作为 **作用域** 名称。</span><span class="sxs-lookup"><span data-stu-id="ed918-150">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="ed918-151">将“谁能同意?”设置为“管理员和用户”。</span><span class="sxs-lookup"><span data-stu-id="ed918-151">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="ed918-152">填写用于配置管理员和用户同意提示的字段，并输入适用于范围的值，使 Office 客户端应用程序能够使用与当前用户相同的权限使用加载项的 Web API。 `access_as_user`</span><span class="sxs-lookup"><span data-stu-id="ed918-152">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office client application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="ed918-153">建议：</span><span class="sxs-lookup"><span data-stu-id="ed918-153">Suggestions:</span></span>

    * <span data-ttu-id="ed918-154">**管理员显示名称：Office** 可以充当用户。</span><span class="sxs-lookup"><span data-stu-id="ed918-154">**Admin consent display name**: Office can act as the user.</span></span>
    * <span data-ttu-id="ed918-155">**管理员许可描述**：使 Office 能够借助与当前用户相同的权限调用加载项的 Web API。</span><span class="sxs-lookup"><span data-stu-id="ed918-155">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    * <span data-ttu-id="ed918-156">**用户同意显示名称：Office** 可以充当你。</span><span class="sxs-lookup"><span data-stu-id="ed918-156">**User consent display name**: Office can act as you.</span></span>
    * <span data-ttu-id="ed918-157">**用户同意** 说明：Office以与您相同的权限调用外接程序的 Web API。</span><span class="sxs-lookup"><span data-stu-id="ed918-157">**User consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="ed918-158">确保将“**状态**”设置为“**已启用**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-158">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="ed918-159">选择“**添加作用域**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-159">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="ed918-160">显示在文本字段正下方的 **作用域** 名称的域部分应自动与你先前设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="ed918-160">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="ed918-161">在“授权客户端应用程序”部分中，确定要授权给加载项 Web 应用程序的应用程序。</span><span class="sxs-lookup"><span data-stu-id="ed918-161">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="ed918-162">下面每个 ID 都需要进行预授权。</span><span class="sxs-lookup"><span data-stu-id="ed918-162">Each of the following IDs needs to be pre-authorized.</span></span>

    * <span data-ttu-id="ed918-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="ed918-163">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    * <span data-ttu-id="ed918-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="ed918-164">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    * <span data-ttu-id="ed918-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="ed918-165">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    * <span data-ttu-id="ed918-166">`08e18876-6177-487e-b8b5-cf950c1e598c`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="ed918-166">`08e18876-6177-487e-b8b5-cf950c1e598c` (Office on the web)</span></span>
    * <span data-ttu-id="ed918-167">`bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Outlook 网页版）</span><span class="sxs-lookup"><span data-stu-id="ed918-167">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Outlook on the web)</span></span>

    <span data-ttu-id="ed918-168">对于每个 ID，执行以下步骤：</span><span class="sxs-lookup"><span data-stu-id="ed918-168">For each ID, take these steps:</span></span>

    <span data-ttu-id="ed918-169">a.</span><span class="sxs-lookup"><span data-stu-id="ed918-169">a.</span></span> <span data-ttu-id="ed918-170">选择“**添加客户端应用程序**”按钮，然后在打开的面板中，将“客户端 ID”设置为相应的 GUID 并勾选 `api://localhost:44355/$App ID GUID$/access_as_user` 框。</span><span class="sxs-lookup"><span data-stu-id="ed918-170">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="ed918-171">b.</span><span class="sxs-lookup"><span data-stu-id="ed918-171">b.</span></span> <span data-ttu-id="ed918-172">选择“添加应用程序”。</span><span class="sxs-lookup"><span data-stu-id="ed918-172">Select **Add application**.</span></span>

1. <span data-ttu-id="ed918-173">在“**管理**”下，选择“**API 权限**”，然后选择“**添加权限**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-173">Under **Manage**, select **API permissions** and then select **Add a permission**.</span></span> <span data-ttu-id="ed918-174">在打开的面板上，选择 **Microsoft Graph**，然后选择“委派权限”。</span><span class="sxs-lookup"><span data-stu-id="ed918-174">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="ed918-175">使用“选择权限”搜索框来搜索加载项需要的权限。</span><span class="sxs-lookup"><span data-stu-id="ed918-175">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="ed918-176">选择以下选项。</span><span class="sxs-lookup"><span data-stu-id="ed918-176">Select the following.</span></span> <span data-ttu-id="ed918-177">外接程序本身确实只需要第一项;但 `profile` 应用程序需要权限Office才能获取外接程序 Web 应用程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-177">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office application to get a token to your add-in web application.</span></span> <span data-ttu-id="ed918-178">（该加载项实际上仅需要 Files.Read.All 和 profile。</span><span class="sxs-lookup"><span data-stu-id="ed918-178">(Only Files.Read.All and profile are actually needed by the add-in.</span></span> <span data-ttu-id="ed918-179">但必须请求其他两个，因为 MSAL.NET 库需要它们。）</span><span class="sxs-lookup"><span data-stu-id="ed918-179">You must request the other two because the MSAL.NET library requires them.)</span></span>

    * <span data-ttu-id="ed918-180">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="ed918-180">Files.Read.All</span></span>
    * <span data-ttu-id="ed918-181">offline_access</span><span class="sxs-lookup"><span data-stu-id="ed918-181">offline_access</span></span>
    * <span data-ttu-id="ed918-182">openid</span><span class="sxs-lookup"><span data-stu-id="ed918-182">openid</span></span>
    * <span data-ttu-id="ed918-183">profile</span><span class="sxs-lookup"><span data-stu-id="ed918-183">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="ed918-184">`User.Read` 权限可能已默认列出。</span><span class="sxs-lookup"><span data-stu-id="ed918-184">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="ed918-185">根据最佳做法，最好不要请求授予不需要的权限，因此，如果加载项实际上不需要此权限，我们建议取消选中此权限对应的框。</span><span class="sxs-lookup"><span data-stu-id="ed918-185">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="ed918-186">选择所显示的每个权限的复选框。</span><span class="sxs-lookup"><span data-stu-id="ed918-186">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="ed918-187">选择加载项需要的权限后，选择面板底部的“**添加权限**”按钮。</span><span class="sxs-lookup"><span data-stu-id="ed918-187">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="ed918-188">在同一页面上，选择“**为[租户名称]授予管理员许可**”按钮，然后在显示的确认中选择“**接受**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-188">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Accept** for the confirmation that appears.</span></span>

    > [!NOTE]
    > <span data-ttu-id="ed918-189">选择“**为[租户名称]授予管理员许可** 后，可能会看到一条横幅消息，要求你在几分钟后再次尝试，以便能够构建许可提示。</span><span class="sxs-lookup"><span data-stu-id="ed918-189">After choosing **Grant admin consent for [tenant name]**, you may see a banner message asking you to try again in a few minutes so that the consent prompt can be constructed.</span></span> <span data-ttu-id="ed918-190">如果是这样，你可以开始下一部分，但不要忘记回到门户并 **_按此按钮_**！</span><span class="sxs-lookup"><span data-stu-id="ed918-190">If so, you can start work on the next section, **_but don't forget to come back to the portal and press this button_**!</span></span>

## <a name="configure-the-solution"></a><span data-ttu-id="ed918-191">配置解决方案</span><span class="sxs-lookup"><span data-stu-id="ed918-191">Configure the solution</span></span>

1. <span data-ttu-id="ed918-192">在 **Before** 文件夹的根部，打开 **Visual Studio** 中的解决方案 (.sln) 文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-192">In the root of the **Before** folder, open the solution (.sln) file in **Visual Studio**.</span></span> <span data-ttu-id="ed918-193">右键单击“**解决方案资源管理器**”最上面的节点（即“解决方案”节点，而非任何项目节点），然后选择“**设置启动项目**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-193">Right-click the top node in **Solution Explorer** (the Solution node, not either of the project nodes), and then select **Set startup projects**.</span></span>

1. <span data-ttu-id="ed918-194">在“**通用属性**”下，选择“**启动项目**”，然后选择“**多个启动项目**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-194">Under **Common Properties**, select **Startup Project**, and then **Multiple startup projects**.</span></span> <span data-ttu-id="ed918-195">确保两个项目的“**操作**”均设置为“**启动**”，并且以“...WebAPI”结尾的项目排在前面。</span><span class="sxs-lookup"><span data-stu-id="ed918-195">Ensure that the **Action** for both projects is set to **Start**, and that the project that ends in "...WebAPI" is listed first.</span></span> <span data-ttu-id="ed918-196">关闭该对话框。</span><span class="sxs-lookup"><span data-stu-id="ed918-196">Close the dialog.</span></span>

1. <span data-ttu-id="ed918-197">返回到"**解决方案资源管理器** (，选择"不要右键) "Office-Add-in-ASPNET-SSO-WebAPI"项目。 </span><span class="sxs-lookup"><span data-stu-id="ed918-197">Back in **Solution Explorer**, select (don't right-click) the **Office-Add-in-ASPNET-SSO-WebAPI** project.</span></span> <span data-ttu-id="ed918-198">随后将打开“**属性**”窗格。</span><span class="sxs-lookup"><span data-stu-id="ed918-198">The **Properties** pane opens.</span></span> <span data-ttu-id="ed918-199">确保“**已启用 SSL**”为“**True**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-199">Ensure that **SSL Enabled** is **True**.</span></span> <span data-ttu-id="ed918-200">验证“**SSL URL**”是否为 `http://localhost:44355/`。</span><span class="sxs-lookup"><span data-stu-id="ed918-200">Verify that the **SSL URL** is `http://localhost:44355/`.</span></span>

1. <span data-ttu-id="ed918-201">在“Web.config”中，使用先前复制的值。</span><span class="sxs-lookup"><span data-stu-id="ed918-201">In "Web.config", use the values that you copied in earlier.</span></span> <span data-ttu-id="ed918-202">将“**ida:ClientID**”和“**ida:Audience**”均设置为“**应用程序(客户端) ID**”，并将“**ida:Password**”设置为客户端密码。</span><span class="sxs-lookup"><span data-stu-id="ed918-202">Set both the **ida:ClientID** and the **ida:Audience** to your **Application (client) ID**, and set **ida:Password** to your client secret.</span></span> <span data-ttu-id="ed918-203">此外，将 **ida：Domain** 设置为 (末尾没有正斜杠 `http://localhost:44355` "/") 。</span><span class="sxs-lookup"><span data-stu-id="ed918-203">Also, set **ida:Domain** to `http://localhost:44355` (no forward slash "/" at the end).</span></span> 

    > [!NOTE]
    > <span data-ttu-id="ed918-204">当其他应用程序（如 Office 客户端应用程序 (例如 PowerPoint、Word、Excel) ）寻求应用程序的授权访问权限时，Application (客户端) **ID** 是"受众"值。</span><span class="sxs-lookup"><span data-stu-id="ed918-204">The **Application (client) ID** is the "audience" value when other applications, such as the Office client application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="ed918-205">当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。</span><span class="sxs-lookup"><span data-stu-id="ed918-205">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="ed918-206">如果在注册该加载项时，“**受支持的帐户类型**”未选择“仅限此组织目录中的帐户”，请保存并关闭 web.config。否则，请保存，但将其保持打开状态。</span><span class="sxs-lookup"><span data-stu-id="ed918-206">If you didn't choose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, save and close the web.config. Otherwise, save but leave it open.</span></span>

1. <span data-ttu-id="ed918-207">仍在"解决方案资源管理器"中，选择 **Office-Add-in-ASPNET-SSO** 项目，打开外接程序清单文件"Office-Add-in-ASPNET-SSO.xml"，然后滚动到文件底部。</span><span class="sxs-lookup"><span data-stu-id="ed918-207">Still in **Solution Explorer**, choose the **Office-Add-in-ASPNET-SSO** project and open the add-in manifest file “Office-Add-in-ASPNET-SSO.xml” and then scroll to the bottom of the file.</span></span> <span data-ttu-id="ed918-208">在结尾的 `</VersionOverrides>` 标记的正上方有以下标记：</span><span class="sxs-lookup"><span data-stu-id="ed918-208">Just above the end `</VersionOverrides>` tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>offline_access</Scope>
          <Scope>openid</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="ed918-209">将标记中的 *两处* 占位符“$application_GUID here$”均替换为在注册加载项时复制的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="ed918-209">Replace the placeholder “$application_GUID here$” *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="ed918-210">由于 ID 并不包含“$”符号，因此请勿添加它们。</span><span class="sxs-lookup"><span data-stu-id="ed918-210">The "$" signs are not part of the ID, so do not include them.</span></span> <span data-ttu-id="ed918-211">这与在 web.config 中对 ClientID 和 Audience 所使用的 ID 相同。</span><span class="sxs-lookup"><span data-stu-id="ed918-211">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

  > [!NOTE]
  > <span data-ttu-id="ed918-212">**资源** 值是注册加载项时设置的 **应用程序 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="ed918-212">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="ed918-213">仅在通过 AppSource 销售加载项时，才使用 **作用域** 部分生成许可对话框。</span><span class="sxs-lookup"><span data-stu-id="ed918-213">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

1. <span data-ttu-id="ed918-214">保存并关闭此文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-214">Save and close the file.</span></span>

### <a name="setup-for-single-tenant"></a><span data-ttu-id="ed918-215">单租户设置</span><span class="sxs-lookup"><span data-stu-id="ed918-215">Setup for single-tenant</span></span>

<span data-ttu-id="ed918-216">如果在注册该加载项时，“**受支持的帐户类型**”选择了“仅限此组织目录中的帐户”，则需要执行以下额外的设置步骤：</span><span class="sxs-lookup"><span data-stu-id="ed918-216">If you chose "Accounts in this organizational directory only" for **SUPPORTED ACCOUNT TYPES** when you registered the add-in, you need to take these additional setup steps:</span></span>

1. <span data-ttu-id="ed918-217">返回 Azure 门户，并打开加载项注册界面的“**概述**”边栏选项卡。</span><span class="sxs-lookup"><span data-stu-id="ed918-217">Go back to the Azure Portal and open the **Overview** blade of the add-in's registration.</span></span> <span data-ttu-id="ed918-218">复制“**目录(租户) ID**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-218">Copy the **Directory (tenant) ID**.</span></span>

1. <span data-ttu-id="ed918-219">在 web.config 中，将“**ida:Authority**”的值中的“common”替换为上一步复制的 GUID。</span><span class="sxs-lookup"><span data-stu-id="ed918-219">In the web.config, replace the "common" in the value of **ida:Authority** with the GUID you copied in the preceding step.</span></span> <span data-ttu-id="ed918-220">完成后，值应如下所示：`<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`。</span><span class="sxs-lookup"><span data-stu-id="ed918-220">When you are finished the value should look similar to this: `<add key="ida:Authority" value="https://login.microsoftonline.com/12345678-91ab-cdef-0123-456789abcdef/oauth2/v2.0" />`.</span></span>

1. <span data-ttu-id="ed918-221">保存并关闭 web.config。</span><span class="sxs-lookup"><span data-stu-id="ed918-221">Save and close the web.config.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="ed918-222">编写客户端代码</span><span class="sxs-lookup"><span data-stu-id="ed918-222">Code the client side</span></span>

1. <span data-ttu-id="ed918-223">打开 **Scripts** 文件夹中的 HomeES6.js 文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-223">Open the HomeES6.js file in the **Scripts** folder.</span></span> <span data-ttu-id="ed918-224">其中已存在一些代码：</span><span class="sxs-lookup"><span data-stu-id="ed918-224">It already has some code in it:</span></span>

    * <span data-ttu-id="ed918-225">有一些填充代码用于向全局窗口对象分配 Office.Promise 对象，以便在 Office 为 UI 使用 Internet Explorer 时可运行该加载项。</span><span class="sxs-lookup"><span data-stu-id="ed918-225">A polyfill that assigns the Office.Promise object to the global window object so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="ed918-226">（有关详细信息，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。）</span><span class="sxs-lookup"><span data-stu-id="ed918-226">(For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).)</span></span>
    * <span data-ttu-id="ed918-227">针对 `Office.initialize` 方法的分配，反过来又将一个处理程序分配给 `getGraphAccessTokenButton` 按钮的 Click 事件。</span><span class="sxs-lookup"><span data-stu-id="ed918-227">An assignment to the `Office.initialize` method that, in turn, assigns a handler to the `getGraphAccessTokenButton` button click event.</span></span>
    * <span data-ttu-id="ed918-228">`showResult` 方法，用于在任务窗格底部显示从 Microsoft Graph 返回的数据（或错误消息）。</span><span class="sxs-lookup"><span data-stu-id="ed918-228">A `showResult` method that will display data returned from Microsoft Graph (or an error message) at the bottom of the task pane.</span></span>
    * <span data-ttu-id="ed918-229">`logErrors` 方法，用于记录最终用户不应看到的控制台错误。</span><span class="sxs-lookup"><span data-stu-id="ed918-229">A `logErrors` method that will log to console errors that are not intended for the end user.</span></span>
    * <span data-ttu-id="ed918-230">一些代码实现了加载项在 SSO 不受支持或有错误的情况下使用的回退授权系统。</span><span class="sxs-lookup"><span data-stu-id="ed918-230">Code that implements the fallback authorization system that the add-in will use in scenarios where SSO is not supported or has errored.</span></span>

1. <span data-ttu-id="ed918-p134">在向 `Office.initialize` 分配函数下方，添加下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-p134">Below the assignment to `Office.initialize`, add the code below. Note the following about this code:</span></span>

    * <span data-ttu-id="ed918-233">加载项中的错误处理有时会自动尝试使用一组不同的选项，重新获取访问令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-233">The error-handling in the add-in will sometimes automatically attempt a second time to get an access token, using a different set of options.</span></span> <span data-ttu-id="ed918-234">计数器变量 `retryGetAccessToken` 用于确保用户不会重复循环失败的尝试来获取令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-234">The counter variable `retryGetAccessToken` is used to ensure that the user isn't cycled repeatedly through failed attempts to get a token.</span></span>
    * <span data-ttu-id="ed918-235">`getGraphData` 函数通过 ES6 `async` 关键字进行定义。</span><span class="sxs-lookup"><span data-stu-id="ed918-235">The `getGraphData` function is defined with the ES6 `async` keyword.</span></span> <span data-ttu-id="ed918-236">使用 ES6 语法可以使 Office 加载项中的 SSO API 更易于使用。</span><span class="sxs-lookup"><span data-stu-id="ed918-236">Using ES6 syntax makes the SSO API in Office Add-ins much easier to to use.</span></span> <span data-ttu-id="ed918-237">此文件是该解决方案中唯一会使用 Internet Explorer 不支持的语法的文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-237">This is the only file in the solution that will use syntax that is not supported by Internet Explorer.</span></span> <span data-ttu-id="ed918-238">我们在文件名中放入“ES6”作为提醒用途。</span><span class="sxs-lookup"><span data-stu-id="ed918-238">We put 'ES6' in the filename as a reminder.</span></span> <span data-ttu-id="ed918-239">该解决方案使用 tsc 转译器将此文件转译为 ES5，以便在 Office 为 UI 使用 Internet Explorer 时可运行加载项。</span><span class="sxs-lookup"><span data-stu-id="ed918-239">The solution uses the tsc transpiler to transpile this file to ES5, so that the add-in can run when Office is using Internet Explorer for the UI.</span></span> <span data-ttu-id="ed918-240">（请查看项目根目录中的 tsconfig.json 文件。）</span><span class="sxs-lookup"><span data-stu-id="ed918-240">(See the tsconfig.json file in the root of the project.)</span></span>

    ```javascript
    var retryGetAccessToken = 0;

    async function getGraphData() {
        await getDataWithToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
    }
    ```

1. <span data-ttu-id="ed918-241">在 `getGraphData` 函数下方，添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="ed918-241">Below the `getGraphData` function add the following function.</span></span> <span data-ttu-id="ed918-242">请注意，你将在稍后的步骤中创建 `handleClientSideErrors` 函数。</span><span class="sxs-lookup"><span data-stu-id="ed918-242">Note that you create the `handleClientSideErrors` function in a later step.</span></span>

    ```javascript
    async function getDataWithToken() {
        try {

            // TODO 1: Get the bootstrap token and send it to the server to exchange
            //         for an access token to Microsoft Graph and then get the data
            //         from Microsoft Graph.

        }
        catch (exception) {
            if (exception.code) {
                handleClientSideErrors(exception);
            }
            else {
                showResult(["EXCEPTION: " + JSON.stringify(exception)]);
            }
        }
    }
    ```

1. <span data-ttu-id="ed918-243">将 `TODO 1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-243">Replace `TODO 1` with the following.</span></span> <span data-ttu-id="ed918-244">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-244">About this code, note:</span></span>

    * <span data-ttu-id="ed918-245">`getAccessToken` 告知 Office 从 Azure AD 获取启动令牌并返回给加载项。</span><span class="sxs-lookup"><span data-stu-id="ed918-245">`getAccessToken` tells Office to get a bootstrap token from Azure AD and return to the add-in.</span></span>
    * <span data-ttu-id="ed918-246">`allowSignInPrompt` 在用户尚未登录 Office 的情况下告知 Office 提示用户进行登录。</span><span class="sxs-lookup"><span data-stu-id="ed918-246">`allowSignInPrompt` tells Office to prompt the user to sign in if the user isn't already signed into Office.</span></span>
    * <span data-ttu-id="ed918-247">`allowConsentPrompt`指示Office如果尚未授予同意，则提示用户同意允许外接程序访问用户的 AAD 配置文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-247">`allowConsentPrompt` tells Office to prompt the user to consent to letting the add-in access the user's AAD profile, if consent has not already been granted.</span></span> <span data-ttu-id="ed918-248"> (生成的提示 *不允许* 用户同意任何 Microsoft Graph作用域。) </span><span class="sxs-lookup"><span data-stu-id="ed918-248">(The resulting prompt does *not* allow the user to consent to any Microsoft Graph scopes.)</span></span>
    * <span data-ttu-id="ed918-249">`forMSGraphAccess` 告知 Office 该加载项打算使用启动令牌来换取 Microsoft Graph 的访问令牌（而不是仅将启动令牌用作用户 ID 令牌）。</span><span class="sxs-lookup"><span data-stu-id="ed918-249">`forMSGraphAccess` tells Office that the add-in intends to swap the bootstrap token for an access token to Microsoft Graph (instead of just using the bootstrap token as a user ID token).</span></span> <span data-ttu-id="ed918-250">通过设置此选项，如果用户的租户管理员尚未向加载项授予许可，则 Office 有机会取消获取启动令牌的过程（并返回错误代码 13012）。</span><span class="sxs-lookup"><span data-stu-id="ed918-250">Setting this option gives Office a chance to cancel the process of getting a bootstrap token (and return error code 13012) if the user's tenant administrator has not granted consent to the add-in.</span></span> <span data-ttu-id="ed918-251">加载项的客户端代码可以通过分支到回退授权系统来响应 13012。</span><span class="sxs-lookup"><span data-stu-id="ed918-251">The add-in's client-side code can respond to the 13012 by branching to a fallback authorization system.</span></span> <span data-ttu-id="ed918-252">如果未使用 且管理员未授予同意，将返回启动令牌，但尝试与代表流交换它将导致 `forMSGraphAccess` 错误。</span><span class="sxs-lookup"><span data-stu-id="ed918-252">If the `forMSGraphAccess` is not used and the admin has not granted consent, the bootstrap token is returned, but the attempt to exchange it with the on-behalf-of flow would result in an error.</span></span> <span data-ttu-id="ed918-253">因此，通过 `forMSGraphAccess` 选项可以快速将加载项分支到回退系统。</span><span class="sxs-lookup"><span data-stu-id="ed918-253">Thus, the `forMSGraphAccess` option enables the add-in to branch to the fallback system quickly.</span></span>
    * <span data-ttu-id="ed918-254">你将在稍后的步骤中创建 `getData` 函数。</span><span class="sxs-lookup"><span data-stu-id="ed918-254">You create the `getData` function in a later step.</span></span>
    * <span data-ttu-id="ed918-255">`/api/values` 参数是服务器端控制器的 URL，它将进行令牌交换并使用它返回的访问令牌来对 Microsoft Graph 执行调用。</span><span class="sxs-lookup"><span data-stu-id="ed918-255">The `/api/values` parameter is the URL of a server-side controller that will make the token exchange and use the access token it gets back to make the call to Microsoft Graph.</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true });

    getData("/api/values", bootstrapToken);
    ```

1. <span data-ttu-id="ed918-256">在 `getGraphData` 函数下方，添加以下内容。</span><span class="sxs-lookup"><span data-stu-id="ed918-256">Below the `getGraphData` function, add the following.</span></span> <span data-ttu-id="ed918-257">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-257">About this code, note:</span></span>

    * <span data-ttu-id="ed918-258">SSO 和回退授权系统均会使用它。</span><span class="sxs-lookup"><span data-stu-id="ed918-258">It is used by both the SSO and the fallback authorization systems.</span></span>
    * <span data-ttu-id="ed918-259">`relativeUrl` 参数是服务器端控制器。</span><span class="sxs-lookup"><span data-stu-id="ed918-259">The `relativeUrl` parameter is a server-side controller.</span></span>
    * <span data-ttu-id="ed918-260">`accessToken` 参数可以是启动令牌或完全访问令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-260">The `accessToken` parameter can be a bootstrap token or a full access token.</span></span>
    * <span data-ttu-id="ed918-261">`writeFileNamesToOfficeDocument` 已是项目的一部分。</span><span class="sxs-lookup"><span data-stu-id="ed918-261">The `writeFileNamesToOfficeDocument` is already part of the project.</span></span>
    * <span data-ttu-id="ed918-262">你将在稍后的步骤中创建 `handleServerSideErrors` 函数。</span><span class="sxs-lookup"><span data-stu-id="ed918-262">You create the `handleServerSideErrors` function in a later step.</span></span>

    ```javascript
    function getData(relativeUrl, accessToken) {

        $.ajax({
            url: relativeUrl,
            headers: { "Authorization": "Bearer " + accessToken },
            type: "GET"
        })
            .done(function (result) {
                writeFileNamesToOfficeDocument(result)
                    .then(function () {
                        showResult(["Your data has been added to the document."]);
                    })
                    .catch(function (error) {
                        showResult([JSON.stringify(error)]);
                    });
            })
            .fail(function (result) {
                handleServerSideErrors(result);
            });
    }
    ```

### <a name="handle-client-side-errors"></a><span data-ttu-id="ed918-263">处理客户端错误</span><span class="sxs-lookup"><span data-stu-id="ed918-263">Handle client-side errors</span></span>

1. <span data-ttu-id="ed918-264">在 `getData` 函数下方，添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="ed918-264">Below the `getData` function, add the following function.</span></span> <span data-ttu-id="ed918-265">请注意，`error.code` 是一个数字，通常处于 13xxx 范围内。</span><span class="sxs-lookup"><span data-stu-id="ed918-265">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 2: Handle errors where the add-in should NOT invoke
            //         the alternative system of authorization.

            // TODO 3: Handle errors where the add-in should invoke
            //         the alternative system of authorization.

        }
    }
    ```

1. <span data-ttu-id="ed918-266">将 `TODO 2` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-266">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="ed918-267">有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="ed918-267">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span>

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one
        // is logged into Office, then the first call of getAccessToken should pass the
        // `allowSignInPrompt: true` option.
        showResult(["No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to sign in, press the Get OneDrive File Names button again."]);
        break;
    case 13002:
        // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
        // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
        showResult(["You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."]);
        break;
    case 13006:
        // Only seen in Office on the web.
        showResult(["Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."]);
        break;
    case 13008:
        // Only seen in Office on the web.
        showResult(["Office is still working on the last operation. When it completes, try this operation again."]);
        break;
    case 13010:
        // Only seen in Office on the web.
        showResult(["Follow the instructions to change your browser's zone configuration."]);
        break;
    ```

1. <span data-ttu-id="ed918-268">将 `TODO 3` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-268">Replace `TODO 3` with the following code.</span></span> <span data-ttu-id="ed918-269">对于所有其他错误，加载项会分支到回退授权系统。</span><span class="sxs-lookup"><span data-stu-id="ed918-269">For all other errors, the add-in branches to the fallback authorization system.</span></span> <span data-ttu-id="ed918-270">有关这些错误的详细信息，请参阅在加载项中Office [SSO 疑难解答](troubleshoot-sso-in-office-add-ins.md)。在此外接程序中，回退系统将打开一个对话框，要求用户登录，即使用户已登录。</span><span class="sxs-lookup"><span data-stu-id="ed918-270">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is.</span></span>

    ```javascript
    default:
        dialogFallback();
        break;
    ```

### <a name="handle-server-side-errors"></a><span data-ttu-id="ed918-271">处理服务器端错误</span><span class="sxs-lookup"><span data-stu-id="ed918-271">Handle server-side errors</span></span>

1. <span data-ttu-id="ed918-272">在 `handleClientSideErrors` 函数下方，添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="ed918-272">Below the `handleClientSideErrors` function, add the following function.</span></span>

    ```javascript
    function handleServerSideErrors(result) {

    // TODO 4: Parse the JSON response.

    // TODO 5: Handle case where Microsoft Graph requires an additional form
    //         of authentication.

    // TODO 6: Handle other Azure AD errors

    }
    ```

1. <span data-ttu-id="ed918-273">将 `TODO 4` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-273">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="ed918-274">关于此代码，请注意，ASP.NET 错误类是在有类似于 MFA 的功能之前创建的。</span><span class="sxs-lookup"><span data-stu-id="ed918-274">About this code, note that ASP.NET error classes were created before there was such a thing as MFA.</span></span> <span data-ttu-id="ed918-275">服务器端逻辑处理针对第二种身份验证因素的请求时有一个副作用，即发送到客户端的服务器端错误有 **Message** 属性，但没有 **ExceptionMessage** 属性。</span><span class="sxs-lookup"><span data-stu-id="ed918-275">As a side-effect of how our server-side logic handles the requests for a second authentication factor, the server-side error sent to the client has a **Message** property but no **ExceptionMessage** property.</span></span> <span data-ttu-id="ed918-276">但是，所有其他错误都有 **ExceptionMessage** 属性，因此客户端代码必须分析这两者的响应。</span><span class="sxs-lookup"><span data-stu-id="ed918-276">But all other errors will have a **ExceptionMessage** property, so the client-side code has to parse the response for both.</span></span> <span data-ttu-id="ed918-277">一个或另一个变量将是未定义的。</span><span class="sxs-lookup"><span data-stu-id="ed918-277">Either one or the other variable will be undefined.</span></span>

    ```javascript
    var message = JSON.parse(result.responseText).Message;
    var exceptionMessage = JSON.parse(result.responseText).ExceptionMessage;
    ```

1. <span data-ttu-id="ed918-278">将 `TODO 5` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-278">Replace `TODO 5` with the following.</span></span> <span data-ttu-id="ed918-279">Microsoft Graph 要求进行其他形式的身份验证时，将发送错误 AADSTS50076。</span><span class="sxs-lookup"><span data-stu-id="ed918-279">When Microsoft Graph requires an additional form of authentication, it sends error AADSTS50076.</span></span> <span data-ttu-id="ed918-280">其中包括 **Message.Claims** 属性中的附加要求的相关信息。</span><span class="sxs-lookup"><span data-stu-id="ed918-280">It includes information about the additional requirement in the **Message.Claims** property.</span></span> <span data-ttu-id="ed918-281">为处理这种情况，该代码会再次尝试获取启动令牌，但这一次还包括请求额外的因素作为 `authChallenge` 选项的值，这会告诉 Azure AD 提示用户输入所有必需的身份验证形式。</span><span class="sxs-lookup"><span data-stu-id="ed918-281">To handle this, the code makes a second attempt to get the bootstrap token, but this time it includes the request for an additional factor as the value of the `authChallenge` option, which tells Azure AD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (message) {
        if (message.indexOf("AADSTS50076") !== -1) {
            var claims = JSON.parse(message).Claims;
            var claimsAsString = JSON.stringify(claims);
            getDataWithToken({ authChallenge: claimsAsString });
            return;
        }
    }
    ```

1. <span data-ttu-id="ed918-282">将 `TODO 6` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-282">Replace `TODO 6` with the following.</span></span>

    ```javascript
    if (exceptionMessage) {

        // TODO 7: Handle case where bootstrap token has expired.

        // TODO 8: Handle all other Azure AD errors.
    }
    ```

1. <span data-ttu-id="ed918-283">将 `TODO 7` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-283">Replace `TODO 7` with the following.</span></span> <span data-ttu-id="ed918-284">请注意，在极少数情况下，启动令牌在由 Office 验证时未过期，但是会在发动到 Azure AD 进行交换时过期。</span><span class="sxs-lookup"><span data-stu-id="ed918-284">Note that on rare occasions the bootstrap token is unexpired when Office validates it, but expires by the time it is sent to Azure AD for exchange.</span></span> <span data-ttu-id="ed918-285">Azure AD 将以错误 AADSTS500133 做出响应。</span><span class="sxs-lookup"><span data-stu-id="ed918-285">Azure AD will respond with error AADSTS500133.</span></span> <span data-ttu-id="ed918-286">发生这种情况时，代码会回调 SSO API（但不超过一次）。</span><span class="sxs-lookup"><span data-stu-id="ed918-286">When this happens, the code  recalls the SSO API (but no more than once).</span></span> <span data-ttu-id="ed918-287">这次，Office 将返回新的未过期的启动令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-287">This time Office returns a new unexpired bootstrap token.</span></span>

    ```javascript
    if ((exceptionMessage.indexOf("AADSTS500133") !== -1)
        && (retryGetAccessToken <= 0)) {

        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="ed918-288">将 `TODO 8` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-288">Replace `TODO 8` with the following.</span></span>

    ```javascript
    else {
        dialogFallback();
    }
    ```

1. <span data-ttu-id="ed918-289">保存文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-289">Save the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="ed918-290">编写服务器端代码</span><span class="sxs-lookup"><span data-stu-id="ed918-290">Code the server side</span></span>

### <a name="configure-the-owin-middleware"></a><span data-ttu-id="ed918-291">配置 OWIN 中间件</span><span class="sxs-lookup"><span data-stu-id="ed918-291">Configure the OWIN middleware</span></span>

1. <span data-ttu-id="ed918-292">在 **Office-Add-in-ASPNET-SSO-WebAPI** 项目的根目录中打开 Startup.cs 文件，并将以下方法添加到 **Startup** 类。</span><span class="sxs-lookup"><span data-stu-id="ed918-292">Open the Startup.cs file in the root of the **Office-Add-in-ASPNET-SSO-WebAPI** project and add the following method to the **Startup** class.</span></span> <span data-ttu-id="ed918-293">请注意，你将在稍后的步骤中创建 `ConfigureAuth` 方法。</span><span class="sxs-lookup"><span data-stu-id="ed918-293">Note that you create the `ConfigureAuth` method in a later step.</span></span>

    ```csharp
    public void Configuration(IAppBuilder app)
    {
        ConfigureAuth(app);
    }
    ```

1. <span data-ttu-id="ed918-294">保存并关闭此文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-294">Save and close the file.</span></span>

1. <span data-ttu-id="ed918-295">右键单击“App_Start”文件夹，并依次选择“添加”>“类”。</span><span class="sxs-lookup"><span data-stu-id="ed918-295">Right-click the **App_Start** folder and select **Add > Class**.</span></span>

1. <span data-ttu-id="ed918-296">在“添加新项”对话框中，命名文件“Startup.Auth.cs”，再单击“添加”。</span><span class="sxs-lookup"><span data-stu-id="ed918-296">In the **Add new item** dialog name the file **Startup.Auth.cs** and then click **Add**.</span></span>

1. <span data-ttu-id="ed918-297">将新文件中的命名空间名称缩短为 `Office_Add_in_ASPNET_SSO_WebAPI`。</span><span class="sxs-lookup"><span data-stu-id="ed918-297">Shorten the namespace name in the new file to `Office_Add_in_ASPNET_SSO_WebAPI`.</span></span>

1. <span data-ttu-id="ed918-298">确保下列所有 `using` 语句都位于文件的顶部。</span><span class="sxs-lookup"><span data-stu-id="ed918-298">Ensure that all of the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Owin;
    using Microsoft.IdentityModel.Tokens;
    using System.Configuration;
    using Microsoft.Owin.Security.OAuth;
    using Microsoft.Owin.Security.Jwt;
    using Office_Add_in_ASPNET_SSO_WebAPI.App_Start;
    ```

1. <span data-ttu-id="ed918-p149">将关键字 `partial` 添加到 `Startup` 类（如果其中尚不存在该关键字）的声明。具体应如下所示：</span><span class="sxs-lookup"><span data-stu-id="ed918-p149">Add the keyword `partial` to the declaration of the `Startup` class, if it is not already there. It should look like this:</span></span>

    `public partial class Startup`

1. <span data-ttu-id="ed918-p150">将下列方法添加到 `Startup` 类。该方法指定 OWIN 中间件如何验证从客户端 Home.js 文件的 `getData` 方法传递给它的访问令牌。每次调用使用 `[Authorize]` 属性修饰的 Web API 终结点时都会触发授权过程。</span><span class="sxs-lookup"><span data-stu-id="ed918-p150">Add the following method to the `Startup` class. This method specifies how the OWIN middleware will validate the access tokens that are passed to it from the `getData` method in the client-side Home.js file. The authorization process is triggered whenever a Web API endpoint that is decorated with the `[Authorize]` attribute is called.</span></span>

    ```csharp
    public void ConfigureAuth(IAppBuilder app)
    {
        // TODO 1: Configure the validation settings

        // TODO 2: Specify the type of authorization and the discovery endpoint
        //        of the secure token service.
    }
    ```

1. <span data-ttu-id="ed918-304">将 `TODO 1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-304">Replace the `TODO 1` with the following.</span></span> <span data-ttu-id="ed918-305">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-305">Note about this code:</span></span>

    * <span data-ttu-id="ed918-306">该代码指示 OWIN 确保在来自 Office 应用程序的启动令牌中指定的访问群体必须与 web.config 中指定的值匹配。</span><span class="sxs-lookup"><span data-stu-id="ed918-306">The code instructs OWIN to ensure that the audience specified in the bootstrap token that comes from the Office application must match the value specified in the web.config.</span></span>
    * <span data-ttu-id="ed918-307">Microsoft 帐户具有不同于任何组织租户 GUID 的颁发者 GUID，因此为了支持这两种类型的帐户，我们不会验证颁发者。</span><span class="sxs-lookup"><span data-stu-id="ed918-307">Microsoft accounts have an issuer GUID that is different from any organizational tenant GUID, so to support both kinds of accounts, we do not validate the issuer.</span></span>
    * <span data-ttu-id="ed918-308">设置为 `SaveSigninToken` `true` 将导致 OWIN 从应用程序保存原始Office令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-308">Setting `SaveSigninToken` to `true` causes OWIN to save the raw bootstrap token from the Office application.</span></span> <span data-ttu-id="ed918-309">加载项需要该令牌来获取具有代理流的 Microsoft Graph 访问令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-309">The add-in needs it to obtain an access token to Microsoft Graph with the on-behalf-of flow.</span></span>
    * <span data-ttu-id="ed918-310">OWIN 中间件不验证作用域。</span><span class="sxs-lookup"><span data-stu-id="ed918-310">Scopes are not validated by the OWIN middleware.</span></span> <span data-ttu-id="ed918-311">启动令牌作用域应包括 `access_as_user`，在控制器中加以验证。</span><span class="sxs-lookup"><span data-stu-id="ed918-311">The scopes of the bootstrap token, which should include `access_as_user`, is validated in the controller.</span></span>

    ```csharp
    TokenValidationParameters tvps = new TokenValidationParameters
    {
        ValidAudience = ConfigurationManager.AppSettings["ida:Audience"],
        ValidateIssuer = false,
        SaveSigninToken = true
    };
    ```

1. <span data-ttu-id="ed918-p154">将 `TODO 2` 替换为以下代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-p154">Replace `TODO 2` with the following. Note about this code:</span></span>

    * <span data-ttu-id="ed918-314">调用的是方法 `UseOAuthBearerAuthentication`，而不是更常见的 `UseWindowsAzureActiveDirectoryBearerAuthentication`，因为后者与 Azure AD V2 终结点不兼容。</span><span class="sxs-lookup"><span data-stu-id="ed918-314">The method `UseOAuthBearerAuthentication` is called instead of the more common `UseWindowsAzureActiveDirectoryBearerAuthentication` because the latter is not compatible with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="ed918-315">传递到 方法的 URL 是 OWIN 中间件获取获取密钥的说明，以验证从 Office 应用程序收到的启动令牌上的签名。</span><span class="sxs-lookup"><span data-stu-id="ed918-315">The URL that is passed to the method is where the OWIN middleware obtains instructions for getting the key it needs to verify the signature on the bootstrap token received from the Office application.</span></span> <span data-ttu-id="ed918-316">此 URL 的 Authority 区段来自 web.config。它可能是“common”字符串，而对于单租户加载项，则是一个 GUID。</span><span class="sxs-lookup"><span data-stu-id="ed918-316">The Authority segment of the URL comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>

    ```csharp
    string[] endAuthoritySegments = { "oauth2/v2.0" };
    string[] parsedAuthority = ConfigurationManager.AppSettings["ida:Authority"].Split(endAuthoritySegments, System.StringSplitOptions.None);
    string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

    app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
    {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
    });
    ```

1. <span data-ttu-id="ed918-317">保存并关闭此文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-317">Save and close the file.</span></span>

### <a name="create-the-apivalues-controller"></a><span data-ttu-id="ed918-318">创建 /api/values 控制器</span><span class="sxs-lookup"><span data-stu-id="ed918-318">Create the /api/values controller</span></span>

1. <span data-ttu-id="ed918-319">打开文件 **Controllers\ValueController.cs**。</span><span class="sxs-lookup"><span data-stu-id="ed918-319">Open the file **Controllers\ValueController.cs**.</span></span> <span data-ttu-id="ed918-320">SSO 系统成功获得启动令牌后，将使用此控制器。</span><span class="sxs-lookup"><span data-stu-id="ed918-320">This controller is used when the SSO system has successfully obtained a bootstrap token.</span></span> <span data-ttu-id="ed918-321">此控制器不用作回退授权系统的一部分。</span><span class="sxs-lookup"><span data-stu-id="ed918-321">It is not used as part of the fallback authorization system.</span></span> <span data-ttu-id="ed918-322">该系统使用的是已为你创建的 AzureADAuthController。</span><span class="sxs-lookup"><span data-stu-id="ed918-322">That system used the AzureADAuthController, which has been created for you.</span></span>

1. <span data-ttu-id="ed918-323">请确保下列 `using` 语句位于文件顶部。</span><span class="sxs-lookup"><span data-stu-id="ed918-323">Ensure that the following `using` statements are at the top of the file.</span></span>

    ```csharp
    using Microsoft.Identity.Client;
    using System.Configuration;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using System.Web.Http;
    using System;
    using System.Net;
    using System.Net.Http;
    using Office_Add_in_ASPNET_SSO_WebAPI.Helpers;
    ```

1. <span data-ttu-id="ed918-p157">在声明 `ValuesController` 的代码行的正上方，添加属性 `[Authorize]`。这可确保只要调用控制器方法时，加载项就会运行在上一过程中配置的授权过程。只有拥有对加载项的有效访问令牌，调用方才能调用控制器的方法。</span><span class="sxs-lookup"><span data-stu-id="ed918-p157">Just above the line that declares the `ValuesController`, add the `[Authorize]` attribute. This ensures that your add-in will run the authorization process that you configured in the last procedure whenever a controller method is called. Only callers with a valid access token to your add-in can invoke the methods of the controller.</span></span>

1. <span data-ttu-id="ed918-327">将下列方法添加到 `ValuesController`。</span><span class="sxs-lookup"><span data-stu-id="ed918-327">Add the following method to the `ValuesController`.</span></span> <span data-ttu-id="ed918-328">请注意，返回值是 `Task<HttpResponseMessage>`（而不是 `Task<IEnumerable<string>>`），这对于 `GET api/values` 方法而言更为常见。</span><span class="sxs-lookup"><span data-stu-id="ed918-328">Note that the return value is `Task<HttpResponseMessage>` instead of `Task<IEnumerable<string>>` as would be more common for a `GET api/values` method.</span></span> <span data-ttu-id="ed918-329">由于 OAuth 授权逻辑必须在控制器中，而不是 ASP.NET 筛选器中，所以这是一种副作用。</span><span class="sxs-lookup"><span data-stu-id="ed918-329">This is a side effect of that fact that the OAuth  authorization logic must be in the controller, instead of in an ASP.NET filter.</span></span> <span data-ttu-id="ed918-330">该逻辑中的一些错误条件要求将 HTTP 响应对象发送到加载项的客户端。</span><span class="sxs-lookup"><span data-stu-id="ed918-330">Some error conditions in that logic require that an HTTP Response object be sent to the add-in's client.</span></span>

    ```csharp
    // GET api/values
    public async Task<HttpResponseMessage> Get()
    {
        // TODO 1: Validate the scopes of the bootstrap token.

        // TODO 2: Assemble all the information that is needed to get a
        //        token for Microsoft Graph using the on-behalf-of flow.

        // TODO 3: Get the access token for Microsoft Graph.

        // TODO 4: Use the token to call Microsoft Graph.
    }
    ```

1. <span data-ttu-id="ed918-331">将 `TODO1` 替换为以下代码，以验证令牌中指定的作用域是否包括 `access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="ed918-331">Replace `TODO1` with the following code to validate that the scopes that are specified in the token include `access_as_user`.</span></span> <span data-ttu-id="ed918-332">请注意，`SendErrorToClient` 方法的第二个参数是 **Exception** 对象。</span><span class="sxs-lookup"><span data-stu-id="ed918-332">Note that the second parameter of the `SendErrorToClient` method is an **Exception** object.</span></span> <span data-ttu-id="ed918-333">在此示例中，代码传递 `null`，因为添加 **Exception** 对象会阻止在生成的 HTTP Response 中添加 **Message** 属性。</span><span class="sxs-lookup"><span data-stu-id="ed918-333">In this case, the code passes `null` because including the **Exception** object blocks the inclusion of the **Message** property in the HTTP Response that is generated.</span></span>


    ```csharp
    string[] addinScopes = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/scope").Value.Split(' ');
    if (!(addinScopes.Contains("access_as_user")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Unauthorized, null, "Missing access_as_user.");
    }
    ```

1. <span data-ttu-id="ed918-334">将 `TODO 2` 替换为以下代码，以便整合在使用代理流来获取 Microsoft Graph 的令牌时所需的所有信息。</span><span class="sxs-lookup"><span data-stu-id="ed918-334">Replace `TODO 2` with the following code to assemble all the information that is needed to get a token for Microsoft Graph using the "on behalf of" flow.</span></span> <span data-ttu-id="ed918-335">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-335">About this code, note:</span></span>

    * <span data-ttu-id="ed918-336">外接程序不再扮演资源用户或 (访问群体) 用户Office访问群体的角色。</span><span class="sxs-lookup"><span data-stu-id="ed918-336">Your add-in is no longer playing the role of a resource (or audience) to which the Office application and user need access.</span></span> <span data-ttu-id="ed918-337">现在它本身就是一个需要访问 Microsoft Graph 的客户端。</span><span class="sxs-lookup"><span data-stu-id="ed918-337">Now it is itself a client that needs access to Microsoft Graph.</span></span> <span data-ttu-id="ed918-338">是 MSAL“客户端上下文”对象。</span><span class="sxs-lookup"><span data-stu-id="ed918-338">`ConfidentialClientApplication` is the MSAL “client context” object.</span></span>
    * <span data-ttu-id="ed918-339">从 MSAL.NET 3.x.x 开始，`bootstrapContext` 仅仅是启动令牌本身。</span><span class="sxs-lookup"><span data-stu-id="ed918-339">Beginning with MSAL.NET 3.x.x, the `bootstrapContext` is just the bootstrap token itself.</span></span>
    * <span data-ttu-id="ed918-340">Authority 来自 web.config。它可能是“common”字符串，而对于单租户加载项，则是一个 GUID。</span><span class="sxs-lookup"><span data-stu-id="ed918-340">The Authority comes from the web.config. It is either the string "common" or, for a single-tenant add-in, a GUID.</span></span>
    * <span data-ttu-id="ed918-341">MSAL 要求 `openid`、`offline_access` 作用域能够发挥作用，但如果代码过多地发出请求，则会抛出错误。</span><span class="sxs-lookup"><span data-stu-id="ed918-341">MSAL requires the `openid` and `offline_access` scopes to function, but it throws an error if your code redundantly requests them.</span></span> <span data-ttu-id="ed918-342">如果你的代码请求 ，它还会引发错误，这仅在 Office 客户端应用程序获取到外接程序的 Web 应用程序的令牌时 `profile` 真正使用。</span><span class="sxs-lookup"><span data-stu-id="ed918-342">It will also throw an error if your code requests `profile`, which is really only used when the Office client application gets the token to your add-in's web application.</span></span> <span data-ttu-id="ed918-343">因此，只会显式请求获取 `Files.Read.All`。</span><span class="sxs-lookup"><span data-stu-id="ed918-343">So only `Files.Read.All` is explicitly requested.</span></span>

    ```csharp
    string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
    UserAssertion userAssertion = new UserAssertion(bootstrapContext);

    var cca = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["ida:ClientID"])
                                                    .WithRedirectUri(ConfigurationManager.AppSettings["ida:Domain"])
                                                    .WithClientSecret(ConfigurationManager.AppSettings["ida:Password"])
                                                    .WithAuthority(ConfigurationManager.AppSettings["ida:Authority"])
                                                    .Build();

    string[] graphScopes = { "https://graph.microsoft.com/Files.Read.All" };
    ```

1. <span data-ttu-id="ed918-p163">将 `TODO 3` 替换为下列代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-p163">Replace `TODO 3` with the following code. Note about this code:</span></span>

    * <span data-ttu-id="ed918-346">`ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` 方法将首先查找内存中的 MSAL 缓存，获取匹配的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-346">The `ConfidentialClientApplication.AcquireTokenOnBehalfOfAsync` method will first look in the MSAL cache, which is in memory, for a matching access token.</span></span> <span data-ttu-id="ed918-347">仅当不存在任何令牌时，该方法才会通过 Azure AD V2 终结点启动代理流。</span><span class="sxs-lookup"><span data-stu-id="ed918-347">Only if there isn't one, does it initiate the on-behalf-of flow with the Azure AD V2 endpoint.</span></span>
    * <span data-ttu-id="ed918-348">任何不属于类型 `MsalServiceException` 的异常都是有意不捕获的，这样才能作为 `500 Server Error` 消息传播到客户端。</span><span class="sxs-lookup"><span data-stu-id="ed918-348">Any exceptions that are not of type `MsalServiceException` are intentionally not caught, so they will propagate to the client as `500 Server Error` messages.</span></span>

    ```csharp
    AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
    AuthenticationResult authResult = null;
    try
    {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
    }
    catch (MsalServiceException e)
    {
        // TODO 3a: Handle request for multi-factor authentication.

        // TODO 3b: Handle lack of consent and invalid scope (permission).

        // TODO 3c: Handle all other MsalServiceExceptions.
    }
    ```

1. <span data-ttu-id="ed918-p165">将 `TODO 3a` 替换为以下代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-p165">Replace `TODO 3a` with the following code. About this code, note:</span></span>

    * <span data-ttu-id="ed918-351">如果 Microsoft Graph 资源要求进行多重身份验证，但用户尚未提供，则 Azure AD 会返回“400 错误请求”以及错误 `AADSTS50076` 和 **Claims** 属性。</span><span class="sxs-lookup"><span data-stu-id="ed918-351">If multi-factor authentication is required by the Microsoft Graph resource and the user has not yet provided it, Azure AD will return "400 Bad Request" with error `AADSTS50076` and a **Claims** property.</span></span> <span data-ttu-id="ed918-352">MSAL 抛出包含此信息的 **MsalUiRequiredException**（继承自 **MsalServiceException**）。</span><span class="sxs-lookup"><span data-stu-id="ed918-352">MSAL throws a **MsalUiRequiredException** (which inherits from **MsalServiceException**) with this information.</span></span>
    * <span data-ttu-id="ed918-353">**必须将 Claims** 属性值传递到客户端，客户端应该将它传递到 Office 应用程序，然后客户端应用程序会向请求新的启动令牌中包含它。</span><span class="sxs-lookup"><span data-stu-id="ed918-353">The **Claims** property value must be passed to the client which should pass it to the Office application, which then includes it in a request for a new bootstrap token.</span></span> <span data-ttu-id="ed918-354">Azure AD 会提示用户进行所有必需形式的身份验证。</span><span class="sxs-lookup"><span data-stu-id="ed918-354">Azure AD will prompt the user for all required forms of authentication.</span></span>
    * <span data-ttu-id="ed918-p168">由于创建异常 HTTP Response 的 API 并不知道 **Claims** 属性，因此它们不会在 Response 对象中添加这个属性。 必须手动创建消息来添加它。 不过，自定义 **Message** 属性会阻止创建 **ExceptionMessage** 属性，因此向客户端发送错误 ID `AADSTS50076` 的唯一方法是，将它添加到自定义 **Message** 中。 客户端中的 JavaScript 需要发现响应是否包含 **Message** 或 **ExceptionMessage**，这样才能了解要读取的内容。</span><span class="sxs-lookup"><span data-stu-id="ed918-p168">The APIs that create HTTP Responses from exceptions don't know about the **Claims** property, so they don't include it in the response object. We have to manually create a message that includes it. A custom **Message** property, however, blocks the creation of an **ExceptionMessage** property, so the only way to get the error ID `AADSTS50076` to the client is to add it to the custom **Message**. JavaScript in the client will need to discover if a response has a **Message** or **ExceptionMessage**, so it knows which to read.</span></span>
    * <span data-ttu-id="ed918-359">自定义消息被格式化为 JSON，以便客户端 JavaScript 能够使用已知的 JavaScript `JSON` 对象方法分析它。</span><span class="sxs-lookup"><span data-stu-id="ed918-359">The custom message is formatted as JSON so that the client-side JavaScript can parse it with well-known JavaScript `JSON` object methods.</span></span>

    ```csharp
    if (e.Message.StartsWith("AADSTS50076"))
    {
        string responseMessage = String.Format("{{\"AADError\":\"AADSTS50076\",\"Claims\":{0}}}", e.Claims);
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, null, responseMessage);
    }
    ```

1. <span data-ttu-id="ed918-p169">将 `TODO 3b` 替换为以下代码。关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ed918-p169">Replace `TODO 3b` with the following code. About this code, note:</span></span>

    * <span data-ttu-id="ed918-362">如果 Azure AD 调用包含至少一个作用域（权限）未获得用户和租户管理员的许可（或许可被撤消），则 Azure AD 将返回“400 错误请求”和错误 `AADSTS65001`。</span><span class="sxs-lookup"><span data-stu-id="ed918-362">If the call to Azure AD contained at least one scope (permission) for which neither the user nor a tenant administrator has consented (or consent was revoked), Azure AD will return "400 Bad Request" with error `AADSTS65001`.</span></span> <span data-ttu-id="ed918-363">MSAL 抛出包含此信息的 **MsalUiRequiredException**。</span><span class="sxs-lookup"><span data-stu-id="ed918-363">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    * <span data-ttu-id="ed918-364">如果 Azure AD 调用包含至少一个 Azure AD 无法识别的作用域，则 AAD 将返回“400 错误请求”和错误 `AADSTS70011`。</span><span class="sxs-lookup"><span data-stu-id="ed918-364">If the call to Azure AD contained at least one scope that Azure AD does not recognize, AAD returns "400 Bad Request" with error `AADSTS70011`.</span></span> <span data-ttu-id="ed918-365">MSAL 抛出包含此信息的 **MsalUiRequiredException**。</span><span class="sxs-lookup"><span data-stu-id="ed918-365">MSAL throws a **MsalUiRequiredException** with this information.</span></span>
    * <span data-ttu-id="ed918-366">其中包含完整说明，因为 70011 也会在其他情况下返回，只有在它表示存在无效范围时，才需要在此加载项中处理它。</span><span class="sxs-lookup"><span data-stu-id="ed918-366">The entire description is included because 70011 is returned in other conditions and it should only be handled in this add-in when it means that there is an invalid scope.</span></span>
    * <span data-ttu-id="ed918-p172">**MsalUiRequiredException** 对象传递给 `SendErrorToClient`。这样可确保 HTTP 响应中有包含错误消息的 **ExceptionMessage** 属性。</span><span class="sxs-lookup"><span data-stu-id="ed918-p172">The **MsalUiRequiredException** object is passed to `SendErrorToClient`. This ensures that an **ExceptionMessage** property that contains the error information is included in the HTTP Response.</span></span>

    ```csharp
    if ((e.Message.StartsWith("AADSTS65001")) || (e.Message.StartsWith("AADSTS70011: The provided value for the input parameter 'scope' is not valid.")))
    {
        return HttpErrorHelper.SendErrorToClient(HttpStatusCode.Forbidden, e, null);
    }
    ```

1. <span data-ttu-id="ed918-369">将 `TODO 3c` 替换为以下代码，以处理所有其他 **MsalServiceException**。</span><span class="sxs-lookup"><span data-stu-id="ed918-369">Replace `TODO 3c` with the following code to handle all other **MsalServiceException** s.</span></span> <span data-ttu-id="ed918-370">正如前文所述，</span><span class="sxs-lookup"><span data-stu-id="ed918-370">As noted earlier,</span></span>

    ```csharp
    else
    {
        throw e;
    }
    ```

1. <span data-ttu-id="ed918-371">将 `TODO 4` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="ed918-371">Replace `TODO 4` with the following code.</span></span> <span data-ttu-id="ed918-372">事先为你创建的 `GraphApiHelper.GetOneDriveFileNames` 方法将向 Microsoft Graph 请求数据并包含访问令牌。</span><span class="sxs-lookup"><span data-stu-id="ed918-372">The `GraphApiHelper.GetOneDriveFileNames` method, which has been created for you, makes the request for data to Microsoft Graph and includes the access token.</span></span>

    ```csharp
    return await GraphApiHelper.GetOneDriveFileNames(authResult.AccessToken);
    ```

1. <span data-ttu-id="ed918-373">保存并关闭文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-373">Save and close the file.</span></span>

## <a name="run-the-solution"></a><span data-ttu-id="ed918-374">运行解决方案</span><span class="sxs-lookup"><span data-stu-id="ed918-374">Run the solution</span></span>

1. <span data-ttu-id="ed918-375">打开 Visual Studio 解决方案文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-375">Open the Visual Studio solution file.</span></span>
1. <span data-ttu-id="ed918-376">在“**生成**”菜单上，选择“**清理解决方案**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-376">On the **Build** menu, select **Clean Solution**.</span></span> <span data-ttu-id="ed918-377">完成后，再次打开“**生成**”菜单，并选择“**生成解决方案**”。</span><span class="sxs-lookup"><span data-stu-id="ed918-377">When it finishes, open the **Build** menu again and select **Build Solution**.</span></span>
1. <span data-ttu-id="ed918-378">在“**解决方案资源管理器**”中，选择“**Office-Add-in-ASPNET-SSO**”项目节点（而不是顶部的解决方案节点，也不是名称以“WebAPI”结尾的项目）。</span><span class="sxs-lookup"><span data-stu-id="ed918-378">In **Solution Explorer**, select the **Office-Add-in-ASPNET-SSO** project node (not the top solution node and not the project whose name ends in "WebAPI").</span></span>
1. <span data-ttu-id="ed918-379">在“**属性**”窗格中，打开“**启动文档**”下拉列表，然后选择三个选项之一（“Excel”、“Word”或“PowerPoint”）。</span><span class="sxs-lookup"><span data-stu-id="ed918-379">In the **Properties** pane, open the **Start Document** drop down and choose one of the three options (Excel, Word, or PowerPoint).</span></span>

    ![选择所需的Office客户端应用程序：Excel、PowerPoint 或 Word](../images/SelectHost.JPG)

1. <span data-ttu-id="ed918-381">按 F5。</span><span class="sxs-lookup"><span data-stu-id="ed918-381">Press F5.</span></span>
1. <span data-ttu-id="ed918-382">在 Office 应用程序的“**主页**”功能区上，选择“**SSO ASP.NET**”组中的“**显示加载项**”以打开任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="ed918-382">In the Office application, on the **Home** ribbon, select the **Show Add-in** in the **SSO ASP.NET** group to open the task pane add-in.</span></span>
1. <span data-ttu-id="ed918-383">单击“**获取 OneDrive 文件名**”按钮。</span><span class="sxs-lookup"><span data-stu-id="ed918-383">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="ed918-384">如果使用 Microsoft 365 教育版 或工作帐户或 Microsoft 帐户登录 Office 并且 SSO 按预期工作，OneDrive for Business 中的前 10 个文件和文件夹名称将显示在任务窗格中。</span><span class="sxs-lookup"><span data-stu-id="ed918-384">If you are logged into Office with either a Microsoft 365 Education or work account, or a Microsoft account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are displayed on the task pane.</span></span> <span data-ttu-id="ed918-385">如果您未登录，或者您位于不支持 SSO 的方案中，或者 SSO 因任何原因无法工作，系统将提示您登录。</span><span class="sxs-lookup"><span data-stu-id="ed918-385">If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to sign in.</span></span> <span data-ttu-id="ed918-386">登录后，将显示文件和文件夹名称。</span><span class="sxs-lookup"><span data-stu-id="ed918-386">After you sign in, the file and folder names appear.</span></span>

## <a name="updating-the-add-in-when-you-go-to-staging-and-production"></a><span data-ttu-id="ed918-387">转到暂存和生产时更新外接程序</span><span class="sxs-lookup"><span data-stu-id="ed918-387">Updating the add-in when you go to staging and production</span></span>

<span data-ttu-id="ed918-388">与Office Web 外接程序一样，当您准备好移动到暂存服务器或生产服务器时，必须使用新域更新清单 `localhost:44355` 中的域。</span><span class="sxs-lookup"><span data-stu-id="ed918-388">Like all Office Web Add-ins, when you are ready to move to a staging or production server, you must update the `localhost:44355` domain in the manifest with the new domain.</span></span> <span data-ttu-id="ed918-389">同样，您必须更新域的 web.config 文件。</span><span class="sxs-lookup"><span data-stu-id="ed918-389">Similarly, you must update the domain in the web.config file.</span></span>

<span data-ttu-id="ed918-390">由于该域出现在 AAD 注册中，因此您需要更新该注册以使用新域，以在它出现 `localhost:44355` 的位置进行更改。</span><span class="sxs-lookup"><span data-stu-id="ed918-390">Since the domain appears in the AAD registration, you need to update that registration to use the new domain in place of `localhost:44355` wherever it appears.</span></span>
