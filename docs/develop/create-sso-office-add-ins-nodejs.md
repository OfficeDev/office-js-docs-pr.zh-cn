---
title: 创建使用单一登录的 Node.js Office 加载项
description: 了解如何创建使用 Office 单一登录的基于 Node.js 的 Office 加载项
ms.date: 01/13/2020
localization_priority: Priority
ms.openlocfilehash: 140a18628ae03d842482721394d976207cb7fa37
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111065"
---
# <a name="create-a-nodejs-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="3b567-103">创建使用单一登录的 Node.js Office 加载项（预览）</span><span class="sxs-lookup"><span data-stu-id="3b567-103">Create a Node.js Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="3b567-p101">用户可以登录 Office，Office Web 加载项能够利用此登录进程，授权用户访问加载项和 Microsoft Graph，而无需要求用户再登录一次。有关概述，请参阅[在 Office 加载项中启用 SSO](sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="3b567-p101">Users can sign in to Office, and your Office Web Add-in can take advantage of this sign-in process to authorize users to your add-in and to Microsoft Graph without requiring users to sign in a second time. For an overview, see [Enable SSO in an Office Add-in](sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="3b567-106">本文将逐步介绍如何在使用 Node.js 和 Express 生成的加载项中启用单一登录 (SSO) 。</span><span class="sxs-lookup"><span data-stu-id="3b567-106">This article walks you through the process of enabling single sign-on (SSO) in an add-in that is built with Node.js and Express.</span></span> <span data-ttu-id="3b567-107">有关与此类似的 ASP.NET 加载项文章，请参阅[创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)。</span><span class="sxs-lookup"><span data-stu-id="3b567-107">For a similar article about an ASP.NET-based add-in, see [Create an ASP.NET Office Add-in that uses single sign-on](create-sso-office-add-ins-aspnet.md).</span></span>

> [!NOTE]
> <span data-ttu-id="3b567-108">作为完成本文中所述步骤的替代方法，可使用 Yeoman 生成器创建启用 SSO 的 Node.js Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="3b567-108">As an alternative to completing the steps described in this article, you can use the Yeoman generator to create an SSO-enabled, Node.js Office Add-in.</span></span> <span data-ttu-id="3b567-109">Yeoman 生成器简化了启用了 SSO 的加载项创建流程，能够自动执行在 Azure 内配置所需的步骤，并生成加载项使用 SSO 所需的代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-109">The Yeoman generator simplifies the process of creating an SSO-enabled add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="3b567-110">有关详细信息，请参阅“[单一登录（SSO）快速入门](../quickstarts/sso-quickstart.md)”。</span><span class="sxs-lookup"><span data-stu-id="3b567-110">For more information, see the [Single sign-on (SSO) quick start](../quickstarts/sso-quickstart.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3b567-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="3b567-111">Prerequisites</span></span>

* <span data-ttu-id="3b567-112">[节点和 npm](https://nodejs.org/)，版本 10.15.0 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="3b567-112">[Node and npm](https://nodejs.org/), version 10.15.0 or later.</span></span>

* <span data-ttu-id="3b567-113">[Git Bash](https://git-scm.com/downloads)（或其他 git 客户端）</span><span class="sxs-lookup"><span data-stu-id="3b567-113">[Git Bash](https://git-scm.com/downloads) (or another git client)</span></span>

* <span data-ttu-id="3b567-114">TypeScript，版本 3.6.2 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3b567-114">TypeScript, version 3.6.2 or later</span></span>

* <span data-ttu-id="3b567-115">Office 365（Office 的订阅版本）帐户，获取方法为加入 [Office 365 开发人员计划](https://aka.ms/devprogramsignup)，其中包含为期 1 年的免费 Office 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="3b567-115">Office 365 (the subscription version of Office) account which you can get by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup) that includes a free 1 year subscription to Office 365.</span></span> <span data-ttu-id="3b567-116">应使用最新的每月版本并从预览体验成员频道构建，但你必须是 Office 预览体验成员才能获取此版本。</span><span class="sxs-lookup"><span data-stu-id="3b567-116">You should use the latest monthly version and build from the Insiders channel but you need to be an Office Insider to get this version.</span></span> <span data-ttu-id="3b567-117">有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。</span><span class="sxs-lookup"><span data-stu-id="3b567-117">For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span> <span data-ttu-id="3b567-118">请注意，当内部版本进入生产半年频道时，将关闭对该内部版本的预览功能（包括 SSO）的支持。</span><span class="sxs-lookup"><span data-stu-id="3b567-118">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

* <span data-ttu-id="3b567-119">一个代码编辑器。</span><span class="sxs-lookup"><span data-stu-id="3b567-119">A code editor.</span></span> <span data-ttu-id="3b567-120">建议使用 Visual Studio Code。</span><span class="sxs-lookup"><span data-stu-id="3b567-120">We recommend Visual Studio Code.</span></span>

* <span data-ttu-id="3b567-121">Office 365 订阅中的 OneDrive for Business 上至少存储了一些文件和文件夹。</span><span class="sxs-lookup"><span data-stu-id="3b567-121">At least a few files and folders stored on OneDrive for Business in your Office 365 subscription.</span></span>

* <span data-ttu-id="3b567-122">一个 Microsoft Azure 订阅。</span><span class="sxs-lookup"><span data-stu-id="3b567-122">Microsoft Azure subscription</span></span> <span data-ttu-id="3b567-123">此加载项需要 Azure Active Directory (AD)。</span><span class="sxs-lookup"><span data-stu-id="3b567-123">This add-in requires Azure Active Directory (AD).</span></span> <span data-ttu-id="3b567-124">Azure AD 为应用程序提供了用于进行身份验证和授权的标识服务。</span><span class="sxs-lookup"><span data-stu-id="3b567-124">Azure AD provides identity services that applications use for authentication and authorization.</span></span> <span data-ttu-id="3b567-125">你还可在 [Microsoft Azure](https://account.windowsazure.com/SignUp) 获得试用订阅。</span><span class="sxs-lookup"><span data-stu-id="3b567-125">A trial subscription can be acquired at [Microsoft Azure](https://account.windowsazure.com/SignUp).</span></span>

## <a name="set-up-the-starter-project"></a><span data-ttu-id="3b567-126">设置初学者项目</span><span class="sxs-lookup"><span data-stu-id="3b567-126">Set up the starter project</span></span>

1. <span data-ttu-id="3b567-127">克隆或下载 [Office 外接程序 NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso) 中的存储库。</span><span class="sxs-lookup"><span data-stu-id="3b567-127">Clone or download the repo at [Office Add-in NodeJS SSO](https://github.com/officedev/office-add-in-nodejs-sso).</span></span>

    > [!NOTE]
    > <span data-ttu-id="3b567-128">示例有三个版本：</span><span class="sxs-lookup"><span data-stu-id="3b567-128">There are three versions of the sample:</span></span>  
    > * <span data-ttu-id="3b567-p107">**Before** 文件夹是初学者项目。未直接连接到 SSO 或授权的外接程序的 UI 和其他方面已经完成。本文后续章节将引导你完成此过程。</span><span class="sxs-lookup"><span data-stu-id="3b567-p107">The **Before** folder is a starter project. The UI and other aspects of the add-in that are not directly connected to SSO or authorization are already done. Later sections of this article walk you through the process of completing it.</span></span>
    > * <span data-ttu-id="3b567-132">如果完成了本文中的过程，该示例的**已完成**版本会与所生成的加载项类似，只不过完成的项目具有对本文文本冗余的代码注释。</span><span class="sxs-lookup"><span data-stu-id="3b567-132">The **Complete** version of the sample is just like the add-in that you would have if you completed the procedures of this article, except that the completed project has code comments that would be redundant with the text of this article.</span></span> <span data-ttu-id="3b567-133">若要使用已完成的版本，请按照本文中的说明进行操作即可，但需要将“Before”替换为“Completed”，并跳过**编写客户端代码**和**编写服务器端代码**部分。</span><span class="sxs-lookup"><span data-stu-id="3b567-133">To use the completed version, just follow the instructions in this article, but replace "Before" with "Completed" and skip the sections **Code the client side** and **Code the server** side.</span></span>
    > * <span data-ttu-id="3b567-134">**SSOAutoSetup** 版本是一个完整示例，可自动执行大多数步骤以在 Azure AD 中注册加载项并对其进行配置。</span><span class="sxs-lookup"><span data-stu-id="3b567-134">The **SSOAutoSetup** version is a completed sample that automates most of the steps to register the add-in with Azure AD and configure it.</span></span> <span data-ttu-id="3b567-135">如果想要快速查看使用 SSO 的加载项，请使用此版本。</span><span class="sxs-lookup"><span data-stu-id="3b567-135">Use this version if you want to see a working add-in with SSO quickly.</span></span> <span data-ttu-id="3b567-136">按照文件夹自述文件中的步骤操作即可。</span><span class="sxs-lookup"><span data-stu-id="3b567-136">Just follow the steps in the Readme of the folder.</span></span> <span data-ttu-id="3b567-137">我们建议你在某些时候完成本文中的手动注册和设置步骤，以更好地了解 Azure AD 与加载项之间的关系。</span><span class="sxs-lookup"><span data-stu-id="3b567-137">We recommend that at some point you go through the manual registration and setup steps in this article to better understand the relationship between Azure AD and an add-in.</span></span> 


1. <span data-ttu-id="3b567-138">在 **Before** 文件夹中打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="3b567-138">Open a command prompt in the **Before** folder.</span></span>

1. <span data-ttu-id="3b567-139">在该控制台中输入 `npm install` 以安装 package.json 文件中列出明细的所有依赖项。</span><span class="sxs-lookup"><span data-stu-id="3b567-139">Enter `npm install` in the console to install all of the dependencies itemized in the package.json file.</span></span>

1. <span data-ttu-id="3b567-140">运行命令 `npm run install-dev-certs`。</span><span class="sxs-lookup"><span data-stu-id="3b567-140">Run the command `npm run install-dev-certs`.</span></span> <span data-ttu-id="3b567-141">为安装证书的提示选择“**是**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-141">Select **Yes** to the prompt to install the certificate.</span></span>

## <a name="register-the-add-in-with-azure-ad-v20-endpoint"></a><span data-ttu-id="3b567-142">向 Azure AD v2.0 终结点注册加载项。</span><span class="sxs-lookup"><span data-stu-id="3b567-142">Register the add-in with Azure AD v2.0 endpoint</span></span>

1. <span data-ttu-id="3b567-143">导航到“Azure 门户 - 应用注册”[](https://go.microsoft.com/fwlink/?linkid=2083908)页面以注册你的应用。</span><span class="sxs-lookup"><span data-stu-id="3b567-143">Navigate to the [Azure portal - App registrations](https://go.microsoft.com/fwlink/?linkid=2083908) page to register your app.</span></span>

1. <span data-ttu-id="3b567-144">使用***管理员***凭据登录 Office 365 租户。</span><span class="sxs-lookup"><span data-stu-id="3b567-144">Sign in with the ***admin*** credentials to your Office 365 tenancy.</span></span> <span data-ttu-id="3b567-145">例如，MyName@contoso.onmicrosoft.com。</span><span class="sxs-lookup"><span data-stu-id="3b567-145">For example, MyName@contoso.onmicrosoft.com.</span></span>

1. <span data-ttu-id="3b567-146">选择“新注册”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-146">Select **New registration**.</span></span> <span data-ttu-id="3b567-147">在“注册应用”\*\*\*\* 页上，按如下方式设置值。</span><span class="sxs-lookup"><span data-stu-id="3b567-147">On the **Register an application** page, set the values as follows.</span></span>

    * <span data-ttu-id="3b567-148">将“名称”\*\*\*\* 设置为“`Office-Add-in-NodeJS-SSO`”。</span><span class="sxs-lookup"><span data-stu-id="3b567-148">Set **Name** to `Office-Add-in-NodeJS-SSO`.</span></span>
    * <span data-ttu-id="3b567-149">将“**受支持的帐户类型**”设置为“**任何组织目录中的帐户和个人 Microsoft 帐户**”（例如，Skype、Xbox、Outlook.com）。</span><span class="sxs-lookup"><span data-stu-id="3b567-149">Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span>
    * <span data-ttu-id="3b567-150">将“**重定向R URI**”设置为 ` https://localhost:44355/dialog.html`。</span><span class="sxs-lookup"><span data-stu-id="3b567-150">Set **Redirect URI** to` https://localhost:44355/dialog.html`.</span></span>
    * <span data-ttu-id="3b567-151">选择“**注册**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-151">Choose **Register**.</span></span>

1. <span data-ttu-id="3b567-152">在 **Office-Add-in-NodeJS-SSO** 页面上，复制并保存“**应用程序（客户端）ID**”和“**目录（租户）ID**”的值。</span><span class="sxs-lookup"><span data-stu-id="3b567-152">On the **Office-Add-in-NodeJS-SSO** page, copy and save the values for the **Application (client) ID** and the **Directory (tenant) ID**.</span></span> <span data-ttu-id="3b567-153">你将在后面的过程中使用它们。</span><span class="sxs-lookup"><span data-stu-id="3b567-153">You'll use both of them in later procedures.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3b567-154">当其他应用程序（例如 PowerPoint、Word、Excel 等 Office 主机应用程序）寻求对应用程序的授权访问权限时，此 ID 是“受众”值。</span><span class="sxs-lookup"><span data-stu-id="3b567-154">This ID is the "audience" value when other applications, such as the Office host application (e.g., PowerPoint, Word, Excel), seek authorized access to the application.</span></span> <span data-ttu-id="3b567-155">当它反过来寻求 Microsoft Graph 的授权访问权限时，它同时也是应用程序的“客户端 ID”。</span><span class="sxs-lookup"><span data-stu-id="3b567-155">It is also the "client ID" of the application when it, in turn, seeks authorized access to Microsoft Graph.</span></span>

1. <span data-ttu-id="3b567-156">选择“**管理**”下的“**身份验证**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-156">Select **Authentication** under **Manage**.</span></span> <span data-ttu-id="3b567-157">在“**隐式授权**”部分中，启用“**访问令牌**”和“**ID 令牌**”的复选框。</span><span class="sxs-lookup"><span data-stu-id="3b567-157">In the **Implict grant** section, enable the checkboxes for both **Access token** and **ID token**.</span></span> <span data-ttu-id="3b567-158">该示例具有一个回退授权系统，当 SSO 不可用时，将调用此系统。</span><span class="sxs-lookup"><span data-stu-id="3b567-158">The sample has a fallback authorization system that is invoked when SSO is not available.</span></span> <span data-ttu-id="3b567-159">该系统使用隐式流。</span><span class="sxs-lookup"><span data-stu-id="3b567-159">This system uses the Implicit Flow.</span></span>

1. <span data-ttu-id="3b567-160">在窗体顶部，选择“**保存**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-160">Select **Save** at the top of the form.</span></span>

1. <span data-ttu-id="3b567-161">选择“管理”\*\*\*\* 下的“证书和密码”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-161">Select **Certificates & secrets** under **Manage**.</span></span> <span data-ttu-id="3b567-162">选择“新客户端密码”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="3b567-162">Select the **New client secret** button.</span></span> <span data-ttu-id="3b567-163">输入“描述”\*\*\*\* 的值，然后选择“到期”\*\*\*\* 的适当选项，并选择“添加”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-163">Enter a value for **Description** then select an appropriate option for **Expires** and choose **Add**.</span></span> <span data-ttu-id="3b567-164">在继续操作前，*立即复制客户端机密码值并使用应用程序 ID 保存它*，因为在后面的过程中，将需要用到它。</span><span class="sxs-lookup"><span data-stu-id="3b567-164">*Copy the client secret value immediately and save it with the application ID* before proceeding as you'll need it in a later procedure.</span></span>

1. <span data-ttu-id="3b567-165">在“管理”\*\*\*\* 下选择“公开 API”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-165">Select **Expose an API** under **Manage**.</span></span> <span data-ttu-id="3b567-166">选择“**设置**”链接以在窗体“api://$App ID GUID$”中生成应用 ID URI，其中 $App ID GUID$ 是**应用程序（客户端）ID**。</span><span class="sxs-lookup"><span data-stu-id="3b567-166">Select the **Set** link to generate the Application ID URI in the form "api://$App ID GUID$", where $App ID GUID$ is the **Application (client) ID**.</span></span> <span data-ttu-id="3b567-167">在双正斜杠和 GUID 之间插入 `localhost:44355/`（注意末尾附加的正斜杠“/”）。</span><span class="sxs-lookup"><span data-stu-id="3b567-167">Insert `localhost:44355/` (note the forward slash "/" appended to the end) between the double forward slashes and the GUID.</span></span> <span data-ttu-id="3b567-168">整个 ID 的格式应为 `api://localhost:44355/$App ID GUID$`；例如 `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`。</span><span class="sxs-lookup"><span data-stu-id="3b567-168">The entire ID should have the form `api://localhost:44355/$App ID GUID$`; for example `api://localhost:44355/c6c1f32b-5e55-4997-881a-753cc1d563b7`.</span></span> 

1. <span data-ttu-id="3b567-169">选择“**添加一个作用域**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3b567-169">Select the **Add a scope** button.</span></span> <span data-ttu-id="3b567-170">在打开的面板中，输入 `access_as_user` 作为**作用域**名称。</span><span class="sxs-lookup"><span data-stu-id="3b567-170">In the panel that opens, enter `access_as_user` as the **Scope** name.</span></span>

1. <span data-ttu-id="3b567-171">将“谁能同意?”\*\*\*\* 设置为“管理员和用户”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-171">Set **Who can consent?** to **Admins and users**.</span></span>

1. <span data-ttu-id="3b567-172">使用适合 `access_as_user` 作用域的值填写用于配置管理员和用户同意提示的字段，使 Office 主机应用能够借助与当前用户具有相同权限使用加载项 Web API。</span><span class="sxs-lookup"><span data-stu-id="3b567-172">Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the `access_as_user` scope which enables the Office host application to use your add-in's web APIs with the same rights as the current user.</span></span> <span data-ttu-id="3b567-173">建议：</span><span class="sxs-lookup"><span data-stu-id="3b567-173">Suggestions:</span></span>

    - <span data-ttu-id="3b567-174">**管理员许可标题**：Office 可以充当用户。</span><span class="sxs-lookup"><span data-stu-id="3b567-174">**Admin consent title**: Office can act as the user.</span></span>
    - <span data-ttu-id="3b567-175">**管理员许可描述**：使 Office 能够借助与当前用户相同的权限调用加载项的 Web API。</span><span class="sxs-lookup"><span data-stu-id="3b567-175">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights as the current user.</span></span>
    - <span data-ttu-id="3b567-176">**用户许可标题**：Office 可以充当你。</span><span class="sxs-lookup"><span data-stu-id="3b567-176">**User consent title**: Office can act as you.</span></span>
    - <span data-ttu-id="3b567-177">**管理员许可描述**：使 Office 能够借助与你相同的权限调用加载项的 Web API。</span><span class="sxs-lookup"><span data-stu-id="3b567-177">**Admin consent description**: Enable Office to call the add-in's web APIs with the same rights that you have.</span></span>

1. <span data-ttu-id="3b567-178">确保将“**状态**”设置为“**已启用**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-178">Ensure that **State** is set to **Enabled**.</span></span>

1. <span data-ttu-id="3b567-179">选择“**添加作用域**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-179">Select **Add scope** .</span></span>

    > [!NOTE]
    > <span data-ttu-id="3b567-180">显示在文本字段正下方的**作用域**名称的域部分应自动与你先前设置的“应用 ID URI”匹配，并将 `/access_as_user` 附加到末尾；例如，`api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="3b567-180">The domain part of the **Scope** name displayed just below the text field should automatically match the Application ID URI that you set earlier, with `/access_as_user` appended to the end; for example, `api://localhost:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`.</span></span>

1. <span data-ttu-id="3b567-181">在“授权客户端应用程序”\*\*\*\* 部分中，确定要授权给加载项 Web 应用程序的应用程序。</span><span class="sxs-lookup"><span data-stu-id="3b567-181">In the **Authorized client applications** section, you identify the applications that you want to authorize to your add-in's web application.</span></span> <span data-ttu-id="3b567-182">下面每个 ID 都需要进行预授权。</span><span class="sxs-lookup"><span data-stu-id="3b567-182">Each of the following IDs needs to be pre-authorized.</span></span>

    - <span data-ttu-id="3b567-183">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="3b567-183">`d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office)</span></span>
    - <span data-ttu-id="3b567-184">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span><span class="sxs-lookup"><span data-stu-id="3b567-184">`ea5a67f6-b6f3-4338-b240-c655ddc3cc8e` (Microsoft Office)</span></span>
    - <span data-ttu-id="3b567-185">`57fb890c-0dab-4253-a5e0-7188c88b2bb4`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="3b567-185">`57fb890c-0dab-4253-a5e0-7188c88b2bb4` (Office on the web)</span></span>
    - <span data-ttu-id="3b567-186">`bc59ab01-8403-45c6-8796-ac3ef710b3e3`（Office 网页版）</span><span class="sxs-lookup"><span data-stu-id="3b567-186">`bc59ab01-8403-45c6-8796-ac3ef710b3e3` (Office on the web)</span></span>

    <span data-ttu-id="3b567-187">对于每个 ID，执行以下步骤：</span><span class="sxs-lookup"><span data-stu-id="3b567-187">For each ID, take these steps:</span></span>

    <span data-ttu-id="3b567-188">a.</span><span class="sxs-lookup"><span data-stu-id="3b567-188">a.</span></span> <span data-ttu-id="3b567-189">选择“**添加客户端应用程序**”按钮，然后在打开的面板中，将“客户端 ID”设置为相应的 GUID 并勾选 `api://localhost:44355/$App ID GUID$/access_as_user` 框。</span><span class="sxs-lookup"><span data-stu-id="3b567-189">Select **Add a client application** button and then, in the panel that opens, set the Client ID to the respective GUID and check the box for `api://localhost:44355/$App ID GUID$/access_as_user`.</span></span>

    <span data-ttu-id="3b567-190">b.</span><span class="sxs-lookup"><span data-stu-id="3b567-190">b.</span></span> <span data-ttu-id="3b567-191">选择“添加应用程序”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-191">Select **Add application**.</span></span>

1. <span data-ttu-id="3b567-192">选择“管理”\*\*\*\* 下的“API 权限”\*\*\*\*，然后选择“添加权限”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-192">Select **API permissions** under **Manage** and select **Add a permission**.</span></span> <span data-ttu-id="3b567-193">在打开的面板上，选择 **Microsoft Graph**，然后选择“委派权限”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3b567-193">On the panel that opens, choose **Microsoft Graph** and then choose **Delegated permissions**.</span></span>

1. <span data-ttu-id="3b567-194">使用“选择权限”\*\*\*\* 搜索框来搜索加载项需要的权限。</span><span class="sxs-lookup"><span data-stu-id="3b567-194">Use the **Select permissions** search box to search for the permissions your add-in needs.</span></span> <span data-ttu-id="3b567-195">选择以下选项。</span><span class="sxs-lookup"><span data-stu-id="3b567-195">Select the following.</span></span> <span data-ttu-id="3b567-196">外接程序本身真正需要的只是第一项权限，但 Office 主机必须有 `profile` 权限，才能获取访问外接程序 Web 应用程序的令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-196">Only the first is really required by your add-in itself; but the `profile` permission is required for the Office host to get a token to your add-in web application.</span></span>

    * <span data-ttu-id="3b567-197">Files.Read.All</span><span class="sxs-lookup"><span data-stu-id="3b567-197">Files.Read.All</span></span>
    * <span data-ttu-id="3b567-198">profile</span><span class="sxs-lookup"><span data-stu-id="3b567-198">profile</span></span>

    > [!NOTE]
    > <span data-ttu-id="3b567-199">`User.Read` 权限可能已默认列出。</span><span class="sxs-lookup"><span data-stu-id="3b567-199">The `User.Read` permission may already be listed by default.</span></span> <span data-ttu-id="3b567-200">根据最佳做法，最好不要请求授予不需要的权限，因此，如果加载项实际上不需要此权限，我们建议取消选中此权限对应的框。</span><span class="sxs-lookup"><span data-stu-id="3b567-200">It is a good practice not to ask for permissions that are not needed, so we recommend that you uncheck the box for this permission if your add-in does not actually need it.</span></span>

1. <span data-ttu-id="3b567-201">选择所显示的每个权限的复选框。</span><span class="sxs-lookup"><span data-stu-id="3b567-201">Select the check box for each permission as it appears.</span></span> <span data-ttu-id="3b567-202">选择加载项需要的权限后，选择面板底部的“**添加权限**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3b567-202">After selecting the permissions that your add-in needs, select the **Add permissions** button at the bottom of the panel.</span></span>

1. <span data-ttu-id="3b567-203">在同一页面上，选择“**为[租户名称]授予管理员许可**”按钮，然后为显示的确认选择“**是**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-203">On the same page, choose the **Grant admin consent for [tenant name]** button, and then select **Yes** for the confirmation that appears.</span></span>

## <a name="configure-the-add-in"></a><span data-ttu-id="3b567-204">配置加载项</span><span class="sxs-lookup"><span data-stu-id="3b567-204">Configure the add-in</span></span>

1. <span data-ttu-id="3b567-205">在代码编辑器中打开克隆项目中的 `\Begin` 文件夹。</span><span class="sxs-lookup"><span data-stu-id="3b567-205">Open the `\Begin` folder in the cloned project in your code editor.</span></span>

1. <span data-ttu-id="3b567-206">打开 `.ENV` 文件，并使用先前复制的值。</span><span class="sxs-lookup"><span data-stu-id="3b567-206">Open the `.ENV` file and use the values that you copied earlier.</span></span> <span data-ttu-id="3b567-207">将 **CLIENT_ID** 设置为**应用程序（客户端）ID**，并将 **CLIENT_SECRET** 设置为客户端密码。</span><span class="sxs-lookup"><span data-stu-id="3b567-207">Set the **CLIENT_ID** to your **Application (client) ID**, and set the **CLIENT_SECRET** to your client secret.</span></span> <span data-ttu-id="3b567-208">该值**不**能用引号引起来。</span><span class="sxs-lookup"><span data-stu-id="3b567-208">The values should **not** be in quotation marks.</span></span> <span data-ttu-id="3b567-209">完成后，文件应当类似于以下示例：</span><span class="sxs-lookup"><span data-stu-id="3b567-209">When you are done, the file should be similar to the following:</span></span> 

    ```javascript
    CLIENT_ID=8791c036-c035-45eb-8b0b-265f43cc4824
    CLIENT_SECRET=X7szTuPwKNts41:-/fa3p.p@l6zsyI/p
    NODE_ENV=development
    ```

1. <span data-ttu-id="3b567-210">打开 `\public\javascripts\fallbackAuthDialog.js` 文件。</span><span class="sxs-lookup"><span data-stu-id="3b567-210">Open the `\public\javascripts\fallbackAuthDialog.js` file.</span></span> <span data-ttu-id="3b567-211">在 `msalConfig` 声明中，将占位符 $application_GUID here$ 替换为在注册加载项时复制的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="3b567-211">In the `msalConfig` declaration, replace the placeholder $application_GUID here$ with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="3b567-212">该值应该用引号引起来。</span><span class="sxs-lookup"><span data-stu-id="3b567-212">The value should be in quotation marks.</span></span>

1. <span data-ttu-id="3b567-213">打开加载项清单文件“manifest\manifest_local.xml”，然后滚动到该文件的底部。</span><span class="sxs-lookup"><span data-stu-id="3b567-213">Open the add-in manifest file "manifest\manifest_local.xml" and then scroll to the bottom of the file.</span></span> <span data-ttu-id="3b567-214">`</VersionOverrides>` 结束标记的正上方有以下标记：</span><span class="sxs-lookup"><span data-stu-id="3b567-214">Just above the `</VersionOverrides>` end tag, you'll find the following markup:</span></span>

    ```xml
    <WebApplicationInfo>
      <Id>$application_GUID here$</Id>
      <Resource>api://localhost:44355/$application_GUID here$</Resource>
      <Scopes>
          <Scope>Files.Read.All</Scope>
          <Scope>profile</Scope>
      </Scopes>
    </WebApplicationInfo>
    ```

1. <span data-ttu-id="3b567-215">将标记中的*两处*占位符“$application_GUID here$”均替换为在注册加载项时复制的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="3b567-215">Replace the placeholder "$application_GUID here$" *in both places* in the markup with the Application ID that you copied when you registered your add-in.</span></span> <span data-ttu-id="3b567-216">由于 ID 并不包含“$”符号，因此请勿包含它们。</span><span class="sxs-lookup"><span data-stu-id="3b567-216">The "$" symbols are not part of the ID, so do not include them.</span></span> <span data-ttu-id="3b567-217">这与在 web.config 中对 ClientID 和 Audience 所使用的 ID 相同。</span><span class="sxs-lookup"><span data-stu-id="3b567-217">This is the same ID you used in for the ClientID and Audience in the web.config.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3b567-218">**资源**值是注册加载项时设置的**应用程序 ID URI**。</span><span class="sxs-lookup"><span data-stu-id="3b567-218">The **Resource** value is the **Application ID URI** you set when you registered the add-in.</span></span> <span data-ttu-id="3b567-219">仅在通过 AppSource 销售加载项时，才使用**作用域**部分生成许可对话框。</span><span class="sxs-lookup"><span data-stu-id="3b567-219">The **Scopes** section is used only to generate a consent dialog box if the add-in is sold through AppSource.</span></span>

## <a name="code-the-client-side"></a><span data-ttu-id="3b567-220">编写客户端代码</span><span class="sxs-lookup"><span data-stu-id="3b567-220">Code the client-side</span></span>

### <a name="create-the-sso-logic"></a><span data-ttu-id="3b567-221">创建 SSO 逻辑</span><span class="sxs-lookup"><span data-stu-id="3b567-221">Create the SSO logic</span></span>

1. <span data-ttu-id="3b567-222">在代码编辑器中，打开文件 `public\javascripts\ssoAuthES6.js`。</span><span class="sxs-lookup"><span data-stu-id="3b567-222">In your code editor, open the file `public\javascripts\ssoAuthES6.js`.</span></span> <span data-ttu-id="3b567-223">它已经具有确保即使在 Internet Explorer 11 中也支持 Promise 的代码，并且具有 `Office.onReady` 调用，可将处理程序分配给加载项的唯一按钮。</span><span class="sxs-lookup"><span data-stu-id="3b567-223">It already has code that ensures that Promises are supported, even in Internet Explorer 11, and an `Office.onReady` call to assign a handler to the add-in's only button.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3b567-224">顾名思义，ssoAuthES6.js 使用 JavaScript ES6 语法，因为使用 `async` 和 `await` 可以最好地显示 SSO API 本质的简单性。</span><span class="sxs-lookup"><span data-stu-id="3b567-224">As the name suggests, the ssoAuthES6.js uses JavaScript ES6 syntax because using `async` and `await` best shows the essential simplicity of the SSO API.</span></span> <span data-ttu-id="3b567-225">启动 localhost 服务器时，此文件将转换为 ES5 语法，以便在 Internet Explorer 11 中运行该示例。</span><span class="sxs-lookup"><span data-stu-id="3b567-225">When the localhost server is started, this file is transpiled to ES5 syntax so that the sample will run in Internet Explorer 11.</span></span> 

1. <span data-ttu-id="3b567-226">将以下代码添加到 Office.onReady 方法：</span><span class="sxs-lookup"><span data-stu-id="3b567-226">Add the following code below the Office.onReady method:</span></span>

    ```javascript
    async function getGraphData() {
        try {
            
            // TODO 1: Tell Office to get a bootstrap token from Azure AD.
            
            // TODO 2: Attempt to exhange the bootstrap token for an 
            //         access token to Microsoft Graph.

            // TODO 3: Handle case where Microsoft Graph requires an 
            //         additional form of authentication.

            // TODO 4: Use the access token in a call to Microsoft Graph 
            //         or handle any error from the attempted token exchange.

        }
        catch(exception) {

            // TODO 5: Respond to exceptions thrown by the
            //         OfficeRuntime.auth.getAccessToken call.

        }
    }
    ```

1. <span data-ttu-id="3b567-227">将 `TODO 1` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-227">Replace `TODO 1` with the following code.</span></span> <span data-ttu-id="3b567-228">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-228">About this code, note:</span></span>

    - <span data-ttu-id="3b567-229">`OfficeRuntime.auth.getAccessToken` 指示 Office 从 Azure AD 获取引导令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-229">`OfficeRuntime.auth.getAccessToken` instructs Office to get a bootstrap token from Azure AD.</span></span> <span data-ttu-id="3b567-230">引导令牌类似于 ID令 牌，但是它具有值为 `access-as-user` 的 `scp`（作用域）属性。</span><span class="sxs-lookup"><span data-stu-id="3b567-230">A bootstrap token is similar to an ID token, but it has a `scp` (scope) property with the value `access-as-user`.</span></span> <span data-ttu-id="3b567-231">Web 应用程序可将此类令牌与 Microsoft Graph 的访问令牌进行交换。</span><span class="sxs-lookup"><span data-stu-id="3b567-231">This kind of token can be exchanged by a web application for an access token to Microsoft Graph.</span></span>
    - <span data-ttu-id="3b567-232">将 `allowSignInPrompt` 选项设置为 true 意味着如果当前没有任何用户登录到 Office，则 Office 将打开弹出窗口登录提示。</span><span class="sxs-lookup"><span data-stu-id="3b567-232">Setting the `allowSignInPrompt`option to true means that if no user is currently signed into Office, then Office will open a popup sign-in prompt.</span></span>
    - <span data-ttu-id="3b567-233">将 `forMSGraphAccess` 选项设置为 true 会向 Office 发出信号，表示加载项打算使用引导令牌来获取 Micrsoft Graph 的访问令牌，而不是仅将其用作 ID 令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-233">Setting the `forMSGraphAccess` option to true signals to Office that the add-in intends to use the bootstrap token to get an access token to Micrsoft Graph, instead of just using it as an ID token.</span></span> <span data-ttu-id="3b567-234">如果租户管理员未向加载项授予对 Microsoft Graph 的访问许可，则 `OfficeRuntime.auth.getAccessToken` 将返回错误 **13012**。</span><span class="sxs-lookup"><span data-stu-id="3b567-234">If the tenant administrator has not granted consent to the add-in's access to Microsoft Graph, then `OfficeRuntime.auth.getAccessToken` returns error **13012**.</span></span> <span data-ttu-id="3b567-235">该加载项可通过回退到备用的授权系统来做出响应，这是必需的，因为 Office 可以提示仅同意访问用户的 Azure AD 配置文件，而不是任何 Microsoft Graph 作用域。</span><span class="sxs-lookup"><span data-stu-id="3b567-235">The add-in can respond by falling back to an alternative system of authorization, which is necessary because Office can prompt only for consent to the user's Azure AD profile, not to any Microsoft Graph scopes.</span></span> <span data-ttu-id="3b567-236">回退授权系统要求用户重新登录，并且系统*会*提示用户同意访问 Microsoft Graph 作用域。</span><span class="sxs-lookup"><span data-stu-id="3b567-236">The fallback authorization system requires the user to sign in again and the user *can* be prompted to consent to Micrsoft Graph scopes.</span></span> <span data-ttu-id="3b567-237">因此，`forMSGraphAccess` 选项可确保加载项不会进行令牌交换，交换会因缺乏许可而失败。</span><span class="sxs-lookup"><span data-stu-id="3b567-237">So, the `forMSGraphAccess` option ensures that the add-in won't make a token exchange that will fail due to lack of consent.</span></span> <span data-ttu-id="3b567-238">（由于先前步骤中已授予管理员许可，此加载项不会发生此情况。</span><span class="sxs-lookup"><span data-stu-id="3b567-238">(Since you granted administrator consent in an earlier step, this scenario won't happen for this add-in.</span></span> <span data-ttu-id="3b567-239">但这里包含了一个选项来说明最佳实践。）</span><span class="sxs-lookup"><span data-stu-id="3b567-239">But the option is included here anyway to illustrate a best practice.)</span></span>

    ```javascript
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, forMSGraphAccess: true }); 
    ```

1. <span data-ttu-id="3b567-240">将 `TODO 2` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-240">Replace `TODO 2` with the following code.</span></span> <span data-ttu-id="3b567-241">将在后续步骤中创建 `getGraphToken` 方法。</span><span class="sxs-lookup"><span data-stu-id="3b567-241">You'll create the `getGraphToken` method in a later step.</span></span>

    ```javascript
    let exchangeResponse = await getGraphToken(bootstrapToken);
    ```

1. <span data-ttu-id="3b567-242">将 `TODO 3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-242">Replace `TODO 3` with the following.</span></span> <span data-ttu-id="3b567-243">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-243">About this code, note:</span></span> 

    - <span data-ttu-id="3b567-244">如果已将 Office 365 租户配置为要求多重身份验证，则 `exchangeResponse` 将包括一个 `claims` 属性，其中包含有关其他所需因素的信息。</span><span class="sxs-lookup"><span data-stu-id="3b567-244">If the Office 365 tenant has been configured to require multifactor authentication, then the `exchangeResponse` will include a `claims` property with information about the additional required factors.</span></span> <span data-ttu-id="3b567-245">在这种情况下，应该再次调用 `OfficeRuntime.auth.getAccessToken`，并将 `authChallenge` 选项设置为 claims 属性的值。</span><span class="sxs-lookup"><span data-stu-id="3b567-245">In that case, `OfficeRuntime.auth.getAccessToken` should be called again with the `authChallenge` option set to the value of the claims property.</span></span> <span data-ttu-id="3b567-246">这就指示 AAD 提示用户进行所有必需形式的身份验证。</span><span class="sxs-lookup"><span data-stu-id="3b567-246">This tells AAD to prompt the user for all required forms of authentication.</span></span>

    ```javascript
    if (exchangeResponse.claims) {
        let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
        exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }
    ```

1. <span data-ttu-id="3b567-247">将 `TODO 4` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-247">Replace `TODO 4` with the following.</span></span> <span data-ttu-id="3b567-248">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-248">About this code, note:</span></span> 

    - <span data-ttu-id="3b567-249">将在后续步骤中创建 `handleAADErrors` 方法。</span><span class="sxs-lookup"><span data-stu-id="3b567-249">You'll create the `handleAADErrors` method in a later step.</span></span> <span data-ttu-id="3b567-250">Azure AD 错误作为 HTTP 代码 200 响应返回给客户端。</span><span class="sxs-lookup"><span data-stu-id="3b567-250">Azure AD errors are returned to the client as HTTP code 200 Responses.</span></span> <span data-ttu-id="3b567-251">它们不会引发错误，因此不会触发 `getGraphData` 方法的 `catch` 块。</span><span class="sxs-lookup"><span data-stu-id="3b567-251">They do not throw errors, so they do not trigger the `catch` block of the `getGraphData` method.</span></span>
    - <span data-ttu-id="3b567-252">将在后续步骤中创建 `makeGraphApiCall` 方法。</span><span class="sxs-lookup"><span data-stu-id="3b567-252">You'll create the `makeGraphApiCall` method in a later step.</span></span> <span data-ttu-id="3b567-253">它将对 MS Graph 终结点进行 AJAX 调用。</span><span class="sxs-lookup"><span data-stu-id="3b567-253">It makes an AJAX call to the MS Graph endpoint.</span></span> <span data-ttu-id="3b567-254">在该调用的 `.fail` 回调中捕获到错误，而不是在 `getGraphData` 方法的 `catch` 块中。</span><span class="sxs-lookup"><span data-stu-id="3b567-254">Errors are caught in the `.fail` callback of that call, not in the `catch` block of the `getGraphData` method.</span></span>

    ```javascript
    if (exchangeResponse.error) {
        handleAADErrors(exchangeResponse);
    } 
    else {
        makeGraphApiCall(exchangeResponse.access_token);
    }
    ```

1. <span data-ttu-id="3b567-255">将 `TODO 5` 替换为以下代码</span><span class="sxs-lookup"><span data-stu-id="3b567-255">Replace `TODO 5` with the following</span></span>

    - <span data-ttu-id="3b567-256">来自 `getAccessToken` 调用的错误将具有 `code` 属性，其错误号通常处于 13xxx 范围内。</span><span class="sxs-lookup"><span data-stu-id="3b567-256">Errors from the call of `getAccessToken` will have a `code` property with an error number, typically in the 13xxx range.</span></span> <span data-ttu-id="3b567-257">将在后续步骤中创建 `handleClientSideErrors` 方法。</span><span class="sxs-lookup"><span data-stu-id="3b567-257">You'll create the `handleClientSideErrors` method in a later step.</span></span>
    - <span data-ttu-id="3b567-258">`showMessage` 方法在任务窗格上显示文本。</span><span class="sxs-lookup"><span data-stu-id="3b567-258">The `showMessage` method displays text on the task pane.</span></span>

    ```javascript
    if (exception.code) { 
        handleClientSideErrors(exception);
    }
    else {
        showMessage("EXCEPTION: " + JSON.stringify(exception));
    }
    ```

1. <span data-ttu-id="3b567-259">在 `getGraphData` 方法下方，添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="3b567-259">Below the `getGraphData` method, add the following function.</span></span> <span data-ttu-id="3b567-260">请注意，`/auth` 是服务器端 Express 路由，用于 Azure AD 引导令牌与 Microsoft Graph 访问令牌进行交换。</span><span class="sxs-lookup"><span data-stu-id="3b567-260">Note that `/auth` is a server-side Express route that exhanges the bootstrap token with Azure AD for an access token to Microsoft Graph.</span></span>

    ```javascript
    async function getGraphToken(bootstrapToken) {
        let response = await $.ajax({type: "GET", 
            url: "/auth",
            headers: {"Authorization": "Bearer " + bootstrapToken }, 
            cache: false
        });
        return response;
    }
    ```

1. <span data-ttu-id="3b567-261">在 `getGraphToken` 方法下方，添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="3b567-261">Below the `getGraphToken` method, add the following function.</span></span> <span data-ttu-id="3b567-262">请注意，`error.code` 是一个数字，通常处于 13xxx 范围内。</span><span class="sxs-lookup"><span data-stu-id="3b567-262">Note that `error.code` is a number, usually in the range 13xxx.</span></span>

    ```javascript
    function handleClientSideErrors(error) {
        switch (error.code) {

            // TODO 6: Handle errors where the add-in should NOT invoke 
            //         the alternative system of authorization.

            // TODO 7: Handle errors where the add-in should invoke 
            //         the alternative system of authorization.

        }
    }
    ```
1. <span data-ttu-id="3b567-263">将 `TODO 6` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-263">Replace `TODO 6` with the following code.</span></span> <span data-ttu-id="3b567-264">有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="3b567-264">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md).</span></span> 

    ```javascript
    case 13001:
        // No one is signed into Office. If the add-in cannot be effectively used when no one 
        // is logged into Office, then the first call of getAccessToken should pass the 
        // `allowSignInPrompt: true` option. Since this add-in does that, you should not see
        // this error. 
        showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
        break;
    case 13002:
        // OfficeRuntime.auth.getAccessToken was called with the allowConsentPrompt 
        // option set to true. But, the user aborted the consent prompt. 
        showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
        break;
    case 13006:
        // Only seen in Office on the Web.
        showMessage("Office on the Web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
        break;
    case 13008:
        // The OfficeRuntime.auth.getAccessToken method has already been called and 
        // that call has not completed yet. Only seen in Office on the web.
        showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
        break;
    case 13010:
        // Only seen in Office on the web.
        showMessage("Follow the instructions to change your browser's zone configuration.");
        break;
    ```

1. <span data-ttu-id="3b567-265">将 `TODO 7` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-265">Replace `TODO 7` with the following code.</span></span> <span data-ttu-id="3b567-266">有关这些错误的详细信息，请参阅[对 Office 加载项中的 SSO 进行故障排除](troubleshoot-sso-in-office-add-ins.md)。函数 `dialogFallback` 用于调用备用授权系统。</span><span class="sxs-lookup"><span data-stu-id="3b567-266">For more information about these errors, see [Troubleshoot SSO in Office Add-ins](troubleshoot-sso-in-office-add-ins.md). The function `dialogFallback` invokes the alternative system of authorization.</span></span> <span data-ttu-id="3b567-267">在此加载项中，回退系统将打开一个对话框，它要求用户登录（即使用户已登录），并使用 msal.js 和隐式流来获取 Microsoft Graph 访问令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-267">In this add-in, the fallback system opens a dialog which requires the user to sign in, even if the user already is, and uses msal.js and the Implicit Flow to get an access token to Microsoft Graph.</span></span>

    ```javascript
    default:
    // For all other errors, including 13000, 13003, 13005, 13007, 13012, 
    // and 50001, fall back to non-SSO sign-in.
    dialogFallback();
    break;
    ```

1. <span data-ttu-id="3b567-268">在 `handleClientSideErrors` 函数下方，添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="3b567-268">Below the `handleClientSideErrors` function, add the following function.</span></span> 

    ```javascript
    function handleAADErrors(exchangeResponse) {

    // TODO 8: Handle case where the bootstrap token is expired.

    // TODO 9: Handle all other Azure AD errors.
    
    }
    ```

1. <span data-ttu-id="3b567-269">在极少数情况下，Office 缓存的引导令牌在 Office 验证时未过期，但是会在到达 Azure AD 进行交换时过期。</span><span class="sxs-lookup"><span data-stu-id="3b567-269">On rare occasions the bootstrap token that Office has cached is unexpired when Office validates it, but expires by the time it reaches Azure AD for exchange.</span></span> <span data-ttu-id="3b567-270">Azure AD 将以错误 **AADSTS500133** 做出响应。</span><span class="sxs-lookup"><span data-stu-id="3b567-270">Azure AD will respond with error **AADSTS500133**.</span></span> <span data-ttu-id="3b567-271">在这种情况下，加载项应仅以递归方式调用 `getGraphData`。</span><span class="sxs-lookup"><span data-stu-id="3b567-271">In this case, the add-in should simply recursively call `getGraphData`.</span></span> <span data-ttu-id="3b567-272">由于缓存的引导令牌现在已过期，Office 将从 Azure AD 获取一个新令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-272">Since the cached bootstrap token is now expired, Office will get a new one from Azure AD.</span></span> <span data-ttu-id="3b567-273">将 `TODO 8` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-273">So replace `TODO 8` with the following.</span></span> 

    ```javascript
    if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)       
    {
        getGraphData();
    }
    ```

1. <span data-ttu-id="3b567-274">若要确保加载项不会进入 `getGraphData` 调用的无限循环，该加载项应跟踪调用 `getGraphData` 的次数，并确保不会多次对它进行递归式调用。</span><span class="sxs-lookup"><span data-stu-id="3b567-274">To ensure that the add-in doesn't enter an infinite loop of calls to `getGraphData`, the add-in should keep track of how many times `getGraphData` has been called and be sure that is not called recursively called more than once.</span></span> <span data-ttu-id="3b567-275">因此，应在 `handleAADErrors` 和 `getGraphData` 函数的全局范围内创建计数器变量。</span><span class="sxs-lookup"><span data-stu-id="3b567-275">So, create a counter variable in a scope that is global to the `handleAADErrors` and `getGraphData` functions.</span></span> <span data-ttu-id="3b567-276">全局变量的理想位置就在 `Office.onReady` 方法调用的正下方。</span><span class="sxs-lookup"><span data-stu-id="3b567-276">A good place for global variables is just below the `Office.onReady` method call.</span></span>

    ```javascript
    let retryGetAccessToken = 0;
    ```

1. <span data-ttu-id="3b567-277">在 `handleAADErrors` 方法中更改 `if` 结构，以使其：</span><span class="sxs-lookup"><span data-stu-id="3b567-277">Change the `if` structure in the `handleAADErrors` method so that it:</span></span>

    - <span data-ttu-id="3b567-278">在调用 `getGraphData` 之前递增计数器。</span><span class="sxs-lookup"><span data-stu-id="3b567-278">Increments the counter just before it calls `getGraphData`.</span></span>
    - <span data-ttu-id="3b567-279">执行测试以确保尚未对 `getGraphData` 进行第二次调用。</span><span class="sxs-lookup"><span data-stu-id="3b567-279">Tests to ensure that `getGraphData` has not already been called a second time.</span></span> 

    <span data-ttu-id="3b567-280">因此，`if` 结构的最终版本应如下所示：</span><span class="sxs-lookup"><span data-stu-id="3b567-280">So the final version of the `if` structure should look like the following:</span></span>

    ```javascript
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    ```

1. <span data-ttu-id="3b567-281">将 `TODO 9` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-281">Replace `TODO 9` with the following.</span></span> 

    ```javascript
    else {                
        dialogFallback();
    }
    ```

1. <span data-ttu-id="3b567-282">保存并关闭此文件。</span><span class="sxs-lookup"><span data-stu-id="3b567-282">Save and close the file.</span></span>

### <a name="get-the-data-and-add-it-to-the-office-document"></a><span data-ttu-id="3b567-283">获取数据并将其添加到 Office 文档</span><span class="sxs-lookup"><span data-stu-id="3b567-283">Get the data and add it to the Office document</span></span>

1. <span data-ttu-id="3b567-284">在 `public\javascripts` 文件夹中，创建名为 `data.js` 的新文件。</span><span class="sxs-lookup"><span data-stu-id="3b567-284">In the `public\javascripts` folder, create a new file named `data.js`.</span></span>

1. <span data-ttu-id="3b567-285">将以下函数添加到文件中。</span><span class="sxs-lookup"><span data-stu-id="3b567-285">Add the following function to the file.</span></span> <span data-ttu-id="3b567-286">这是 `getGraphData` 函数在获得 Microsoft Graph 访问令牌后调用的函数。</span><span class="sxs-lookup"><span data-stu-id="3b567-286">This is the function that is called by the `getGraphData` function when it has acquired an access token to Microsoft Graph.</span></span> 

    ```javascript
    function makeGraphApiCall(accessToken) {
        $.ajax(

            // TODO 10: Call an Express route on the add-in's server-side 
            //          code and pass the access token to Microsoft Graph.

        )
        .done(function (response) {

            // TODO 11: Write the data received from Microsoft Graph to 
            //          the Office document.

        })
        .fail(function (errorResult) {
            showMessage("Error from Microsoft Graph: " + JSON.stringify(errorResult));
        });
    }
    ```

1. <span data-ttu-id="3b567-287">将 `TODO 10` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-287">Replace `TODO 10` with the following.</span></span> <span data-ttu-id="3b567-288">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-288">About this code, note:</span></span> 

    - <span data-ttu-id="3b567-289">此对象是 `$.ajax` 方法的参数。</span><span class="sxs-lookup"><span data-stu-id="3b567-289">This object is the parameter to the `$.ajax` method.</span></span>
    - <span data-ttu-id="3b567-290">`/getuserdata` 是你在后续步骤中创建的加载项服务器上的 Express 路由。</span><span class="sxs-lookup"><span data-stu-id="3b567-290">The `/getuserdata` is an Express route on the add-in's server that you create in a later step.</span></span> <span data-ttu-id="3b567-291">它将调用 Microsoft Graph 终结点，并在其调用中包含访问令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-291">It will call a Microsoft Graph endpoint and include the access token in its call.</span></span> 

    ```javascript
    {
        type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }
    ```

1. <span data-ttu-id="3b567-292">将 `TODO11` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-292">Replace `TODO11` with the following.</span></span> <span data-ttu-id="3b567-293">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-293">About this code, note:</span></span>

    - <span data-ttu-id="3b567-294">`writeFileNamesToOfficeDocument` 会将来自 Graph 的数据插入到 Office 文档中。</span><span class="sxs-lookup"><span data-stu-id="3b567-294">The `writeFileNamesToOfficeDocument` will insert the data from Graph into the Office document.</span></span> <span data-ttu-id="3b567-295">它在 `public\javascripts\document.js` 文件中定义。</span><span class="sxs-lookup"><span data-stu-id="3b567-295">It is defined in the `public\javascripts\document.js` file.</span></span> 
    - <span data-ttu-id="3b567-296">如果 `writeFileNamesToOfficeDocument` 返回错误，它将以“无法将文件名添加到文档中”开头。</span><span class="sxs-lookup"><span data-stu-id="3b567-296">If `writeFileNamesToOfficeDocument` returns an error, it will begin with "Unable to add filenames to document."</span></span>

    ```javascript
    writeFileNamesToOfficeDocument(response)
    .then(function () { 
        showMessage("Your data has been added to the document."); 
    })
    .catch(function (error) {        
        showMessage(error);
    });
    ```

1. <span data-ttu-id="3b567-297">保存并关闭此文件。</span><span class="sxs-lookup"><span data-stu-id="3b567-297">Save and close the file.</span></span>

## <a name="code-the-server-side"></a><span data-ttu-id="3b567-298">编写服务器端代码</span><span class="sxs-lookup"><span data-stu-id="3b567-298">Code the server-side</span></span>

### <a name="create-the-auth-router-and-the-token-exchange-logic"></a><span data-ttu-id="3b567-299">创建身份验证路由器和令牌交换逻辑</span><span class="sxs-lookup"><span data-stu-id="3b567-299">Create the auth router and the token exchange logic</span></span>

1. <span data-ttu-id="3b567-300">打开文件 `routes\authRoute.js`，然后在 `require` 语句正下方和 `module.exports` 语句上方添加以下路由函数。</span><span class="sxs-lookup"><span data-stu-id="3b567-300">Open the file `routes\authRoute.js` and add the following route function just below the `require` statements and above the `module.exports` statement.</span></span> <span data-ttu-id="3b567-301">请注意，`router.get` 的 URL 参数是“/”。</span><span class="sxs-lookup"><span data-stu-id="3b567-301">Note that the URL parameter of `router.get` is '/'.</span></span> <span data-ttu-id="3b567-302">由于此路由是在负责处理 URL“/auth”的所有 HTTP 请求的路由器中定义的，因此该路由可有效处理“/auth”的所有请求。</span><span class="sxs-lookup"><span data-stu-id="3b567-302">Since this route is being defined in a router that will handle all HTTP Requests for the URL '/auth', this route effectively handles all requests for '/auth'.</span></span> <span data-ttu-id="3b567-303">先前创建的客户端 `getGraphToken` 函数将调用此路由。</span><span class="sxs-lookup"><span data-stu-id="3b567-303">The client-side `getGraphToken` function that you created earlier calls this route.</span></span>  

    ```javascript
    router.get('/', async function(req, res, next) {

        // TODO 12: Test for the presence of the Authorization header.

        // TODO 13: Create the hidden form that will be sent to Azure AD 
        //          to request the access token in exhange for the 
        //          bootstrap token.

        // TODO 14: Send the POST request to Azure AD and relay the 
        //          access token (or an error) to the client.

    });
    ```

1. <span data-ttu-id="3b567-304">将 `TODO 12` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-304">Replace `TODO 12` with the following code.</span></span>

    ```javascript
    const authorization = req.get('Authorization');
    if (authorization == null) {
        let error = new Error('No Authorization header was found.');
        next(error);
    } 
    ```

1. <span data-ttu-id="3b567-305">将 `TODO 13` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-305">Replace `TODO 13` with the following code.</span></span> <span data-ttu-id="3b567-306">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-306">About this code, note:</span></span> 

    - <span data-ttu-id="3b567-307">这是一个长 `else` 块的开头，但是结尾 `}` 尚未结束，因为你将向其添加更多代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-307">This is the beginning of a long `else` block, but the closing `}` is not at the end yet because you will be adding more code to it.</span></span> 
    - <span data-ttu-id="3b567-308">`authorization` 字符串是“持有者”，后跟引导令牌，因此 `else` 块的第一行将令牌分配给 `jwt`。</span><span class="sxs-lookup"><span data-stu-id="3b567-308">The `authorization` string is "Bearer " followed by the bootstrap token, so the first line of the `else` block is assigning the token to the `jwt`.</span></span> <span data-ttu-id="3b567-309">（“JWT”代表“JSON Web 令牌”。）</span><span class="sxs-lookup"><span data-stu-id="3b567-309">("JWT" stands for "JSON Web Token".)</span></span>
    - <span data-ttu-id="3b567-310">两个 `process.env.*` 值是你配置加载项时分配的常量。</span><span class="sxs-lookup"><span data-stu-id="3b567-310">The two `process.env.*` values are the constants that you assigned when you configured the add-in.</span></span> 
    - <span data-ttu-id="3b567-311">`requested_token_use` 窗体参数设置为“on_behalf_of”。</span><span class="sxs-lookup"><span data-stu-id="3b567-311">The `requested_token_use` form parameter is set to 'on_behalf_of'.</span></span> <span data-ttu-id="3b567-312">它告知 Azure AD 加载项正在使用“代理流”请求 Microsoft Graph 访问令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-312">This tells Azure AD that the add-in is requesting an access token to Microsoft Graph using the On-Behalf-Of Flow.</span></span> <span data-ttu-id="3b567-313">通过验证分配给 `assertion` 窗体参数的引导令牌是否具有设置为 `access-as-user` 的 `scp` 属性，Azure 将对此做出响应。</span><span class="sxs-lookup"><span data-stu-id="3b567-313">Azure will respond by validating that the bootstrap token, which is assigned to `assertion` form parameter, has a `scp` property that is set to `access-as-user`.</span></span>
    - <span data-ttu-id="3b567-314">`scope` 窗体参数设置为“Files.Read.All”，这是加载项唯一需要的 Microsoft Graph 作用域。</span><span class="sxs-lookup"><span data-stu-id="3b567-314">The `scope` form parameter is set to 'Files.Read.All' which is the only Microsoft Graph scope that the add-in needs.</span></span>

    ```javascript
     else {
        const [schema, jwt] = authorization.split(' ');
        const formParams = {
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        assertion: jwt,
        requested_token_use: 'on_behalf_of',
        scope: ['Files.Read.All'].join(' ')
        };
    ```

1. <span data-ttu-id="3b567-315">将 `TODO 14` 替换为以下代码，它将完成 `else` 块。</span><span class="sxs-lookup"><span data-stu-id="3b567-315">Replace `TODO 14` with the following code, which completes the `else` block.</span></span> <span data-ttu-id="3b567-316">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-316">About this code, note:</span></span>

    - <span data-ttu-id="3b567-317">常量 `tenant` 设置为“通用”，因为你在 Azure AD 中注册加载项时已将其配置为多租户；特别是当你将“**支持的帐户类型**”设置为“**任何组织目录中的帐户和个人 Microsoft 帐户（例如，Skype、Xbox、Outlook.com）**”时。</span><span class="sxs-lookup"><span data-stu-id="3b567-317">The const `tenant` is set to 'common' because you configured the add-in as multitenant when you registered it with Azure AD; specifically when you set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)**.</span></span> <span data-ttu-id="3b567-318">如果改为选择仅支持在其中注册加载项的同一 Office 365 租户中的帐户，则此代码 `tenant` 将设置为租户的 GUID。</span><span class="sxs-lookup"><span data-stu-id="3b567-318">If you had instead chosen to support only accounts in the same Office 365 tenancy where the add-in is registered, then in this code `tenant` would be set to the GUID of the tenant.</span></span> 
    - <span data-ttu-id="3b567-319">如果 POST 请求没有错误，那么 Azure AD 的响应将转换为 JSON 并发送到客户端。</span><span class="sxs-lookup"><span data-stu-id="3b567-319">If the POST request does not error, then the response from Azure AD is converted to JSON and sent to the client.</span></span> <span data-ttu-id="3b567-320">此 JSON 对象具有 `access_token` 属性，Azure AD 已为其分配 Microsoft Graph 访问令牌。</span><span class="sxs-lookup"><span data-stu-id="3b567-320">This JSON object has an `access_token` property to which Azure AD has assigned the access token to Microsoft Graph.</span></span>

    ```javascript
        const stsDomain = 'https://login.microsoftonline.com';
        const tenant = 'common';
        const tokenURLSegment = 'oauth2/v2.0/token';

        try {
            const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: 'POST',
                body: form(formParams),
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/x-www-form-urlencoded'
                }
            });
            const json = await tokenResponse.json();
            
            res.send(json);
        }
        catch(error) {
            res.status(500).send(error);
        }
    }
    ```

1. <span data-ttu-id="3b567-321">保存并关闭此文件。</span><span class="sxs-lookup"><span data-stu-id="3b567-321">Save and close the file.</span></span>

### <a name="create-the-route-that-will-fetch-the-data-from-microsoft-graph"></a><span data-ttu-id="3b567-322">创建将从 Microsoft Graph 获取数据的路由</span><span class="sxs-lookup"><span data-stu-id="3b567-322">Create the route that will fetch the data from Microsoft Graph</span></span>

1. <span data-ttu-id="3b567-323">打开项目根目录中的 `app.js` 文件。</span><span class="sxs-lookup"><span data-stu-id="3b567-323">Open the file `app.js` in the root of the project.</span></span> <span data-ttu-id="3b567-324">在“/dialog.html”路由的正下方，添加以下路由。</span><span class="sxs-lookup"><span data-stu-id="3b567-324">Just below the route for '/dialog.html', add the following route.</span></span> <span data-ttu-id="3b567-325">此路由由你在前面步骤中创建的 `makeGraphApiCall` 函数调用。</span><span class="sxs-lookup"><span data-stu-id="3b567-325">This route is called by the `makeGraphApiCall` function that you created in an earlier step.</span></span>

    ```javascript
    app.get('/getuserdata', async function(req, res, next) {
        
        // TODO 15: Send a request to the Microsoft Graph REST endpoint.

        // TODO 16: Trim excess information from the returned data and relay it
        //          to the client.
        
    });
    ```

1. <span data-ttu-id="3b567-326">将 `TODO 15` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-326">Replace `TODO 15` with the following.</span></span> <span data-ttu-id="3b567-327">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-327">About this code, note:</span></span>

    - <span data-ttu-id="3b567-328">此路由的调用方 `makeGraphApiCall` 将 Microsoft Graph 访问令牌作为名为“access_token”的标头添加到 HTTP 请求中。</span><span class="sxs-lookup"><span data-stu-id="3b567-328">The caller of this route, `makeGraphApiCall`, added the access token to Microsoft Graph to the HTTP Request as a header named "access_token".</span></span>
    - <span data-ttu-id="3b567-329">`getGraphData` 函数在 `msgraph-helper.js` 文件中定义。</span><span class="sxs-lookup"><span data-stu-id="3b567-329">The `getGraphData` function is defined in the `msgraph-helper.js` file.</span></span> <span data-ttu-id="3b567-330">（此函数与在 `ssoAuthES6.js` 文件中定义的客户端 `getGraphData` 函数不同。）</span><span class="sxs-lookup"><span data-stu-id="3b567-330">(This is not the same function as the client-side `getGraphData` function that you defined in the `ssoAuthES6.js` file.)</span></span>
    - <span data-ttu-id="3b567-331">`queryParamsSegment` 的最后一个参数是硬编码值。</span><span class="sxs-lookup"><span data-stu-id="3b567-331">The last parameter, for `queryParamsSegment`, is hardcoded.</span></span> <span data-ttu-id="3b567-332">如果你在生产加载项中重复使用此代码，并且 `queryParamsSegment` 的任何部分均来自用户输入，请确保它已被清理，以便它不能用于响应标头注入攻击。</span><span class="sxs-lookup"><span data-stu-id="3b567-332">If you reuse this code in a production add-in and any part of `queryParamsSegment` comes from user input, be sure that it is sanitized so that it cannot be used in a Response header injection attack.</span></span>
    - <span data-ttu-id="3b567-333">通过仅指定所需的属性（“名称”）以及仅前 10 个文件夹或文件名，该代码可最大限度地减少来自 Microsoft Graph 的数据量。</span><span class="sxs-lookup"><span data-stu-id="3b567-333">The code minimizes the data that must come from Microsoft Graph by specifying only the property we need ("name") and only the top 10 folder or file names.</span></span>

    ```javascript
    const graphToken = req.get('access_token');    
    const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");
    ```

1. <span data-ttu-id="3b567-334">将 `TODO 16` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="3b567-334">Replace `TODO 16` with the following.</span></span> <span data-ttu-id="3b567-335">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="3b567-335">About this code, note:</span></span>

    - <span data-ttu-id="3b567-336">如果 Microsoft Graph 返回错误（例如无效或过期的令牌），则返回的对象中将有一个 code 属性设置为 HTTP 状态（例如 401）。</span><span class="sxs-lookup"><span data-stu-id="3b567-336">If Microsoft Graph returns an error, such as invalid or expired token, there will be a code property in the returned object set to a HTTP status (e.g., 401).</span></span> <span data-ttu-id="3b567-337">代码会将错误转发给客户端。</span><span class="sxs-lookup"><span data-stu-id="3b567-337">The code relays the error to the client.</span></span> <span data-ttu-id="3b567-338">它将在 `makeGraphApiCall` 的 `.fail` 回调中被捕获。</span><span class="sxs-lookup"><span data-stu-id="3b567-338">It will be caught in the `.fail` callback of `makeGraphApiCall`.</span></span>
    - <span data-ttu-id="3b567-339">Microsoft Graph 数据包含该加载项不需要的 OData 元数据和 eTag，因此代码将构造一个新数组，其中仅包含要发送到客户端的文件名。</span><span class="sxs-lookup"><span data-stu-id="3b567-339">Microsoft Graph data includes OData metadata and eTags that the add-in does not need, so the code constructs a new array containing only the file names to send to the client.</span></span>

    ```javascript
    if (graphData.code) {
        next(createError(graphData.code, "Microsoft Graph error: " + JSON.stringify(graphData)));
    }
    else {
        const itemNames = [];
        const oneDriveItems = graphData['value'];
        for (let item of oneDriveItems) {
            itemNames.push(item['name']);
        }

        res.send(itemNames)
    }
    ```

1. <span data-ttu-id="3b567-340">保存并关闭此文件。</span><span class="sxs-lookup"><span data-stu-id="3b567-340">Save and close the file.</span></span>

## <a name="run-the-project"></a><span data-ttu-id="3b567-341">运行项目</span><span class="sxs-lookup"><span data-stu-id="3b567-341">Run the project</span></span>

1. <span data-ttu-id="3b567-342">请确保 OneDrive 中有一些文件，以便可以验证结果。</span><span class="sxs-lookup"><span data-stu-id="3b567-342">Ensure that you have some files in your OneDrive so that you can verify the results.</span></span>

1. <span data-ttu-id="3b567-343">在 `\Complete` 文件夹的根目录中打开命令提示符。</span><span class="sxs-lookup"><span data-stu-id="3b567-343">Open a command prompt in the root of the `\Complete` folder.</span></span> 

1. <span data-ttu-id="3b567-344">运行命令 `npm start`。</span><span class="sxs-lookup"><span data-stu-id="3b567-344">Run the command `npm start`.</span></span> 

1. <span data-ttu-id="3b567-345">需要将加载项旁加载到 Office 应用程序（Excel、Word 或 PowerPoint），以便对其进行测试。</span><span class="sxs-lookup"><span data-stu-id="3b567-345">You need to sideload the add-in into an Office application (Excel, Word, or PowerPoint) to test it.</span></span> <span data-ttu-id="3b567-346">说明取决于你的平台。</span><span class="sxs-lookup"><span data-stu-id="3b567-346">The instructions depend on your platform.</span></span> <span data-ttu-id="3b567-347">在[旁加载 Office 加载项以供测试](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing)中有指向说明的链接。</span><span class="sxs-lookup"><span data-stu-id="3b567-347">There are links to instructions at [Sideload an Office Add-in for Testing](../testing/test-debug-office-add-ins.md#sideload-an-office-add-in-for-testing).</span></span>

1. <span data-ttu-id="3b567-348">在 Office 应用程序的“**主页**”功能区上，选择“**SSO Node.js**”组中的“**显示加载项**”按钮以打开任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="3b567-348">In the Office application, on the **Home** ribbon, select the **Show Add-in** button in the **SSO Node.js** group to open the task pane add-in.</span></span>

1. <span data-ttu-id="3b567-349">单击“**获取 OneDrive 文件名**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3b567-349">Click the **Get OneDrive File Names** button.</span></span> <span data-ttu-id="3b567-350">如果你使用工作或学校 (Office 365) 帐户或 Microsoft 帐户登录 Office，并且 SSO 工作正常，则 OneDrive for Business 中的前 10 个文件和文件夹名称将插入文档中。</span><span class="sxs-lookup"><span data-stu-id="3b567-350">If you are logged into Office with either a Work or School (Office 365) account or Microsoft Account, and SSO is working as expected, the first 10 file and folder names in your OneDrive for Business are inserted into the document.</span></span> <span data-ttu-id="3b567-351">（第一次登录可能需要长达 15 秒钟。）如果你未登录，或者处于不支持 SSO 的情形中，或者 SSO 出于任何原因无法正常工作，则系统将提示你登录。</span><span class="sxs-lookup"><span data-stu-id="3b567-351">(It may take as much as 15 seconds the first time.) If you are not logged in, or you are in a scenario that does not support SSO, or SSO is not working for any reason, you will be prompted to log in.</span></span> <span data-ttu-id="3b567-352">登录后，将显示文件和文件夹名称。</span><span class="sxs-lookup"><span data-stu-id="3b567-352">After you log in, the file and folder names appear.</span></span>

> [!NOTE]
> <span data-ttu-id="3b567-353">如果先前使用其他 ID 登录过 Office，并且当时打开的一些 Office 应用现在仍处于打开状态，Office 可能无法可靠地更改 ID，即使看似已更改过，也不例外。</span><span class="sxs-lookup"><span data-stu-id="3b567-353">If you were previously signed into Office with a different ID, and some Office applications that were open at the time are still open, Office may not reliably change your ID even if it appears to have done so.</span></span> <span data-ttu-id="3b567-354">在这种情况下，可能无法调用 Microsoft Graph，或者可能返回以前 ID 的数据。</span><span class="sxs-lookup"><span data-stu-id="3b567-354">If this happens, the call to Microsoft Graph may fail or data from the previous ID may be returned.</span></span> <span data-ttu-id="3b567-355">为了防止发生这种情况，请务必先*关闭其他所有 Office 应用程序*，然后再按“**获取 OneDrive 文件名**”。</span><span class="sxs-lookup"><span data-stu-id="3b567-355">To prevent this, be sure to *close all other Office applications* before you press **Get OneDrive File Names**.</span></span>
