---
title: Office 加载项中的身份验证和授权概述
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 2733f8af9f236347e52269c9e73b322b4310e2a9
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/13/2019
ms.locfileid: "36302933"
---
# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a><span data-ttu-id="b5fe3-102">Office 加载项中的身份验证和授权概述</span><span class="sxs-lookup"><span data-stu-id="b5fe3-102">Overview of identity, authentication, and authorization in Office 2013</span></span>

<span data-ttu-id="b5fe3-103">Web 应用程序和 Office 加载项默认允许匿名访问，但你可要求用户通过登录进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-103">Web applications and, hence, Office Add-ins allow anonymous access by default, but you can require users to authenticate with a login.</span></span> <span data-ttu-id="b5fe3-104">具体而言，你可要求用户使用 Microsoft 帐户或工作/学校 (Office 365) 帐户登录。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-104">In particular, you can require that your users be logged in with either a Microsoft Account or a Work or School (Office 365) account.</span></span> <span data-ttu-id="b5fe3-105">此任务被称为“用户身份验证”，因为它让加载项能够知道用户的身份。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-105">This task is called user authentication because it enables the add-in to know who the user is.</span></span>

<span data-ttu-id="b5fe3-106">你的加载项还能够获得用户的同意来访问其 Microsoft Graph 数据（例如其 Office 365 个人资料、OneDrive 文件和 SharePoint 数据），或者访问 Google、Facebook、领英、SalesForce 和 GitHub 等其他外部源中的数据。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-106">Your add-in can also get the user's consent to access their Microsoft Graph data (such as their Office 365 profile, OneDrive files, and SharePoint data) or to data in other external sources such as Google, Facebook, LinkedIn, SalesForce, and GitHub.</span></span> <span data-ttu-id="b5fe3-107">此任务被称为“加载项（或应用）授权”，因为要获得授权的是*加载项*，而不是用户。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-107">This task is called add-in (or app) authorization, because it is the *add-in* that is being authorized, not the user.</span></span>

<span data-ttu-id="b5fe3-108">有两种方式可用来完成这些身份验证。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-108">You have a choice of two ways to accomplish these authentications.</span></span>

- <span data-ttu-id="b5fe3-109">**Office 单一登录 (SSO)**：此系统*当前为预览版*，它让用户能在登录到 Office 的同时登录到加载项。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-109">**Office Single Sign-on (SSO)**: A system, *currently in preview*, that enables the user's login to Office to also function as a login to the add-in.</span></span> <span data-ttu-id="b5fe3-110">此外，此加载项还可使用用户的 Office 凭据向加载项授予对 Microsoft Graph 的权限。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-110">Optionally, the add-in can also use the user's Office credentials to authorize the add-in to Microsoft Graph.</span></span> <span data-ttu-id="b5fe3-111">（不可通过此系统访问非 Microsoft 源。）</span><span class="sxs-lookup"><span data-stu-id="b5fe3-111">(Non-Microsoft sources are not accessible through this system.)</span></span>
- <span data-ttu-id="b5fe3-112">**通过 Azure Active Directory 进行 Web 身份验证和授权**：这是老生常谈，没有特别之处。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-112">**Web Application Authentication and Authorization with Azure Active Directory**: This isn't something new or special.</span></span> <span data-ttu-id="b5fe3-113">它只是在出现 Office SSO 系统之前 Office 加载项（及其他 Web 应用）对用户进行身份验证和授权应用的方式，现在仍在 Office SSO 不可用的场景中使用。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-113">It's just the way Office add-in (and other web apps) authenticated users and authorized apps before there was an Office SSO system and is still used in scenarios where Office SSO cannot be.</span></span>

<span data-ttu-id="b5fe3-114">下列流程图展示了需要如同加载项开发人员一样作出的决策。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-114">The following flowchart shows you the decisions that you need to make as an add-in developer.</span></span> <span data-ttu-id="b5fe3-115">详细信息请参见本文稍后部分。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-115">Authentication options are discussed later in this article.</span></span>

![一张图像，它显示了在 Office 加载项中实现身份验证和授权的决策流程图](../images/auth-decisions-flowchart.gif)

## <a name="user-authentication-without-sso"></a><span data-ttu-id="b5fe3-117">在不使用 SSO 的情况下进行用户身份验证</span><span class="sxs-lookup"><span data-stu-id="b5fe3-117">User authentication without SSO</span></span>

<span data-ttu-id="b5fe3-118">你可如同在任何其他 Web 应用程序中操作一样使用 Azure Active Directory (AAD) 在 Office 加载项中对用户进行身份验证，但存在一个例外：AAD 禁止其登录页在 iFrame 中打开。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-118">You can authenticate a user in an Office Add-in with Azure Active Directory (AAD) as you would any in any other web application with one exception: AAD does not allow its login page to open in an iframe.</span></span> <span data-ttu-id="b5fe3-119">当 Office 加载项在 *Office 网页版*中运行时，任务窗格是一个 iFrame。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-119">When an Office Add-in is running on *Office on the web*, the task pane is an iframe.</span></span> <span data-ttu-id="b5fe3-120">这意味着你将需要在通过 Office 对话框 API 打开的对话框中打开 AAD 登录屏幕。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-120">This means that you will need to open the AAD login screen in a dialog opened with the Office Dialog API.</span></span> <span data-ttu-id="b5fe3-121">这会影响你使用身份验证帮助程序库的方式。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-121">This affects how you use authentication helper libraries.</span></span> <span data-ttu-id="b5fe3-122">有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-122">For more information, see [Authentication with the Office Dialog API](auth-with-office-dialog-api.md).</span></span>

<span data-ttu-id="b5fe3-123">要了解如何使用 AAD 对身份验证进行编程，首先请查看 [Microsoft 标识平台 (v2.0) 概述](/azure/active-directory/develop/v2-overview)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-123">For information about programming authentication with AAD, begin with [Microsoft identity platform (v2.0) overview](/azure/active-directory/develop/v2-overview).</span></span> <span data-ttu-id="b5fe3-124">该文档集中有很多教程和指南，还有相关示例和库的链接。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-124">There are many tutorials and guides in that documentation set, as well as links to relevant samples and libraries.</span></span> <span data-ttu-id="b5fe3-125">正如[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)中所述，你可能需要调整示例中的代码以在 Office 对话框中运行。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-125">As explained in [Authentication with the Office Dialog API](auth-with-office-dialog-api.md), you may need to adjust the code in the samples to run in the Office Dialog.</span></span>

## <a name="access-to-microsoft-graph-without-sso"></a><span data-ttu-id="b5fe3-126">在不使用 SSO 的情况下访问 Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="b5fe3-126">Access to Microsoft Graph without SSO</span></span>

<span data-ttu-id="b5fe3-127">可通过从 Azure Active Directory (AAD) 获取到 Microsoft Graph 的访问令牌，为加载项获得到 Graph 数据的授权。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-127">You can get authorization to Microsoft Graph data for your add-in by obtaining an access token to Graph from Azure Active Directory (AAD).</span></span> <span data-ttu-id="b5fe3-128">可在不依赖 Office SSO 的情况下执行此操作。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-128">You can do this without relying on Office SSO.</span></span> <span data-ttu-id="b5fe3-129">要详细了解操作方式，请参阅[在不使用 SSO 的情况下访问 Microsoft Graph](authorize-to-microsoft-graph-without-sso.md)（此文中有更多详细信息和示例链接）。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-129">For more information about how, see [Access to Microsoft Graph without SSO](authorize-to-microsoft-graph-without-sso.md) which has more details and links to samples.</span></span>

## <a name="user-authentication-with-sso"></a><span data-ttu-id="b5fe3-130">在使用 SSO 的情况下进行用户身份验证</span><span class="sxs-lookup"><span data-stu-id="b5fe3-130">User authentication with SSO</span></span>

<span data-ttu-id="b5fe3-131">要使用 SSO 来验证用户的身份，任务窗格或函数文件中的代码会调用 [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-) 方法。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-131">To use SSO to authenticate the user, your code in a task pane or function file calls the [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-) method.</span></span> <span data-ttu-id="b5fe3-132">如果用户未登录 Office，则 Office 将打开一个对话框，并将其导航到 Azure Active Directory 登录页面。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-132">If the user is not signed into Office, Office will open a dialog and navigate it to the Azure Active Directory login page.</span></span> <span data-ttu-id="b5fe3-133">用户登录后或者在用户已登录时，该方法会返回一个访问令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-133">After the user is signed in, or if the user is already signed in, the method returns an access token.</span></span> <span data-ttu-id="b5fe3-134">此令牌是**代理**流中的启动令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-134">The token is a bootstrap token in the **On Behalf Of** flow.</span></span> <span data-ttu-id="b5fe3-135">（详见[使用 SSO 访问 Microsoft Graph](#access-to-microsoft-graph-with-sso)。）但是，它也可用作 ID 令牌，因为它包含多个对当前用户而言唯一的声明，例如 `preferred_username`、`name`、`sub` 和 `oid`。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-135">(See [Access to Microsoft Graph with SSO](#access-to-microsoft-graph-with-sso).) However, it can be used as an ID token as well, because it contains several claims that are unique to the current user, including `preferred_username`, `name`, `sub`, and `oid`.</span></span> <span data-ttu-id="b5fe3-136">要查看指南了解将哪个属性用作最终用户 ID，请参阅 [Microsoft 标识平台访问令牌](https://docs.microsoft.com/zh-CN/azure/active-directory/develop/access-tokens#payload-claims)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-136">For guidance on which property to use as the ultimate user ID, see [Microsoft identity platform access tokens](https://docs.microsoft.com/en-us/azure/active-directory/develop/access-tokens#payload-claims).</span></span> <span data-ttu-id="b5fe3-137">有关上述某一令牌的示例，请参阅[访问令牌示例](sso-in-office-add-ins.md#example-access-token)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-137">For an example of a one of these tokens, see the [Example access token](sso-in-office-add-ins.md#example-access-token).</span></span>

<span data-ttu-id="b5fe3-138">代码从令牌中提取所需的声明后，它将使用该值在你保留的用户表或用户数据库中查找用户。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-138">After your code has extracted the desired claim from the token, it uses that value to look up the user in a user table or user database that you maintain.</span></span> <span data-ttu-id="b5fe3-139">使用数据库来用户用户首选项或用户帐户状态等用户相关信息。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-139">Use the database to store user-relative information such as the user's preferences or the state of the user's account.</span></span> <span data-ttu-id="b5fe3-140">由于你在使用 SSO，因此你的用户不单独登录到你的加载项，你无需存储用户的密码。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-140">Since you are using SSO, your users don't sign-in separately to your add-in, so you do not need to store a password for the user.</span></span>

<span data-ttu-id="b5fe3-141">在开始使用 SSO 实现用户身份验证之前，请确保你完全熟悉[为 Office 加载项启用单一登录](sso-in-office-add-ins.md)一文。另请注意下述示例：</span><span class="sxs-lookup"><span data-stu-id="b5fe3-141">Before you begin implementing user authentication with SSO, be sure that you are thoroughly familiar with the article [Enable single sign-on for Office Add-ins](sso-in-office-add-ins.md). Note also these samples:</span></span>

- <span data-ttu-id="b5fe3-142">[NodeJS SSO 中的 Office 加载项](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)，尤其是 [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) 文件，它使用 [jswebtoken](https://github.com/auth0/node-jsonwebtoken) 库来解码和分析令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-142">[Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO), especially the file [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) which uses the library [jswebtoken](https://github.com/auth0/node-jsonwebtoken) to decode and parse the token.</span></span> <span data-ttu-id="b5fe3-143">（但是，此示例不将令牌用作 ID 令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-143">(This sample, however, does not use the token as an ID token.</span></span> <span data-ttu-id="b5fe3-144">它使用此令牌通过**代理**流获得对 Microsoft Graph 的访问权限。）</span><span class="sxs-lookup"><span data-stu-id="b5fe3-144">It uses it to get access to Microsoft Graph with the **On Behalf Of** flow.)</span></span>
- <span data-ttu-id="b5fe3-145">[ASP.NET SSO 中的 Office 加载项](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)，尤其是 [ValuesController.ts](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs) 文件，它使用库 [System.Security.Claims.ClaimsPrincipal](https://docs.microsoft.com/dotnet/api/system.security.claims.claimsprincipal) 类从令牌中提取声明。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-145">[Office Add-in ASP.NET SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO), especially the file [ValuesController.ts](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Controllers/ValuesController.cs) which uses the library [System.Security.Claims.ClaimsPrincipal](https://docs.microsoft.com/dotnet/api/system.security.claims.claimsprincipal) class to extract claims from the token.</span></span> <span data-ttu-id="b5fe3-146">（但是，此示例不将令牌用作 ID 令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-146">(This sample, however, does not use the token as an ID token.</span></span> <span data-ttu-id="b5fe3-147">它从令牌中提取 `scope`，并用它通过**代理**流获得访问 Microsoft Graph 的权限。）</span><span class="sxs-lookup"><span data-stu-id="b5fe3-147">It extracts a `scope` claim from the token and uses it to get access to Microsoft Graph with the **On Behalf Of** flow.)</span></span>

## <a name="access-to-microsoft-graph-with-sso"></a><span data-ttu-id="b5fe3-148">在使用 SSO 的情况下访问 Microsoft Graph</span><span class="sxs-lookup"><span data-stu-id="b5fe3-148">Access to Microsoft Graph with SSO</span></span>

<span data-ttu-id="b5fe3-149">要使用 SSO 来获取访问 Microsoft Graph 的权限，任务窗格或函数文件中的加载项会调用 [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-) 方法。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-149">To use SSO to get access to Microsoft Graph, your add-in in a task pane or function file calls the [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-) method.</span></span> <span data-ttu-id="b5fe3-150">如果用户未登录 Office，则 Office 将打开一个对话框，并将其导航到 Azure Active Directory 登录页面。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-150">If the user is not signed into Office, Office will open a dialog and navigate it to the Azure Active Directory login page.</span></span> <span data-ttu-id="b5fe3-151">用户登录后或者在用户已登录时，该方法会返回一个访问令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-151">After the user is signed in, or if the user is already signed in, the method returns an access token.</span></span> <span data-ttu-id="b5fe3-152">此令牌是**代理**流中的启动令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-152">The token is a bootstrap token in the **On Behalf Of** flow.</span></span> <span data-ttu-id="b5fe3-153">具体而言，它有一个带 `access_as_user` 值的 `scope` 声明。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-153">Specifically, it has a `scope` claim with the value `access_as_user`.</span></span> <span data-ttu-id="b5fe3-154">要在指南中了解令牌中的声明，请参阅 [Microsoft 标识平台访问令牌](https://docs.microsoft.com/zh-CN/azure/active-directory/develop/access-tokens#payload-claims)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-154">For guidance about the claims in the token, see [Microsoft identity platform access tokens](https://docs.microsoft.com/en-us/azure/active-directory/develop/access-tokens#payload-claims).</span></span> <span data-ttu-id="b5fe3-155">有关上述某一令牌的示例，请参阅[访问令牌示例](sso-in-office-add-ins.md#example-access-token)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-155">For an example of a one of these tokens, see the [Example access token](sso-in-office-add-ins.md#example-access-token).</span></span>

<span data-ttu-id="b5fe3-156">在代码获取令牌后，它会在**代理**流中使用该令牌来获取第二个令牌，即到 Microsoft Graph 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-156">After your code obtains the token, it uses it in the **On Behalf Of** flow to obtain a second token: an access token to Microsoft Graph.</span></span>

<span data-ttu-id="b5fe3-157">在开始实现 Office SSO 之前，请确保你完全熟悉下面两篇文章：</span><span class="sxs-lookup"><span data-stu-id="b5fe3-157">Before you begin implementing Office SSO, be sure that you are thoroughly familiar with these two articles:</span></span>

- [<span data-ttu-id="b5fe3-158">为 Office 加载项启用单一登录</span><span class="sxs-lookup"><span data-stu-id="b5fe3-158">Enable single sign-on for Office Add-ins</span></span>](sso-in-office-add-ins.md)
- [<span data-ttu-id="b5fe3-159">使用 SSO 对 Microsoft Graph 授权</span><span class="sxs-lookup"><span data-stu-id="b5fe3-159">Authorize to Microsoft Graph with SSO</span></span>](authorize-to-microsoft-graph.md)

<span data-ttu-id="b5fe3-160">你还应至少阅读此处所列的其中一篇演示文章。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-160">You should also read at least one of the walkthrough articles listed here.</span></span> <span data-ttu-id="b5fe3-161">即使你不执行这些步骤，也可在其中了解有关如何实现 Office SSO 和**代理**流的宝贵信息。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-161">Even if you don't carry out the steps, these contain valuable information about how you implement Office SSO and the **On Behalf Of** flow.</span></span> 

- [<span data-ttu-id="b5fe3-162">创建使用单一登录的 ASP.NET Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b5fe3-162">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)
- [<span data-ttu-id="b5fe3-163">创建使用单一登录的 Node.js Office 加载项</span><span class="sxs-lookup"><span data-stu-id="b5fe3-163">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)

<span data-ttu-id="b5fe3-164">另请注意下述示例：</span><span class="sxs-lookup"><span data-stu-id="b5fe3-164">Note also these samples:</span></span>

- [<span data-ttu-id="b5fe3-165">Office 加载项 NodeJS SSO</span><span class="sxs-lookup"><span data-stu-id="b5fe3-165">Office Add-in NodeJS SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)
- [<span data-ttu-id="b5fe3-166">Office 加载项 ASP.NET SSO</span><span class="sxs-lookup"><span data-stu-id="b5fe3-166">Office Add-in ASP.NET SSO</span></span>](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)

## <a name="access-to-non-microsoft-data-sources"></a><span data-ttu-id="b5fe3-167">访问非 Microsoft 数据源</span><span class="sxs-lookup"><span data-stu-id="b5fe3-167">Access to non-Microsoft data sources</span></span>

<span data-ttu-id="b5fe3-168">借助 Google、Facebook、领英、SalesForce 和 GitHub 等热门在线服务，开发人员可授权用户访问自己在其他应用中的帐户。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-168">Popular online services, including Office 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications.</span></span> <span data-ttu-id="b5fe3-169">这样，便可在 Office 加载项中添加这些服务。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-169">This gives you the ability to include these services in your Office Add-in.</span></span> <span data-ttu-id="b5fe3-170">要概述了解加载项可执行此操作的方法，请参阅[在 Office 加载项中授权外部服务](auth-external-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-170">For an overview of the ways that your add-in can do this, see [Authorize external services in your Office Add-in](auth-external-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b5fe3-171">开始编码之前，请了解数据源是否允许在 iFrame 中打开其登录屏幕。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-171">Before you begin coding, find out if the data source allows its login in screen to be opened in an iFrame.</span></span> <span data-ttu-id="b5fe3-172">当 Office 加载项在 *Office 网页版*中运行时，任务窗格是一个 iFrame。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-172">When an Office Add-in is running on *Office on the web*, the task pane is an iFrame.</span></span> <span data-ttu-id="b5fe3-173">如果数据源禁止在 iFrame 中打开其登录屏幕，则你需要在通过 Office 对话框 API 打开的对话框中打开登录屏幕。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-173">If the data source does not allow its login screen to be opened in an iFrame, then you will need to open the login screen in a dialog opened with the Office Dialog API.</span></span> <span data-ttu-id="b5fe3-174">有关详细信息，请参阅[使用 Office 对话框 API 进行身份验证](auth-with-office-dialog-api.md)。</span><span class="sxs-lookup"><span data-stu-id="b5fe3-174">For more information, see [Authentication with the Office Dialog API](auth-with-office-dialog-api.md).</span></span>
