---
title: 为 Office 加载项启用单一登录
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: ca8280b72ab863d0e34330585fb307475e3aa9b9
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298563"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="30c9b-102">为 Office 加载项启用单一登录（预览）</span><span class="sxs-lookup"><span data-stu-id="30c9b-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="30c9b-103">用户可以使用自己的个人 Microsoft 帐户/工作或学校 (Office 365) 帐户，登录 Office（在线、移动和桌面平台）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="30c9b-104">可以利用此功能并使用单一登录 (SSO) 授权用户访问加载项（用户无需再次登录）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![显示加载项登录过程的图像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="30c9b-106">预览状态</span><span class="sxs-lookup"><span data-stu-id="30c9b-106">Preview Status</span></span>

<span data-ttu-id="30c9b-107">当前只在预览中支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="30c9b-107">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="30c9b-108">它可供开发人员进行实验，但不应用于生产加载项。</span><span class="sxs-lookup"><span data-stu-id="30c9b-108">It is available to developers for experimentation; but it should not be used in a production add-in.</span></span> <span data-ttu-id="30c9b-109">此外，在 [AppSource](https://appsource.microsoft.com) 中不接受使用 SSO 的加载项。</span><span class="sxs-lookup"><span data-stu-id="30c9b-109">In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="30c9b-110">并非所有 Office 应用程序都支持 SSO 预览。</span><span class="sxs-lookup"><span data-stu-id="30c9b-110">Not all Office applications support the SSO preview.</span></span> <span data-ttu-id="30c9b-111">可以在 Word、Excel、Outlook 和 PowerPoint 中使用此加载项。</span><span class="sxs-lookup"><span data-stu-id="30c9b-111">It is available in Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="30c9b-112">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-112">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="30c9b-113">要求和最佳做法</span><span class="sxs-lookup"><span data-stu-id="30c9b-113">Requirements and Best Practices</span></span>

<span data-ttu-id="30c9b-114">若要使用 SSO，必须从加载项的启动 HTML 页面中的 `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` 加载 Office JavaScript 库的 Beta 版。</span><span class="sxs-lookup"><span data-stu-id="30c9b-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="30c9b-115">如果使用的是 **Outlook** 加载项，请务必为 Office 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="30c9b-115">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="30c9b-116">若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-116">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="30c9b-117">*不应*依赖 SSO 作为加载项的唯一身份验证方法。</span><span class="sxs-lookup"><span data-stu-id="30c9b-117">You should *not* rely on SSO as your add-in's only method of authentication.</span></span> <span data-ttu-id="30c9b-118">应实现备用身份验证系统，在某些错误情况下，加载项可以返回到该系统。</span><span class="sxs-lookup"><span data-stu-id="30c9b-118">You should implement an alternate authentication system that your add-in can fall back to in certain error situations.</span></span> <span data-ttu-id="30c9b-119">可以使用包含用户表和身份验证的系统，也可以利用其中某个社交登录提供者。</span><span class="sxs-lookup"><span data-stu-id="30c9b-119">You can use a system of user tables and authentication, or you can leverage one of the social login providers.</span></span> <span data-ttu-id="30c9b-120">有关如何使用 Office 加载项执行此操作的详细信息，请参阅 [Authorize external services in your Office Add-in](https://docs.microsoft.com/zh-CN/office/dev/add-ins/develop/auth-external-add-ins)（对 Office 加载项中的外部服务授权）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-120">For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/zh-CN/office/dev/add-ins/develop/auth-external-add-ins).</span></span> <span data-ttu-id="30c9b-121">对于 *Outlook*，建议使用后备系统。</span><span class="sxs-lookup"><span data-stu-id="30c9b-121">For *Outlook*, there is a recommended fall back system.</span></span> <span data-ttu-id="30c9b-122">有关详细信息，请参阅[应用场景：在 Outlook 加载项中对服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-122">For more details, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="30c9b-123">运行时 SSO 的工作方式</span><span class="sxs-lookup"><span data-stu-id="30c9b-123">How it works at runtime</span></span>

<span data-ttu-id="30c9b-124">以下关系图显示了 SSO 流程的工作方式。</span><span class="sxs-lookup"><span data-stu-id="30c9b-124">The following diagram shows how the SSO process works.</span></span>

![SSO 过程关系图](../images/sso-overview-diagram.png)

1. <span data-ttu-id="30c9b-126">在加载项中，JavaScript 调用新的 Office.js API [getAccessTokenAsync](#sso-api-reference)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-126">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference).</span></span> <span data-ttu-id="30c9b-127">这会指示 Office 主机应用程序获取对加载项的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="30c9b-127">This tells the Office host application to obtain an access token to the add-in.</span></span> <span data-ttu-id="30c9b-128">请参阅[示例访问令牌](#example-access-token)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-128">See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="30c9b-129">如果用户未登录，Office 主机应用会打开弹出窗口，以供用户登录。</span><span class="sxs-lookup"><span data-stu-id="30c9b-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="30c9b-130">如果当前用户是首次使用加载项，则会看到同意提示。</span><span class="sxs-lookup"><span data-stu-id="30c9b-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="30c9b-131">Office 主机应用程序从当前用户的 Azure AD v2.0 终结点请求获取**加载项令牌**。</span><span class="sxs-lookup"><span data-stu-id="30c9b-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="30c9b-132">Azure AD 将加载项令牌发送给 Office 主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="30c9b-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="30c9b-133">Office 主机应用程序在 `getAccessTokenAsync` 调用返回的结果对象中，将“**加载项令牌**”发送给加载项。</span><span class="sxs-lookup"><span data-stu-id="30c9b-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="30c9b-134">加载项中的 JavaScript 可以解析令牌并提取所需信息，如用户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="30c9b-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="30c9b-135">（可选）加载项可以向其服务器端发送 HTTP 请求以获取关于用户的更多数据，如用户的偏好。</span><span class="sxs-lookup"><span data-stu-id="30c9b-135">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="30c9b-136">此外，访问令牌本身也可发送到服务器端以进行解析和验证。</span><span class="sxs-lookup"><span data-stu-id="30c9b-136">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="30c9b-137">开发 SSO 加载项</span><span class="sxs-lookup"><span data-stu-id="30c9b-137">Develop an SSO add-in</span></span>

<span data-ttu-id="30c9b-138">此部分介绍了创建启用 SSO 的 Office 加载项所需完成的任务。</span><span class="sxs-lookup"><span data-stu-id="30c9b-138">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="30c9b-139">其中介绍的这些任务与语言和框架无关。</span><span class="sxs-lookup"><span data-stu-id="30c9b-139">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="30c9b-140">有关详细演练的示例，请参阅：</span><span class="sxs-lookup"><span data-stu-id="30c9b-140">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="30c9b-141">创建使用单一登录的 Node.js Office 加载项</span><span class="sxs-lookup"><span data-stu-id="30c9b-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="30c9b-142">创建使用单一登录的 ASP.NET Office 加载项</span><span class="sxs-lookup"><span data-stu-id="30c9b-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="30c9b-143">创建服务应用程序</span><span class="sxs-lookup"><span data-stu-id="30c9b-143">Create the service application</span></span>

<span data-ttu-id="30c9b-144">在 Azure v2.0 端点的注册门户注册加载项：https://apps.dev.microsoft.com。</span><span class="sxs-lookup"><span data-stu-id="30c9b-144">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="30c9b-145">该流程用时 5-10 分钟，包括以下任务：</span><span class="sxs-lookup"><span data-stu-id="30c9b-145">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="30c9b-146">获取加载项的客户端 ID 和机密。</span><span class="sxs-lookup"><span data-stu-id="30c9b-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="30c9b-147">指定加载项访问 AAD v 所需的权限。</span><span class="sxs-lookup"><span data-stu-id="30c9b-147">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="30c9b-148">2.0 端点（可选 Microsoft Graph）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-148">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="30c9b-149">始终需要“profile”权限。</span><span class="sxs-lookup"><span data-stu-id="30c9b-149">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="30c9b-150">授予 Office 主机应用程序信任加载项。</span><span class="sxs-lookup"><span data-stu-id="30c9b-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="30c9b-151">将 Office 主机应用程序预授权给具有 *access_as_user* 默认权限的加载项。</span><span class="sxs-lookup"><span data-stu-id="30c9b-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="30c9b-152">有关此过程的详细信息，请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项](register-sso-add-in-aad-v2.md)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-152">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="30c9b-153">配置加载项</span><span class="sxs-lookup"><span data-stu-id="30c9b-153">Configure the add-in</span></span>

<span data-ttu-id="30c9b-154">向外接程序清单添加新标记：</span><span class="sxs-lookup"><span data-stu-id="30c9b-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="30c9b-155">**WebApplicationInfo** - 下列元素的父元素。</span><span class="sxs-lookup"><span data-stu-id="30c9b-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="30c9b-156">**ID** - 加载项的客户端 ID。这是在注册加载项时获得的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="30c9b-156">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="30c9b-157">请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项](register-sso-add-in-aad-v2.md)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-157">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="30c9b-158">**Resource** - 加载项 URL。</span><span class="sxs-lookup"><span data-stu-id="30c9b-158">**Resource** - The URL of the add-in.</span></span> <span data-ttu-id="30c9b-159">这是在 AAD 中注册加载项时使用的相同 URI（包括 `api:` 协议）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-159">This is the same URI (including the `api:` protocol) that you used when registering the add-in in AAD.</span></span> <span data-ttu-id="30c9b-160">此 URI 的域部分应与加载项清单的 `<Resources>` 部分中的 URL 中使用的域（包括任何子域）匹配。</span><span class="sxs-lookup"><span data-stu-id="30c9b-160">The domain part of this URI should match the domain, including any subdomains, used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>
* <span data-ttu-id="30c9b-161">**Scopes** - 一个或多个“**Scope**”元素的父元素。</span><span class="sxs-lookup"><span data-stu-id="30c9b-161">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="30c9b-162">**Scope** - 指定加载项访问 AAD 所需的权限。</span><span class="sxs-lookup"><span data-stu-id="30c9b-162">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="30c9b-163">如果加载项无法访问 Microsoft Graph，则始终需要 `profile` 权限，并且它可能是唯一需要的权限。</span><span class="sxs-lookup"><span data-stu-id="30c9b-163">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="30c9b-164">如果可以访问，则还需要“**Scope**”元素来获取所需的 Microsoft Graph 权限（如 `User.Read``Mail.Read`）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-164">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="30c9b-165">在代码中用于访问 Microsoft Graph 的库可能需要其他权限。</span><span class="sxs-lookup"><span data-stu-id="30c9b-165">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="30c9b-166">例如，用于 .NET 的 Microsoft 身份验证库 (MSAL) 需要 `offline_access` 权限。</span><span class="sxs-lookup"><span data-stu-id="30c9b-166">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="30c9b-167">有关详细信息，请参阅[向 Office 加载项中的 Microsoft Graph 授权](authorize-to-microsoft-graph.md)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-167">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="30c9b-p114">对于除 Outlook 之外的 Office 主机，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 部分的末尾。对 Outlook，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 部分的末尾。</span><span class="sxs-lookup"><span data-stu-id="30c9b-p114">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="30c9b-170">下面的示例展示了标记：</span><span class="sxs-lookup"><span data-stu-id="30c9b-170">The following is an example of the markup:</span></span>

```xml
<WebApplicationInfo>
    <Id>5661fed9-f33d-4e95-b6cf-624a34a2f51d</Id>
    <Resource>api://addin.contoso.com/5661fed9-f33d-4e95-b6cf-624a34a2f51d</Resource>
    <Scopes>
        <Scope>user.read</Scope>
        <Scope>files.read</Scope>
        <Scope>profile</Scope>
    </Scopes>
</WebApplicationInfo>
```

### <a name="add-client-side-code"></a><span data-ttu-id="30c9b-171">添加客户端代码</span><span class="sxs-lookup"><span data-stu-id="30c9b-171">Add client-side code</span></span>

<span data-ttu-id="30c9b-172">将 JavaScript 添加到加载项，以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="30c9b-172">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="30c9b-173">调用 [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span><span class="sxs-lookup"><span data-stu-id="30c9b-173">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="30c9b-174">解析访问令牌或将其传递到加载项的服务器端代码。</span><span class="sxs-lookup"><span data-stu-id="30c9b-174">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="30c9b-175">下面是调用 `getAccessTokenAsync` 的简单示例。</span><span class="sxs-lookup"><span data-stu-id="30c9b-175">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="30c9b-176">此示例只显式处理一种错误。</span><span class="sxs-lookup"><span data-stu-id="30c9b-176">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="30c9b-177">有关更详细的错误处理的示例，请参阅 [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) 和 [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-177">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="30c9b-178">另请参阅[排查单一登录 (SSO) 错误消息](troubleshoot-sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="30c9b-178">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

```js
Office.context.auth.getAccessTokenAsync(function (result) {
    if (result.status === "succeeded") {
        // Use this token to call Web API
        var ssoToken = result.value;
        ...
    } else {
        if (result.error.code === 13003) {
            // SSO is not supported for domain user accounts, only
            // work or school (Office 365) or Microsoft Account IDs.
        } else {
            // Handle error
        }
    }
});
```

<span data-ttu-id="30c9b-179">下面是一个将加载项令牌传递到服务器端的简单示例。</span><span class="sxs-lookup"><span data-stu-id="30c9b-179">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="30c9b-180">将请求发送回服务器端时，令牌作为 `Authorization` 标头包含在内。</span><span class="sxs-lookup"><span data-stu-id="30c9b-180">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="30c9b-181">此示例设想发送 JSON 数据，因此它使用 `POST` 方法，但使用 `GET` 就足以在未写入服务器时发送访问令牌。</span><span class="sxs-lookup"><span data-stu-id="30c9b-181">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

```js
$.ajax({
    type: "POST",
    url: "/api/DoSomething",
    headers: {
        "Authorization": "Bearer " + ssoToken
    },
    data: { /* some JSON payload */ },
    contentType: "application/json; charset=utf-8"
}).done(function (data) {
    // Handle success
}).fail(function (error) {
    // Handle error
}).always(function () {
    // Cleanup
});
```

#### <a name="when-to-call-the-method"></a><span data-ttu-id="30c9b-182">何时调用方法</span><span class="sxs-lookup"><span data-stu-id="30c9b-182">When to call the method</span></span>

<span data-ttu-id="30c9b-183">如果因没有用户登录 Office 而无法使用加载项，则应*在加载项启动时*调用 `getAccessTokenAsync`。</span><span class="sxs-lookup"><span data-stu-id="30c9b-183">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="30c9b-184">如果加载项具有一些无需用户登录的功能，那么*当用户执行需要用户登录的操作时*，请调用 `getAccessTokenAsync`。</span><span class="sxs-lookup"><span data-stu-id="30c9b-184">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="30c9b-185">`getAccessTokenAsync` 的冗余调用不会导致性能严重下降，因为 Office 缓存并重用访问没有过期的令牌，无需每次调用 AAD v。</span><span class="sxs-lookup"><span data-stu-id="30c9b-185">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="30c9b-186">`getAccessTokenAsync` 都重新调用 AAD V 2.0 端点。</span><span class="sxs-lookup"><span data-stu-id="30c9b-186">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="30c9b-187">因此，可以将 `getAccessTokenAsync` 调用添加到所有在需要令牌时启动操作的函数和处理程序。</span><span class="sxs-lookup"><span data-stu-id="30c9b-187">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="30c9b-188">添加服务器端代码</span><span class="sxs-lookup"><span data-stu-id="30c9b-188">Add server-side code</span></span>

<span data-ttu-id="30c9b-189">大多数情况下，如果加载项没有将访问令牌传递到服务器端并在其中使用它，那么获取访问令牌的意义就不大。</span><span class="sxs-lookup"><span data-stu-id="30c9b-189">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="30c9b-190">加载项可以执行的一些服务器端任务：</span><span class="sxs-lookup"><span data-stu-id="30c9b-190">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="30c9b-191">创建一种或多种 Web API 方法（例如，一种在托管数据库中查找用户首选项的方法），使用有关从令牌中提取的用户的信息。</span><span class="sxs-lookup"><span data-stu-id="30c9b-191">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="30c9b-192">（请参阅下文“**使用 SSO 令牌作为标识**”。）可以使用一些库简化需要编写的代码，具体视语言和框架而定。</span><span class="sxs-lookup"><span data-stu-id="30c9b-192">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="30c9b-193">获取 Microsoft Graph 数据。</span><span class="sxs-lookup"><span data-stu-id="30c9b-193">Get Microsoft Graph data.</span></span> <span data-ttu-id="30c9b-194">服务器端代码应执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="30c9b-194">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="30c9b-195">验证访问令牌（请参阅下文“**验证访问令牌**”）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-195">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="30c9b-196">通过调用 Azure AD v2.0 端点启动“代表”流，该端点包括访问令牌、关于用户的一些元数据以及加载项的凭据（其 ID 和机密）。</span><span class="sxs-lookup"><span data-stu-id="30c9b-196">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="30c9b-197">在此上下文中，访问令牌称为启动令牌。</span><span class="sxs-lookup"><span data-stu-id="30c9b-197">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="30c9b-198">缓存代表流返回的新访问令牌。</span><span class="sxs-lookup"><span data-stu-id="30c9b-198">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="30c9b-199">使用新的令牌从 Microsoft Graph 获取数据。</span><span class="sxs-lookup"><span data-stu-id="30c9b-199">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="30c9b-200">如需深入了解如何获得对用户的 Microsoft Graph 数据的授权访问，请参阅[向 Office 加载项中的 Microsoft Graph 授权](authorize-to-microsoft-graph.md)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-200">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="30c9b-201">验证访问令牌</span><span class="sxs-lookup"><span data-stu-id="30c9b-201">Validate the token</span></span>

<span data-ttu-id="30c9b-202">Web API 收到访问令牌后，必须在使用该令牌前对其进行验证。</span><span class="sxs-lookup"><span data-stu-id="30c9b-202">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="30c9b-203">该令牌是 JSON Web 令牌 (JWT)，这意味着验证方式与大多数标准 OAuth 流中的令牌验证方式类似。</span><span class="sxs-lookup"><span data-stu-id="30c9b-203">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="30c9b-204">有许多可用于处理 JWT 验证的库，而它们的基本内容为：</span><span class="sxs-lookup"><span data-stu-id="30c9b-204">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="30c9b-205">检查令牌的格式是否正确</span><span class="sxs-lookup"><span data-stu-id="30c9b-205">Checking that the token is well-formed</span></span>
- <span data-ttu-id="30c9b-206">检查令牌是否由预期的颁发机构颁发</span><span class="sxs-lookup"><span data-stu-id="30c9b-206">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="30c9b-207">检查令牌是否是针对 Web API</span><span class="sxs-lookup"><span data-stu-id="30c9b-207">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="30c9b-208">验证令牌时，请牢记以下准则：</span><span class="sxs-lookup"><span data-stu-id="30c9b-208">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="30c9b-209">有效的 SSO 令牌是由 Azure 颁发机构 `https://login.microsoftonline.com` 的。</span><span class="sxs-lookup"><span data-stu-id="30c9b-209">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="30c9b-210">令牌中的 `iss` 声明应以此值开头。</span><span class="sxs-lookup"><span data-stu-id="30c9b-210">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="30c9b-211">令牌的 `aud` 参数将被设置为加载项注册的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="30c9b-211">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="30c9b-212">令牌的 `scp` 参数将被设置为 `access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="30c9b-212">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="30c9b-213">将 SSO 令牌用作标识</span><span class="sxs-lookup"><span data-stu-id="30c9b-213">Using the SSO token as an identity</span></span>

<span data-ttu-id="30c9b-214">如果加载项需要验证用户标识，则 SSO 令牌包含的信息可用于创建此标识。</span><span class="sxs-lookup"><span data-stu-id="30c9b-214">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="30c9b-215">令牌中的以下声明与标识相关。</span><span class="sxs-lookup"><span data-stu-id="30c9b-215">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="30c9b-216">`name` - 用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="30c9b-216">`name` - The user's display name.</span></span>
- <span data-ttu-id="30c9b-217">`preferred_username` - 用户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="30c9b-217">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="30c9b-218">`oid` - 表示 Azure Active Directory 中的用户 ID 的 GUID。</span><span class="sxs-lookup"><span data-stu-id="30c9b-218">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="30c9b-219">`tid` - 表示 Azure Active Directory 中的用户组织 ID 的 GUID。</span><span class="sxs-lookup"><span data-stu-id="30c9b-219">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="30c9b-220">由于 `name` 和 `preferred_username` 值可以更改，因此建议使用 `oid` 和 `tid` 值将标识与后端的授权服务关联。</span><span class="sxs-lookup"><span data-stu-id="30c9b-220">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="30c9b-221">例如，你的服务可以将这些值组合在一起，并设置为类似 `{oid-value}@{tid-value}` 的格式，然后将其存储为内部用户数据库中的用户记录值。</span><span class="sxs-lookup"><span data-stu-id="30c9b-221">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="30c9b-222">然后，在后续的请求中，可以使用同一值检索此用户，并可基于现有访问控制机制确定对特定资源的访问。</span><span class="sxs-lookup"><span data-stu-id="30c9b-222">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="30c9b-223">示例访问令牌</span><span class="sxs-lookup"><span data-stu-id="30c9b-223">Example access token</span></span>

<span data-ttu-id="30c9b-224">以下是访问令牌的典型解码有效负载。</span><span class="sxs-lookup"><span data-stu-id="30c9b-224">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="30c9b-225">有关属性的详细信息，请参阅 [Azure Active Directory v2.0 令牌参考](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-225">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


```js
{
    aud: "2c3caa80-93f9-425e-8b85-0745f50c0d24",         
    iss: "https://login.microsoftonline.com/fec4f964-8bc9-4fac-b972-1c1da35adbcd/v2.0",         
    iat: 1521143967,         
    nbf: 1521143967,         
    exp: 1521147867,         
    aio: "ATQAy/8GAAAA0agfnU4DTJUlEqGLisMtBk5q6z+6DB+sgiRjB/Ni73q83y0B86yBHU/WFJnlMQJ8",         
    azp: "e4590ed6-62b3-5102-beff-bad2292ab01c",         
    azpacr: "0",         
    e_exp: 262800,         
    name: "Mila Nikolova",         
    oid: "6467882c-fdfd-4354-a1ed-4e13f064be25",         
    preferred_username: "milan@contoso.com",         
    scp: "access_as_user",         
    sub: "XkjgWjdmaZ-_xDmhgN1BMP2vL2YOfeVxfPT_o8GRWaw",         
    tid: "fec4f964-8bc9-4fac-b972-1c1da35adbcd",         
    uti: "MICAQyhrH02ov54bCtIDAA",         
    ver: "2.0"
}
```

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="30c9b-226">将 SSO 与 Outlook 加载项一起使用</span><span class="sxs-lookup"><span data-stu-id="30c9b-226">Using SSO with an Outlook add-in</span></span>

<span data-ttu-id="30c9b-227">在 Outlook 加载项中使用 SSO 与在 Excel、PowerPoint 或 Word 加载项中使用 SSO 存在一些细微但却重要的差别。</span><span class="sxs-lookup"><span data-stu-id="30c9b-227">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in.</span></span> <span data-ttu-id="30c9b-228">请务必阅读[使用 Outlook 加载项的单一登录对用户进行身份验证](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)和[：在 Outlook 加载项中为服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。</span><span class="sxs-lookup"><span data-stu-id="30c9b-228">Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="30c9b-229">SSO API 参考</span><span class="sxs-lookup"><span data-stu-id="30c9b-229">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="30c9b-230">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="30c9b-230">getAccessTokenAsync</span></span>

<span data-ttu-id="30c9b-231">Office Auth 命名空间 `Office.context.auth` 提供了一种方法 `getAccessTokenAsync`，使 Office 主机能够获得加载项的 Web 应用程序的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="30c9b-231">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="30c9b-232">这也使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户第二次登录。</span><span class="sxs-lookup"><span data-stu-id="30c9b-232">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="30c9b-233">该方法调用 Azure Active Directory V 2.0 端点以获取令牌来访问加载项的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="30c9b-233">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="30c9b-234">这样可以使加载项识别用户。</span><span class="sxs-lookup"><span data-stu-id="30c9b-234">This enables add-ins to identify users.</span></span> <span data-ttu-id="30c9b-235">通过[“代表”OAuth 流](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)，服务器端代码可以使用此令牌访问加载项 Web 应用程序的 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="30c9b-235">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="30c9b-236">在 Outlook 中，如果加载项加载到 Outlook.com 或 Gmail 邮箱中，则此 API 不受支持。</span><span class="sxs-lookup"><span data-stu-id="30c9b-236">In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.</span></span>

<table><tr><td><span data-ttu-id="30c9b-237">主机</span><span class="sxs-lookup"><span data-stu-id="30c9b-237">Hosts</span></span></td><td><span data-ttu-id="30c9b-238">Excel, OneNote, Outlook, PowerPoint, Word</span><span class="sxs-lookup"><span data-stu-id="30c9b-238">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td>[<span data-ttu-id="30c9b-239">要求集</span><span class="sxs-lookup"><span data-stu-id="30c9b-239">Requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td><td>[<span data-ttu-id="30c9b-240">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="30c9b-240">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="30c9b-241">参数</span><span class="sxs-lookup"><span data-stu-id="30c9b-241">Parameters</span></span>

<span data-ttu-id="30c9b-242">`options` - 可选。</span><span class="sxs-lookup"><span data-stu-id="30c9b-242">`options` - Optional.</span></span> <span data-ttu-id="30c9b-243">接受 `AuthOptions` 对象（参见下文）以定义登录行为。</span><span class="sxs-lookup"><span data-stu-id="30c9b-243">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="30c9b-244">`callback` - 可选。</span><span class="sxs-lookup"><span data-stu-id="30c9b-244">`callback` - Optional.</span></span> <span data-ttu-id="30c9b-245">接受可以解析用户 ID 的令牌或使用“代表”流中的令牌来访问 Microsoft Graph 的回调方法。</span><span class="sxs-lookup"><span data-stu-id="30c9b-245">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="30c9b-246">如果 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult) `.status`为“成功”，则 `AsyncResult.value` 是原始 AAD v。</span><span class="sxs-lookup"><span data-stu-id="30c9b-246">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="30c9b-247">2.0 格式的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="30c9b-247">2.0-formatted access token.</span></span>

<span data-ttu-id="30c9b-248">当 Office 从 AAD v 获取加载项的访问令牌时，`AuthOptions` 接口提供用户体验选项。</span><span class="sxs-lookup"><span data-stu-id="30c9b-248">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="30c9b-249">2.0 使用 `getAccessTokenAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="30c9b-249">2.0 with the `getAccessTokenAsync` method.</span></span>

```typescript
interface AuthOptions {
    /**
        * Causes Office to display the add-in consent experience. Useful if the add-in's Azure permissions have changed or if the user's consent has 
        * been revoked.
        */
    forceConsent?: boolean,
    /**
        * Prompts the user to add their Office account (or to switch to it, if it is already added).
        */
    forceAddAccount?: boolean,
    /**
        * Causes Office to prompt the user to provide the additional factor when the tenancy being targeted by Microsoft Graph requires multifactor 
        * authentication. The string value identifies the type of additional factor that is required. In most cases, you won't know at development 
        * time whether the user's tenant requires an additional factor or what the string should be. So this option would be used in a "second try" 
        * call of getAccessTokenAsync after Microsoft Graph has sent an error requesting the additional factor and containing the string that should 
        * be used with the authChallenge option.
        */
    authChallenge?: string
    /**
        * A user-defined item of any type that is returned, unchanged, in the asyncContext property of the AsyncResult object that is passed to a callback.
        */
    asyncContext?: any
}
```



