---
title: 为 Office 加载项启用单一登录
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: ce57c5d70e2c48a89b2fd84c30ac7b8580650896
ms.sourcegitcommit: 8333ede51307513312d3078cb072f856f5bef8a2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/07/2018
ms.locfileid: "23876604"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="0f815-102">为 Office 加载项启用单一登录（预览）</span><span class="sxs-lookup"><span data-stu-id="0f815-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="0f815-103">用户可以使用自己的个人 Microsoft 帐户/工作或学校 (Office 365) 帐户，登录 Office（在线、移动和桌面平台）。</span><span class="sxs-lookup"><span data-stu-id="0f815-103">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account.</span></span> <span data-ttu-id="0f815-104">你可以利用这一优势并使用单一登录（SSO）授权用户访问你的加载项，无需用户再次登录。</span><span class="sxs-lookup"><span data-stu-id="0f815-104">You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>


![显示加载项登录过程的图像](../images/office-host-title-bar-sign-in.png)

> [!NOTE]
> <span data-ttu-id="0f815-106">目前，Word、Excel、Outlook 和 PowerPoint 预览版支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="0f815-106">The Single Sign-on API is currently supported in preview for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="0f815-107">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="0f815-107">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://dev.office.com/reference/add-ins/requirement-sets/identity-api-requirement-sets).</span></span>
> <span data-ttu-id="0f815-108">若要使用 SSO，您必须从加载项启动 HTML 页的 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 中加载 beta 版的 Office JavaScript 库。</span><span class="sxs-lookup"><span data-stu-id="0f815-108">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>
> <span data-ttu-id="0f815-109">如果使用的是 Outlook 加载项，请务必为 Office 365 租户启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="0f815-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="0f815-110">若要了解如何执行此操作，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。</span><span class="sxs-lookup"><span data-stu-id="0f815-110">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="0f815-111">对于用户来说，这只涉及一次登录，因而使得运行加载项的体验更流畅。</span><span class="sxs-lookup"><span data-stu-id="0f815-111">For users, this makes running your add-in a smooth experience that involves at signing in only once.</span></span> <span data-ttu-id="0f815-112">对于开发人员来说，这意味着你的加载项不必使用加密的密码维护自己的用户表。</span><span class="sxs-lookup"><span data-stu-id="0f815-112">For developers, this means that your add-in does not have to maintain it's own user tables with encrypted passwords.</span></span>

### <a name="how-it-works-at-runtime"></a><span data-ttu-id="0f815-113">运行时的工作方式</span><span class="sxs-lookup"><span data-stu-id="0f815-113">How it works at runtime</span></span>

<span data-ttu-id="0f815-114">以下关系图显示了 SSO 流程的工作方式。</span><span class="sxs-lookup"><span data-stu-id="0f815-114">The following diagram shows how the SSO process works.</span></span>

![SSO 过程关系图](../images/sso-overview-diagram.png)

1. <span data-ttu-id="0f815-p104">在外接程序 JavaScript 调用新的 Office.js API [getAccessTokenAsync](#sso-api-reference)。这会告诉 Office 主机应用程序获取访问令牌到外接程序。请参阅 [示例访问令牌](#example-access-token)。</span><span class="sxs-lookup"><span data-stu-id="0f815-p104">In the add-in, JavaScript calls a new Office.js API [](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. (Hereafter, this is called the [add-in token](#example-access-token).)</span></span>
2. <span data-ttu-id="0f815-119">如果用户未登录，Office 主机应用会打开弹出窗口，以供用户登录。</span><span class="sxs-lookup"><span data-stu-id="0f815-119">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="0f815-120">如果当前用户是首次使用加载项，则会看到同意提示。</span><span class="sxs-lookup"><span data-stu-id="0f815-120">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="0f815-121">Office 主机应用程序从当前用户的 Azure AD v2.0 终结点请求获取**加载项令牌**。</span><span class="sxs-lookup"><span data-stu-id="0f815-121">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="0f815-122">Azure AD 将加载项令牌发送给 Office 主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="0f815-122">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="0f815-123">Office 主机应用程序在 `getAccessTokenAsync` 调用返回的结果对象中，将**加载项令牌**发送给加载项。</span><span class="sxs-lookup"><span data-stu-id="0f815-123">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="0f815-124">加载项中的 JavaScript 可以分析令牌并提取它所需的信息，如用户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="0f815-124">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="0f815-125">可选地，加载项可以向其服务器端发送 HTTP 请求以获取关于用户的更多数据；如用户的首选项。</span><span class="sxs-lookup"><span data-stu-id="0f815-125">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences.</span></span> <span data-ttu-id="0f815-126">或者，访问令牌本身可以发送到服务器端进行分析和验证。</span><span class="sxs-lookup"><span data-stu-id="0f815-126">Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="0f815-127">开发 SSO 加载项</span><span class="sxs-lookup"><span data-stu-id="0f815-127">Develop an SSO add-in</span></span>

<span data-ttu-id="0f815-128">此部分介绍了创建启用 SSO 的 Office 加载项所需完成的任务。</span><span class="sxs-lookup"><span data-stu-id="0f815-128">This section describes the tasks involved in creating an Office Add-in that uses SSO.</span></span> <span data-ttu-id="0f815-129">其中介绍的这些任务与语言和框架无关。</span><span class="sxs-lookup"><span data-stu-id="0f815-129">These tasks are described here in a language- and framework-agnostic way.</span></span> <span data-ttu-id="0f815-130">有关详细演练的示例，请参阅：</span><span class="sxs-lookup"><span data-stu-id="0f815-130">For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="0f815-131">创建使用单一登录的 Node.js Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0f815-131">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="0f815-132">创建使用单一登录的 ASP.NET Office 加载项</span><span class="sxs-lookup"><span data-stu-id="0f815-132">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="0f815-133">创建服务应用程序</span><span class="sxs-lookup"><span data-stu-id="0f815-133">Create the service application</span></span>

<span data-ttu-id="0f815-134">在 Azure v2.0 端点的注册门户注册加载项：https://apps.dev.microsoft.com。</span><span class="sxs-lookup"><span data-stu-id="0f815-134">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span> <span data-ttu-id="0f815-135">这是一个 5 – 10 分钟过程，包括以下任务：</span><span class="sxs-lookup"><span data-stu-id="0f815-135">This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="0f815-136">获取加载项的客户端 ID 和机密。</span><span class="sxs-lookup"><span data-stu-id="0f815-136">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="0f815-137">指定加载项访问 AAD v.</span><span class="sxs-lookup"><span data-stu-id="0f815-137">Specify the permissions that your add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="0f815-138">2.0 端点（以及可选的 Microsoft Graph）所需的权限。</span><span class="sxs-lookup"><span data-stu-id="0f815-138">2.0 endpoint (and optionally to Microsoft Graph).</span></span> <span data-ttu-id="0f815-139">总是需要“个人资料”权限。</span><span class="sxs-lookup"><span data-stu-id="0f815-139">The "profile" permission is always needed.</span></span>
* <span data-ttu-id="0f815-140">向加载项授予 Office 主机应用程序信任。</span><span class="sxs-lookup"><span data-stu-id="0f815-140">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="0f815-141">将 Office 主机应用程序预授权给具有 *access_as_user* 默认权限的加载项。</span><span class="sxs-lookup"><span data-stu-id="0f815-141">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="0f815-142">有关此过程的更多详细信息，请参阅[注册使用 SSO 和 Azure AD v2.0 端点的 Office 加载项](register-sso-add-in-aad-v2.md)。</span><span class="sxs-lookup"><span data-stu-id="0f815-142">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="0f815-143">配置加载项</span><span class="sxs-lookup"><span data-stu-id="0f815-143">Configure the add-in</span></span>

<span data-ttu-id="0f815-144">向外接程序清单添加新标记：</span><span class="sxs-lookup"><span data-stu-id="0f815-144">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="0f815-145">**WebApplicationInfo** - 下列元素的父项。</span><span class="sxs-lookup"><span data-stu-id="0f815-145">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="0f815-146">**ID** - 加载项的客户端 ID 这是你在注册加载项的过程中获得的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="0f815-146">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in.</span></span> <span data-ttu-id="0f815-147">请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项](register-sso-add-in-aad-v2.md)。</span><span class="sxs-lookup"><span data-stu-id="0f815-147">Details are at: [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="0f815-148">**Resource** - 加载项 URL。</span><span class="sxs-lookup"><span data-stu-id="0f815-148">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="0f815-149">**Scopes** - 一个或多个 **Scope** 元素的父元素。</span><span class="sxs-lookup"><span data-stu-id="0f815-149">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="0f815-150">**Scope** - 指定加载项访问 AAD 所需的权限。</span><span class="sxs-lookup"><span data-stu-id="0f815-150">**Scope** - Specifies a permission that the add-in needs to Microsoft Graph.</span></span> <span data-ttu-id="0f815-151">权限总是需要的，如果加载项不能访问 Microsoft Graph，它可能是唯一需要的权限。`profile`</span><span class="sxs-lookup"><span data-stu-id="0f815-151">The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph.</span></span> <span data-ttu-id="0f815-152">如果是这样，你也需要用于所需 Microsoft Graph 权限的 **作用域** 元素；例如：`User.Read`、`Mail.Read`。</span><span class="sxs-lookup"><span data-stu-id="0f815-152">If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`.</span></span> <span data-ttu-id="0f815-153">你在代码中用于访问 Microsoft Graph 的库可能需要额外的权限。</span><span class="sxs-lookup"><span data-stu-id="0f815-153">Libraries that you use in your code to access Microsoft Graph may need additional permissions.</span></span> <span data-ttu-id="0f815-154">例如，用于 .NET 的 Microsoft 身份验证库（MSAL）需要 `offline_access` 权限。</span><span class="sxs-lookup"><span data-stu-id="0f815-154">For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission.</span></span> <span data-ttu-id="0f815-155">有关更多信息，请参阅 [从 Office 加载项授权给 Microsoft Graph](authorize-to-microsoft-graph.md)。</span><span class="sxs-lookup"><span data-stu-id="0f815-155">For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="0f815-p111">对于除 Outlook 之外的 Office 主机，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 部分的末尾。对 Outlook，请将此标记添加到 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 部分的末尾。</span><span class="sxs-lookup"><span data-stu-id="0f815-p111">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="0f815-158">下面的示例展示了标记：</span><span class="sxs-lookup"><span data-stu-id="0f815-158">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="0f815-159">添加客户端代码</span><span class="sxs-lookup"><span data-stu-id="0f815-159">Add client-side code</span></span>

<span data-ttu-id="0f815-160">将 JavaScript 添加到加载项，以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="0f815-160">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="0f815-161">调用 [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync)。</span><span class="sxs-lookup"><span data-stu-id="0f815-161">Call [Office.context.auth.getAccessTokenAsync](https://dev.office.com/reference/add-ins/shared/office.context.auth.getAccessTokenAsync).</span></span>
* <span data-ttu-id="0f815-162">分析访问令牌或将其传递给加载项的服务器端代码。</span><span class="sxs-lookup"><span data-stu-id="0f815-162">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="0f815-163">下面是调用 `getAccessTokenAsync` 的简单例子。</span><span class="sxs-lookup"><span data-stu-id="0f815-163">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!Note]
> <span data-ttu-id="0f815-164">这个例子明确地只处理一种错误。</span><span class="sxs-lookup"><span data-stu-id="0f815-164">This example handles only one kind of error explicitly.</span></span> <span data-ttu-id="0f815-165">有关更详细的错误处理示例，请参阅 [Office-Add-in-ASPNET-SSO 中的 Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) 和 [Office-Add-in-NodeJS-SSO 中的 program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)。</span><span class="sxs-lookup"><span data-stu-id="0f815-165">For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js).</span></span> <span data-ttu-id="0f815-166">并参阅[单一登录 (SSO) 错误消息故障排除](troubleshoot-sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="0f815-166">Troubleshoot error messages for single sign-on (SSO)</span></span>
 

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

<span data-ttu-id="0f815-167">下面是将加载项令牌传递给服务器端的一个简单示例。</span><span class="sxs-lookup"><span data-stu-id="0f815-167">Here's a simple example of passing the add-in token to the server-side.</span></span> <span data-ttu-id="0f815-168">令牌包含在向服务器端发回请求时的 `Authorization` 标头中。</span><span class="sxs-lookup"><span data-stu-id="0f815-168">The token is included as an `Authorization` header when sending a request back to the server-side.</span></span> <span data-ttu-id="0f815-169">这个例子设想发送 JSON 数据，所以使用了 `POST` 方法，但是当不写入服务器时 `GET` 足以发送访问令牌。</span><span class="sxs-lookup"><span data-stu-id="0f815-169">This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="0f815-170">何时调用方法</span><span class="sxs-lookup"><span data-stu-id="0f815-170">When to call the method</span></span>

<span data-ttu-id="0f815-171">如果在没有用户登录到 Office 时无法使用加载项，那么*当加载项启动时*你应该调用 `getAccessTokenAsync`。</span><span class="sxs-lookup"><span data-stu-id="0f815-171">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="0f815-172">如果加载项具有某些不需要登录用户的功能，那么*当用户采取需要登录用户的操作时*你可以调用 `getAccessTokenAsync`。</span><span class="sxs-lookup"><span data-stu-id="0f815-172">If the add-in has some functionality that doesn't require access to Microsoft Graph or even a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires access to Microsoft Graph or, at least, a logged in user*.</span></span> <span data-ttu-id="0f815-173">`getAccessTokenAsync`  的冗余调用不会导致性能严重下降，因为 Office 缓存并重用没有过期的访问令牌，无需</span><span class="sxs-lookup"><span data-stu-id="0f815-173">There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD V. 2.0 endpoint whenever  is called.</span></span> <span data-ttu-id="0f815-174">在 调用 `getAccessTokenAsync` 时重新调用 AAD v.2.0 端点。</span><span class="sxs-lookup"><span data-stu-id="0f815-174">2.0 endpoint whenever `getAccessTokenAsync` is called.</span></span> <span data-ttu-id="0f815-175">因此，可以将 `getAccessTokenAsync` 调用添加到所有在需要令牌时启动操作的函数和处理程序。</span><span class="sxs-lookup"><span data-stu-id="0f815-175">So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="0f815-176">添加服务器端代码</span><span class="sxs-lookup"><span data-stu-id="0f815-176">Add server-side code</span></span>

<span data-ttu-id="0f815-177">大多数情况下，如果加载项没有将访问令牌传递到服务器端并在其中使用它，那么获取访问令牌的意义就不大。</span><span class="sxs-lookup"><span data-stu-id="0f815-177">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="0f815-178">加载项可以执行的一些服务器端任务：</span><span class="sxs-lookup"><span data-stu-id="0f815-178">Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="0f815-179">创建一个或多个 Web API 方法，使用从令牌中提取的有关用户的信息；例如，在托管数据库中查找用户首选项的方法。</span><span class="sxs-lookup"><span data-stu-id="0f815-179">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base.</span></span> <span data-ttu-id="0f815-180">（请参阅下面的 **使用 SSO 令牌作为标识** 。）根据你的语言和框架，可能会有库提供，因而简化你必须编写的代码。</span><span class="sxs-lookup"><span data-stu-id="0f815-180">(See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="0f815-181">获取 Microsoft Graph 数据。</span><span class="sxs-lookup"><span data-stu-id="0f815-181">Get Microsoft Graph data.</span></span> <span data-ttu-id="0f815-182">服务器端代码应执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="0f815-182">Your server-side code should do the following:</span></span>

    * <span data-ttu-id="0f815-183">验证访问令牌（请参阅下面的 **验证访问令牌** ）。</span><span class="sxs-lookup"><span data-stu-id="0f815-183">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="0f815-184">通过调用 Azure AD v2.0 端点来启动“代表”流程，此端点包含访问令牌、一些关于用户的元数据以及加载项凭据（ID 和机密）。</span><span class="sxs-lookup"><span data-stu-id="0f815-184">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the add-in access token, some metadata about the user, and the credentials of the add-in (its ID and secret).</span></span> <span data-ttu-id="0f815-185">在这种情况下，访问令牌称为引导令牌。</span><span class="sxs-lookup"><span data-stu-id="0f815-185">In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="0f815-186">缓存从代表流程返回的新访问令牌。</span><span class="sxs-lookup"><span data-stu-id="0f815-186">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="0f815-187">使用新令牌从 Microsoft Graph 获取数据。</span><span class="sxs-lookup"><span data-stu-id="0f815-187">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="0f815-188">有关获得用户 Microsoft Graph 数据的授权访问的更多详细信息，请参阅[在 Office 加载项中授权给 Microsoft Graph](authorize-to-microsoft-graph.md)。</span><span class="sxs-lookup"><span data-stu-id="0f815-188">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="0f815-189">验证访问令牌</span><span class="sxs-lookup"><span data-stu-id="0f815-189">For more information, see Validate the access token.</span></span>

<span data-ttu-id="0f815-190">Web API 收到访问令牌后，必须在使用该令牌前对其进行验证。</span><span class="sxs-lookup"><span data-stu-id="0f815-190">Once the Web API receives the access token, it must validate it before using it.</span></span> <span data-ttu-id="0f815-191">该令牌是 JSON Web 令牌 (JWT)，这意味着验证方式与大多数标准 OAuth 流中的令牌验证方式类似。</span><span class="sxs-lookup"><span data-stu-id="0f815-191">The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows.</span></span> <span data-ttu-id="0f815-192">有许多可用于处理 JWT 验证的库，而它们的基本内容为：</span><span class="sxs-lookup"><span data-stu-id="0f815-192">There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="0f815-193">检查令牌的格式是否正确</span><span class="sxs-lookup"><span data-stu-id="0f815-193">Checking that the token is well-formed</span></span>
- <span data-ttu-id="0f815-194">检查令牌是否由预期的颁发机构颁发</span><span class="sxs-lookup"><span data-stu-id="0f815-194">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="0f815-195">检查令牌是否是针对 Web API</span><span class="sxs-lookup"><span data-stu-id="0f815-195">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="0f815-196">验证令牌时，请牢记以下准则：</span><span class="sxs-lookup"><span data-stu-id="0f815-196">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="0f815-197">有效的 SSO 令牌是由 Azure 颁发机构 `https://login.microsoftonline.com` 的。</span><span class="sxs-lookup"><span data-stu-id="0f815-197">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`.</span></span> <span data-ttu-id="0f815-198">令牌中的 `iss` 声明应以此值开头。</span><span class="sxs-lookup"><span data-stu-id="0f815-198">The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="0f815-199">令牌的 `aud` 参数将被设置为加载项注册的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="0f815-199">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="0f815-200">令牌的 `scp` 参数将被设置为 `access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="0f815-200">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="0f815-201">将 SSO 令牌用作标识</span><span class="sxs-lookup"><span data-stu-id="0f815-201">Using the SSO token as an identity</span></span>

<span data-ttu-id="0f815-202">如果加载项需要验证用户标识，则 SSO 令牌包含的信息可用于创建此标识。</span><span class="sxs-lookup"><span data-stu-id="0f815-202">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity.</span></span> <span data-ttu-id="0f815-203">令牌中的以下声明与标识相关。</span><span class="sxs-lookup"><span data-stu-id="0f815-203">The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="0f815-204">`name` - 用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="0f815-204">`name` - The user's display name.</span></span>
- <span data-ttu-id="0f815-205">`preferred_username` - 用户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="0f815-205">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="0f815-206">`oid` - 表示 Azure Active Directory 中的用户 ID 的 GUID。</span><span class="sxs-lookup"><span data-stu-id="0f815-206">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="0f815-207">`tid` - 表示 Azure Active Directory 中的用户组织 ID 的 GUID。</span><span class="sxs-lookup"><span data-stu-id="0f815-207">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="0f815-208">因为 `name` 和 `preferred_username` 值可以更改，我们建议使用 `oid` 和 `tid` 值将标识与后端的授权服务关联。</span><span class="sxs-lookup"><span data-stu-id="0f815-208">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="0f815-209">例如，你的服务可以将这些值组合在一起，并设置为类似 `{oid-value}@{tid-value}` 的格式，然后将其存储为内部用户数据库中的用户记录值。</span><span class="sxs-lookup"><span data-stu-id="0f815-209">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database.</span></span> <span data-ttu-id="0f815-210">然后，在后续的请求中，可以使用同一值检索此用户，并可基于现有访问控制机制确定对特定资源的访问。</span><span class="sxs-lookup"><span data-stu-id="0f815-210">Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="0f815-211">示例访问令牌</span><span class="sxs-lookup"><span data-stu-id="0f815-211">Example access token</span></span>

<span data-ttu-id="0f815-212">以下是访问令牌的典型解码有效负载。</span><span class="sxs-lookup"><span data-stu-id="0f815-212">The following is a typical decoded payload of an access token.</span></span> <span data-ttu-id="0f815-213">有关这些属性的信息，请参阅 [Azure Active Directory v2.0 令牌引用](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens)。</span><span class="sxs-lookup"><span data-stu-id="0f815-213">For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="0f815-214">在 Outlook 加载项中使用 SSO</span><span class="sxs-lookup"><span data-stu-id="0f815-214">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="0f815-p124">在 Outlook 加载项中使用与在 Excel、 PowerPoint、或 Word 加载项中使用 SSO 之间有一些细微、但很重要的差别。请务必阅读 [使用 Outlook 加载项中的单一登录令牌对用户进行身份验证](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)和[方案：在 Outlook 加载项中对你的服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。</span><span class="sxs-lookup"><span data-stu-id="0f815-p124">There are some small, but important differences in using SSO in and Outlook add-in from using it in as Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="0f815-217">SSO API 引用</span><span class="sxs-lookup"><span data-stu-id="0f815-217">SSO API reference</span></span>

<span data-ttu-id="0f815-218">Office 身份验证命名空间 `Office.context.auth` 提供了方法 `getAccessTokenAsync`，使 Office 主机可以获取加载项 Web 应用程序的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="0f815-218">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="0f815-219">这样间接地使加载项能够间接访问已登录用户的 Microsoft Graph 数据，而不需要用户二次登录。</span><span class="sxs-lookup"><span data-stu-id="0f815-219">Indirectly, enable the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="0f815-220">调用 Azure Active Directory V 2.0 端点以获取令牌来访问加载项的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="0f815-220">Calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application.</span></span> <span data-ttu-id="0f815-221">这样加载项便能识别用户。</span><span class="sxs-lookup"><span data-stu-id="0f815-221">This enables add-ins to identify users.</span></span> <span data-ttu-id="0f815-222">通过[“代表”OAuth 流](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)，服务器端代码可以使用此令牌访问加载项 Web 应用程序的 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="0f815-222">Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!注意在 Outlook 中，如果加载项加载于 Outlook.com 或 Gmail 邮箱，此 API 将不受支持。]

<table><tr><td><span data-ttu-id="0f815-224">主机</span><span class="sxs-lookup"><span data-stu-id="0f815-224">Hosts</span></span></td><td><span data-ttu-id="0f815-225">Excel、Outlook、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="0f815-225">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="0f815-226">要求集</span><span class="sxs-lookup"><span data-stu-id="0f815-226">Requirement sets</span></span></td><td>[<span data-ttu-id="0f815-227">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="0f815-227">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

### <a name="parameters"></a><span data-ttu-id="0f815-228">参数</span><span class="sxs-lookup"><span data-stu-id="0f815-228">Parameters</span></span>

<span data-ttu-id="0f815-229">`options` - 可选。</span><span class="sxs-lookup"><span data-stu-id="0f815-229">`options` - Optional.</span></span> <span data-ttu-id="0f815-230">接受 `AuthOptions` 对象 （请参阅下面）以定义登录行为。</span><span class="sxs-lookup"><span data-stu-id="0f815-230">Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="0f815-231">`callback` - 可选。</span><span class="sxs-lookup"><span data-stu-id="0f815-231">`callback` - Optional.</span></span> <span data-ttu-id="0f815-232">接受一个回调方法，可为用户的 ID 分析令牌或使用“代表”流中的令牌获取 Microsoft Graph 访问权限。</span><span class="sxs-lookup"><span data-stu-id="0f815-232">Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph.</span></span> <span data-ttu-id="0f815-233">|||UNTRANSLATED_CONTENT_START|||If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="0f815-233">If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v.</span></span> <span data-ttu-id="0f815-234">2.0 格式的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="0f815-234">2.0-formatted access token.</span></span>

<span data-ttu-id="0f815-235"> `AuthOptions` 接口提供了 Office 从 AAD v.</span><span class="sxs-lookup"><span data-stu-id="0f815-235">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v.</span></span> <span data-ttu-id="0f815-236">2.0 通过 `getAccessTokenAsync` 方法获取加载项访问令牌时的用户体验选项。</span><span class="sxs-lookup"><span data-stu-id="0f815-236">2.0 with the `getAccessTokenAsync` method.</span></span>

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



