---
title: 为 Office 加载项启用单一登录
description: ''
ms.date: 09/26/2018
ms.openlocfilehash: fb4eacee9419339116e15ef3fccc03b291faf3ec
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506026"
---
# <a name="enable-single-sign-on-for-office-add-ins-preview"></a><span data-ttu-id="91de6-102">为 Office 加载项启用单一登录（预览）</span><span class="sxs-lookup"><span data-stu-id="91de6-102">Enable single sign-on for Office Add-ins (preview)</span></span>

<span data-ttu-id="91de6-p101">用户使用个人 Microsoft 账户或工作/学校（Office 365）账户登录到 Office （在线、移动设备和桌面平台） 。可以趁机使用单一登录 (SSO) 来授权用户使用加载项而无需用户二次登陆。</span><span class="sxs-lookup"><span data-stu-id="91de6-p101">Users sign in to Office (online, mobile, and desktop platforms) using either their personal Microsoft account or their work or school (Office 365) account. You can take advantage of this and use single sign-on (SSO) to authorize the user to your add-in without requiring the user to sign in a second time.</span></span>

![显示加载项登录过程的图像](../images/office-host-title-bar-sign-in.png)

### <a name="preview-status"></a><span data-ttu-id="91de6-106">预览状态</span><span class="sxs-lookup"><span data-stu-id="91de6-106">Preview Status</span></span>

<span data-ttu-id="91de6-p102">单一登录 API 只支持预览。可供开发人员实验；但是，不应用于生产加载项。此外, 使用 SSO 的加载项不被 [AppSource](https://appsource.microsoft.com) 接受。</span><span class="sxs-lookup"><span data-stu-id="91de6-p102">The Single Sign-on API is currently supported in preview only. It is available to developers for experimentation; but it should not be used in a production add-in. In addition, add-ins that use SSO are not accepted in [AppSource](https://appsource.microsoft.com).</span></span>

<span data-ttu-id="91de6-p103">并非所有 Office 应用都支持 SSO 预览。在 Word、 Excel、 Outlook 和 PowerPoint 中可用。目前关于哪里支持单一登录 API 的详细信息，请参阅 [IdentityAPI 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js) 。</span><span class="sxs-lookup"><span data-stu-id="91de6-p103">Not all Office applications support the SSO preview. It is available in Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js).</span></span>

### <a name="requirements-and-best-practices"></a><span data-ttu-id="91de6-113">要求和最佳做法</span><span class="sxs-lookup"><span data-stu-id="91de6-113">Requirements and Best Practices</span></span>

<span data-ttu-id="91de6-114">若要使用 SSO，必须从加载项启动 HTML 页的 `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` 中加载 beta 版的 Office JavaScript 库。</span><span class="sxs-lookup"><span data-stu-id="91de6-114">To use SSO, you must load the beta version of the Office JavaScript Library from `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` in the startup HTML page of the add-in.</span></span>

<span data-ttu-id="91de6-p104">若正在使用 **Outlook** 加载项，须为 Office 365 租户启用新式身份验证。有关如何执行此操作的信息，请参阅 [Exchange Online：如何为租户启用新式身份验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx) 。</span><span class="sxs-lookup"><span data-stu-id="91de6-p104">If you are working with an **Outlook** add-in, be sure to enable Modern Authentication for the Office 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="91de6-p105">应该 *不* 依赖于 SSO 作为加载项的唯一的身份验证方法。当加载项在某种错误情况回退时，应执行一个备用的身份验证系统。可以使用用户表和身份验证系统，或利用一个社交登录服务商。欲知如何使用 Office 加载项执行此操作的详细信息，请参阅 [在 Office 加载项中授权外部服务](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins) 。对 *Outlook* 建议一个回退系统。欲知详情，请参阅 [方案：在 Outlook 加载项中实现服务单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) 。</span><span class="sxs-lookup"><span data-stu-id="91de6-p105">You should *not* rely on SSO as your add-in's only method of authentication. You should implement an alternate authentication system that your add-in can fall back to in certain error situations. You can use a system of user tables and authentication, or you can leverage one of the social login providers. For more information about how to do this with an Office add-in, see [Authorize external services in your Office Add-in](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/auth-external-add-ins). For *Outlook*, there is a recommended fall back system. For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

### <a name="how-sso-works-at-runtime"></a><span data-ttu-id="91de6-123">运行时 SSO 工作方式</span><span class="sxs-lookup"><span data-stu-id="91de6-123">How it works at runtime</span></span>

<span data-ttu-id="91de6-124">以下关系图显示了 SSO 流程的工作方式。</span><span class="sxs-lookup"><span data-stu-id="91de6-124">The following diagram shows how the SSO process works.</span></span>

![SSO 过程关系图](../images/sso-overview-diagram.png)

1. <span data-ttu-id="91de6-p106">在外接程序 JavaScript 调用新的 Office.js API [getAccessTokenAsync](#sso-api-reference)。这会告诉 Office 主机应用程序获取访问令牌到外接程序。请参阅 [示例访问令牌](#example-access-token)。</span><span class="sxs-lookup"><span data-stu-id="91de6-p106">In the add-in, JavaScript calls a new Office.js API [getAccessTokenAsync](#sso-api-reference). This tells the Office host application to obtain an access token to the add-in. See [Example access token](#example-access-token).</span></span>
2. <span data-ttu-id="91de6-129">如果用户未登录，Office 主机应用会打开弹出窗口，以供用户登录。</span><span class="sxs-lookup"><span data-stu-id="91de6-129">If the user is not signed in, the Office host application opens a pop-up window for the user to sign in.</span></span>
3. <span data-ttu-id="91de6-130">如果当前用户是首次使用加载项，则会看到同意提示。</span><span class="sxs-lookup"><span data-stu-id="91de6-130">If this is the first time the current user has used your add-in, he or she is prompted to consent.</span></span>
4. <span data-ttu-id="91de6-131">Office 主机应用程序从当前用户的 Azure AD v2.0 端点请求获取**加载项令牌**。</span><span class="sxs-lookup"><span data-stu-id="91de6-131">The Office host application requests the **add-in token** from the Azure AD v2.0 endpoint for the current user.</span></span>
5. <span data-ttu-id="91de6-132">Azure AD 将加载项令牌发送给 Office 主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="91de6-132">Azure AD sends the add-in token to the Office host application.</span></span>
6. <span data-ttu-id="91de6-133">作为 `getAccessTokenAsync` 调用返回的结果对象的一部分，Office 主机应用程序将**加载项令牌**发送给加载项。</span><span class="sxs-lookup"><span data-stu-id="91de6-133">The Office host application sends the **add-in token** to the add-in as part of the result object returned by the `getAccessTokenAsync` call.</span></span>
7. <span data-ttu-id="91de6-134">加载项中的 JavaScript 可以分析令牌并提取它所需的信息，如用户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="91de6-134">JavaScript in the add-in can parse the token and extract the information it needs, such as the user's email address.</span></span> 
8. <span data-ttu-id="91de6-p107">加载项可以向服务器端发送 HTTP 请求来获取更多用户数据；比如用户偏好。此外，访问令牌本身可被发送到服务器端进行分析和检验。</span><span class="sxs-lookup"><span data-stu-id="91de6-p107">Optionally, the add-in can send HTTP request to its server-side for more data about the user; such as the user's preferences. Alternatively, the access token itself could be sent to the server-side for parsing and validation there.</span></span> 

## <a name="develop-an-sso-add-in"></a><span data-ttu-id="91de6-137">开发 SSO 加载项</span><span class="sxs-lookup"><span data-stu-id="91de6-137">Develop an SSO add-in</span></span>

<span data-ttu-id="91de6-p108">本节介绍创建使用 SSO Office 的加载项时的任务。这些任务再次以语言-框架-不可知论证的方式描述。详细步骤，请参阅：</span><span class="sxs-lookup"><span data-stu-id="91de6-p108">This section describes the tasks involved in creating an Office Add-in that uses SSO. These tasks are described here in a language- and framework-agnostic way. For examples of detailed walkthroughs, see:</span></span>

* [<span data-ttu-id="91de6-141">创建使用单一登录的 Node.js Office 加载项</span><span class="sxs-lookup"><span data-stu-id="91de6-141">Create a Node.js Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-nodejs.md)
* [<span data-ttu-id="91de6-142">创建使用单一登录的 ASP.NET Office 加载项</span><span class="sxs-lookup"><span data-stu-id="91de6-142">Create an ASP.NET Office Add-in that uses single sign-on</span></span>](create-sso-office-add-ins-aspnet.md)

### <a name="create-the-service-application"></a><span data-ttu-id="91de6-143">创建服务应用程序</span><span class="sxs-lookup"><span data-stu-id="91de6-143">Create the service application</span></span>

<span data-ttu-id="91de6-p109">在 Azure v2.0 端点的注册门户注册加载项：https://apps.dev.microsoft.com。此过程用时 5-10 分钟，包括以下任务：</span><span class="sxs-lookup"><span data-stu-id="91de6-p109">Register the add-in at the registration portal for the Azure v2.0 endpoint: https://apps.dev.microsoft.com. This is a 5–10 minute process that includes the following tasks:</span></span>

* <span data-ttu-id="91de6-146">为加载项获取客户 ID 和机密。</span><span class="sxs-lookup"><span data-stu-id="91de6-146">Get a client ID and secret for the add-in.</span></span>
* <span data-ttu-id="91de6-p110">加载项需要 AAD v. 2.0 终点 （可选 Microsoft Graph） 许可。始终需要"配置文件"权限。</span><span class="sxs-lookup"><span data-stu-id="91de6-p110">Specify the permissions that your add-in needs to AAD v. 2.0 endpoint (and optionally to Microsoft Graph). The "profile" permission is always needed.</span></span>
* <span data-ttu-id="91de6-150">向加载项授予 Office 主机应用信任。</span><span class="sxs-lookup"><span data-stu-id="91de6-150">Grant the Office host application trust to the add-in.</span></span>
* <span data-ttu-id="91de6-151">将 Office 主机应用程序预授权给具有 *access_as_user* 默认权限的加载项。</span><span class="sxs-lookup"><span data-stu-id="91de6-151">Preauthorize the Office host application to the add-in with the default permission *access_as_user*.</span></span>

<span data-ttu-id="91de6-152">有关此过程的更多详细信息，请参阅[注册使用 SSO 和 Azure AD v2.0 端点的 Office 加载项](register-sso-add-in-aad-v2.md)。</span><span class="sxs-lookup"><span data-stu-id="91de6-152">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>

### <a name="configure-the-add-in"></a><span data-ttu-id="91de6-153">配置加载项</span><span class="sxs-lookup"><span data-stu-id="91de6-153">Configure the add-in</span></span>

<span data-ttu-id="91de6-154">向加载项清单添加新标记：</span><span class="sxs-lookup"><span data-stu-id="91de6-154">Add new markup to the add-in manifest:</span></span>

* <span data-ttu-id="91de6-155">**WebApplicationInfo** - 下列元素的母元素。</span><span class="sxs-lookup"><span data-stu-id="91de6-155">**WebApplicationInfo** - The parent of the following elements.</span></span>
* <span data-ttu-id="91de6-p111">**Id** - 客户加载项 ID ,这是注册加载项的申请 ID。参阅 [注册使用 Azure AD v2.0 终点的加载项](register-sso-add-in-aad-v2.md) 。</span><span class="sxs-lookup"><span data-stu-id="91de6-p111">**Id** - The client ID of the add-in This is an application ID that you obtain as part of registering the add-in. See [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](register-sso-add-in-aad-v2.md).</span></span>
* <span data-ttu-id="91de6-158">**Resource** - 加载项 URL。</span><span class="sxs-lookup"><span data-stu-id="91de6-158">**Resource** - The URL of the add-in.</span></span>
* <span data-ttu-id="91de6-159">**Scopes** - 一个或多个 **Scope** 元素的母元素.</span><span class="sxs-lookup"><span data-stu-id="91de6-159">**Scopes** - The parent of one or more **Scope** elements.</span></span>
* <span data-ttu-id="91de6-p112">**范围** 指加载项需要 AAD 的许可。如果加载项不访问 Microsoft Graph的话， `profile` 就是始终并唯一需要的权限。如果是，可能还需要 **范围** 元素以获取所需的 Microsoft Graph 许可; 例如， `User.Read`， `Mail.Read`。用于访问 Microsoft Graph 的代码库可能需要其他权限。例如，适用于.NET 的 Microsoft 身份验证库 (MSAL) 需要 `offline_access` 权限。详细信息，参阅 [ Office 加载项授权 Microsoft Graph](authorize-to-microsoft-graph.md)。</span><span class="sxs-lookup"><span data-stu-id="91de6-p112">**Scope** - Specifies a permission that the add-in needs to AAD. The `profile` permission is always needed and it may be the only permission needed, if your add-in does not access Microsoft Graph. If it does, you also need **Scope** elements for the required Microsoft Graph permissions; for example, `User.Read`, `Mail.Read`. Libraries that you use in your code to access Microsoft Graph may need additional permissions. For example, Microsoft Authentication Library (MSAL) for .NET requires `offline_access` permission. For more information, see [Authorize to Microsoft Graph from an Office Add-in](authorize-to-microsoft-graph.md).</span></span>

<span data-ttu-id="91de6-p113">对于除 Outlook 之外的 Office 主机，在 `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` 文末添加标记。对于 Outlook，在 `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` 文末添加标记。</span><span class="sxs-lookup"><span data-stu-id="91de6-p113">For Office hosts other than Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_0">` section. For Outlook, add the markup to the end of the `<VersionOverrides ... xsi:type="VersionOverridesV1_1">` section.</span></span>

<span data-ttu-id="91de6-168">以下是标记的示例：</span><span class="sxs-lookup"><span data-stu-id="91de6-168">The following is an example of the markup:</span></span>

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

### <a name="add-client-side-code"></a><span data-ttu-id="91de6-169">添加客户端代码</span><span class="sxs-lookup"><span data-stu-id="91de6-169">Add client-side code</span></span>

<span data-ttu-id="91de6-170">将 JavaScript 添加到加载项，以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="91de6-170">Add JavaScript to the add-in to:</span></span>

* <span data-ttu-id="91de6-171">调用 [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)。</span><span class="sxs-lookup"><span data-stu-id="91de6-171">Call [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference).</span></span>

* <span data-ttu-id="91de6-172">分析访问令牌或将其传递给加载项的服务器端代码。</span><span class="sxs-lookup"><span data-stu-id="91de6-172">Parse the access token or pass it to the add-in’s server-side code.</span></span> 

<span data-ttu-id="91de6-173">下面是调用 `getAccessTokenAsync` 的简单例子。</span><span class="sxs-lookup"><span data-stu-id="91de6-173">Here's a simple example of a call to `getAccessTokenAsync`.</span></span> 

> [!NOTE]
> <span data-ttu-id="91de6-p114">此例只明显处理一种错误类型。更精细的错误处理的示例，请参阅 [Office-加载项-SPNET- SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) 和 [program.js in Office-加载项-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)。请参阅 [诊断错误消息的单一登录 (SSO)](troubleshoot-sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="91de6-p114">This example handles only one kind of error explicitly. For examples of more elaborate error handling, see [Home.js in Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js) and [program.js in Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js). And see [Troubleshoot error messages for single sign-on (SSO)](troubleshoot-sso-in-office-add-ins.md).</span></span>
 

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

<span data-ttu-id="91de6-p115">下面是将加载项令牌传递到服务器端的简单示例。需求被送回服务器端时，令牌标题为 `Authorization` 。本示例预期发送 JSON 数据，因此使用 `POST` 方法，但 `GET` 就足够在不写入服务器时发送访问令牌。</span><span class="sxs-lookup"><span data-stu-id="91de6-p115">Here's a simple example of passing the add-in token to the server-side. The token is included as an `Authorization` header when sending a request back to the server-side. This example envisions sending JSON data, so it uses the `POST` method, but `GET` is sufficient to send the access token when you are not writing to the server.</span></span>

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

#### <a name="when-to-call-the-method"></a><span data-ttu-id="91de6-180">何时调用方法</span><span class="sxs-lookup"><span data-stu-id="91de6-180">When to call the method</span></span>

<span data-ttu-id="91de6-181">如果在没有用户登录到 Office 时无法使用加载项，应调用 `getAccessTokenAsync` *当加载项启动时* 。</span><span class="sxs-lookup"><span data-stu-id="91de6-181">If your add-in cannot be used when a no user is logged into Office and Office does not have an access token to your add-in, then you should call `getAccessTokenAsync` *when the add-in launches*.</span></span>

<span data-ttu-id="91de6-p116">如果加载项不需要登录某些功能，然后调用`getAccessTokenAsync`*如果用户执行需要登录的用户的操作* 。没有冗余的调用不会显著的性能下降`getAccessTokenAsync`因为 Office 缓存访问令牌并将反复使用，直到过期，而不会每次`getAccessTokenAsync`被调用时都尝试调用 AAD v. 2.0 终结点。这样可将添加调用`getAccessTokenAsync`到所有需要令牌的功能和处理程序来启动操作。</span><span class="sxs-lookup"><span data-stu-id="91de6-p116">If the add-in has some functionality that doesn't require a logged in user, then you call `getAccessTokenAsync` *when the user takes an action that requires a logged in user*. There is no significant performance degradation with redundant calls of `getAccessTokenAsync` because Office caches the access token and will reuse it, until it expires, without making another call to the AAD v. 2.0 endpoint whenever `getAccessTokenAsync` is called. So you can add calls of `getAccessTokenAsync` to all functions and handlers that initiate an action where the token is needed.</span></span>

### <a name="add-server-side-code"></a><span data-ttu-id="91de6-186">添加服务器端代码</span><span class="sxs-lookup"><span data-stu-id="91de6-186">Add server-side code</span></span>

<span data-ttu-id="91de6-p117">如果加载项没有将访问令牌传递到服务器端并在服务器端使用它，多数情况下获取访问令牌意义极小。加载项可以在服务器端做如下事情：</span><span class="sxs-lookup"><span data-stu-id="91de6-p117">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there. Some server-side tasks your add-in could do:</span></span>

* <span data-ttu-id="91de6-p118">多创建一个或多个 Web API 方法，该方法从令牌中提取客户信息；比如，从托管数据库查看用户偏好（参阅如下 **使用 SSO 令牌作为身份** ）。根据语言和框架，库可用来简化代码编写。</span><span class="sxs-lookup"><span data-stu-id="91de6-p118">Create one or more Web API methods that use information about the user that is extracted from the token; for example, a method that looks up the user's preferences in your hosted data base. (See **Using the SSO token as an identity** below.) Depending on your language and framework, libraries might be available that will simplify the code you have to write.</span></span>
* <span data-ttu-id="91de6-p119">获取 Microsoft Graph 数据。服务器端代码应执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="91de6-p119">Get Microsoft Graph data. Your server-side code should do the following:</span></span>

    * <span data-ttu-id="91de6-193">验证访问令牌（请参阅如下 **验证访问令牌** ）。</span><span class="sxs-lookup"><span data-stu-id="91de6-193">Validate the access token (see **Validate the access token** below).</span></span>
    * <span data-ttu-id="91de6-p120">开始“代表”流调用 Azure AD v2.0 结点，包括访问令牌、 用户元数据，以及加载项身份 （其 ID 和机密）。在此上下文中，访问令牌被称为 bootstrap 令牌。</span><span class="sxs-lookup"><span data-stu-id="91de6-p120">Initiate the “on behalf of” flow with a call to the Azure AD v2.0 endpoint that includes the access token, some metadata about the user, and the credentials of the add-in (its ID and secret). In this context, the access token is called the bootstrap token.</span></span>
    * <span data-ttu-id="91de6-196">缓存从代表流程返回的新访问令牌。</span><span class="sxs-lookup"><span data-stu-id="91de6-196">Cache the new access token that is returned from the on-behalf-of flow.</span></span>
    * <span data-ttu-id="91de6-197">使用新令牌从 Microsoft Graph 获取数据。</span><span class="sxs-lookup"><span data-stu-id="91de6-197">Get data from Microsoft Graph by using the MSG token.</span></span>

 <span data-ttu-id="91de6-198">有关获得对用户 Microsoft Graph 数据的授权访问的更多详细信息，请参阅[在 Office 加载项中授权给 Microsoft Graph](authorize-to-microsoft-graph.md)。</span><span class="sxs-lookup"><span data-stu-id="91de6-198">For more details about getting authorized access to the user's Microsoft Graph data, see [Authorize to Microsoft Graph in your Office Add-in](authorize-to-microsoft-graph.md).</span></span>

#### <a name="validate-the-access-token"></a><span data-ttu-id="91de6-199">验证访问令牌</span><span class="sxs-lookup"><span data-stu-id="91de6-199">For more information, see Validate the access token.</span></span>

<span data-ttu-id="91de6-p121">一旦 Web API 接收访问令牌，必须在使用它之前进行验证。该标记是 JSON Web 令牌 (JWT)，这意味着验证与多数标准的 OAuth 流中的令牌验证一致。有许多库可用来处理 JWT 验证，但基础库包括：</span><span class="sxs-lookup"><span data-stu-id="91de6-p121">Once the Web API receives the access token, it must validate it before using it. The token is a JSON Web Token (JWT), which means that validation works just like token validation in most standard OAuth flows. There are a number of libraries available that can handle JWT validation, but the basics include:</span></span>

- <span data-ttu-id="91de6-203">检查令牌的格式是否正确</span><span class="sxs-lookup"><span data-stu-id="91de6-203">Checking that the token is well-formed</span></span>
- <span data-ttu-id="91de6-204">检查令牌是否由预期的颁发机构颁发</span><span class="sxs-lookup"><span data-stu-id="91de6-204">Checking that the token was issued by the intended authority</span></span>
- <span data-ttu-id="91de6-205">检查令牌是否是针对 Web API</span><span class="sxs-lookup"><span data-stu-id="91de6-205">Checking that the token is targeted to the Web API</span></span>

<span data-ttu-id="91de6-206">验证令牌时，请记住以下准则：</span><span class="sxs-lookup"><span data-stu-id="91de6-206">Keep in mind the following guidelines when validating the token:</span></span>

- <span data-ttu-id="91de6-p122">有效的 SSO 令牌于 Azure 权威颁发， `https://login.microsoftonline.com`。 `iss` 声明令牌的值应由此开始。</span><span class="sxs-lookup"><span data-stu-id="91de6-p122">Valid SSO tokens will be issued by the Azure authority, `https://login.microsoftonline.com`. The `iss` claim in the token should start with this value.</span></span>
- <span data-ttu-id="91de6-209">令牌的 `aud` 参数将用来设置加载项注册的应用程序 ID。</span><span class="sxs-lookup"><span data-stu-id="91de6-209">The token's `aud` parameter will be set to the application ID of the add-in's registration.</span></span>
- <span data-ttu-id="91de6-210">将令牌的 `scp` 参数设置为 `access_as_user`。</span><span class="sxs-lookup"><span data-stu-id="91de6-210">The token's `scp` parameter will be set to `access_as_user`.</span></span>

#### <a name="using-the-sso-token-as-an-identity"></a><span data-ttu-id="91de6-211">将 SSO 令牌用作标识</span><span class="sxs-lookup"><span data-stu-id="91de6-211">Using the SSO token as an identity</span></span>

<span data-ttu-id="91de6-p123">如果加载项需要验证用户的身份，SSO 令牌包含可以用于建立标识的信息。令牌中与标识相关的有如下几点。</span><span class="sxs-lookup"><span data-stu-id="91de6-p123">If your add-in needs to verify the user's identity, the SSO token contains information that can be used to establish the identity. The following claims in the token relate to identity.</span></span>

- <span data-ttu-id="91de6-214">`name` - 用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="91de6-214">`name` - The user's display name.</span></span>
- <span data-ttu-id="91de6-215">`preferred_username` - 用户的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="91de6-215">`preferred_username`The user's email address.</span></span>
- <span data-ttu-id="91de6-216">`oid` - 表示 Azure Active Directory 中的用户 ID 的 GUID。</span><span class="sxs-lookup"><span data-stu-id="91de6-216">`oid` - A GUID representing the ID of the user in the Azure Active Directory.</span></span>
- <span data-ttu-id="91de6-217">`tid` - 表示 Azure Active Directory 中的用户组织 ID 的 GUID。</span><span class="sxs-lookup"><span data-stu-id="91de6-217">`tid` - A GUID representing the ID of the user's organization in the Azure Active Directory.</span></span>

<span data-ttu-id="91de6-218">因为 `name` 和 `preferred_username`值可以更改，我们建议使用 `oid`和 `tid` 值将标识与后端的授权服务关联。</span><span class="sxs-lookup"><span data-stu-id="91de6-218">Since the `name` and `preferred_username` values could change, it's recommended that the `oid` and `tid` values be used to correlate the identity with your back-end's authorization service.</span></span>

<span data-ttu-id="91de6-p124">例如，服务器无法格式化这些值如 `{oid-value}@{tid-value}`，那么将这些值存储为内部用户数据库的记录。后续若有请求时，可以通过检索相同的值来找到用户，并根据现有访问机制决定获取特定资源的访问权限。</span><span class="sxs-lookup"><span data-stu-id="91de6-p124">For example, your service could format those values together like `{oid-value}@{tid-value}`, then store that as a value on the user's record in your internal user database. Then on subsequent requests, the user could be retrieved by using the same value, and access to specific resources could be determined based on your existing access control mechanisms.</span></span>

### <a name="example-access-token"></a><span data-ttu-id="91de6-221">访问令牌样本</span><span class="sxs-lookup"><span data-stu-id="91de6-221">Example access token</span></span>

<span data-ttu-id="91de6-p125">下面是典型的解码有效负载的访问令牌。欲知属性的信息，请参阅 [Azure Active Directory v2.0 令牌参照](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens) 。</span><span class="sxs-lookup"><span data-stu-id="91de6-p125">The following is a typical decoded payload of an access token. For information about the properties, see [Azure Active Directory v2.0 tokens reference](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-tokens).</span></span>


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

## <a name="using-sso-with-an-outlook-add-in"></a><span data-ttu-id="91de6-224">在 Outlook 加载项中使用 SSO</span><span class="sxs-lookup"><span data-stu-id="91de6-224">Using SSO with and Outlook add-in</span></span>

<span data-ttu-id="91de6-p126">在 Outlook 加载项中使用与在 Excel、 PowerPoint、或 Word 加载项中使用 SSO 之间有一些细微、但很重要的差别。请务必阅读 [使用 Outlook 加载项中的单一登录令牌对用户进行身份验证](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)和[方案：在 Outlook 加载项中对你的服务实现单一登录](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)。</span><span class="sxs-lookup"><span data-stu-id="91de6-p126">There are some small, but important differences in using SSO in an Outlook add-in from using it in an Excel, PowerPoint, or Word add-in. Be sure to read [Authenticate a user with a single sign-on token in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) and [Scenario: Implement single sign-on to your service in an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in).</span></span>

## <a name="sso-api-reference"></a><span data-ttu-id="91de6-227">SSO API 参照</span><span class="sxs-lookup"><span data-stu-id="91de6-227">SSO API reference</span></span>

### <a name="getaccesstokenasync"></a><span data-ttu-id="91de6-228">getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="91de6-228">getAccessTokenAsync</span></span>

<span data-ttu-id="91de6-p127">Office Auth 命名空间， `Office.context.auth`，提供了 Office 主机获取登陆令牌以完成加载项网申的方法 `getAccessTokenAsync` 。并且非直接地允许加载项访问已登录的用户的 Microsoft Graph 数据而无需用户再次登录。</span><span class="sxs-lookup"><span data-stu-id="91de6-p127">The Office Auth namespace, `Office.context.auth`, provides a method, `getAccessTokenAsync` that enables the Office host to obtain an access token to the add-in's web application. Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

```typescript
getAccessTokenAsync(options?: AuthOptions, callback?: (result: AsyncResult<string>) => void): void;
```

<span data-ttu-id="91de6-p128">该方法调用了 Azure Active Directory V 2.0 终结点来获取加载项网络申请的登陆令牌。使加载项可以识别用户。服务器端代码可以使用此令牌，用["代表" OAuth 流](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) 接入 Microsoft Graph。</span><span class="sxs-lookup"><span data-stu-id="91de6-p128">The method calls the Azure Active Directory V 2.0 endpoint to get an access token to your add-in's web application. This enables add-ins to identify users. Server side code can use this token to access Microsoft Graph for the add-in's web application by using the ["on behalf of" OAuth flow](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span>

> [!NOTE]
> <span data-ttu-id="91de6-234">在 Outlook 中，如果加载项加载于 Outlook.com 或 Gmail 邮箱，此 API 不受支持。</span><span class="sxs-lookup"><span data-stu-id="91de6-234">[!Note In Outlook, this API is not supported if the add-in is loaded in an Outlook.com or Gmail mailbox.]</span></span>

<table><tr><td><span data-ttu-id="91de6-235">Hosts</span><span class="sxs-lookup"><span data-stu-id="91de6-235">Hosts</span></span></td><td><span data-ttu-id="91de6-236">Excel、Outlook、PowerPoint、Word</span><span class="sxs-lookup"><span data-stu-id="91de6-236">Excel, Outlook, PowerPoint, Word</span></span></td></tr>

 <tr><td><span data-ttu-id="91de6-237">要求集</span><span class="sxs-lookup"><span data-stu-id="91de6-237">Requirement sets</span></span></td><td>[<span data-ttu-id="91de6-238">IdentityAPI</span><span class="sxs-lookup"><span data-stu-id="91de6-238">IdentityAPI</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</td></tr></table>

#### <a name="parameters"></a><span data-ttu-id="91de6-239">参数</span><span class="sxs-lookup"><span data-stu-id="91de6-239">Parameters</span></span>

<span data-ttu-id="91de6-p129">`options` - 可选。 接受 `AuthOptions` 对象 （请参阅下方）以定义登录行为。</span><span class="sxs-lookup"><span data-stu-id="91de6-p129">`options` - Optional. Accepts an `AuthOptions` object (see below) to define sign-on behaviors.</span></span>

<span data-ttu-id="91de6-p130">`callback` - 可选。接受回调方法分析令牌，用用户 ID 或用“代表”流进入 Microsoft Graph。如果 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status`  为"成功"，则`AsyncResult.value` 是原始 AAD v. 2.0-被格式化的登陆令牌。</span><span class="sxs-lookup"><span data-stu-id="91de6-p130">`callback` - Optional. Accepts a callback method that can parse the token for the user's ID or use the token in the "on behalf of" flow to get access to Microsoft Graph. If [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult)`.status` is "succeeded", then `AsyncResult.value` is the raw AAD v. 2.0-formatted access token.</span></span>

<span data-ttu-id="91de6-p131">当 Office 得到 AAD v. 2.0 通过 `getAccessTokenAsync` 方法获取的加载项登陆令牌， `AuthOptions` 交互就有了不同的用户体验。</span><span class="sxs-lookup"><span data-stu-id="91de6-p131">The `AuthOptions` interface provides options for the user experience when Office obtains an access token to the add-in from AAD v. 2.0 with the `getAccessTokenAsync` method.</span></span>

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



