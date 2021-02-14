---
title: Outlook 加载项中的身份验证选项
description: Outlook 加载项 根据特定场景提供了多种不同的身份验证方法。
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: b53647c8a38051096ef59e7b784f8a00964aa73f
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234280"
---
# <a name="authentication-options-in-outlook-add-ins"></a><span data-ttu-id="5467f-103">Outlook 加载项中的身份验证选项</span><span class="sxs-lookup"><span data-stu-id="5467f-103">Authentication options in Outlook add-ins</span></span>

<span data-ttu-id="5467f-104">Outlook 加载项可以访问 Internet 上任意位置的信息，无论是托管加载项的服务器、内部网络，还是云中的其他位置。</span><span class="sxs-lookup"><span data-stu-id="5467f-104">Your Outlook add-in can access information from anywhere on the Internet, whether from the server that hosts the add-in, from your internal network, or from somewhere else in the cloud.</span></span> <span data-ttu-id="5467f-105">如果相应信息受保护，加载项需要能够验证用户身份。</span><span class="sxs-lookup"><span data-stu-id="5467f-105">If that information is protected, your add-in needs a way to authenticate your user.</span></span> <span data-ttu-id="5467f-106">Outlook 加载项 根据特定场景提供了多种不同的身份验证方法。</span><span class="sxs-lookup"><span data-stu-id="5467f-106">Outlook add-ins provide a number of different methods to authenticate, depending on your specific scenario.</span></span>

## <a name="single-sign-on-access-token"></a><span data-ttu-id="5467f-107">单一登录访问令牌</span><span class="sxs-lookup"><span data-stu-id="5467f-107">Single sign-on access token</span></span>

<span data-ttu-id="5467f-108">单一登录访问令牌为你的加载项提供了进行身份验证和获取访问令牌以调用 [Microsoft Graph API](/graph/overview) 的无缝方法。</span><span class="sxs-lookup"><span data-stu-id="5467f-108">Single sign-on access tokens provide a seamless way for your add-in to authenticate and obtain access tokens to call the [Microsoft Graph API](/graph/overview).</span></span> <span data-ttu-id="5467f-109">由于不需要用户输入其凭据，此功能可以减少摩擦。</span><span class="sxs-lookup"><span data-stu-id="5467f-109">This capability reduces friction since the user is not required to enter their credentials.</span></span>

> [!NOTE]
> <span data-ttu-id="5467f-110">目前，Word、Excel、Outlook 和 PowerPoint 支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="5467f-110">The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="5467f-111">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="5467f-111">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span>
> <span data-ttu-id="5467f-112">如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="5467f-112">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="5467f-113">若要了解如何这样做，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。</span><span class="sxs-lookup"><span data-stu-id="5467f-113">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

<span data-ttu-id="5467f-114">如果加载项符合以下情况，请考虑使用 SSO 访问令牌：</span><span class="sxs-lookup"><span data-stu-id="5467f-114">Consider using SSO access tokens if your add-in:</span></span>

- <span data-ttu-id="5467f-115">主要由 Microsoft 365 用户使用</span><span class="sxs-lookup"><span data-stu-id="5467f-115">Is used primarily by Microsoft 365 users</span></span>
- <span data-ttu-id="5467f-116">需要访问以下服务：</span><span class="sxs-lookup"><span data-stu-id="5467f-116">Needs access to:</span></span>
  - <span data-ttu-id="5467f-117">作为 Microsoft Graph 的一部分公开的 Microsoft 服务</span><span class="sxs-lookup"><span data-stu-id="5467f-117">Microsoft services that are exposed as part of Microsoft Graph</span></span>
  - <span data-ttu-id="5467f-118">你控制的非 Microsoft 服务</span><span class="sxs-lookup"><span data-stu-id="5467f-118">A non-Microsoft service that you control</span></span>

<span data-ttu-id="5467f-119">SSO 身份验证方法使用 [Azure Active Directory 提供的 OAuth2 代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。</span><span class="sxs-lookup"><span data-stu-id="5467f-119">The SSO authentication method uses the [OAuth2 On-Behalf-Of flow provided by Azure Active Directory](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).</span></span> <span data-ttu-id="5467f-120">它要求加载项在[应用程序注册门户](https://apps.dev.microsoft.com/)中进行注册并在其清单中指定任何所需的 Microsoft Graph 作用域。</span><span class="sxs-lookup"><span data-stu-id="5467f-120">It requires that the add-in register in the [Application Registration Portal](https://apps.dev.microsoft.com/) and specify any required Microsoft Graph scopes in its manifest.</span></span>

<span data-ttu-id="5467f-121">借助此方法，加载项可以获取作用域为你的服务器后端 API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="5467f-121">Using this method, your add-in can obtain an access token scoped to your server back-end API.</span></span> <span data-ttu-id="5467f-122">加载项将此令牌用作 `Authorization` 标头中的持有者令牌，来对 API 回调进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="5467f-122">The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API.</span></span> <span data-ttu-id="5467f-123">此时，服务器可以：</span><span class="sxs-lookup"><span data-stu-id="5467f-123">At that point your server can:</span></span>

- <span data-ttu-id="5467f-124">完成“代表”流来获取作用域为 Microsoft Graph API 的访问令牌</span><span class="sxs-lookup"><span data-stu-id="5467f-124">Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API</span></span>
- <span data-ttu-id="5467f-125">使用令牌中的标识信息创建用户标识并对自己的后端服务进行身份验证</span><span class="sxs-lookup"><span data-stu-id="5467f-125">Use the identity information in the token to establish the user's identity and authenticate to your own back-end services</span></span>

<span data-ttu-id="5467f-126">有关更详细的概述，请参阅 [SSO 身份验证方法的完整概述](../develop/sso-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="5467f-126">For a more detailed overview, see the [full overview of the SSO authentication method](../develop/sso-in-office-add-ins.md).</span></span>

<span data-ttu-id="5467f-127">有关在 Outlook 加载项中使用 SSO 令牌的详细信息，请参阅[在 Outlook 加载项中使用单一登录令牌对用户进行身份验证](authenticate-a-user-with-an-sso-token.md)。</span><span class="sxs-lookup"><span data-stu-id="5467f-127">For details on using the SSO token in an Outlook add-in, see [Authenticate a user with an single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md).</span></span>

<span data-ttu-id="5467f-128">有关使用 SSO 令牌的加载项示例，请参阅 [Outlook 加载项 SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO)。</span><span class="sxs-lookup"><span data-stu-id="5467f-128">For a sample add-in that uses the SSO token, see [Outlook Add-in SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO).</span></span>

## <a name="exchange-user-identity-token"></a><span data-ttu-id="5467f-129">Exchange 用户标识令牌</span><span class="sxs-lookup"><span data-stu-id="5467f-129">Exchange user identity token</span></span>

<span data-ttu-id="5467f-130">Exchange 用户标识令牌为加载项提供了一种创建用户标识的方法。</span><span class="sxs-lookup"><span data-stu-id="5467f-130">Exchange user identity tokens provide a way for your add-in to establish the identity of the user.</span></span> <span data-ttu-id="5467f-131">通过验证用户标识，可以对后端系统执行一次性身份验证，然后接受用户标识令牌，来作为对未来请求的授权。</span><span class="sxs-lookup"><span data-stu-id="5467f-131">By verifying the user's identity, you can then perform a one-time authentication into your back-end system, then accept the user identity token as an authorization for future requests.</span></span> <span data-ttu-id="5467f-132">使用 Exchange 用户标识令牌：</span><span class="sxs-lookup"><span data-stu-id="5467f-132">Use the Exchange user identity token:</span></span>

- <span data-ttu-id="5467f-133">当加载项主要由 Exchange 本地用户使用时。</span><span class="sxs-lookup"><span data-stu-id="5467f-133">When the add-in is used primarily by Exchange on-premises users.</span></span>
- <span data-ttu-id="5467f-134">当加载项需要访问你控制的非 Microsoft 服务时。</span><span class="sxs-lookup"><span data-stu-id="5467f-134">When the add-in needs access to a non-Microsoft service that you control.</span></span>
- <span data-ttu-id="5467f-135">作为回退身份验证（和对 Microsoft Graph 的授权），当加载项在不支持 SSO 的 Office 版本上运行时。</span><span class="sxs-lookup"><span data-stu-id="5467f-135">As a fallback authentication (and authorization to Microsoft Graph) when the add-in is running on a version of Office that doesn't support SSO.</span></span>

<span data-ttu-id="5467f-136">加载项可以调用 [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#getuseridentitytokenasync-callback--usercontext-) 以获取 Exchange 用户标识令牌。</span><span class="sxs-lookup"><span data-stu-id="5467f-136">Your add-in can call [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#getuseridentitytokenasync-callback--usercontext-) to get Exchange user identity tokens.</span></span> <span data-ttu-id="5467f-137">有关使用这些令牌的详细信息，请参阅[使用 Exchange 标识令牌对用户进行身份验证](authenticate-a-user-with-an-identity-token.md)。</span><span class="sxs-lookup"><span data-stu-id="5467f-137">For details on using these tokens, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).</span></span>

## <a name="access-tokens-obtained-via-oauth2-flows"></a><span data-ttu-id="5467f-138">通过 OAuth2 流获取的访问令牌</span><span class="sxs-lookup"><span data-stu-id="5467f-138">Access tokens obtained via OAuth2 flows</span></span>

<span data-ttu-id="5467f-139">加载项也可以访问支持 OAuth2 进行授权的第三方服务。</span><span class="sxs-lookup"><span data-stu-id="5467f-139">Add-ins can also access third-party services that support OAuth2 for authorization.</span></span> <span data-ttu-id="5467f-140">如果你的加载项符合以下情况，请考虑使用 OAuth2 令牌：</span><span class="sxs-lookup"><span data-stu-id="5467f-140">Consider using OAuth2 tokens if your add-in:</span></span>

- <span data-ttu-id="5467f-141">需要访问不受你控制的第三方服务</span><span class="sxs-lookup"><span data-stu-id="5467f-141">Needs access to a third-party service outside of your control</span></span>

<span data-ttu-id="5467f-142">使用此方法，加载项会提示用户通过使用 [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) 方法初始化 OAuth2 流或使用 [office-js-helpers 库](https://github.com/OfficeDev/office-js-helpers) 转到 OAuth2 隐式流来登录到服务。</span><span class="sxs-lookup"><span data-stu-id="5467f-142">Using this method, your add-in prompts the user to sign-in to the service either by using the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) method to initialize the OAuth2 flow, or by using the [office-js-helpers library](https://github.com/OfficeDev/office-js-helpers) to the OAuth2 Implicit flow.</span></span>

## <a name="callback-tokens"></a><span data-ttu-id="5467f-143">回调令牌</span><span class="sxs-lookup"><span data-stu-id="5467f-143">Callback tokens</span></span>

<span data-ttu-id="5467f-144">借助回调令牌，可以使用 [Exchange Web 服务 (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange) 或 [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api) 从服务器后端访问用户邮箱。</span><span class="sxs-lookup"><span data-stu-id="5467f-144">Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api).</span></span> <span data-ttu-id="5467f-145">如果你的加载项符合以下情况，请考虑使用回调令牌：</span><span class="sxs-lookup"><span data-stu-id="5467f-145">Consider using callback tokens if your add-in:</span></span>

- <span data-ttu-id="5467f-146">需要从服务器后端访问用户邮箱。</span><span class="sxs-lookup"><span data-stu-id="5467f-146">Needs access to the user's mailbox from your server back-end.</span></span>

<span data-ttu-id="5467f-147">加载项使用 [getCallbackTokenAsync ](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)方法之一获取回调令牌。</span><span class="sxs-lookup"><span data-stu-id="5467f-147">Add-ins obtain callback tokens using one of the [getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) methods.</span></span> <span data-ttu-id="5467f-148">访问权限级别由加载项清单中指定的权限控制。</span><span class="sxs-lookup"><span data-stu-id="5467f-148">The level of access is controlled by the permissions specified in the add-in manifest.</span></span>
