---
title: 使用单一登录令牌对用户进行身份验证
description: 了解如何使用 Outlook 外接程序提供的单一登录令牌为服务实现 SSO。
ms.date: 11/19/2019
localization_priority: Normal
ms.openlocfilehash: 9ee3ece5929df602a35ddd9883c08e25164d8a22
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721028"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in-preview"></a><span data-ttu-id="9ebc2-103">在 Outlook 加载项中使用单一登录令牌验证用户（预览版）</span><span class="sxs-lookup"><span data-stu-id="9ebc2-103">Authenticate a user with a single-sign-on token in an Outlook add-in (preview)</span></span>

<span data-ttu-id="9ebc2-104">使用单一登录 (SSO)，加载项可以无缝方式验证用户（并根据需要获取访问令牌来调用 [Microsoft Graph API](/graph/overview)）。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-104">Single sign-on (SSO) provides a seamless way for your add-in to authenticate users (and optionally to obtain access tokens to call the [Microsoft Graph API](/graph/overview)).</span></span>

<span data-ttu-id="9ebc2-105">借助此方法，加载项可以获取范围限定为服务器后端 API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-105">Using this method, your add-in can obtain an access token scoped to your server back-end API.</span></span> <span data-ttu-id="9ebc2-106">加载项将此令牌用作 `Authorization` 头中的持有者令牌，以验证 API 回调。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-106">The add-in uses this as a bearer token in the `Authorization` header to authenticate a call back to your API.</span></span> <span data-ttu-id="9ebc2-107">也可以使用服务器端代码执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="9ebc2-107">Optionally, you can also have your server-side code:</span></span>

- <span data-ttu-id="9ebc2-108">完成“代表”流，以获取范围限定为 Microsoft Graph API 的访问令牌</span><span class="sxs-lookup"><span data-stu-id="9ebc2-108">Complete the On-Behalf-Of flow to obtain an access token scoped to the Microsoft Graph API</span></span>
- <span data-ttu-id="9ebc2-109">使用令牌中的标识信息，以创建用户标识并验证自己的后端服务</span><span class="sxs-lookup"><span data-stu-id="9ebc2-109">Use the identity information in the token to establish the user's identity and authenticate to your own back-end services</span></span>

<span data-ttu-id="9ebc2-110">有关 Office 外接程序中的 SSO 的概述，请参阅[为 Office 外接程序启用单一登录](../develop/sso-in-office-add-ins.md)和[在 Office 外接程序中授予对 Microsoft Graph 的访问权限](../develop/authorize-to-microsoft-graph.md)。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-110">For an overview of SSO in Office Add-ins, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) and [Authorize to Microsoft Graph in your Office Add-in](../develop/authorize-to-microsoft-graph.md).</span></span>

> [!NOTE]
> <span data-ttu-id="9ebc2-111">若要使用 SSO，必须从外接程序的启动 HTML 页面中的 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 加载 Office JavaScript 库的 Beta 版。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-111">To use SSO, you must load the beta version of the Office JavaScript Library from https://appsforoffice.microsoft.com/lib/beta/hosted/office.js in the startup HTML page of the add-in.</span></span>

## <a name="enable-modern-authentication-in-your-office-365-tenancy"></a><span data-ttu-id="9ebc2-112">在 Office 365 租赁中启用新式验证</span><span class="sxs-lookup"><span data-stu-id="9ebc2-112">Enable modern authentication in your Office 365 tenancy</span></span>

<span data-ttu-id="9ebc2-113">必须为 Office 365 租赁启用新式验证，才能将 SSO 与 Outlook 外接程序结合使用。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-113">To use SSO with an Outlook add-in, you must enable Modern Authentication for the Office 365 tenancy.</span></span> <span data-ttu-id="9ebc2-114">若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-114">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>

## <a name="register-your-add-in"></a><span data-ttu-id="9ebc2-115">注册外接程序</span><span class="sxs-lookup"><span data-stu-id="9ebc2-115">Register your add-in</span></span>

<span data-ttu-id="9ebc2-116">若要使用 SSO，Outlook 外接程序需要有已向 Azure Active Directory (AAD) v2.0 注册的服务器端 Web API。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-116">To use SSO, your Outlook add-in will need to have a server-side web API that is registered with Azure Active Directory (AAD) v2.0.</span></span> <span data-ttu-id="9ebc2-117">有关详细信息，请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 外接程序](../develop/register-sso-add-in-aad-v2.md)。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-117">For more information, see [Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint](../develop/register-sso-add-in-aad-v2.md).</span></span>

### <a name="provide-consent-when-sideloading-an-add-in"></a><span data-ttu-id="9ebc2-118">旁加载加载项时授予许可</span><span class="sxs-lookup"><span data-stu-id="9ebc2-118">Provide consent when sideloading an add-in</span></span>

<span data-ttu-id="9ebc2-119">从 AppSource 获取使用 SSO 的加载项时，应用商店 UI 将负责提示用户同意授予所请求的 Graph 权限。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-119">When an add-in that uses SSO is acquired from AppSource, the store UI handles prompting the user for consent to the requested Graph permissions.</span></span> <span data-ttu-id="9ebc2-120">但是，在开发加载项时，需要提前提供授权。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-120">However, when you are developing an add-in, you have to provide consent in advance.</span></span> <span data-ttu-id="9ebc2-121">有关详细信息，请参阅[向加载项授予管理员许可](../develop/grant-admin-consent-to-an-add-in.md)</span><span class="sxs-lookup"><span data-stu-id="9ebc2-121">For more information, see [Grant administrator consent to the add-in](../develop/grant-admin-consent-to-an-add-in.md)</span></span>

## <a name="update-the-add-in-manifest"></a><span data-ttu-id="9ebc2-122">更新加载项清单</span><span class="sxs-lookup"><span data-stu-id="9ebc2-122">Update the add-in manifest</span></span>

<span data-ttu-id="9ebc2-123">若要在加载项中启用 SSO，下一步在 `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md) 元素末尾添加 `WebApplicationInfo` 元素。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-123">The next step to enable SSO in the add-in is to add a `WebApplicationInfo` element at the end of the `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md) element.</span></span> <span data-ttu-id="9ebc2-124">有关详细信息，请参阅[配置加载项](../develop/sso-in-office-add-ins.md#configure-the-add-in)。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-124">For more information, see [Configure the add-in](../develop/sso-in-office-add-ins.md#configure-the-add-in).</span></span>

## <a name="get-the-sso-token"></a><span data-ttu-id="9ebc2-125">获取 SSO 令牌</span><span class="sxs-lookup"><span data-stu-id="9ebc2-125">Get the SSO token</span></span>

<span data-ttu-id="9ebc2-126">加载项使用客户端脚本获取 SSO 令牌。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-126">The add-in gets an SSO token with client-side script.</span></span> <span data-ttu-id="9ebc2-127">有关详细信息，请参阅[添加客户端代码](../develop/sso-in-office-add-ins.md#add-client-side-code)。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-127">For more information, see [Add client-side code](../develop/sso-in-office-add-ins.md#add-client-side-code).</span></span>

## <a name="use-the-sso-token-at-the-back-end"></a><span data-ttu-id="9ebc2-128">在后端使用 SSO 令牌</span><span class="sxs-lookup"><span data-stu-id="9ebc2-128">Use the SSO token at the back-end</span></span>

<span data-ttu-id="9ebc2-129">大多数情况下，如果加载项没有将访问令牌传递到服务器端并在其中使用它，那么获取访问令牌的意义就不大。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-129">In most scenarios, there would be little point to obtaining the access token, if your add-in does not pass it on to a server-side and use it there.</span></span> <span data-ttu-id="9ebc2-130">若要详细了解服务器端可以和应该执行的操作，请参阅[添加服务器端代码](../develop/sso-in-office-add-ins.md#add-server-side-code)。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-130">For details on what your server-side could and should do, see [Add server-side code](../develop/sso-in-office-add-ins.md#add-server-side-code).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9ebc2-131">若要将 SSO 令牌用作 *Outlook* 加载项中的标识，建议还[使用 Exchange 标识令牌](authenticate-a-user-with-an-identity-token.md)作为备用标识。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-131">When using the SSO token as an identity in an *Outlook* add-in, we recommend that you also [use the Exchange identity token](authenticate-a-user-with-an-identity-token.md) as an alternate identity.</span></span> <span data-ttu-id="9ebc2-132">加载项用户可能使用多个客户端，而有些客户端可能不支持提供 SSO 令牌。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-132">Users of your add-in may use multiple clients, and some may not support providing an SSO token.</span></span> <span data-ttu-id="9ebc2-133">通过将 Exchange 标识令牌用作备用令牌，就不用多次提示这些用户输入凭据了。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-133">By using the Exchange identity token as an alternate, you can avoid having to prompt these users for credentials multiple times.</span></span> <span data-ttu-id="9ebc2-134">有关详细信息，请参阅[应用场景：在 Outlook 外接程序中对服务实现单一登录](implement-sso-in-outlook-add-in.md)。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-134">For more information, see [Scenario: Implement single sign-on to your service in an Outlook add-in](implement-sso-in-outlook-add-in.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="9ebc2-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9ebc2-135">See also</span></span>

- <span data-ttu-id="9ebc2-136">有关使用 SSO 令牌访问 Microsoft Graph API 的 Outlook 外接程序示例，请参阅 [AttachmentsDemo 示例外接程序](https://github.com/OfficeDev/outlook-add-in-attachments-demo)。</span><span class="sxs-lookup"><span data-stu-id="9ebc2-136">For a sample Outlook add-in that uses the SSO token to access the Microsoft Graph API, see [AttachmentsDemo Sample Add-in](https://github.com/OfficeDev/outlook-add-in-attachments-demo).</span></span>
- [<span data-ttu-id="9ebc2-137">SSO API 参考</span><span class="sxs-lookup"><span data-stu-id="9ebc2-137">SSO API reference</span></span>](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [<span data-ttu-id="9ebc2-138">IdentityAPI 要求集</span><span class="sxs-lookup"><span data-stu-id="9ebc2-138">IdentityAPI requirement set</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)
