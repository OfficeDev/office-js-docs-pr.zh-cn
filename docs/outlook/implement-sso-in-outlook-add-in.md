---
title: 应用场景 - 为服务实施单一登录
description: 了解如何使用 Outlook 加载项提供的单一登录令牌和 Exchange 标识令牌为服务实现 SSO。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 44a1ee10af3f49a3738526b0ee7daf6cada3774b
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234231"
---
# <a name="scenario-implement-single-sign-on-to-your-service-in-an-outlook-add-in"></a><span data-ttu-id="7b2ae-103">应用场景：为 Outlook 加载项中的服务实现单一登录</span><span class="sxs-lookup"><span data-stu-id="7b2ae-103">Scenario: Implement single sign-on to your service in an Outlook add-in</span></span>

<span data-ttu-id="7b2ae-104">在本文中，我们将探讨结合使用[单一登录访问令牌](authenticate-a-user-with-an-sso-token.md)和 [Exchange 标识令牌](authenticate-a-user-with-an-identity-token.md)为自己的后端服务提供单一登录实现的建议方法。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-104">In this article we'll explore a recommended method of using the [single sign-on access token](authenticate-a-user-with-an-sso-token.md) and the [Exchange identity token](authenticate-a-user-with-an-identity-token.md) together to provide a single-sign on implementation to your own backend service.</span></span> <span data-ttu-id="7b2ae-105">通过结合使用这两种令牌，可以在 SSO 访问令牌可用时利用其优势，并在其不可用时确保加载项仍能正常工作（例如，当用户切换到不支持这些令牌的客户端时，或当用户的邮箱位于本地 Exchange 服务器时）。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-105">By using both tokens together, you can take advantage of the benefits of the SSO access token when it is available, while ensuring that your add-in will work when it is not, such as when the user switches to a client that does not support them, or if the user's mailbox is on an on-premises Exchange server.</span></span>

<span data-ttu-id="7b2ae-106">有关实现本文中想法的示例外接程序，请参阅[Outlook 外接程序 SSO。](https://github.com/OfficeDev/Outlook-Add-in-SSO)</span><span class="sxs-lookup"><span data-stu-id="7b2ae-106">For a sample add-in that implements the ideas in this article, see [Outlook Add-in SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO).</span></span>


> [!NOTE]
> <span data-ttu-id="7b2ae-107">目前，Word、Excel、Outlook 和 PowerPoint 支持单一登录 API。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-107">The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint.</span></span> <span data-ttu-id="7b2ae-108">若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-108">For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](../reference/requirement-sets/identity-api-requirement-sets.md).</span></span>
> <span data-ttu-id="7b2ae-109">如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-109">If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy.</span></span> <span data-ttu-id="7b2ae-110">若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-110">For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).</span></span>


## <a name="why-use-the-sso-access-token"></a><span data-ttu-id="7b2ae-111">为什么使用 SSO 访问令牌？</span><span class="sxs-lookup"><span data-stu-id="7b2ae-111">Why use the SSO access token?</span></span>

<span data-ttu-id="7b2ae-112">Exchange 标识令牌适用于加载项 API 的所有要求集，因此，仅依赖此令牌并完全忽略 SSO 令牌似乎是更好的做法。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-112">The Exchange identity token is available in all requirement sets of the add-in APIs, so it may be tempting to just rely on this token and ignore the SSO token altogether.</span></span> <span data-ttu-id="7b2ae-113">但是，与 Exchange 标识令牌相比，SSO 令牌具有某些优势，因此，当该令牌可用时会建议使用此方法。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-113">However, the SSO token offers some advantages over the Exchange identity token which make it the recommended method to use when it is available.</span></span>

- <span data-ttu-id="7b2ae-114">SSO 令牌使用标准 OpenID 格式并由 Azure 颁发。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-114">The SSO token uses a standard OpenID format and is issued by Azure.</span></span> <span data-ttu-id="7b2ae-115">这极大地简化了验证这些令牌的过程。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-115">This greatly simplifies the process of validating these tokens.</span></span> <span data-ttu-id="7b2ae-116">与之相比，Exchange 标识令牌使用基于 JSON Web 令牌标准的自定义格式，因此需要使用自定义操作来验证此令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-116">In comparison, Exchange identity tokens use a custom format based on the JSON Web Token standard, requiring custom work to validate the token.</span></span>
- <span data-ttu-id="7b2ae-117">后端可以使用 SSO 令牌来检索 Microsoft Graph 访问令牌，而用户无需执行任何其他登录操作。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-117">The SSO token can be used by your backend to retrieve an access token for Microsoft Graph without the user having to do any additional sign in action.</span></span>
- <span data-ttu-id="7b2ae-118">SSO 令牌提供的标识信息更为丰富，例如用户的显示名称。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-118">The SSO token provides richer identity information, such as the user's display name.</span></span>

## <a name="add-in-scenario"></a><span data-ttu-id="7b2ae-119">加载项应用场景</span><span class="sxs-lookup"><span data-stu-id="7b2ae-119">Add-in scenario</span></span>

<span data-ttu-id="7b2ae-120">鉴于此示例的目的，请考虑使用包含加载项 UI 和脚本 (HTML + JavaScript) 以及加载项调用的后端 Web API 的加载项。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-120">For the purposes of this example, consider an add-in that consists of both the add-in UI and scripts (HTML + JavaScript) and a backend Web API that is called by the add-in.</span></span> <span data-ttu-id="7b2ae-121">后端 Web API 将同时调用 [Microsoft Graph API](/graph/overview) 和 Contoso 数据 API（由第三方创建的虚拟 API）。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-121">The backend Web API makes calls both to the [Microsoft Graph API](/graph/overview) and the Contoso Data API, a fictional API created by a third party.</span></span> <span data-ttu-id="7b2ae-122">与 Microsoft Graph API 类似，Contoso 数据 API 也需要进行 OAuth 身份验证。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-122">Like the Microsoft Graph API, the Contoso Data API requires OAuth authentication.</span></span> <span data-ttu-id="7b2ae-123">要求是，后端 Web API 应能够同时调用这两个 API，而无需在每次访问令牌过期时提示用户提供凭据。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-123">The requirement is that the backend Web API should be able to call both APIs without having to prompt the user for their credentials every time an access token expires.</span></span>

<span data-ttu-id="7b2ae-124">为了实现这一目的，后端 API 创建了一个安全的用户数据库。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-124">To do this, the backend API creates a secure database of users.</span></span> <span data-ttu-id="7b2ae-125">每个用户都将在该数据库中获得一个条目，后端可以在其中存储 Microsoft Graph API 和 Contoso 数据 API 的长期刷新令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-125">Each user will get an entry in the database where the backend can store long-lived refresh tokens for both the Microsoft Graph API and the Contoso Data API.</span></span> <span data-ttu-id="7b2ae-126">以下 JSON 标记表示用户在数据库中的条目。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-126">The following JSON markup represents a user's entry in the database.</span></span>

```JSON
{
  "userDisplayName": "...",
  "ssoId": "...",
  "exchangeId": "...",
  "graphRefreshToken": "...",
  "contosoRefreshToken": "..."
}
```

<span data-ttu-id="7b2ae-127">加载项会在对后端 Web API 的每个调用中包含 SSO 访问令牌（如果可用）或 Exchange 标识令牌（如果 SSO 令牌不可用）。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-127">The add-in includes either the SSO access token (if it is available) or the Exchange identity token (if the SSO token is not available) with every call it makes to the backend Web API.</span></span>

### <a name="add-in-startup"></a><span data-ttu-id="7b2ae-128">加载项启动</span><span class="sxs-lookup"><span data-stu-id="7b2ae-128">Add-in startup</span></span>

1. <span data-ttu-id="7b2ae-129">当加载项启动时，它向后端 Web API 发送请求，以确定用户是否已注册（即在用户数据库中是否有相关联的记录）以及 API 是否同时具有 Graph 和 Contoso 的刷新令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-129">When the add-in starts, it sends a request to the backend Web API to determine if the user is registered (i.e. has an associated record in the user database) and that the API has refresh tokens for both Graph and Contoso.</span></span> <span data-ttu-id="7b2ae-130">在此调用中，加载项同时包含 SSO 令牌（如果可用）和标识令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-130">In this call, the add-in includes both the SSO token (if available) and the identity token.</span></span>

1. <span data-ttu-id="7b2ae-131">Web API 使用[使用 Outlook 加载项中的单一登录令牌对用户进行身份验证](authenticate-a-user-with-an-sso-token.md)和[使用 Exchange 标识令牌对用户进行身份验证](authenticate-a-user-with-an-identity-token.md)中的方法进行验证并从这两种令牌中生成唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-131">The Web API uses the methods in [Authenticate a user with an single-sign-on token in an Outlook add-in](authenticate-a-user-with-an-sso-token.md) and [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md) to validate and generate a unique identifier from both tokens.</span></span>

1. <span data-ttu-id="7b2ae-132">如果提供了 SSO 令牌，则 Web API 会查询用户数据库中是否存在具有 `ssoId` 值（该值与从 SSO 令牌生成的唯一标识符相匹配）的条目。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-132">If an SSO token was provided, the Web API then queries the user database for an entry that has an `ssoId` value that matches the unique identifier generated from the SSO token.</span></span>
   - <span data-ttu-id="7b2ae-133">如果条目不存在，则继续执行下一步。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-133">If an entry did not exist, continue to the next step.</span></span>
   - <span data-ttu-id="7b2ae-134">如果条目存在，则继续执行步骤 5。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-134">If an entry exists, proceed to step 5.</span></span>

1. <span data-ttu-id="7b2ae-135">Web API 查询数据库中是否存在具有 `exchangeId` 值（该值与从 Exchange 标识令牌生成的唯一标识符相匹配）的条目。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-135">The Web API queries the database for an entry that has an `exchangeId` value that matches the unique identifier generated from the Exchange identity token.</span></span>
   - <span data-ttu-id="7b2ae-136">如果该条目存在且 SSO 令牌可用，则更新该数据库中的用户记录，以将 `ssoId` 值设置为从 SSO 令牌生成的唯一标识符，并继续执行步骤 5。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-136">If an entry exists and an SSO token was provided, update the user's record in the database to set the `ssoId` value to the unique identifier generated from the SSO token and proceed to step 5.</span></span>
   - <span data-ttu-id="7b2ae-137">如果该条目存在但 SSO 令牌不可用，则继续执行步骤 5。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-137">If an entry exists and no SSO token was provided, proceed to step 5.</span></span>
   - <span data-ttu-id="7b2ae-138">如果该条目不存在，则新建一个条目。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-138">If no entry exists, create a new entry.</span></span> <span data-ttu-id="7b2ae-139">将 `ssoId` 设置为从 SSO 令牌生成的唯一标识符（如果可用），并将 `exchangeId` 设置为从 Exchange 标识令牌生成的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-139">Set `ssoId` to the unique identifier generated from the SSO token (if available), and set `exchangeId` to the unique identifier generated from the Exchange identity token.</span></span>

1. <span data-ttu-id="7b2ae-140">检查用户的 `graphRefreshToken` 值中是否存在有效的刷新令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-140">Check for a valid refresh token in the user's `graphRefreshToken` value.</span></span>
   - <span data-ttu-id="7b2ae-141">如果此值缺失或无效且 SSO 令牌可用，则使用 [OAuth2代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)获取 Graph 的访问令牌和刷新令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-141">If the value is missing or invalid and an SSO token was provided, use the [OAuth2 On-Behalf-Of flow](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of) to obtain an access token and refresh token for Graph.</span></span> <span data-ttu-id="7b2ae-142">将刷新令牌保存在用户的 `graphRefreshToken` 值中。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-142">Save the refresh token in the `graphRefreshToken` value for the user.</span></span>

1. <span data-ttu-id="7b2ae-143">检查 `graphRefreshToken` 和 `contosoRefreshToken` 中是否存在有效的刷新令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-143">Check for valid refresh tokens in both `graphRefreshToken` and `contosoRefreshToken`.</span></span>
   - <span data-ttu-id="7b2ae-144">如果两个值均有效，则对加载项做出响应，来指示用户已注册且已进行了配置。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-144">If both values are valid, respond to the add-in to indicate that the user is already registered and configured.</span></span>
   - <span data-ttu-id="7b2ae-145">如果任一值无效，则对加载项做出响应，来指示需要进行用户设置，并指示需要配置的服务（Graph 和 Contoso）。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-145">If either value is invalid, respond to the add-in to indicate that user setup is required, along with which services (Graph or Contoso) need to be configured.</span></span>

1. <span data-ttu-id="7b2ae-146">加载项将检查响应。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-146">The add-in checks the response.</span></span>
   - <span data-ttu-id="7b2ae-147">如果用户已注册并已进行了配置，则加载项将继续正常运行。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-147">If the user is already registered and configured, the add-in continues with normal operation.</span></span>
   - <span data-ttu-id="7b2ae-148">如需进行用户设置，则加载项进入“设置”模式并提示用户向加载项授权。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-148">If user setup is required, the add-in enters "setup" mode and prompts the user to authorize the add-in.</span></span>

### <a name="authorize-the-backend-web-api"></a><span data-ttu-id="7b2ae-149">授权后端 Web API</span><span class="sxs-lookup"><span data-stu-id="7b2ae-149">Authorize the backend Web API</span></span>

<span data-ttu-id="7b2ae-150">理想情况下，授权后端 Web API 调用 Microsoft Graph API 和 Contoso 数据 API 这一过程应仅进行一次，以尽量减少提示用户进行登录的次数。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-150">The procedure for authorizing the backend Web API to call the Microsoft Graph API and Contoso Data API should ideally only have to happen once, to minimize having to prompt the user for sign-in.</span></span>

<span data-ttu-id="7b2ae-151">基于后端 Web API 的响应，加载项可能需要授权用户使用 Microsoft Graph API 和/或 Contoso 数据 API。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-151">Based on the response from the backend Web API, the add-in may need to authorize the user for the Microsoft Graph API, the Contoso Data API, or both.</span></span> <span data-ttu-id="7b2ae-152">因为这两种 API 都使用 OAuth2 身份验证，所以它们的授权方法类似。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-152">Since both APIs use OAuth2 authentication, the method is similar for both.</span></span>

1. <span data-ttu-id="7b2ae-153">加载项通知用户需要授权其使用 API 并让用户单击一个链接或按钮来启动这一过程。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-153">The add-in notifies the user that it needs them to authorize their use of the API and asks them to click a link or button to start the process.</span></span>

    > [!NOTE]
    > <span data-ttu-id="7b2ae-154">Outlook 外接程序 [SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO) 中的示例外接程序演示如何使用 [对话框 API](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) 和 [office-js-helpers](https://github.com/OfficeDev/office-js-helpers) 库作为启动 API 的 [OAuth2 授权](/azure/active-directory/develop/active-directory-protocols-oauth-code) 代码流的选项。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-154">The example add-in at [Outlook Add-in SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO) shows how to use the [Dialog API](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) and the [office-js-helpers library](https://github.com/OfficeDev/office-js-helpers) as options to start the [OAuth2 Authorization Code flow](/azure/active-directory/develop/active-directory-protocols-oauth-code) for the API.</span></span>

1. <span data-ttu-id="7b2ae-155">此流完成后，加载项向后端 Web API 发送刷新令牌并包含 SSO 令牌（如果可用）或 Exchange 标识令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-155">Once the flow completes, the add-in sends the refresh token to the backend Web API and includes the SSO token (if available) or the Exchange identity token.</span></span>

1. <span data-ttu-id="7b2ae-156">后端 Web API 在数据库中查找用户并更新相应的刷新令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-156">The backend Web API locates the user in the database and updates the appropriate refresh token.</span></span>

1. <span data-ttu-id="7b2ae-157">加载项将继续正常运行。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-157">The add-in continues with normal operation.</span></span>

### <a name="normal-operation"></a><span data-ttu-id="7b2ae-158">正常运行</span><span class="sxs-lookup"><span data-stu-id="7b2ae-158">Normal operation</span></span>

<span data-ttu-id="7b2ae-159">每当加载项调用后端 Web API 时，它都将包含 SSO 令牌或 Exchange 标识令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-159">Whenever the add-in calls the backend Web API, it includes either the SSO token or the Exchange identity token.</span></span> <span data-ttu-id="7b2ae-160">后端 Web API 根据此令牌查找用户，然后使用存储的刷新令牌来获取 Microsoft Graph API 和 Contoso 数据 API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-160">The backend Web API locates the user by this token, then uses the stored refresh tokens to obtain access tokens for the Microsoft Graph API and the Contoso Data API.</span></span> <span data-ttu-id="7b2ae-161">只要刷新令牌有效，用户就无需再次登录。</span><span class="sxs-lookup"><span data-stu-id="7b2ae-161">As long as the refresh tokens are valid, the user will not have to sign in again.</span></span>
