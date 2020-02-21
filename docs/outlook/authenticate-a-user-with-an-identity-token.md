---
title: 使用加载项中的标识令牌对用户进行身份验证
description: 了解如何使用 Outlook 加载项提供的标识令牌对服务实施 SSO。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4134aa8ff21262f2f384d141db002b56a4a32f0a
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165954"
---
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a><span data-ttu-id="6b0e1-103">使用 Exchange 的标识令牌对用户进行身份验证</span><span class="sxs-lookup"><span data-stu-id="6b0e1-103">Authenticate a user with an identity token for Exchange</span></span>

<span data-ttu-id="6b0e1-104">Exchange 用户标识令牌为加载项提供了一种以唯一的方式标识加载项用户的方法。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-104">Exchange user identity tokens provide a way for your add-in to uniquely identify an add-in user.</span></span> <span data-ttu-id="6b0e1-105">通过创建用户标识，可以为后端服务实现单一登录 (SSO) 身份验证方案，此方案使使用 Outlook 加载项的客户无需登录即可连接服务。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-105">By establishing the user's identity, you can implement a single sign-on (SSO) authentication scheme for your back-end service that enables customers who are using Outlook add-ins to connect to your service without logging in.</span></span> <span data-ttu-id="6b0e1-106">有关何时使用此令牌类型的更多详细信息，请参阅 [Exchange 用户标识令牌](authentication.md#exchange-user-identity-token)。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-106">See [Exchange user identity token](authentication.md#exchange-user-identity-token) for more about when to use this token type.</span></span> <span data-ttu-id="6b0e1-107">本文介绍了使用 Exchange 标识令牌对访问后端的用户进行身份验证的简单方法。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-107">In this article, we'll take a look at a simplistic method of using the Exchange identity token to authenticate a user to your back-end.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6b0e1-108">这只是 SSO 实现的一个简单示例。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-108">This is just a simple example of an SSO implementation.</span></span> <span data-ttu-id="6b0e1-109">和以往一样，在处理标识和身份验证事宜时，一定要确保代码符合组织的安全要求。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-109">As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.</span></span>

## <a name="send-the-id-token-with-each-request"></a><span data-ttu-id="6b0e1-110">通过每个请求发送 ID 令牌</span><span class="sxs-lookup"><span data-stu-id="6b0e1-110">Send the ID token with each request</span></span>

<span data-ttu-id="6b0e1-111">第一步是通过调用 [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 使加载项获取服务器中的 Exchange 用户标识令牌。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-111">The first step is for your add-in to obtain the Exchange user identity token from the server by calling [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span></span> <span data-ttu-id="6b0e1-112">然后加载项通过向后端发出的每个请求发送该令牌。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-112">Then the add-in sends this token with every request it makes to your back-end.</span></span> <span data-ttu-id="6b0e1-113">它可能是在标头中，或在请求正文中。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-113">This could be in a header, or as part of the request body.</span></span>

## <a name="validate-the-token"></a><span data-ttu-id="6b0e1-114">验证令牌</span><span class="sxs-lookup"><span data-stu-id="6b0e1-114">Validate the token</span></span>

<span data-ttu-id="6b0e1-115">后端必须在接受令牌之前对其进行验证。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-115">The back-end MUST validate the token before accepting it.</span></span> <span data-ttu-id="6b0e1-116">这是确保令牌是由用户的 Exchange 服务器颁发的重要步骤。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-116">This is an important step to ensure that the token was issued by the user's Exchange server.</span></span> <span data-ttu-id="6b0e1-117">有关验证 Exchange 用户标识令牌的信息，请参阅[验证 Exchange 标识令牌](validate-an-identity-token.md)。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-117">For information on validating Exchange user identity tokens, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span>

<span data-ttu-id="6b0e1-118">验证并解码之后，令牌的有效负载如下所示。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-118">Once validated and decoded, the payload of the token looks something like the following.</span></span>

```json
{ 
    "aud" : "https://mailhost.contoso.com/IdentityTest.html",
    "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com",
    "nbf" : "1505749527",
    "exp" : "1505778327",
    "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
    "isbrowserhostedapp":"true",
    "appctx" : {
        "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389",
        "version" : "ExIdTok.V1",
        "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
    }
}
```

## <a name="map-the-token-to-a-user-in-your-backend"></a><span data-ttu-id="6b0e1-119">将令牌映射到后端的用户</span><span class="sxs-lookup"><span data-stu-id="6b0e1-119">Map the token to a user in your backend</span></span>

<span data-ttu-id="6b0e1-120">后端服务可以根据令牌计算唯一的用户 ID 并将其映射到内部用户系统中的用户。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-120">Your back-end service can calculate a unique user ID from the token and map it to a user in your internal user system.</span></span> <span data-ttu-id="6b0e1-121">例如，如果使用数据库来存储用户，可以在数据库中将此唯一 ID 添加到用户的记录中。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-121">For example, if you use a database to store users, you could add this unique ID to the user's record in your database.</span></span>

### <a name="generate-a-unique-id"></a><span data-ttu-id="6b0e1-122">生成唯一 ID</span><span class="sxs-lookup"><span data-stu-id="6b0e1-122">Generate a unique ID</span></span>

<span data-ttu-id="6b0e1-123">建议结合使用 `msexchuid` 和 `amurl` 属性。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-123">We recommend that you use a combination of the `msexchuid` and `amurl` properties.</span></span> <span data-ttu-id="6b0e1-124">例如，可以将两个值连接在一起，生成 Base 64 编码的字符串。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-124">For example, you could concatenate the two values together and generate a base 64-encoded string.</span></span> <span data-ttu-id="6b0e1-125">每次均可通过令牌生成此值，因此你可以将 Exchange 用户标识令牌映射回系统中的用户。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-125">This value can be reliably generated from the token every time, so you can map an Exchange user identity token back to the user in your system.</span></span>

### <a name="check-the-user"></a><span data-ttu-id="6b0e1-126">检查用户</span><span class="sxs-lookup"><span data-stu-id="6b0e1-126">Check the user</span></span>

<span data-ttu-id="6b0e1-127">生成唯一 ID 后，下一步就是查找系统中使用此关联 ID 的用户。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-127">With the unique ID generated, the next step is to check for a user in your system with that associated ID.</span></span>

- <span data-ttu-id="6b0e1-128">如果能找到该用户，后端会将请求视为已经过身份验证并允许请求继续。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-128">If the user is found, the back-end treats the request as authenticated, and allows the request to proceed.</span></span>

- <span data-ttu-id="6b0e1-129">如果找不到用户，后端将返回一个错误，指示用户需要登录。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-129">If the user is not found, then the back-end returns an error indicating that the user needs to sign in.</span></span> <span data-ttu-id="6b0e1-130">然后加载项会提示用户使用现有的身份验证方法登录到后端。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-130">The add-in then prompts the user to sign in to the back-end using your existing authentication method.</span></span> <span data-ttu-id="6b0e1-131">一旦用户经过身份验证，将提交含用户身份验证详细信息的 Exchange 用户标识令牌。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-131">Once the user is authenticated, the Exchange user identity token is submitted with the user authentication details.</span></span> <span data-ttu-id="6b0e1-132">然后后端可以使用唯一 ID 更新系统中的用户记录。</span><span class="sxs-lookup"><span data-stu-id="6b0e1-132">The back-end can then update the user's record in your system with the unique ID.</span></span>
