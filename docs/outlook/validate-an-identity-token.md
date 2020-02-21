---
title: 验证 Outlook 加载项标识令牌
description: Outlook 加载项可以向你发送 Exchange 用户标识令牌，但是在你信任此请求之前，必须验证该令牌以确保它来自预期的 Exchange 服务器。
ms.date: 11/07/2019
localization_priority: Normal
ms.openlocfilehash: b412756a980d54a20a1c8deab43cd7634c0188cb
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165986"
---
# <a name="validate-an-exchange-identity-token"></a><span data-ttu-id="51722-103">验证 Exchange 标识令牌</span><span class="sxs-lookup"><span data-stu-id="51722-103">Validate an Exchange identity token</span></span>

<span data-ttu-id="51722-104">Outlook 加载项可以向你发送 Exchange 用户标识令牌，但是在你信任此请求之前，必须验证该令牌以确保它来自预期的 Exchange 服务器。</span><span class="sxs-lookup"><span data-stu-id="51722-104">Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect.</span></span> <span data-ttu-id="51722-105">Exchange 用户标识令牌均为 JSON Web 令牌 (JWT)。</span><span class="sxs-lookup"><span data-stu-id="51722-105">Exchange user identity tokens are JSON Web Tokens (JWT).</span></span> <span data-ttu-id="51722-106">[RFC 7519 JSON Web 令牌 (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt) 中介绍了验证 JWT 所需的步骤。</span><span class="sxs-lookup"><span data-stu-id="51722-106">The steps required to validate a JWT are described in [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

<span data-ttu-id="51722-107">建议使用四个步骤验证标识令牌并获取用户的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="51722-107">We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier.</span></span> <span data-ttu-id="51722-108">首先，从 base64 URL 编码的字符串中提取 JSON Web 令牌 (JWT)。</span><span class="sxs-lookup"><span data-stu-id="51722-108">First, extract the JSON Web Token (JWT) from a base64 URL-encoded string.</span></span> <span data-ttu-id="51722-109">然后，确保该令牌格式正确、它是用于 Outlook 外接程序的令牌、它未过期且你可以提取身份验证元数据文档的有效 URL。</span><span class="sxs-lookup"><span data-stu-id="51722-109">Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document.</span></span> <span data-ttu-id="51722-110">接下来，从 Exchange 服务器中检索身份验证元数据文档并验证附加到标识令牌的签名。</span><span class="sxs-lookup"><span data-stu-id="51722-110">Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token.</span></span> <span data-ttu-id="51722-111">最后，通过将用户的 Exchange ID 与身份验证元数据文档的 URL 连接，来计算用户的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="51722-111">Finally, compute a unique identifier for the user by concatenating the user's Exchange ID with the URL of the authentication metadata document.</span></span>

## <a name="extract-the-json-web-token"></a><span data-ttu-id="51722-112">提取 JSON Web 令牌</span><span class="sxs-lookup"><span data-stu-id="51722-112">Extract the JSON Web Token</span></span>

<span data-ttu-id="51722-113">[getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 返回的令牌是令牌的编码字符串表示形式。</span><span class="sxs-lookup"><span data-stu-id="51722-113">The token returned from [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) is an encoded string representation of the token.</span></span> <span data-ttu-id="51722-114">在此表示形式下，根据 RFC 7519，所有 JWT 都有三个部分，以句点分隔。</span><span class="sxs-lookup"><span data-stu-id="51722-114">In this form, per RFC 7519, all JWTs have three parts, separated by a period.</span></span> <span data-ttu-id="51722-115">格式如下所示。</span><span class="sxs-lookup"><span data-stu-id="51722-115">The format is as follows.</span></span>

```json
{header}.{payload}.{signature}
```

<span data-ttu-id="51722-116">标头和有效负载应进行 Base64 解码，以获取每一部分的 JSON 表示形式。</span><span class="sxs-lookup"><span data-stu-id="51722-116">The header and payload should be base64-decoded to obtain a JSON representation of each part.</span></span> <span data-ttu-id="51722-117">签名应进行 base64 解码，以获取包含二进制签名的字节数组。</span><span class="sxs-lookup"><span data-stu-id="51722-117">The signature should be base64-decoded to obtain a byte array containing the binary signature.</span></span>

<span data-ttu-id="51722-118">有关令牌内容的详细信息，请参阅 [Exchange 标识令牌揭秘](inside-the-identity-token.md)。</span><span class="sxs-lookup"><span data-stu-id="51722-118">For more information about the contents of the token, see [Inside the Exchange identity token](inside-the-identity-token.md).</span></span>

<span data-ttu-id="51722-119">三个组件都解码后，可以继续验证该令牌的内容。</span><span class="sxs-lookup"><span data-stu-id="51722-119">After you have the three decoded components, you can proceed with validating the content of the token.</span></span>

## <a name="validate-token-contents"></a><span data-ttu-id="51722-120">验证令牌内容</span><span class="sxs-lookup"><span data-stu-id="51722-120">Validate token contents</span></span>

<span data-ttu-id="51722-121">若要验证令牌内容，还应检查以下项目。</span><span class="sxs-lookup"><span data-stu-id="51722-121">To validate the token contents, you should check the following.</span></span>

- <span data-ttu-id="51722-122">检查标头并验证：</span><span class="sxs-lookup"><span data-stu-id="51722-122">Check the header and verify that the:</span></span>
    - <span data-ttu-id="51722-123">`typ`"声明" 设置`JWT`为。</span><span class="sxs-lookup"><span data-stu-id="51722-123">`typ` claim is set to `JWT`.</span></span>
    - <span data-ttu-id="51722-124">`alg`"声明" 设置`RS256`为。</span><span class="sxs-lookup"><span data-stu-id="51722-124">`alg` claim is set to `RS256`.</span></span>
    - <span data-ttu-id="51722-125">`x5t`声明存在。</span><span class="sxs-lookup"><span data-stu-id="51722-125">`x5t` claim is present.</span></span>

- <span data-ttu-id="51722-126">检查有效负载并验证：</span><span class="sxs-lookup"><span data-stu-id="51722-126">Check the payload and verify that the:</span></span>
    - <span data-ttu-id="51722-127">`amurl`中的`appctx`声明已设置为授权令牌签名密钥清单文件的位置。</span><span class="sxs-lookup"><span data-stu-id="51722-127">`amurl` claim inside the `appctx` is set to the location of an authorized token signing key manifest file.</span></span> <span data-ttu-id="51722-128">例如，Office 365 的`amurl`预期值为https://outlook.office365.com:443/autodiscover/metadata/json/1。</span><span class="sxs-lookup"><span data-stu-id="51722-128">For example, the expected `amurl` value for Office 365 is https://outlook.office365.com:443/autodiscover/metadata/json/1.</span></span> <span data-ttu-id="51722-129">有关详细信息，请参阅下一节[验证域](#verify-the-domain)。</span><span class="sxs-lookup"><span data-stu-id="51722-129">See the next section [Verify the domain](#verify-the-domain) for additional information.</span></span>
    - <span data-ttu-id="51722-130">当前时间介于`nbf`和`exp`声明中指定的时间之间。</span><span class="sxs-lookup"><span data-stu-id="51722-130">Current time is between the times specified in the `nbf` and `exp` claims.</span></span> <span data-ttu-id="51722-131">`nbf` 声明指定了令牌被视为有效的最早时间，而 `exp` 声明指定了令牌的失效时间。</span><span class="sxs-lookup"><span data-stu-id="51722-131">The `nbf` claim specifies the earliest time that the token is considered valid, and the `exp` claim specifies the expiration time for the token.</span></span> <span data-ttu-id="51722-132">建议将服务器之间的时钟设置差异考虑在内。</span><span class="sxs-lookup"><span data-stu-id="51722-132">It is recommended to allow for some variation in clock settings between servers.</span></span>
    - <span data-ttu-id="51722-133">`aud`声明是你的外接程序的预期 URL。</span><span class="sxs-lookup"><span data-stu-id="51722-133">`aud` claim is the expected URL for your add-in.</span></span>
    - <span data-ttu-id="51722-134">`version`声明内的`appctx`声明已设置为`ExIdTok.V1`。</span><span class="sxs-lookup"><span data-stu-id="51722-134">`version` claim inside the `appctx` claim is set to `ExIdTok.V1`.</span></span>

### <a name="verify-the-domain"></a><span data-ttu-id="51722-135">验证域</span><span class="sxs-lookup"><span data-stu-id="51722-135">Verify the domain</span></span>

<span data-ttu-id="51722-136">在实现本节前面所述的验证逻辑时，您还应要求`amurl`声明的域与用户的自动发现域相匹配。</span><span class="sxs-lookup"><span data-stu-id="51722-136">When implementing the verification logic described previously in this section, you should also require that the domain of the `amurl` claim matches the Autodiscover domain for the user.</span></span> <span data-ttu-id="51722-137">若要执行此操作，您需要使用或实现自动发现。</span><span class="sxs-lookup"><span data-stu-id="51722-137">To do so, you'll need to use or implement Autodiscover.</span></span> <span data-ttu-id="51722-138">若要了解详细信息，可以从[Exchange 自动发现](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange)开始。</span><span class="sxs-lookup"><span data-stu-id="51722-138">To learn more, you can start with [Autodiscover for Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).</span></span>

## <a name="validate-the-identity-token-signature"></a><span data-ttu-id="51722-139">验证标识令牌签名</span><span class="sxs-lookup"><span data-stu-id="51722-139">Validate the identity token signature</span></span>

<span data-ttu-id="51722-140">知道 JWT 包含必需的声明后，便可以继续验证令牌签名。</span><span class="sxs-lookup"><span data-stu-id="51722-140">After you know that the JWT contains the required claims, you can proceed with validating the token signature.</span></span>

### <a name="retrieve-the-public-signing-key"></a><span data-ttu-id="51722-141">检索公用签名密钥</span><span class="sxs-lookup"><span data-stu-id="51722-141">Retrieve the public signing key</span></span>

<span data-ttu-id="51722-142">第一步是检索 Exchange 服务器用于为令牌签名的证书对应的公钥。</span><span class="sxs-lookup"><span data-stu-id="51722-142">The first step is to retrieve the public key that corresponds to the certificate that the Exchange server used to sign the token.</span></span> <span data-ttu-id="51722-143">可在身份验证元数据文档中找到此公钥。</span><span class="sxs-lookup"><span data-stu-id="51722-143">The key is found in the authentication metadata document.</span></span> <span data-ttu-id="51722-144">此文档是托管在 `amurl` 声明中指定的 URL 上的一个 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="51722-144">This document is a JSON file hosted at the URL specified in the `amurl` claim.</span></span>

<span data-ttu-id="51722-145">身份验证元数据文档使用以下格式。</span><span class="sxs-lookup"><span data-stu-id="51722-145">The authentication metadata document uses the following format.</span></span>

```json
{
    "id": "_70b34511-d105-4e2b-9675-39f53305bb01",
    "version": "1.0",
    "name": "Exchange",
    "realm": "*",
    "serviceName": "00000002-0000-0ff1-ce00-000000000000",
    "issuer": "00000002-0000-0ff1-ce00-000000000000@*",
    "allowedAudiences": [
        "00000002-0000-0ff1-ce00-000000000000@*"
    ],
    "keys": [
        {
            "usage": "signing",
            "keyinfo": {
                "x5t": "enh9BJrVPU5ijV1qjZjV-fL2bco"
            },
            "keyvalue": {
                "type": "x509Certificate",
                "value": "MIIHNTCC..."
            }
        }
    ],
    "endpoints": [
        {
            "location": "https://by2pr06mb2229.namprd06.prod.outlook.com:444/autodiscover/metadata/json/1",
            "protocol": "OAuth2",
            "usage": "metadata"
        }
    ]
}
```

<span data-ttu-id="51722-146">可用签名密钥位于 `keys` 数组中。</span><span class="sxs-lookup"><span data-stu-id="51722-146">The available signing keys are in the `keys` array.</span></span> <span data-ttu-id="51722-147">通过确保 `keyinfo` 属性中的 `x5t` 值与令牌标头中的 `x5t` 值相匹配，来选择正确的密钥。</span><span class="sxs-lookup"><span data-stu-id="51722-147">Select the correct key by ensuring that the `x5t` value in the `keyinfo` property matches the `x5t` value in the header of the token.</span></span> <span data-ttu-id="51722-148">公钥位于 `keyvalue` 属性中的 `value` 属性内，被存储为 Base64 编码的字节数组。</span><span class="sxs-lookup"><span data-stu-id="51722-148">The public key is inside the `value` property in the `keyvalue` property, stored as a base64-encoded byte array.</span></span>

<span data-ttu-id="51722-149">拥有正确的公钥后，验证此签名。</span><span class="sxs-lookup"><span data-stu-id="51722-149">After you have the correct public key, verify the signature.</span></span> <span data-ttu-id="51722-150">签名数据是已编码的令牌的前两个部分，用句点分隔：</span><span class="sxs-lookup"><span data-stu-id="51722-150">The signed data is the first two parts of the encoded token, separated by a period:</span></span>

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a><span data-ttu-id="51722-151">计算 Exchange 帐户的唯一 ID</span><span class="sxs-lookup"><span data-stu-id="51722-151">Compute the unique ID for an Exchange account</span></span>

<span data-ttu-id="51722-152">您可以通过将身份验证元数据文档 URL 与帐户的 Exchange 标识符连接来创建 Exchange 帐户的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="51722-152">You can create a unique identifier for an Exchange account by concatenating the authentication metadata document URL with the Exchange identifier for the account.</span></span> <span data-ttu-id="51722-153">如果你拥有此唯一标识符，则可以使用它为 Outlook 加载项 Web 服务创建单一登录 (SSO) 系统。</span><span class="sxs-lookup"><span data-stu-id="51722-153">When you have this unique identifier, you can use it to create a single sign-on (SSO) system for your Outlook add-in web service.</span></span> <span data-ttu-id="51722-154">有关将此唯一标识符用于 SSO 的详细信息，请参阅[对具有 Exchange 标识令牌的用户进行身份验证](authenticate-a-user-with-an-identity-token.md)。</span><span class="sxs-lookup"><span data-stu-id="51722-154">For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).</span></span>

## <a name="use-a-library-to-validate-the-token"></a><span data-ttu-id="51722-155">使用库验证令牌</span><span class="sxs-lookup"><span data-stu-id="51722-155">Use a library to validate the token</span></span>

<span data-ttu-id="51722-156">有许多库可以执行常规 JWT 解析和验证。</span><span class="sxs-lookup"><span data-stu-id="51722-156">There are a number of libraries that can do general JWT parsing and validation.</span></span> <span data-ttu-id="51722-157">Microsoft 提供了两个可用于验证 Exchange 用户标识令牌的库。</span><span class="sxs-lookup"><span data-stu-id="51722-157">Microsoft provides two libraries that can be used to validate Exchange user identity tokens.</span></span>

### <a name="systemidentitymodeltokensjwt"></a><span data-ttu-id="51722-158">System.IdentityModel.Tokens.Jwt</span><span class="sxs-lookup"><span data-stu-id="51722-158">System.IdentityModel.Tokens.Jwt</span></span>

<span data-ttu-id="51722-159">[System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) 库可以解析令牌，也可以执行验证，但需要自行解析 `appctx` 声明并检索公用签名密钥。</span><span class="sxs-lookup"><span data-stu-id="51722-159">The [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) library can parse the token and also perform the validation, though you will need to parse the `appctx` claim yourself and retrieve the public signing key.</span></span>

```cs
// Load the encoded token
string encodedToken = "...";
JwtSecurityToken jwt = new JwtSecurityToken(encodedToken);

// Parse the appctx claim to get the auth metadata url
string authMetadataUrl = string.Empty;
var appctx = jwt.Claims.FirstOrDefault(claim => claim.Type == "appctx");
if (appctx != null)
{
    var AppContext = JsonConvert.DeserializeObject<ExchangeAppContext>(appctx.Value);

    // Token version check
    if (string.Compare(AppContext.Version, "ExIdTok.V1", StringComparison.InvariantCulture) != 0) {
        // Fail validation
    }

    authMetadataUrl = AppContext.MetadataUrl;
}

// Use System.IdentityModel.Tokens.Jwt library to validate standard parts
JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
TokenValidationParameters tvp = new TokenValidationParameters();

tvp.ValidateIssuer = false;
tvp.ValidateAudience = true;
tvp.ValidAudience = "{URL to add-in}";
tvp.ValidateIssuerSigningKey = true;
// GetSigningKeys downloads the auth metadata doc and
// returns a List<SecurityKey>
tvp.IssuerSigningKeys = GetSigningKeys(authMetadataUrl);
tvp.ValidateLifetime = true;

try
{
    var claimsPrincipal = tokenHandler.ValidateToken(encodedToken, tvp, out SecurityToken validatedToken);

    // If no exception, all standard checks passed
}
catch (SecurityTokenValidationException ex)
{
    // Validation failed
}
```

<br/>

<span data-ttu-id="51722-160">`ExchangeAppContext` 类定义如下：</span><span class="sxs-lookup"><span data-stu-id="51722-160">The `ExchangeAppContext` class is defined as follows:</span></span>

```cs
using Newtonsoft.Json;

/// <summary>
/// Representation of the appctx claim in an Exchange user identity token.
/// </summary>
public class ExchangeAppContext
{
    /// <summary>
    /// The Exchange identifier for the user
    /// </summary>
    [JsonProperty("msexchuid")]
    public string ExchangeUid { get; set; }

    /// <summary>
    /// The token version
    /// </summary>
    public string Version { get; set; }

    /// <summary>
    /// The URL to download authentication metadata
    /// </summary>
    [JsonProperty("amurl")]
    public string MetadataUrl { get; set; }
}
```

<span data-ttu-id="51722-161">有关使用此库验证 Exchange 令牌并拥有 `GetSigningKeys` 实现的示例，请参阅 [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)。</span><span class="sxs-lookup"><span data-stu-id="51722-161">For an example that uses this library to validate Exchange tokens and has an implementation of `GetSigningKeys`, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>

### <a name="microsoftexchangewebservices"></a><span data-ttu-id="51722-162">Microsoft.Exchange.WebServices</span><span class="sxs-lookup"><span data-stu-id="51722-162">Microsoft.Exchange.WebServices</span></span>

<span data-ttu-id="51722-163">[Exchange Web 服务托管 API](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/) 也可以验证 Exchange 用户标识令牌。</span><span class="sxs-lookup"><span data-stu-id="51722-163">The [Exchange Web Services Managed API](https://www.nuget.org/packages/Microsoft.Exchange.WebServices/) can also validate Exchange user identity tokens.</span></span> <span data-ttu-id="51722-164">由于它是 Exchange 专用，因此它会实现所有必要逻辑，以解析 `appctx` 声明并验证令牌版本。</span><span class="sxs-lookup"><span data-stu-id="51722-164">Because it is Exchange-specific, it implements all of the necessary logic to parse the `appctx` claim and verify the token version.</span></span>

```cs
using Microsoft.Exchange.WebServices.Auth.Validation;

AppIdentityToken ValidateIdentityToken(string rawToken, string expectedAudience)
{
    try
    {
        AppIdentityToken appIdToken = AuthToken.Parse(rawToken) as AppIdentityToken;
        appIdToken.Validate(new Uri(expectedAudience));

        // No exception, validation succeeded
        return appIdToken;
    }
    catch (TokenValidationException ex)
    {
        throw new Exception(string.Format("Token validation failed: {0}", ex.Message));
    }
}
```

## <a name="see-also"></a><span data-ttu-id="51722-165">另请参阅</span><span class="sxs-lookup"><span data-stu-id="51722-165">See also</span></span>

- [<span data-ttu-id="51722-166">Outlook-Add-In-Token-Viewer</span><span class="sxs-lookup"><span data-stu-id="51722-166">Outlook-Add-In-Token-Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="51722-167">Outlook-Add-in-JavaScript-ValidateIdentityToken</span><span class="sxs-lookup"><span data-stu-id="51722-167">Outlook-Add-in-JavaScript-ValidateIdentityToken</span></span>](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
