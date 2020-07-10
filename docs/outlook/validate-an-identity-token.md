---
title: 验证 Outlook 加载项标识令牌
description: Outlook 加载项可以向你发送 Exchange 用户标识令牌，但是在你信任此请求之前，必须验证该令牌以确保它来自预期的 Exchange 服务器。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 6ad5f99093530528ec83cfc7a6e3a2571e0df491
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094104"
---
# <a name="validate-an-exchange-identity-token"></a><span data-ttu-id="d7a17-103">验证 Exchange 标识令牌</span><span class="sxs-lookup"><span data-stu-id="d7a17-103">Validate an Exchange identity token</span></span>

<span data-ttu-id="d7a17-104">Outlook 加载项可以向你发送 Exchange 用户标识令牌，但是在你信任此请求之前，必须验证该令牌以确保它来自预期的 Exchange 服务器。</span><span class="sxs-lookup"><span data-stu-id="d7a17-104">Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect.</span></span> <span data-ttu-id="d7a17-105">Exchange 用户标识令牌均为 JSON Web 令牌 (JWT)。</span><span class="sxs-lookup"><span data-stu-id="d7a17-105">Exchange user identity tokens are JSON Web Tokens (JWT).</span></span> <span data-ttu-id="d7a17-106">[RFC 7519 JSON Web 令牌 (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt) 中介绍了验证 JWT 所需的步骤。</span><span class="sxs-lookup"><span data-stu-id="d7a17-106">The steps required to validate a JWT are described in [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

<span data-ttu-id="d7a17-107">建议使用四个步骤验证标识令牌并获取用户的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="d7a17-107">We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier.</span></span> <span data-ttu-id="d7a17-108">首先，从 base64 URL 编码的字符串中提取 JSON Web 令牌 (JWT)。</span><span class="sxs-lookup"><span data-stu-id="d7a17-108">First, extract the JSON Web Token (JWT) from a base64 URL-encoded string.</span></span> <span data-ttu-id="d7a17-109">然后，确保该令牌格式正确、它是用于 Outlook 外接程序的令牌、它未过期且你可以提取身份验证元数据文档的有效 URL。</span><span class="sxs-lookup"><span data-stu-id="d7a17-109">Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document.</span></span> <span data-ttu-id="d7a17-110">接下来，从 Exchange 服务器中检索身份验证元数据文档并验证附加到标识令牌的签名。</span><span class="sxs-lookup"><span data-stu-id="d7a17-110">Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token.</span></span> <span data-ttu-id="d7a17-111">最后，通过将用户的 Exchange ID 与身份验证元数据文档的 URL 连接，来计算用户的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="d7a17-111">Finally, compute a unique identifier for the user by concatenating the user's Exchange ID with the URL of the authentication metadata document.</span></span>

## <a name="extract-the-json-web-token"></a><span data-ttu-id="d7a17-112">提取 JSON Web 令牌</span><span class="sxs-lookup"><span data-stu-id="d7a17-112">Extract the JSON Web Token</span></span>

<span data-ttu-id="d7a17-113">[getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 返回的令牌是令牌的编码字符串表示形式。</span><span class="sxs-lookup"><span data-stu-id="d7a17-113">The token returned from [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) is an encoded string representation of the token.</span></span> <span data-ttu-id="d7a17-114">在此表示形式下，根据 RFC 7519，所有 JWT 都有三个部分，以句点分隔。</span><span class="sxs-lookup"><span data-stu-id="d7a17-114">In this form, per RFC 7519, all JWTs have three parts, separated by a period.</span></span> <span data-ttu-id="d7a17-115">格式如下所示。</span><span class="sxs-lookup"><span data-stu-id="d7a17-115">The format is as follows.</span></span>

```json
{header}.{payload}.{signature}
```

<span data-ttu-id="d7a17-116">标头和有效负载应进行 Base64 解码，以获取每一部分的 JSON 表示形式。</span><span class="sxs-lookup"><span data-stu-id="d7a17-116">The header and payload should be base64-decoded to obtain a JSON representation of each part.</span></span> <span data-ttu-id="d7a17-117">签名应进行 base64 解码，以获取包含二进制签名的字节数组。</span><span class="sxs-lookup"><span data-stu-id="d7a17-117">The signature should be base64-decoded to obtain a byte array containing the binary signature.</span></span>

<span data-ttu-id="d7a17-118">有关令牌内容的详细信息，请参阅 [Exchange 标识令牌揭秘](inside-the-identity-token.md)。</span><span class="sxs-lookup"><span data-stu-id="d7a17-118">For more information about the contents of the token, see [Inside the Exchange identity token](inside-the-identity-token.md).</span></span>

<span data-ttu-id="d7a17-119">三个组件都解码后，可以继续验证该令牌的内容。</span><span class="sxs-lookup"><span data-stu-id="d7a17-119">After you have the three decoded components, you can proceed with validating the content of the token.</span></span>

## <a name="validate-token-contents"></a><span data-ttu-id="d7a17-120">验证令牌内容</span><span class="sxs-lookup"><span data-stu-id="d7a17-120">Validate token contents</span></span>

<span data-ttu-id="d7a17-121">若要验证令牌内容，还应检查以下项目。</span><span class="sxs-lookup"><span data-stu-id="d7a17-121">To validate the token contents, you should check the following.</span></span>

- <span data-ttu-id="d7a17-122">检查标头并验证：</span><span class="sxs-lookup"><span data-stu-id="d7a17-122">Check the header and verify that the:</span></span>
    - <span data-ttu-id="d7a17-123">`typ`"声明" 设置为 `JWT` 。</span><span class="sxs-lookup"><span data-stu-id="d7a17-123">`typ` claim is set to `JWT`.</span></span>
    - <span data-ttu-id="d7a17-124">`alg`"声明" 设置为 `RS256` 。</span><span class="sxs-lookup"><span data-stu-id="d7a17-124">`alg` claim is set to `RS256`.</span></span>
    - <span data-ttu-id="d7a17-125">`x5t`声明存在。</span><span class="sxs-lookup"><span data-stu-id="d7a17-125">`x5t` claim is present.</span></span>

- <span data-ttu-id="d7a17-126">检查有效负载并验证：</span><span class="sxs-lookup"><span data-stu-id="d7a17-126">Check the payload and verify that the:</span></span>
    - <span data-ttu-id="d7a17-127">`amurl`中的声明 `appctx` 已设置为授权令牌签名密钥清单文件的位置。</span><span class="sxs-lookup"><span data-stu-id="d7a17-127">`amurl` claim inside the `appctx` is set to the location of an authorized token signing key manifest file.</span></span> <span data-ttu-id="d7a17-128">例如， `amurl` Microsoft 365 的预期值为 https://outlook.office365.com:443/autodiscover/metadata/json/1 。</span><span class="sxs-lookup"><span data-stu-id="d7a17-128">For example, the expected `amurl` value for Microsoft 365 is https://outlook.office365.com:443/autodiscover/metadata/json/1.</span></span> <span data-ttu-id="d7a17-129">有关详细信息，请参阅下一节[验证域](#verify-the-domain)。</span><span class="sxs-lookup"><span data-stu-id="d7a17-129">See the next section [Verify the domain](#verify-the-domain) for additional information.</span></span>
    - <span data-ttu-id="d7a17-130">当前时间介于和声明中指定的时间 `nbf` 之间 `exp` 。</span><span class="sxs-lookup"><span data-stu-id="d7a17-130">Current time is between the times specified in the `nbf` and `exp` claims.</span></span> <span data-ttu-id="d7a17-131">`nbf` 声明指定了令牌被视为有效的最早时间，而 `exp` 声明指定了令牌的失效时间。</span><span class="sxs-lookup"><span data-stu-id="d7a17-131">The `nbf` claim specifies the earliest time that the token is considered valid, and the `exp` claim specifies the expiration time for the token.</span></span> <span data-ttu-id="d7a17-132">建议将服务器之间的时钟设置差异考虑在内。</span><span class="sxs-lookup"><span data-stu-id="d7a17-132">It is recommended to allow for some variation in clock settings between servers.</span></span>
    - <span data-ttu-id="d7a17-133">`aud`声明是你的外接程序的预期 URL。</span><span class="sxs-lookup"><span data-stu-id="d7a17-133">`aud` claim is the expected URL for your add-in.</span></span>
    - <span data-ttu-id="d7a17-134">`version`声明内的声明 `appctx` 已设置为 `ExIdTok.V1` 。</span><span class="sxs-lookup"><span data-stu-id="d7a17-134">`version` claim inside the `appctx` claim is set to `ExIdTok.V1`.</span></span>

### <a name="verify-the-domain"></a><span data-ttu-id="d7a17-135">验证域</span><span class="sxs-lookup"><span data-stu-id="d7a17-135">Verify the domain</span></span>

<span data-ttu-id="d7a17-136">在实现本节前面所述的验证逻辑时，您还应要求声明的域 `amurl` 与用户的自动发现域相匹配。</span><span class="sxs-lookup"><span data-stu-id="d7a17-136">When implementing the verification logic described previously in this section, you should also require that the domain of the `amurl` claim matches the Autodiscover domain for the user.</span></span> <span data-ttu-id="d7a17-137">若要执行此操作，您需要使用或实现自动发现。</span><span class="sxs-lookup"><span data-stu-id="d7a17-137">To do so, you'll need to use or implement Autodiscover.</span></span> <span data-ttu-id="d7a17-138">若要了解详细信息，可以从[Exchange 自动发现](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange)开始。</span><span class="sxs-lookup"><span data-stu-id="d7a17-138">To learn more, you can start with [Autodiscover for Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).</span></span>

## <a name="validate-the-identity-token-signature"></a><span data-ttu-id="d7a17-139">验证标识令牌签名</span><span class="sxs-lookup"><span data-stu-id="d7a17-139">Validate the identity token signature</span></span>

<span data-ttu-id="d7a17-140">知道 JWT 包含必需的声明后，便可以继续验证令牌签名。</span><span class="sxs-lookup"><span data-stu-id="d7a17-140">After you know that the JWT contains the required claims, you can proceed with validating the token signature.</span></span>

### <a name="retrieve-the-public-signing-key"></a><span data-ttu-id="d7a17-141">检索公用签名密钥</span><span class="sxs-lookup"><span data-stu-id="d7a17-141">Retrieve the public signing key</span></span>

<span data-ttu-id="d7a17-142">第一步是检索 Exchange 服务器用于为令牌签名的证书对应的公钥。</span><span class="sxs-lookup"><span data-stu-id="d7a17-142">The first step is to retrieve the public key that corresponds to the certificate that the Exchange server used to sign the token.</span></span> <span data-ttu-id="d7a17-143">可在身份验证元数据文档中找到此公钥。</span><span class="sxs-lookup"><span data-stu-id="d7a17-143">The key is found in the authentication metadata document.</span></span> <span data-ttu-id="d7a17-144">此文档是托管在 `amurl` 声明中指定的 URL 上的一个 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="d7a17-144">This document is a JSON file hosted at the URL specified in the `amurl` claim.</span></span>

<span data-ttu-id="d7a17-145">身份验证元数据文档使用以下格式。</span><span class="sxs-lookup"><span data-stu-id="d7a17-145">The authentication metadata document uses the following format.</span></span>

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

<span data-ttu-id="d7a17-146">可用签名密钥位于 `keys` 数组中。</span><span class="sxs-lookup"><span data-stu-id="d7a17-146">The available signing keys are in the `keys` array.</span></span> <span data-ttu-id="d7a17-147">通过确保 `keyinfo` 属性中的 `x5t` 值与令牌标头中的 `x5t` 值相匹配，来选择正确的密钥。</span><span class="sxs-lookup"><span data-stu-id="d7a17-147">Select the correct key by ensuring that the `x5t` value in the `keyinfo` property matches the `x5t` value in the header of the token.</span></span> <span data-ttu-id="d7a17-148">公钥位于 `keyvalue` 属性中的 `value` 属性内，被存储为 Base64 编码的字节数组。</span><span class="sxs-lookup"><span data-stu-id="d7a17-148">The public key is inside the `value` property in the `keyvalue` property, stored as a base64-encoded byte array.</span></span>

<span data-ttu-id="d7a17-149">拥有正确的公钥后，验证此签名。</span><span class="sxs-lookup"><span data-stu-id="d7a17-149">After you have the correct public key, verify the signature.</span></span> <span data-ttu-id="d7a17-150">签名数据是已编码的令牌的前两个部分，用句点分隔：</span><span class="sxs-lookup"><span data-stu-id="d7a17-150">The signed data is the first two parts of the encoded token, separated by a period:</span></span>

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a><span data-ttu-id="d7a17-151">计算 Exchange 帐户的唯一 ID</span><span class="sxs-lookup"><span data-stu-id="d7a17-151">Compute the unique ID for an Exchange account</span></span>

<span data-ttu-id="d7a17-152">您可以通过将身份验证元数据文档 URL 与帐户的 Exchange 标识符连接来创建 Exchange 帐户的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="d7a17-152">You can create a unique identifier for an Exchange account by concatenating the authentication metadata document URL with the Exchange identifier for the account.</span></span> <span data-ttu-id="d7a17-153">如果你拥有此唯一标识符，则可以使用它为 Outlook 加载项 Web 服务创建单一登录 (SSO) 系统。</span><span class="sxs-lookup"><span data-stu-id="d7a17-153">When you have this unique identifier, you can use it to create a single sign-on (SSO) system for your Outlook add-in web service.</span></span> <span data-ttu-id="d7a17-154">有关将此唯一标识符用于 SSO 的详细信息，请参阅[对具有 Exchange 标识令牌的用户进行身份验证](authenticate-a-user-with-an-identity-token.md)。</span><span class="sxs-lookup"><span data-stu-id="d7a17-154">For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).</span></span>

## <a name="use-a-library-to-validate-the-token"></a><span data-ttu-id="d7a17-155">使用库验证令牌</span><span class="sxs-lookup"><span data-stu-id="d7a17-155">Use a library to validate the token</span></span>

<span data-ttu-id="d7a17-156">有许多库可以执行常规 JWT 解析和验证。</span><span class="sxs-lookup"><span data-stu-id="d7a17-156">There are a number of libraries that can do general JWT parsing and validation.</span></span> <span data-ttu-id="d7a17-157">Microsoft 提供 `System.IdentityModel.Tokens.Jwt` 可用于验证 Exchange 用户标识令牌的库。</span><span class="sxs-lookup"><span data-stu-id="d7a17-157">Microsoft provides the `System.IdentityModel.Tokens.Jwt` library that can be used to validate Exchange user identity tokens.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d7a17-158">我们不再建议使用 Exchange Web 服务托管 API，因为 Microsoft.Exchange.WebServices.Auth.dll 现在仍然可用，但它依赖于不受支持的库（如 Microsoft.IdentityModel.Extensions.dll）。</span><span class="sxs-lookup"><span data-stu-id="d7a17-158">We no longer recommend the Exchange Web Services Managed API because the Microsoft.Exchange.WebServices.Auth.dll, though still available, is now obsolete and relies on unsupported libraries like Microsoft.IdentityModel.Extensions.dll.</span></span>

### <a name="systemidentitymodeltokensjwt"></a><span data-ttu-id="d7a17-159">System.IdentityModel.Tokens.Jwt</span><span class="sxs-lookup"><span data-stu-id="d7a17-159">System.IdentityModel.Tokens.Jwt</span></span>

<span data-ttu-id="d7a17-160">[System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) 库可以解析令牌，也可以执行验证，但需要自行解析 `appctx` 声明并检索公用签名密钥。</span><span class="sxs-lookup"><span data-stu-id="d7a17-160">The [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) library can parse the token and also perform the validation, though you will need to parse the `appctx` claim yourself and retrieve the public signing key.</span></span>

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

<span data-ttu-id="d7a17-161">`ExchangeAppContext` 类定义如下：</span><span class="sxs-lookup"><span data-stu-id="d7a17-161">The `ExchangeAppContext` class is defined as follows:</span></span>

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

<span data-ttu-id="d7a17-162">有关使用此库验证 Exchange 令牌并拥有 `GetSigningKeys` 实现的示例，请参阅 [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)。</span><span class="sxs-lookup"><span data-stu-id="d7a17-162">For an example that uses this library to validate Exchange tokens and has an implementation of `GetSigningKeys`, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>

## <a name="see-also"></a><span data-ttu-id="d7a17-163">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d7a17-163">See also</span></span>

- [<span data-ttu-id="d7a17-164">Outlook-Add-In-Token-Viewer</span><span class="sxs-lookup"><span data-stu-id="d7a17-164">Outlook-Add-In-Token-Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="d7a17-165">Outlook-Add-in-JavaScript-ValidateIdentityToken</span><span class="sxs-lookup"><span data-stu-id="d7a17-165">Outlook-Add-in-JavaScript-ValidateIdentityToken</span></span>](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
