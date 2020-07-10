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
# <a name="validate-an-exchange-identity-token"></a>验证 Exchange 标识令牌

Outlook 加载项可以向你发送 Exchange 用户标识令牌，但是在你信任此请求之前，必须验证该令牌以确保它来自预期的 Exchange 服务器。 Exchange 用户标识令牌均为 JSON Web 令牌 (JWT)。 [RFC 7519 JSON Web 令牌 (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt) 中介绍了验证 JWT 所需的步骤。

建议使用四个步骤验证标识令牌并获取用户的唯一标识符。 首先，从 base64 URL 编码的字符串中提取 JSON Web 令牌 (JWT)。 然后，确保该令牌格式正确、它是用于 Outlook 外接程序的令牌、它未过期且你可以提取身份验证元数据文档的有效 URL。 接下来，从 Exchange 服务器中检索身份验证元数据文档并验证附加到标识令牌的签名。 最后，通过将用户的 Exchange ID 与身份验证元数据文档的 URL 连接，来计算用户的唯一标识符。

## <a name="extract-the-json-web-token"></a>提取 JSON Web 令牌

[getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 返回的令牌是令牌的编码字符串表示形式。 在此表示形式下，根据 RFC 7519，所有 JWT 都有三个部分，以句点分隔。 格式如下所示。

```json
{header}.{payload}.{signature}
```

标头和有效负载应进行 Base64 解码，以获取每一部分的 JSON 表示形式。 签名应进行 base64 解码，以获取包含二进制签名的字节数组。

有关令牌内容的详细信息，请参阅 [Exchange 标识令牌揭秘](inside-the-identity-token.md)。

三个组件都解码后，可以继续验证该令牌的内容。

## <a name="validate-token-contents"></a>验证令牌内容

若要验证令牌内容，还应检查以下项目。

- 检查标头并验证：
    - `typ`"声明" 设置为 `JWT` 。
    - `alg`"声明" 设置为 `RS256` 。
    - `x5t`声明存在。

- 检查有效负载并验证：
    - `amurl`中的声明 `appctx` 已设置为授权令牌签名密钥清单文件的位置。 例如， `amurl` Microsoft 365 的预期值为 https://outlook.office365.com:443/autodiscover/metadata/json/1 。 有关详细信息，请参阅下一节[验证域](#verify-the-domain)。
    - 当前时间介于和声明中指定的时间 `nbf` 之间 `exp` 。 `nbf` 声明指定了令牌被视为有效的最早时间，而 `exp` 声明指定了令牌的失效时间。 建议将服务器之间的时钟设置差异考虑在内。
    - `aud`声明是你的外接程序的预期 URL。
    - `version`声明内的声明 `appctx` 已设置为 `ExIdTok.V1` 。

### <a name="verify-the-domain"></a>验证域

在实现本节前面所述的验证逻辑时，您还应要求声明的域 `amurl` 与用户的自动发现域相匹配。 若要执行此操作，您需要使用或实现自动发现。 若要了解详细信息，可以从[Exchange 自动发现](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange)开始。

## <a name="validate-the-identity-token-signature"></a>验证标识令牌签名

知道 JWT 包含必需的声明后，便可以继续验证令牌签名。

### <a name="retrieve-the-public-signing-key"></a>检索公用签名密钥

第一步是检索 Exchange 服务器用于为令牌签名的证书对应的公钥。 可在身份验证元数据文档中找到此公钥。 此文档是托管在 `amurl` 声明中指定的 URL 上的一个 JSON 文件。

身份验证元数据文档使用以下格式。

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

可用签名密钥位于 `keys` 数组中。 通过确保 `keyinfo` 属性中的 `x5t` 值与令牌标头中的 `x5t` 值相匹配，来选择正确的密钥。 公钥位于 `keyvalue` 属性中的 `value` 属性内，被存储为 Base64 编码的字节数组。

拥有正确的公钥后，验证此签名。 签名数据是已编码的令牌的前两个部分，用句点分隔：

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a>计算 Exchange 帐户的唯一 ID

您可以通过将身份验证元数据文档 URL 与帐户的 Exchange 标识符连接来创建 Exchange 帐户的唯一标识符。 如果你拥有此唯一标识符，则可以使用它为 Outlook 加载项 Web 服务创建单一登录 (SSO) 系统。 有关将此唯一标识符用于 SSO 的详细信息，请参阅[对具有 Exchange 标识令牌的用户进行身份验证](authenticate-a-user-with-an-identity-token.md)。

## <a name="use-a-library-to-validate-the-token"></a>使用库验证令牌

有许多库可以执行常规 JWT 解析和验证。 Microsoft 提供 `System.IdentityModel.Tokens.Jwt` 可用于验证 Exchange 用户标识令牌的库。

> [!IMPORTANT]
> 我们不再建议使用 Exchange Web 服务托管 API，因为 Microsoft.Exchange.WebServices.Auth.dll 现在仍然可用，但它依赖于不受支持的库（如 Microsoft.IdentityModel.Extensions.dll）。

### <a name="systemidentitymodeltokensjwt"></a>System.IdentityModel.Tokens.Jwt

[System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) 库可以解析令牌，也可以执行验证，但需要自行解析 `appctx` 声明并检索公用签名密钥。

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

`ExchangeAppContext` 类定义如下：

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

有关使用此库验证 Exchange 令牌并拥有 `GetSigningKeys` 实现的示例，请参阅 [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)。

## <a name="see-also"></a>另请参阅

- [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
