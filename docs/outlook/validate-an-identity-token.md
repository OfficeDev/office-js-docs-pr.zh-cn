---
title: 验证 Outlook 加载项标识令牌
description: Outlook 加载项可以向你发送 Exchange 用户标识令牌，但是在你信任此请求之前，必须验证该令牌以确保它来自预期的 Exchange 服务器。
ms.date: 10/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6b903b582fee59fd1c5ff0aa949d614c4ee1dff7
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484413"
---
# <a name="validate-an-exchange-identity-token"></a>验证 Exchange 标识令牌

Outlook 加载项可以向你发送 Exchange 用户标识令牌，但是在你信任此请求之前，必须验证该令牌以确保它来自预期的 Exchange 服务器。 Exchange 用户标识令牌均为 JSON Web 令牌 (JWT)。 [RFC 7519 JSON Web 令牌 (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt) 中介绍了验证 JWT 所需的步骤。

建议使用四个步骤验证标识令牌并获取用户的唯一标识符。 首先，从 base64 URL 编码的字符串中提取 JSON Web 令牌 (JWT)。 然后，确保该令牌格式正确、它是用于 Outlook 外接程序的令牌、它未过期且你可以提取身份验证元数据文档的有效 URL。 接下来，从 Exchange 服务器中检索身份验证元数据文档并验证附加到标识令牌的签名。 最后，将用户的 ID 与身份验证元数据文档的 URL Exchange连接来计算用户的唯一标识符。

## <a name="extract-the-json-web-token"></a>提取 JSON Web 令牌

[getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 返回的令牌是令牌的编码字符串表示形式。 在此表示形式下，根据 RFC 7519，所有 JWT 都有三个部分，以句点分隔。 格式如下所示。

```json
{header}.{payload}.{signature}
```

标头和有效负载应进行 Base64 解码，以获取每一部分的 JSON 表示形式。 签名应进行 base64 解码，以获取包含二进制签名的字节数组。

有关令牌内容的详细信息，请参阅 [Exchange 标识令牌揭秘](inside-the-identity-token.md)。

三个组件都解码后，可以继续验证该令牌的内容。

## <a name="validate-token-contents"></a>验证令牌内容

若要验证令牌内容，应检查以下内容：

- 检查标头并验证：
  - `typ` 声明设置为 `JWT`。
  - `alg` 声明设置为 `RS256`。
  - `x5t` 声明存在。

- 检查有效负载并验证：
  - `amurl` 中的 声明 `appctx` 设置为授权令牌签名密钥清单文件的位置。 例如，Microsoft 365`amurl`值为 https://outlook.office365.com:443/autodiscover/metadata/json/1。 有关其他信息，请参阅下 [一部分验证](#verify-the-domain) 域。
  - 当前时间介于 和 声明中指定的`nbf``exp`时间之间。 `nbf` 声明指定了令牌被视为有效的最早时间，而 `exp` 声明指定了令牌的失效时间。 建议将服务器之间的时钟设置差异考虑在内。
  - `aud` claim 是外接程序的预期 URL。
  - `version` 声明内的 `appctx` 声明设置为 `ExIdTok.V1`。

### <a name="verify-the-domain"></a>验证域

实现上一节中所述的 `amurl` 验证逻辑时，还必须要求声明的域与用户的自动发现域匹配。 为此，你需要使用或实现自动发现[Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange)。

- For Exchange Online， confirm that the `amurl` is a well-known domain (https://outlook.office365.com:443/autodiscover/metadata/json/1)， or belongs to a geo-specific or 专业 cloud ([Office 365 URLs and IP address ranges](/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide&preserve-view=true)) .

- 如果您的外接程序服务预先具有用户租户的配置，则您可以确定这是否 `amurl` 受信任。

- 对于[Exchange部署](/microsoft-365/enterprise/configure-exchange-server-for-hybrid-modern-authentication?view=o365-worldwide&preserve-view=true)，使用基于 OAuth 的自动发现验证用户预期的域。 但是，虽然用户需要作为自动发现流的一部分进行身份验证，但外接程序永远不应收集用户的凭据和执行基本身份验证。

`amurl`如果加载项无法验证是否使用了这些选项中的任一选项，可以选择让加载项正常关闭，如果加载项工作流需要身份验证，则向用户发送相应通知。

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

通过连接身份验证元数据文档 URL 和Exchange的身份验证元数据文档 URL，为Exchange创建唯一标识符。 具有此唯一标识符时，使用它为 (Web 服务) SSO Outlook单一登录。 有关将此唯一标识符用于 SSO 的详细信息，请参阅[对具有 Exchange 标识令牌的用户进行身份验证](authenticate-a-user-with-an-identity-token.md)。

## <a name="use-a-library-to-validate-the-token"></a>使用库验证令牌

有许多库可以执行常规 JWT 解析和验证。 Microsoft 提供了`System.IdentityModel.Tokens.Jwt`可用于验证用户标识Exchange库。

> [!IMPORTANT]
> 我们不再建议使用 Exchange Web 服务托管 API，因为 Microsoft.Exchange.WebServices.Auth.dll 尽管仍然可用，但现在已过时，并且依赖于不受支持库（如 Microsoft.IdentityModel.Extensions.dll）。

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
