---
title: Outlook 加载项中的 Exchange 标识令牌揭秘
description: 了解从 Outlook 加载项生成的 Exchange 用户标识令牌的内容。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: dee8416660386c25a55caa42b6e5ee8685ee8852
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609088"
---
# <a name="inside-the-exchange-identity-token"></a>Exchange 标识令牌揭秘

由 [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法返回的 Exchange 用户标识令牌为加载项代码提供了一种将用户的标识包含在后端服务调用中的方法。 本文将探讨令牌的格式和内容。

Exchange 用户标识令牌是一个 Base 64 URL 编码的字符串，由发送它的 Exchange 服务器签名。 该令牌未加密，用于验证签名的公钥存储在颁发该令牌的 Exchange 服务器上。 该令牌由三部分组成：标头、有效负载和签名。 在令牌字符串中，各部分由句点字符 (`.`) 分隔，以便于拆分令牌。

Exchange 使用标识令牌的 JSON Web 令牌 (JWT) 格式。 有关 JWT 令牌的信息，请参阅 [RFC 7519 JSON Web 令牌 (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt)。

## <a name="identity-token-header"></a>标识令牌标头

标头提供令牌的格式和签名的相关信息。 令牌标头如以下示例所示。

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
下表描述了令牌标头的各个部分。

| 声明 | 值 | 说明 |
|:-----|:-----|:-----|
| `typ` | `JWT` | 将令牌识别为 JSON Web 令牌。 Exchange 服务器提供的所有标识令牌均是 JWT 令牌。 |
| `alg` | `RS256` | 用于创建签名的哈希算法。 Exchange 服务器提供的所有令牌均结合使用了 RSASSA-PKCS1-v1_5 和 SHA-256 哈希算法。 |
| `x5t` | 证书指纹 | 令牌的 X.509 指纹。 |

## <a name="identity-token-payload"></a>标识令牌有效负载

有效负载包含身份验证声明，标识电子邮件帐户和发送令牌的 Exchange 服务器。下面的示例显示有效负载部分的形式。

```JSON
{ 
  "aud": "https://mailhost.contoso.com/IdentityTest.html", 
  "iss": "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
  "nbf": "1331579055", 
  "exp": "1331607855", 
  "appctxsender": "00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
  "isbrowserhostedapp": "true",
  "appctx": { 
    "msexchuid": "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com",
    "version": "ExIdTok.V1",
    "amurl": "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
  } 
}
```

<br/>
 
下表列出标识令牌有效负载的各个部分。

| 声明 | 说明 |
|:-----|:-----|
| `aud` | 请求该令牌的加载项的 URL。 只有客户端的浏览器运行的加载项发送的令牌有效。 如果加载项使用 Office 加载项清单 v1.1，则此 URL 为加载项清单的 [FormSettings](../reference/manifest/formsettings.md) 元素中首先出现的 `ItemRead` 或 `ItemEdit` 窗体类型下的第一个 `SourceLocation` 元素指定的 URL。 |
| `iss` | 颁发令牌的 Exchange 服务器的唯一标识符。此 Exchange 服务器颁发的所有令牌将具有相同标识符。 |
| `nbf` | 令牌开始生效的日期和时间。值是自 1970 年 1 月 1 日以来的秒数。 |
| `exp` | 标记失效的日期和时间，值是自 1970 年 1 月 1 日以来的秒数。 |
| `appctxsender` | 发送应用程序上下文的 Exchange 服务器的唯一标识符。 |
| `isbrowserhostedapp` | 指示加载项是否托管在浏览器中。 |
| `appctx` | 令牌的应用程序上下文。 |

appctx 声明中的信息提供了帐户的唯一标识符和用于为令牌签名的公钥的位置。 下表列出了 `appctx` 声明的各部分。

| 应用程序上下文属性 | 说明 |
|:-----|:-----|
| `msexchuid` | 与电子邮件帐户和 Exchange 服务器关联的唯一标识符。 |
| `version` | 令牌的版本号。 对于 Exchange 提供的所有令牌，值为 `ExIdTok.V1`。 |
| `amurl` | 身份验证元数据文档（包含用于登录该令牌的 X.509 证书的公钥）的 URL。<br/><br/>有关如何使用身份验证元数据文档的详细信息，请参阅[验证 Exchange 标识令牌](validate-an-identity-token.md)。 |

## <a name="identity-token-signature"></a>标识令牌签名

通过使用标头中指定的算法，并使用有效负载中指定的服务器位置处的自签名 X 509 证书，对标头和有效负载部分进行哈希处理来创建签名。Web 服务可以验证此签名，以帮助确保标识令牌来自预期的服务器。

## <a name="see-also"></a>另请参阅

有关解析 Exchange 用户标识令牌的示例，请参阅 [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)。
