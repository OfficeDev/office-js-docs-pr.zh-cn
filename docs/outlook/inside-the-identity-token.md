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
# <a name="inside-the-exchange-identity-token"></a><span data-ttu-id="9c50b-103">Exchange 标识令牌揭秘</span><span class="sxs-lookup"><span data-stu-id="9c50b-103">Inside the Exchange identity token</span></span>

<span data-ttu-id="9c50b-104">由 [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 方法返回的 Exchange 用户标识令牌为加载项代码提供了一种将用户的标识包含在后端服务调用中的方法。</span><span class="sxs-lookup"><span data-stu-id="9c50b-104">The Exchange user identity token returned by the [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method provides a way for your add-in code to include the user's identity with calls to your back-end service.</span></span> <span data-ttu-id="9c50b-105">本文将探讨令牌的格式和内容。</span><span class="sxs-lookup"><span data-stu-id="9c50b-105">This article will discuss the format and contents of the token.</span></span>

<span data-ttu-id="9c50b-106">Exchange 用户标识令牌是一个 Base 64 URL 编码的字符串，由发送它的 Exchange 服务器签名。</span><span class="sxs-lookup"><span data-stu-id="9c50b-106">An Exchange user identity token is a base-64 URL-encoded string that is signed by the Exchange server that sent it.</span></span> <span data-ttu-id="9c50b-107">该令牌未加密，用于验证签名的公钥存储在颁发该令牌的 Exchange 服务器上。</span><span class="sxs-lookup"><span data-stu-id="9c50b-107">The token is not encrypted, and the public key that you use to validate the signature is stored on the Exchange server that issued the token.</span></span> <span data-ttu-id="9c50b-108">该令牌由三部分组成：标头、有效负载和签名。</span><span class="sxs-lookup"><span data-stu-id="9c50b-108">The token has three parts: a header, a payload, and a signature.</span></span> <span data-ttu-id="9c50b-109">在令牌字符串中，各部分由句点字符 (`.`) 分隔，以便于拆分令牌。</span><span class="sxs-lookup"><span data-stu-id="9c50b-109">In the token string, the parts are separated by a period character (`.`) to make it easy for you to split the token.</span></span>

<span data-ttu-id="9c50b-110">Exchange 使用标识令牌的 JSON Web 令牌 (JWT) 格式。</span><span class="sxs-lookup"><span data-stu-id="9c50b-110">Exchange uses a the JSON Web Token (JWT) format for the identity token.</span></span> <span data-ttu-id="9c50b-111">有关 JWT 令牌的信息，请参阅 [RFC 7519 JSON Web 令牌 (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt)。</span><span class="sxs-lookup"><span data-stu-id="9c50b-111">For information about JWT tokens, see [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

## <a name="identity-token-header"></a><span data-ttu-id="9c50b-112">标识令牌标头</span><span class="sxs-lookup"><span data-stu-id="9c50b-112">Identity token header</span></span>

<span data-ttu-id="9c50b-113">标头提供令牌的格式和签名的相关信息。</span><span class="sxs-lookup"><span data-stu-id="9c50b-113">The header provides information about the format and signature information of the token.</span></span> <span data-ttu-id="9c50b-114">令牌标头如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="9c50b-114">The following example shows what the header of the token looks like.</span></span>

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
<span data-ttu-id="9c50b-115">下表描述了令牌标头的各个部分。</span><span class="sxs-lookup"><span data-stu-id="9c50b-115">The following table describes the parts of the token header.</span></span>

| <span data-ttu-id="9c50b-116">声明</span><span class="sxs-lookup"><span data-stu-id="9c50b-116">Claim</span></span> | <span data-ttu-id="9c50b-117">值</span><span class="sxs-lookup"><span data-stu-id="9c50b-117">Value</span></span> | <span data-ttu-id="9c50b-118">说明</span><span class="sxs-lookup"><span data-stu-id="9c50b-118">Description</span></span> |
|:-----|:-----|:-----|
| `typ` | `JWT` | <span data-ttu-id="9c50b-119">将令牌识别为 JSON Web 令牌。</span><span class="sxs-lookup"><span data-stu-id="9c50b-119">Identifies the token as a JSON Web Token.</span></span> <span data-ttu-id="9c50b-120">Exchange 服务器提供的所有标识令牌均是 JWT 令牌。</span><span class="sxs-lookup"><span data-stu-id="9c50b-120">All identity tokens provided by Exchange server are JWT tokens.</span></span> |
| `alg` | `RS256` | <span data-ttu-id="9c50b-121">用于创建签名的哈希算法。</span><span class="sxs-lookup"><span data-stu-id="9c50b-121">The hashing algorithm that is used to create the signature.</span></span> <span data-ttu-id="9c50b-122">Exchange 服务器提供的所有令牌均结合使用了 RSASSA-PKCS1-v1_5 和 SHA-256 哈希算法。</span><span class="sxs-lookup"><span data-stu-id="9c50b-122">All tokens provided by Exchange server use the RSASSA-PKCS1-v1_5 with SHA-256 hash algorithm.</span></span> |
| `x5t` | <span data-ttu-id="9c50b-123">证书指纹</span><span class="sxs-lookup"><span data-stu-id="9c50b-123">Certificate thumbprint</span></span> | <span data-ttu-id="9c50b-124">令牌的 X.509 指纹。</span><span class="sxs-lookup"><span data-stu-id="9c50b-124">The X.509 thumbprint of the token.</span></span> |

## <a name="identity-token-payload"></a><span data-ttu-id="9c50b-125">标识令牌有效负载</span><span class="sxs-lookup"><span data-stu-id="9c50b-125">Identity token payload</span></span>

<span data-ttu-id="9c50b-p107">有效负载包含身份验证声明，标识电子邮件帐户和发送令牌的 Exchange 服务器。下面的示例显示有效负载部分的形式。</span><span class="sxs-lookup"><span data-stu-id="9c50b-p107">The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.</span></span>

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
 
<span data-ttu-id="9c50b-128">下表列出标识令牌有效负载的各个部分。</span><span class="sxs-lookup"><span data-stu-id="9c50b-128">The following table lists the parts of the identity token payload.</span></span>

| <span data-ttu-id="9c50b-129">声明</span><span class="sxs-lookup"><span data-stu-id="9c50b-129">Claim</span></span> | <span data-ttu-id="9c50b-130">说明</span><span class="sxs-lookup"><span data-stu-id="9c50b-130">Description</span></span> |
|:-----|:-----|
| `aud` | <span data-ttu-id="9c50b-131">请求该令牌的加载项的 URL。</span><span class="sxs-lookup"><span data-stu-id="9c50b-131">The URL of the add-in that requested the token.</span></span> <span data-ttu-id="9c50b-132">只有客户端的浏览器运行的加载项发送的令牌有效。</span><span class="sxs-lookup"><span data-stu-id="9c50b-132">A token is only valid if it is sent from the add-in that is running in the client's browser.</span></span> <span data-ttu-id="9c50b-133">如果加载项使用 Office 加载项清单 v1.1，则此 URL 为加载项清单的 [FormSettings](../reference/manifest/formsettings.md) 元素中首先出现的 `ItemRead` 或 `ItemEdit` 窗体类型下的第一个 `SourceLocation` 元素指定的 URL。</span><span class="sxs-lookup"><span data-stu-id="9c50b-133">If the add-in uses the Office Add-ins manifests schema v1.1, this URL is the URL specified in the first `SourceLocation` element, under the form type `ItemRead` or `ItemEdit`, whichever occurs first as part of the [FormSettings](../reference/manifest/formsettings.md) element in the add-in manifest.</span></span> |
| `iss` | <span data-ttu-id="9c50b-p109">颁发令牌的 Exchange 服务器的唯一标识符。此 Exchange 服务器颁发的所有令牌将具有相同标识符。</span><span class="sxs-lookup"><span data-stu-id="9c50b-p109">A unique identifier for the Exchange server that issued the token. All tokens issued by this Exchange server will have the same identifier.</span></span> |
| `nbf` | <span data-ttu-id="9c50b-p110">令牌开始生效的日期和时间。值是自 1970 年 1 月 1 日以来的秒数。</span><span class="sxs-lookup"><span data-stu-id="9c50b-p110">The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970.</span></span> |
| `exp` | <span data-ttu-id="9c50b-p111">标记失效的日期和时间，值是自 1970 年 1 月 1 日以来的秒数。</span><span class="sxs-lookup"><span data-stu-id="9c50b-p111">The date and time that the token is valid until. The value is the number of seconds since January 1, 1970.</span></span> |
| `appctxsender` | <span data-ttu-id="9c50b-140">发送应用程序上下文的 Exchange 服务器的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="9c50b-140">A unique identifier for the Exchange server that sent the application context.</span></span> |
| `isbrowserhostedapp` | <span data-ttu-id="9c50b-141">指示加载项是否托管在浏览器中。</span><span class="sxs-lookup"><span data-stu-id="9c50b-141">Indicates whether the add-in is hosted in a browser.</span></span> |
| `appctx` | <span data-ttu-id="9c50b-142">令牌的应用程序上下文。</span><span class="sxs-lookup"><span data-stu-id="9c50b-142">The application context for the token.</span></span> |

<span data-ttu-id="9c50b-143">appctx 声明中的信息提供了帐户的唯一标识符和用于为令牌签名的公钥的位置。</span><span class="sxs-lookup"><span data-stu-id="9c50b-143">The information in the appctx claim provides you with the unique identifier for the account and the location of the public key used to sign the token.</span></span> <span data-ttu-id="9c50b-144">下表列出了 `appctx` 声明的各部分。</span><span class="sxs-lookup"><span data-stu-id="9c50b-144">The following table lists the parts of the `appctx` claim.</span></span>

| <span data-ttu-id="9c50b-145">应用程序上下文属性</span><span class="sxs-lookup"><span data-stu-id="9c50b-145">Application context property</span></span> | <span data-ttu-id="9c50b-146">说明</span><span class="sxs-lookup"><span data-stu-id="9c50b-146">Description</span></span> |
|:-----|:-----|
| `msexchuid` | <span data-ttu-id="9c50b-147">与电子邮件帐户和 Exchange 服务器关联的唯一标识符。</span><span class="sxs-lookup"><span data-stu-id="9c50b-147">A unique identifier associated with the email account and the Exchange server.</span></span> |
| `version` | <span data-ttu-id="9c50b-148">令牌的版本号。</span><span class="sxs-lookup"><span data-stu-id="9c50b-148">The version number of the token.</span></span> <span data-ttu-id="9c50b-149">对于 Exchange 提供的所有令牌，值为 `ExIdTok.V1`。</span><span class="sxs-lookup"><span data-stu-id="9c50b-149">For all tokens provided by Exchange, the value is `ExIdTok.V1`.</span></span> |
| `amurl` | <span data-ttu-id="9c50b-150">身份验证元数据文档（包含用于登录该令牌的 X.509 证书的公钥）的 URL。</span><span class="sxs-lookup"><span data-stu-id="9c50b-150">The URL of the authentication metadata document that contains the public key of the X.509 certificate that was used to sign the token.</span></span><br/><br/><span data-ttu-id="9c50b-151">有关如何使用身份验证元数据文档的详细信息，请参阅[验证 Exchange 标识令牌](validate-an-identity-token.md)。</span><span class="sxs-lookup"><span data-stu-id="9c50b-151">For more information about how to use the authentication metadata document, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span> |

## <a name="identity-token-signature"></a><span data-ttu-id="9c50b-152">标识令牌签名</span><span class="sxs-lookup"><span data-stu-id="9c50b-152">Identity token signature</span></span>

<span data-ttu-id="9c50b-p114">通过使用标头中指定的算法，并使用有效负载中指定的服务器位置处的自签名 X 509 证书，对标头和有效负载部分进行哈希处理来创建签名。Web 服务可以验证此签名，以帮助确保标识令牌来自预期的服务器。</span><span class="sxs-lookup"><span data-stu-id="9c50b-p114">The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.</span></span>

## <a name="see-also"></a><span data-ttu-id="9c50b-155">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9c50b-155">See also</span></span>

<span data-ttu-id="9c50b-156">有关解析 Exchange 用户标识令牌的示例，请参阅 [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)。</span><span class="sxs-lookup"><span data-stu-id="9c50b-156">For an example that parses the Exchange user identity token, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>
