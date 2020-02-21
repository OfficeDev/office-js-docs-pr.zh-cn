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
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a>使用 Exchange 的标识令牌对用户进行身份验证

Exchange 用户标识令牌为加载项提供了一种以唯一的方式标识加载项用户的方法。 通过创建用户标识，可以为后端服务实现单一登录 (SSO) 身份验证方案，此方案使使用 Outlook 加载项的客户无需登录即可连接服务。 有关何时使用此令牌类型的更多详细信息，请参阅 [Exchange 用户标识令牌](authentication.md#exchange-user-identity-token)。 本文介绍了使用 Exchange 标识令牌对访问后端的用户进行身份验证的简单方法。

> [!IMPORTANT]
> 这只是 SSO 实现的一个简单示例。 和以往一样，在处理标识和身份验证事宜时，一定要确保代码符合组织的安全要求。

## <a name="send-the-id-token-with-each-request"></a>通过每个请求发送 ID 令牌

第一步是通过调用 [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) 使加载项获取服务器中的 Exchange 用户标识令牌。 然后加载项通过向后端发出的每个请求发送该令牌。 它可能是在标头中，或在请求正文中。

## <a name="validate-the-token"></a>验证令牌

后端必须在接受令牌之前对其进行验证。 这是确保令牌是由用户的 Exchange 服务器颁发的重要步骤。 有关验证 Exchange 用户标识令牌的信息，请参阅[验证 Exchange 标识令牌](validate-an-identity-token.md)。

验证并解码之后，令牌的有效负载如下所示。

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

## <a name="map-the-token-to-a-user-in-your-backend"></a>将令牌映射到后端的用户

后端服务可以根据令牌计算唯一的用户 ID 并将其映射到内部用户系统中的用户。 例如，如果使用数据库来存储用户，可以在数据库中将此唯一 ID 添加到用户的记录中。

### <a name="generate-a-unique-id"></a>生成唯一 ID

建议结合使用 `msexchuid` 和 `amurl` 属性。 例如，可以将两个值连接在一起，生成 Base 64 编码的字符串。 每次均可通过令牌生成此值，因此你可以将 Exchange 用户标识令牌映射回系统中的用户。

### <a name="check-the-user"></a>检查用户

生成唯一 ID 后，下一步就是查找系统中使用此关联 ID 的用户。

- 如果能找到该用户，后端会将请求视为已经过身份验证并允许请求继续。

- 如果找不到用户，后端将返回一个错误，指示用户需要登录。 然后加载项会提示用户使用现有的身份验证方法登录到后端。 一旦用户经过身份验证，将提交含用户身份验证详细信息的 Exchange 用户标识令牌。 然后后端可以使用唯一 ID 更新系统中的用户记录。
