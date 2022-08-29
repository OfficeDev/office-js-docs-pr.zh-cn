---
title: '在使用基于事件的激活的 Outlook 加载项中启用单一登录 (SSO) '
description: 了解如何在基于事件的激活加载项中工作时启用 SSO。
ms.date: 06/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 10fd973c0476878443d7238e8805aa4db9f62953
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423116"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>在使用基于事件的激活的 Outlook 加载项中启用单一登录 (SSO) 

当 Outlook 加载项使用基于事件的激活时，事件会在单独的 [运行时中运行](../testing/runtimes.md)。 完成在 Outlook 加载项 [中使用单一登录令牌对用户进行身份验证](authenticate-a-user-with-an-sso-token.md)的步骤后，请按照本文中所述的其他步骤为事件处理代码启用 SSO。 启用 SSO 后，可以调用 [getAccessToken () API](/javascript/api/office-runtime/officeruntime.auth) 以获取具有用户标识的访问令牌。

> [!IMPORTANT]
> 虽然 `OfficeRuntime.auth.getAccessToken` 检索访问令牌并 `Office.auth.getAccessToken` 执行相同的功能，但我们建议在基于事件的加载项中调用 `OfficeRuntime.auth.getAccessToken` 。 支持基于事件的激活和 SSO 的所有 Outlook 客户端版本都支持此 API。 另一方面， `Office.auth.getAccessToken` 仅从版本 2111 (内部版本 14701.20000) 开始，Outlook on Windows 才受支持。

对于 Outlook on Windows，在 Outlook 外接程序的清单中，标识要加载的单个 JavaScript 文件以进行基于事件的激活。 还需要向 Office 指定允许此文件支持 SSO。 为此，请创建所有加载项及其 JavaScript 文件的列表，以便通过已知的 URI 提供给 Office。

> [!NOTE]
> 本文中的步骤仅适用于在 Windows 上运行 Outlook 加载项时。 这是因为 Windows 上的 Outlook 使用 JavaScript 文件，而Outlook 网页版使用可引用同一 JavaScript 文件的 HTML 文件。

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>使用已知 URI 列出允许的加载项

若要列出允许使用 SSO 的加载项，请创建一个 JSON 文件，用于标识每个加载项的每个 JavaScript 文件。 然后在已知 URI 中托管该 JSON 文件。 众所周知的 URI 允许对所有已授权获取当前 Web 源令牌的托管 JS 文件进行规范。 这可确保源的所有者能够完全控制哪些托管 JS 文件应用于加载项，哪些文件不是，从而防止了模拟周围的任何安全漏洞，例如。

以下示例演示如何为主版本和 beta 版本)  (两个加载项启用 SSO。 可以根据需要列出任意数量的加载项，具体取决于从 Web 服务器提供的加载项数。

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

在源根目录的 URI 中命名 `.well-known` 的位置下托管 JSON 文件。 例如，如果源是 `https://addin.contoso.com:8000/`，则已知的 URI 为 `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`。

源是指方案 + 子域 + 域 + 端口的模式。 位置的名称 **必须** 是 `.well-known`，资源文件的名称 **必须** 是 `microsoft-officeaddins-allowed.json`。 此文件必须包含一个 JSON 对象，其属性名为 `allowed` 其值是为其各自的外接程序为 SSO 授权的所有 JavaScript 文件的数组。

## <a name="see-also"></a>另请参阅

- [在 Outlook 加载项中使用单一登录令牌对用户进行身份验证](authenticate-a-user-with-an-sso-token.md)
- [配置 Outlook 外接程序以进行基于事件的激活](autolaunch.md)
