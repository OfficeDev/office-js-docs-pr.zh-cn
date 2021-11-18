---
title: 在使用基于事件的 () Outlook加载项中启用单一登录或 SSO 登录
description: 了解如何在基于事件的激活加载项中操作时启用 SSO。
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 66d1edb8b7b0092ee107b73af24d5420caee8677
ms.sourcegitcommit: 6e6c4803fdc0a3cc2c1bcd275288485a987551ff
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2021
ms.locfileid: "61066651"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>在使用基于事件的 () Outlook加载项中启用单一登录或 SSO 登录

当Outlook加载项使用基于事件的激活时，事件在单独的 JavaScript 运行时中运行。 完成在[Outlook](authenticate-a-user-with-an-sso-token.md)加载项中使用单一登录令牌对用户进行身份验证中的步骤后，请按照本文中所述的其他步骤操作，为事件处理代码启用 SSO。 启用 SSO 后，可以调用 `getAccessToken()` API 获取具有用户标识的访问令牌。

> [!NOTE]
> 本文中的步骤仅适用于在加载项Outlook加载项Windows。 这是因为，Outlook上的Windows使用 JavaScript 文件，而 Outlook 网页版 使用可引用同一 JavaScript 文件的 HTML 文件。

For Outlook on Windows， in the manifest for your Outlook add-in， you identify a single JavaScript file to load for event-based activation. 还需要指定是否Office此文件支持 SSO。 有两种方法可以这样做。 您可以创建所有加载项及其 JavaScript 文件的列表，以Office已知的 URI 访问加载项。 也可以添加自定义响应标头以启用 SSO。

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>列出具有已知 URI 的允许加载项

若要列出允许哪些加载项使用 SSO，请创建一个 JSON 文件，用于标识每个加载项的每个 JavaScript 文件。 然后，在已知 URI 上托管该 JSON 文件。 已知 URI 允许指定授权获取当前 Web 源令牌的所有托管 JS 文件。 这将确保源所有者对哪些托管 JS 文件应用于外接程序以及哪些不用于外接程序具有完全控制权，例如，防止有关模拟的任何安全漏洞。

以下示例演示如何在主版本和 beta (中为两个外接程序启用 SSO) 。 您可以列出所需多的加载项，具体取决于从 Web 服务器提供的加载项数。

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

将 JSON 文件托管在 `.well-known` 源根目录的 URI 中命名的位置下。 例如，如果原点为 `https://addin.contoso.com:8000/` ，则已知 URI 为 `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json` 。

源引用方案 + 子域 + 域 + 端口的模式。 位置的名称 **必须为** `.well-known` ，资源文件的名称必须为 `microsoft-officeaddins-allowed.json` 。 此文件必须包含一个 JSON 对象，其属性名为 ，其值是授权 SSO 用于其各自外接程序的所有 `allowed` JavaScript 文件的数组。

## <a name="add-a-custom-response-header"></a>添加自定义响应头

第二个方法是添加名为 的自定义响应标头 `MS-OfficeAddins-Allowed-Origin` 。 标头的值必须是 JavaScript 文件的来源。

例如，如果 JavaScript 文件位于 `https://addin.contoso.com:8000/main/js/autorun.js` ，则添加以下响应标头。

`MS-OfficeAddins-Allowed-Origin : https://addin.contoso.com:8000`

您需要参考特定 Web 服务器文档，了解如何添加自定义响应标头。

## <a name="see-also"></a>另请参阅

- [使用加载项中的单一登录令牌Outlook用户](authenticate-a-user-with-an-sso-token.md)
- [配置Outlook加载项进行基于事件的激活](autolaunch.md)
