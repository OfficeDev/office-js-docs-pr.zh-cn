---
title: 开发 Office 加载项以在使用第三方 Cookie 时使用 ITP
description: 使用第三方 Cookie 时如何使用 ITP 和 Office 加载项
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: b01051fa39441fddb2453b0bd95a0629ebf3ef65
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423088"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>开发 Office 加载项以在使用第三方 Cookie 时使用 ITP

如果 Office 加载项需要第三方 Cookie，则如果加载加载项的 [运行时](../testing/runtimes.md) 使用智能跟踪防护 (ITP) ，则会阻止这些 Cookie。 可以使用第三方 Cookie 对用户进行身份验证，或者针对其他方案（例如存储设置）进行身份验证。

如果 Office 加载项和网站必须依赖于第三方 Cookie，请使用以下步骤来处理 ITP。

1. 设置 [OAuth 2.0 授权](https://tools.ietf.org/html/rfc6749) ，以便在你的情况下 (身份验证域，需要 Cookie 的第三方) 将授权令牌转发到您的网站。 使用令牌与服务器集 Secure 和 [HttpOnly Cookie](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies) 建立第一方登录会话。
1. 使用 [存储访问 API](https://webkit.org/blog/8124/introducing-storage-access-api/) ，以便第三方可以请求访问其第一方 Cookie 的权限。 当前版本的 Office on Mac 和 Office web 版 都支持此 API。
    > [!NOTE]
    > 如果将 Cookie 用于身份验证以外的目的，请考虑 `localStorage` 用于方案。

以下代码示例演示如何使用存储访问 API。

```javascript
function displayLoginButton() {
  const button = createLoginButton();
  button.addEventListener("click", function(ev) {
    document.requestStorageAccess().then(function() {
      authenticateWithCookies(); 
    }).catch(function() {
      // User must have previously interacted with this domain loaded in a top frame
      // Also you should have previously written a cookie when domain was loaded in the top frame
      console.error("User cancelled or requirements were not met.");
    });
  });
}

if (document.hasStorageAccess) { 
  document.hasStorageAccess().then(function(hasStorageAccess) { 
    if (!hasStorageAccess) { 
      displayLoginButton(); 
    } else { 
      authenticateWithCookies(); 
    } 
  }); 
} else { 
    authenticateWithCookies(); 
} 
```

## <a name="about-itp-and-third-party-cookies"></a>关于 ITP 和第三方 Cookie

第三方 Cookie 是 iframe 中加载的 Cookie，其中域与顶级帧不同。 ITP 可能会影响复杂的身份验证方案，其中弹出对话框用于输入凭据，然后加载项 iframe 需要 Cookie 访问才能完成身份验证流。 ITP 也可能影响无提示身份验证方案，你以前曾使用弹出对话来进行身份验证，但随后使用外接程序会尝试通过隐藏的 iframe 进行身份验证。

在 Mac 上开发 Office 加载项时，MacOS Big Sur SDK 会阻止访问第三方 Cookie。 这是因为默认情况下，Safari 浏览器上启用了 WKWebView ITP，而 WKWebView 会阻止所有第三方 Cookie。 Office on Mac 版本 16.44 或更高版本与 MacOS Big Sur SDK 集成。

在 Safari 浏览器中，最终用户可以切换 **“首选项** > **隐私**”下 **的“防止跨站点跟踪”** 复选框以关闭 ITP。 但是，无法关闭嵌入式 WKWebView 控件的 ITP。

## <a name="see-also"></a>另请参阅

- [在 Safari 和其他阻止第三方 Cookie 的浏览器中处理 ITP](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [在 WebKit 中跟踪防护](https://webkit.org/tracking-prevention/)
- [Chrome 的“隐私沙盒”](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [介绍存储访问 API](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)
