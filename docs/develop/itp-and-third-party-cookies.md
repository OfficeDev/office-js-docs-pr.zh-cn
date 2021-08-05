---
title: 开发Office外接程序以使用第三方 Cookie 时与 ITP 一起使用
description: 使用第三方 cookie 时Office ITP 和加载项
ms.date: 07/8/2021
localization_priority: Normal
ms.openlocfilehash: 6a9452f24cb1cbd76c4f6cc3f39fab1f9310ec97
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773473"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>开发Office外接程序以使用第三方 Cookie 时与 ITP 一起使用

如果您的Office外接程序需要第三方 Cookie，则加载外接程序的浏览器运行时使用智能跟踪防护 (ITP) 时，将阻止这些 Cookie。 你可能会使用第三方 Cookie 对用户进行身份验证，或者用于存储设置等其他方案。

如果您的Office和网站必须依赖第三方 Cookie，请使用以下步骤来使用 ITP。

1. 设置[OAuth 2.0](https://tools.ietf.org/html/rfc6749)授权，以便验证域 (在这种情况下，需要 cookie 的第三) 将授权令牌转发到   您的网站。 使用令牌通过服务器集 Secure 和 HttpOnly Cookie 建立第一 [方登录会话](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)。
2. 使用[存储 Access API，](https://webkit.org/blog/8124/introducing-storage-access-api/)以便第三方可以请求获取访问其第一方   Cookie 的权限。 Mac 和 Office 上当前版本的 Office web 版都支持此 API。
    > [!NOTE]
    > 如果你将 Cookie 用于除身份验证外的其他目的，请考虑将 `localStorage` 用于你的方案。

下面的代码示例演示如何使用 存储 Access API。

```javascript
function displayLoginButton() {
  var button = createLoginButton();
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

第三方 Cookie 是在 iframe 中加载的 Cookie，其中域不同于顶级框架。 ITP 可能会影响复杂的身份验证方案，其中弹出对话框用于输入凭据，然后外接程序 iframe 需要 Cookie 访问才能完成身份验证流。 ITP 还可能会影响静默身份验证方案，其中您之前曾使用弹出对话框进行身份验证，但外接程序的后续使用会尝试通过隐藏的 iframe 进行身份验证。

在Office Mac 上开发外接程序时，对第三方 Cookie 的访问将被 MacOS Big Sur SDK 阻止。 这是因为默认情况下，在 Safari 浏览器中启用 WKWebView ITP，并且 WKWebView 会阻止所有第三方 Cookie。 Office Mac 版本 16.44 或更高版本上的版本与 MacOS 大 Sur SDK 集成。

在 Safari 浏览器中，最终用户可以切换首选项隐私下的"阻止 **跨** 网站跟踪"复选框  >  以关闭 ITP。 但是，无法为嵌入的 WKWebView 控件关闭 ITP。

## <a name="see-also"></a>另请参阅

- [在 Safari 和其他阻止第三方 Cookie 的浏览器中处理 ITP](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [WebKit 中的跟踪防护](https://webkit.org/tracking-prevention/)
- [Chrome 的"隐私沙盒"](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [存储 Access API](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)