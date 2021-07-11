---
title: 开发Office外接程序以使用第三方 Cookie 时与 ITP 一起使用
description: 使用第三方 cookie 时Office ITP 和加载项
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: dbc23e4ead0abc94ffa173ffc22919342c4fca6d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349859"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a><span data-ttu-id="979f8-103">开发Office外接程序以使用第三方 Cookie 时与 ITP 一起使用</span><span class="sxs-lookup"><span data-stu-id="979f8-103">Develop your Office Add-in to work with ITP when using third-party cookies</span></span>

<span data-ttu-id="979f8-104">如果您的Office外接程序需要第三方 Cookie，则加载外接程序的浏览器运行时使用智能跟踪防护 (ITP) 时，将阻止这些 Cookie。</span><span class="sxs-lookup"><span data-stu-id="979f8-104">If your Office Add-in requires third-party cookies, those cookies are blocked if Intelligent Tracking Prevention (ITP) is used by the browser runtime that loaded your add-in.</span></span> <span data-ttu-id="979f8-105">你可能会使用第三方 Cookie 对用户进行身份验证，或者用于存储设置等其他方案。</span><span class="sxs-lookup"><span data-stu-id="979f8-105">You may be using third-party cookies to authenticate users, or for other scenarios, such as storing settings.</span></span>

<span data-ttu-id="979f8-106">如果您的Office和网站必须依赖第三方 Cookie，请使用以下步骤来使用 ITP：</span><span class="sxs-lookup"><span data-stu-id="979f8-106">If your Office Add-in and website must rely on third-party cookies, use the following steps to work with ITP:</span></span>

1. <span data-ttu-id="979f8-107">设置[OAuth 2.0](https://tools.ietf.org/html/rfc6749)授权，以便验证域 (在这种情况下，需要 cookie 的第三) 将授权令牌转发到   您的网站。</span><span class="sxs-lookup"><span data-stu-id="979f8-107">Set up [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749) so that the authenticating domain (in your case, the third-party that expects cookies) forwards an authorization token to your website.</span></span> <span data-ttu-id="979f8-108">使用令牌通过服务器集 Secure 和 HttpOnly Cookie 建立第一 [方登录会话](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)。</span><span class="sxs-lookup"><span data-stu-id="979f8-108">Use the token to establish a first-party login session with a server-set Secure and [HttpOnly cookie](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).</span></span>
2. <span data-ttu-id="979f8-109">使用[存储 Access API，](https://webkit.org/blog/8124/introducing-storage-access-api/)以便第三方可以请求获取访问其第一方   Cookie 的权限。</span><span class="sxs-lookup"><span data-stu-id="979f8-109">Use the [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/) so that the third-party can request permission to get access to its first-party cookies.</span></span> <span data-ttu-id="979f8-110">Mac 和 Office 上当前版本的 Office web 版都支持此 API。</span><span class="sxs-lookup"><span data-stu-id="979f8-110">Current versions of Office on Mac and Office on the web both support this API.</span></span>
    > [!NOTE]
    > <span data-ttu-id="979f8-111">如果你将 Cookie 用于除身份验证外的其他目的，请考虑将 `localStorage` 用于你的方案。</span><span class="sxs-lookup"><span data-stu-id="979f8-111">If you're using cookies for purposes other than authentication, then consider using `localStorage` for your scenario.</span></span>

<span data-ttu-id="979f8-112">下面的代码示例演示如何使用 存储 Access API。</span><span class="sxs-lookup"><span data-stu-id="979f8-112">The following code sample shows how to use the Storage Access API.</span></span>

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

## <a name="about-itp-and-third-party-cookies"></a><span data-ttu-id="979f8-113">关于 ITP 和第三方 Cookie</span><span class="sxs-lookup"><span data-stu-id="979f8-113">About ITP and third-party cookies</span></span>

<span data-ttu-id="979f8-114">第三方 Cookie 是在 iframe 中加载的 Cookie，其中域不同于顶级框架。</span><span class="sxs-lookup"><span data-stu-id="979f8-114">Third-party cookies are cookies that are loaded in an iframe, where the domain is different from the top level frame.</span></span> <span data-ttu-id="979f8-115">ITP 可能会影响复杂的身份验证方案，其中弹出对话框用于输入凭据，然后外接程序 iframe 需要 Cookie 访问才能完成身份验证流。</span><span class="sxs-lookup"><span data-stu-id="979f8-115">ITP could affect complex authentication scenarios, where a popup dialog is used to enter credentials and then the cookie access is needed by an add-in iframe to complete the authentication flow.</span></span> <span data-ttu-id="979f8-116">ITP 还可能会影响静默身份验证方案，其中您之前曾使用弹出对话框进行身份验证，但外接程序的后续使用会尝试通过隐藏的 iframe 进行身份验证。</span><span class="sxs-lookup"><span data-stu-id="979f8-116">ITP could also affect silent authentication scenarios, where you have previously used a popup dialog to authenticate, but subsequent use of the add-in tries to authenticate through a hidden iframe.</span></span>

<span data-ttu-id="979f8-117">在Office Mac 上开发外接程序时，对第三方 Cookie 的访问将被 MacOS Big Sur SDK 阻止。</span><span class="sxs-lookup"><span data-stu-id="979f8-117">When developing Office Add-ins on Mac, access to third-party cookies is blocked by the MacOS Big Sur SDK.</span></span> <span data-ttu-id="979f8-118">这是因为默认情况下，在 Safari 浏览器中启用 WKWebView ITP，并且 WKWebView 会阻止所有第三方 Cookie。</span><span class="sxs-lookup"><span data-stu-id="979f8-118">This is because WKWebView ITP is enabled by default on the Safari browser, and WKWebView blocks all third-party cookies.</span></span> <span data-ttu-id="979f8-119">Office Mac 版本 16.44 或更高版本上的版本与 MacOS 大 Sur SDK 集成。</span><span class="sxs-lookup"><span data-stu-id="979f8-119">Office on Mac version 16.44 or later is integrated with the MacOS Big Sur SDK.</span></span>

<span data-ttu-id="979f8-120">在 Safari 浏览器中，最终用户可以切换首选项隐私下的"阻止 **跨** 网站跟踪"复选框  >  以关闭 ITP。</span><span class="sxs-lookup"><span data-stu-id="979f8-120">In the Safari browser, end users can toggle the **Prevent cross-site tracking** checkbox under **Preference** > **Privacy** to turn off ITP.</span></span> <span data-ttu-id="979f8-121">但是，无法为嵌入的 WKWebView 控件关闭 ITP。</span><span class="sxs-lookup"><span data-stu-id="979f8-121">However, ITP cannot be turned off for the embedded WKWebView control.</span></span>

## <a name="see-also"></a><span data-ttu-id="979f8-122">另请参阅</span><span class="sxs-lookup"><span data-stu-id="979f8-122">See also</span></span>

- [<span data-ttu-id="979f8-123">在 Safari 和其他阻止第三方 Cookie 的浏览器中处理 ITP</span><span class="sxs-lookup"><span data-stu-id="979f8-123">Handle ITP in Safari and other browsers where third-party cookies are blocked</span></span>](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [<span data-ttu-id="979f8-124">WebKit 中的跟踪防护</span><span class="sxs-lookup"><span data-stu-id="979f8-124">Tracking Prevention in WebKit</span></span>](https://webkit.org/tracking-prevention/)
- [<span data-ttu-id="979f8-125">Chrome 的"隐私沙盒"</span><span class="sxs-lookup"><span data-stu-id="979f8-125">Chrome’s “Privacy Sandbox”</span></span>](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [<span data-ttu-id="979f8-126">存储 Access API</span><span class="sxs-lookup"><span data-stu-id="979f8-126">Introducing the Storage Access API</span></span>](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)