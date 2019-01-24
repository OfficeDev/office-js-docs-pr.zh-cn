---
title: 解决 Office 加载项中的同源策略限制
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 75bc42cd7d2a7acc8cb57ee08807a8486e21f467
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387753"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="27b5d-102">解决 Office 加载项中的同源策略限制</span><span class="sxs-lookup"><span data-stu-id="27b5d-102">Addressing same-origin policy limitations in Office Add-ins</span></span>


<span data-ttu-id="27b5d-p101">浏览器强制的同源策略可防止从一个域加载的脚本获取或操控来自另一个域的网页的属性。即，默认情况下，请求 URL 的域必须与当前网页的域相同。例如，此策略将阻止一个域中的网页对非托管该网页的域执行 [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web 服务调用。</span><span class="sxs-lookup"><span data-stu-id="27b5d-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="27b5d-106">由于 Office 外接程序在浏览器控件中托管，因此同源策略也适用于在其网页中运行的脚本。</span><span class="sxs-lookup"><span data-stu-id="27b5d-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="27b5d-107">开发加载项时，要解决同源策略强制，您可以执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="27b5d-107">To overcome same-origin policy enforcement when you develop add-ins, you can:</span></span>

- <span data-ttu-id="27b5d-108">针对匿名访问使用 JSON/P。</span><span class="sxs-lookup"><span data-stu-id="27b5d-108">Use JSON/P for anonymous access.</span></span> 
    
- <span data-ttu-id="27b5d-109">使用基于令牌的身份验证架构实施服务器端脚本。</span><span class="sxs-lookup"><span data-stu-id="27b5d-109">Implement server-side script using a token-based authentication scheme.</span></span>
    
- <span data-ttu-id="27b5d-110">使用跨源资源共享 (CORS)。</span><span class="sxs-lookup"><span data-stu-id="27b5d-110">Using cross-origin resource sharing (CORS).</span></span>
    
- <span data-ttu-id="27b5d-111">使用 IFRAME 和 POST MESSAGE 生成您自己的代理。</span><span class="sxs-lookup"><span data-stu-id="27b5d-111">Build your own proxy using IFRAME and POST MESSAGE.</span></span>
    

## <a name="using-jsonp-for-anonymous-access"></a><span data-ttu-id="27b5d-112">针对匿名访问使用 JSON/P</span><span class="sxs-lookup"><span data-stu-id="27b5d-112">Using JSON/P for anonymous access</span></span>


<span data-ttu-id="27b5d-p102">解决此限制的一个方法是使用 JSON/P 提供 Web 服务的代理。可以通过包括指向任何域上托管的某些脚本的 `script` 标签（带有 `src` 属性）实现此过程。可以使用编程的方法创建 `script` 标签，动态创建 `src` 属性所指向的 URL，然后通过 URI 查询参数将参数传递给 URL。Web 服务提供程序在特定的 URL 位置创建和托管 JavaScript 代码，并根据 URI 查询参数返回不同的脚本。这些脚本然后在插入位置执行并按照预期的方式工作。</span><span class="sxs-lookup"><span data-stu-id="27b5d-p102">One way to overcome this limitation is to use JSON/P to provide a proxy for the web service. You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain. You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="27b5d-118">下面是使用可在任何 Office 外接程序中工作的技术的 JSON/P 示例。</span><span class="sxs-lookup"><span data-stu-id="27b5d-118">The following is an example of JSON/P that uses a technique that will work in any Office Add-in.</span></span>

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a><span data-ttu-id="27b5d-119">使用基于令牌的身份验证架构实施服务器端脚本</span><span class="sxs-lookup"><span data-stu-id="27b5d-119">Implementing server-side script using a token-based authentication scheme</span></span>


<span data-ttu-id="27b5d-120">解决同源策略限制的另一个方法是将加载项网页作为在 Cookie 中使用 OAuth 或缓存凭据的 ASP 页来实施。</span><span class="sxs-lookup"><span data-stu-id="27b5d-120">Another way to address same-origin policy limitations is to implement the add-in's webpage as an ASP page that uses OAuth or caches credentials in cookies.</span></span>

<span data-ttu-id="27b5d-121">有关演示如何使用 `System.Net` 中的 `Cookie` 对象获取和设置 cookie 值的服务器端代码示例，请参阅 [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) 属性。</span><span class="sxs-lookup"><span data-stu-id="27b5d-121">For an example of server-side code that shows how to use the  `Cookie` object in `System.Net` to get and set cookie values, see the [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) property.</span></span>


## <a name="using-cross-origin-resource-sharing-cors"></a><span data-ttu-id="27b5d-122">使用跨源资源共享 (CORS)</span><span class="sxs-lookup"><span data-stu-id="27b5d-122">Using cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="27b5d-123">有关使用 [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) 的跨源资源共享功能的示例，请参阅 [XMLHttpRequest2 中的新技巧](https://www.html5rocks.com/en/tutorials/file/xhr2/)的“跨源资源共享 (CORS)”部分。</span><span class="sxs-lookup"><span data-stu-id="27b5d-123">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a><span data-ttu-id="27b5d-124">使用 IFRAME 和 POST MESSAGE 生成您自己的代理</span><span class="sxs-lookup"><span data-stu-id="27b5d-124">Building your own proxy using IFRAME and POST MESSAGE</span></span>


<span data-ttu-id="27b5d-125">有关如何使用 IFRAME 和 POST MESSAGE 生成自己代理的示例，请参阅[跨窗口消息传送](http://ejohn.org/blog/cross-window-messaging/)。</span><span class="sxs-lookup"><span data-stu-id="27b5d-125">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="27b5d-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="27b5d-126">See also</span></span>

- [<span data-ttu-id="27b5d-127">Office 加载项的隐私和安全</span><span class="sxs-lookup"><span data-stu-id="27b5d-127">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
