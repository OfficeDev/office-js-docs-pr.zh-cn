---
title: 解决 Office 加载项中的同源策略限制
description: 了解如何使用 JSONP、CORS、Iframe 和其他技术来适应相同来源的策略限制。
ms.date: 10/17/2019
localization_priority: Normal
ms.openlocfilehash: e50292c30d77856c896f892c930038c1e19d7af7
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293336"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="b7935-103">解决 Office 加载项中的同源策略限制</span><span class="sxs-lookup"><span data-stu-id="b7935-103">Addressing same-origin policy limitations in Office Add-ins</span></span>

<span data-ttu-id="b7935-p101">浏览器强制的同源策略可防止从一个域加载的脚本获取或操控来自另一个域的网页的属性。即，默认情况下，请求 URL 的域必须与当前网页的域相同。例如，此策略将阻止一个域中的网页对非托管该网页的域执行 [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web 服务调用。</span><span class="sxs-lookup"><span data-stu-id="b7935-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="b7935-107">由于 Office 外接程序在浏览器控件中托管，因此同源策略也适用于在其网页中运行的脚本。</span><span class="sxs-lookup"><span data-stu-id="b7935-107">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="b7935-108">同一来源的策略可能在许多情况下是不必要的障碍，例如当 web 应用程序跨多个子域托管内容和 API 时。</span><span class="sxs-lookup"><span data-stu-id="b7935-108">The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains.</span></span> <span data-ttu-id="b7935-109">有一些常见技术可以安全解决同一来源策略执行的问题。</span><span class="sxs-lookup"><span data-stu-id="b7935-109">There are a few common techniques for securely overcoming same-origin policy enforcement.</span></span> <span data-ttu-id="b7935-110">本文仅提供有关部分内容的最简洁的介绍。</span><span class="sxs-lookup"><span data-stu-id="b7935-110">This article can only provide the briefest introduction to some of them.</span></span> <span data-ttu-id="b7935-111">请使用提供的链接开始对这些技术进行研究。</span><span class="sxs-lookup"><span data-stu-id="b7935-111">Please use the links provided to get started in your research of these techniques.</span></span>

## <a name="use-jsonp-for-anonymous-access"></a><span data-ttu-id="b7935-112">针对匿名访问使用 JSONP</span><span class="sxs-lookup"><span data-stu-id="b7935-112">Use JSONP for anonymous access</span></span>

<span data-ttu-id="b7935-113">解决同一来源策略限制的一个方法是使用 [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) 为 web 服务提供代理。</span><span class="sxs-lookup"><span data-stu-id="b7935-113">One way to overcome same-origin policy limitations is to use [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) to provide a proxy for the web service.</span></span> <span data-ttu-id="b7935-114">可以通过包括指向任何域上托管的某些脚本的 `script` 标签（带有 `src` 属性）实现此过程。</span><span class="sxs-lookup"><span data-stu-id="b7935-114">You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain.</span></span> <span data-ttu-id="b7935-115">可以使用编程的方法创建 `script` 标签，动态创建 `src` 属性所指向的 URL，然后通过 URI 查询参数将参数传递给 URL。</span><span class="sxs-lookup"><span data-stu-id="b7935-115">You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters.</span></span> <span data-ttu-id="b7935-116">Web 服务提供程序在特定的 URL 位置创建和托管 JavaScript 代码，并根据 URI 查询参数返回不同的脚本。</span><span class="sxs-lookup"><span data-stu-id="b7935-116">Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters.</span></span> <span data-ttu-id="b7935-117">这些脚本然后在插入位置执行并按照预期的方式工作。</span><span class="sxs-lookup"><span data-stu-id="b7935-117">These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="b7935-118">下面是使用可在任何 Office 外接程序中工作的技术的 JSONP 示例。</span><span class="sxs-lookup"><span data-stu-id="b7935-118">The following is an example of JSONP that uses a technique that will work in any Office Add-in.</span></span>

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a><span data-ttu-id="b7935-119">使用基于令牌的授权架构实施服务器端代码</span><span class="sxs-lookup"><span data-stu-id="b7935-119">Implement server-side code using a token-based authorization scheme</span></span>

<span data-ttu-id="b7935-120">解决同一来源策略限制的另一个方法是提供使用 [OAuth 2.0](https://oauth.net/2/) 流的服务器端代码，让一个域获取对另一个域上托管的资源的授权访问。</span><span class="sxs-lookup"><span data-stu-id="b7935-120">Another way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.</span></span> 


## <a name="use-cross-origin-resource-sharing-cors"></a><span data-ttu-id="b7935-121">使用跨源资源共享 (CORS)</span><span class="sxs-lookup"><span data-stu-id="b7935-121">Use cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="b7935-122">有关使用 [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) 的跨源资源共享功能的示例，请参阅 [XMLHttpRequest2 中的新技巧](https://www.html5rocks.com/en/tutorials/file/xhr2/)的“跨源资源共享 (CORS)”部分。</span><span class="sxs-lookup"><span data-stu-id="b7935-122">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a><span data-ttu-id="b7935-123">使用 IFRAME 和 POST MESSAGE 生成您自己的代理（跨 Window 消息传递）。</span><span class="sxs-lookup"><span data-stu-id="b7935-123">Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)</span></span>


<span data-ttu-id="b7935-124">有关如何使用 IFRAME 和 POST MESSAGE 生成自己代理的示例，请参阅[跨窗口消息传送](http://ejohn.org/blog/cross-window-messaging/)。</span><span class="sxs-lookup"><span data-stu-id="b7935-124">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="b7935-125">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b7935-125">See also</span></span>

- [<span data-ttu-id="b7935-126">Office 加载项的隐私和安全</span><span class="sxs-lookup"><span data-stu-id="b7935-126">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
