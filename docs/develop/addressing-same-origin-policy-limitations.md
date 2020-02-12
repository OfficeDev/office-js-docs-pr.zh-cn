---
title: 解决 Office 加载项中的同源策略限制
description: ''
ms.date: 10/17/2019
localization_priority: Normal
ms.openlocfilehash: 2a47339bd5cc0b0bf919152b7078d5373382124f
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950444"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>解决 Office 加载项中的同源策略限制

浏览器强制的同源策略可防止从一个域加载的脚本获取或操控来自另一个域的网页的属性。即，默认情况下，请求 URL 的域必须与当前网页的域相同。例如，此策略将阻止一个域中的网页对非托管该网页的域执行 [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web 服务调用。

由于 Office 外接程序在浏览器控件中托管，因此同源策略也适用于在其网页中运行的脚本。

同一来源的策略可能在许多情况下是不必要的障碍，例如当 web 应用程序跨多个子域托管内容和 API 时。 有一些常见技术可以安全解决同一来源策略执行的问题。 本文仅提供有关部分内容的最简洁的介绍。 请使用提供的链接开始对这些技术进行研究。

## <a name="use-jsonp-for-anonymous-access"></a>针对匿名访问使用 JSONP

解决同一来源策略限制的一个方法是使用 [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) 为 web 服务提供代理。 可以通过包括指向任何域上托管的某些脚本的 `script` 标签（带有 `src` 属性）实现此过程。 可以使用编程的方法创建 `script` 标签，动态创建 `src` 属性所指向的 URL，然后通过 URI 查询参数将参数传递给 URL。 Web 服务提供程序在特定的 URL 位置创建和托管 JavaScript 代码，并根据 URI 查询参数返回不同的脚本。 这些脚本然后在插入位置执行并按照预期的方式工作。

下面是使用可在任何 Office 外接程序中工作的技术的 JSONP 示例。

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a>使用基于令牌的授权架构实施服务器端代码

解决同一来源策略限制的另一个方法是提供使用 [OAuth 2.0](https://oauth.net/2/) 流的服务器端代码，让一个域获取对另一个域上托管的资源的授权访问。 


## <a name="use-cross-origin-resource-sharing-cors"></a>使用跨源资源共享 (CORS)


有关使用 [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) 的跨源资源共享功能的示例，请参阅 [XMLHttpRequest2 中的新技巧](https://www.html5rocks.com/en/tutorials/file/xhr2/)的“跨源资源共享 (CORS)”部分。


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a>使用 IFRAME 和 POST MESSAGE 生成您自己的代理（跨 Window 消息传递）。


有关如何使用 IFRAME 和 POST MESSAGE 生成自己代理的示例，请参阅[跨窗口消息传送](http://ejohn.org/blog/cross-window-messaging/)。


## <a name="see-also"></a>另请参阅

- [Office 加载项的隐私和安全](../concepts/privacy-and-security.md)
    
