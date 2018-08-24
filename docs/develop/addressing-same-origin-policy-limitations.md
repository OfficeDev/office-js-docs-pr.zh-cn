---
title: 解决 Office 加载项中的同源策略限制
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 536e02d2367bef81d4a6e49098d66833c99f5e50
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925106"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>解决 Office 加载项中的同源策略限制


浏览器强制的同源策略可防止从一个域加载的脚本获取或操控来自另一个域的网页的属性。即，默认情况下，请求 URL 的域必须与当前网页的域相同。例如，此策略将阻止一个域中的网页对非托管该网页的域执行 [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) Web 服务调用。

由于 Office 外接程序在浏览器控件中托管，因此同源策略也适用于在其网页中运行的脚本。

开发加载项时，要解决同源策略强制，您可以执行以下操作：

- 针对匿名访问使用 JSON/P。 
    
- 使用基于令牌的身份验证架构实施服务器端脚本。
    
- 使用跨源资源共享 (CORS)。
    
- 使用 IFRAME 和 POST MESSAGE 生成您自己的代理。
    

## <a name="using-jsonp-for-anonymous-access"></a>针对匿名访问使用 JSON/P


解决此限制的一个方法是使用 JSON/P 提供 Web 服务的代理。可以通过包括指向任何域上托管的某些脚本的 `script` 标签（带有 `src` 属性）实现此过程。可以使用编程的方法创建 `script` 标签，动态创建 `src` 属性所指向的 URL，然后通过 URI 查询参数将参数传递给 URL。Web 服务提供程序在特定的 URL 位置创建和托管 JavaScript 代码，并根据 URI 查询参数返回不同的脚本。这些脚本然后在插入位置执行并按照预期的方式工作。

下面是使用可在任何 Office 外接程序中工作的技术的 JSON/P 示例。

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


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a>使用基于令牌的身份验证架构实施服务器端脚本


解决同源策略限制的另一个方法是将加载项网页作为在 Cookie 中使用 OAuth 或缓存凭据的 ASP 页来实施。

有关演示如何使用 `System.Net` 中的 `Cookie` 对象获取和设置 cookie 值的服务器端代码示例，请参阅 [Value](https://msdn.microsoft.com/library/4f772twc)(#value) 属性。


## <a name="using-cross-origin-resource-sharing-cors"></a>使用跨源资源共享 (CORS)


有关使用 [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) 的跨源资源共享功能的示例，请参阅 [XMLHttpRequest2 中的新技巧](http://www.html5rocks.com/en/tutorials/file/xhr2/)的“跨源资源共享 (CORS)”部分。


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a>使用 IFRAME 和 POST MESSAGE 生成您自己的代理


有关如何使用 IFRAME 和 POST MESSAGE 生成自己代理的示例，请参阅[跨窗口消息传送](http://ejohn.org/blog/cross-window-messaging/)。


## <a name="see-also"></a>另请参阅

- [Office 加载项的隐私和安全](../concepts/privacy-and-security.md)
    
