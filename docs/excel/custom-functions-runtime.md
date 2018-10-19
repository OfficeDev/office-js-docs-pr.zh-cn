---
ms.date: 10/17/2018
description: 了解开发使用新 JavaScript 运行时的 Excel 自定义函数方面的主要方案。
title: Excel 自定义函数运行时
ms.openlocfilehash: 333816c3916af1490d14b8344c4bb49094f9a7f9
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640013"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="a9d1f-103">Excel 自定义函数的运行时（预览）</span><span class="sxs-lookup"><span data-stu-id="a9d1f-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="a9d1f-p101">自定义函数使用新的JavaScript运行时，该运行时与加载项的其他部分（例如任务窗格或其他UI元素）使用的运行时不同。此JavaScript运行时旨在优化自定义函数中的计算性能，并公开可用于在自定义函数中执行常见的基于Web的操作（如请求外部数据或通过与服务器的持久连接以交换数据）的新API。JavaScript运行时还提供对 `OfficeRuntime`  命名空间中的新 API 的访问，这些 API 可以在自定义函数中使用，也可以由加载项的其他部分用于存储数据或显示对话框。 本文介绍如何在自定义函数中使用这些 API，并概述在开发自定义函数时要记住的其他注意事项。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p101">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements. This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server. The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box. This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="a9d1f-108">请求外部数据</span><span class="sxs-lookup"><span data-stu-id="a9d1f-108">Requesting external data</span></span>

<span data-ttu-id="a9d1f-p102">在自定义函数中，您可以使用 [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) 等 API 或使用 [XmlHttpRequest  (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)  （发布 HTTP 请求以与服务器交互的一个标准 Web API）来请求外部数据。在 JavaScript  运行时中，XHR 通过要求 [同源策略](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) 和简单的 [CORS](https://www.w3.org/TR/cors/) 来实现其他安全措施。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p102">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers. Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="a9d1f-111">XHR 示例</span><span class="sxs-lookup"><span data-stu-id="a9d1f-111">XHR example</span></span>

<span data-ttu-id="a9d1f-p103">在下面的代码示例中， `getTemperature` 函数调用 `sendWebRequest` 函数以根据温度计 ID 获取特定区域的温度。 `sendWebRequest` 函数使用 XHR 向可提供数据的端点发出 `GET` 请求。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p103">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID. The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="a9d1f-p104">当使用 fetch 或 XHR 时，会返回一个新的 JavaScript `Promise` 。在2018年9月之前，您必须指定 `OfficeExtension.Promise` 在 Office  JavaScript API 中使用 promises，但现在您只需使用 JavaScript `Promise` 即可。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p104">When using fetch or XHR, a new JavaScript `Promise` is returned. Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ 
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions  
function sendWebRequest(thermometerID, data) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
           data.temperature = JSON.parse(xhttp.responseText).temperature
        };
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="a9d1f-116">通过 WebSockets 接收数据</span><span class="sxs-lookup"><span data-stu-id="a9d1f-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="a9d1f-p105">在自定义函数中，您可以使用 [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) 通过与服务器的持久连接来交换数据。 通过使用 WebSockets，您的自定义函数可以打开与服务器的连接，然后在发生特定事件时自动从服务器接收消息，而无需显式轮询服务器以获取数据。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p105">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server. By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="a9d1f-119">WebSockets 示例</span><span class="sxs-lookup"><span data-stu-id="a9d1f-119">WebSockets example</span></span>

<span data-ttu-id="a9d1f-120">下面的代码示例建立了一个 `WebSocket` 连接，然后记录来自服务器的每一封传入的邮件。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="a9d1f-121">存储和访问数据</span><span class="sxs-lookup"><span data-stu-id="a9d1f-121">Storing and accessing data</span></span>

<span data-ttu-id="a9d1f-p106">在自定义函数内（或在加载项的任何其他部分内），您可以使用 `OfficeRuntime.AsyncStorage` 对象来存储和访问数据。 `AsyncStorage` 是一个持久的，未加密的键值存储系统，它提供了 [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) 的替代方法，它不能在自定义函数中使用。加载项可以使用 `AsyncStorage` 来存储多达 10 MB 的数据。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p106">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object. `AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used within custom functions. An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="a9d1f-125">下面的方法在 `AsyncStorage` 对象上可用：</span><span class="sxs-lookup"><span data-stu-id="a9d1f-125">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a><span data-ttu-id="a9d1f-126">AsyncStorage 示例</span><span class="sxs-lookup"><span data-stu-id="a9d1f-126">AsyncStorage example</span></span> 

<span data-ttu-id="a9d1f-127">下面的代码示例调用 `AsyncStorage.getItem` 函数以从存储检索值。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-127">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```typescript
_goGetData = async () => {
    try {
        const value = await AsyncStorage.getItem('toDoItem');
        if (value !== null) {
            //data exists and you can do something with it here
        }
    } catch (error) {
        //handle errors here
    }
}
```

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="a9d1f-128">显示对话框</span><span class="sxs-lookup"><span data-stu-id="a9d1f-128">Open a dialog box</span></span>

<span data-ttu-id="a9d1f-p107">在自定义函数内（或在加载项的任何其他部分内），您可以使用 `OfficeRuntime.displayWebDialogOptions`  API 以显示对话框。 此对话框 API 提供了 [Dialog API](../develop/dialog-api-in-office-add-ins.md) 的替代方法，可以在任务窗格和加载项命令中使用，但不能在自定义函数中使用。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p107">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box. This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="a9d1f-131">Dialog API 示例</span><span class="sxs-lookup"><span data-stu-id="a9d1f-131">Dialog API example</span></span> 

<span data-ttu-id="a9d1f-132">在下面的代码示例中，该函数 `getTokenViaDialog` 使用 Dialog API 的 `displayWebDialogOptions` 函数显示一个对话框。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-132">In the following code sample, the `getTokenViaDialog` method uses the Dialog API’s `displayWebDialogOptions` method to open a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
        let timeout = 5;
        let count = 0;
        var intervalId = setInterval(function () {
          count++;
          if(_cachedToken) {
            resolve(_cachedToken);
            clearInterval(intervalId);
          }
          if(count >= timeout) {
            reject("Timeout while waiting for token");
            clearInterval(intervalId);
          }
        }, 1000);
      } else {
        _dialogOpen = true;
        OfficeRuntime.displayWebDialogOptions(url, {
          height: '50%',
          width: '50%',
          onMessage: function (message, dialog) {
            _cachedToken = message;
            resolve(message);
            dialog.closeDialog();
            return;
          },
          onRuntimeError: function(error, dialog) {
            reject(error);
          },
        }).catch(function (e) {
          reject(e);
        });
      }
    });
  }
}
```

## <a name="additional-considerations"></a><span data-ttu-id="a9d1f-133">其他注意事项</span><span class="sxs-lookup"><span data-stu-id="a9d1f-133">Additional considerations</span></span>

<span data-ttu-id="a9d1f-p108">为了创建一个将在多个平台（Office 加载项的关键租户之一）上运行的加载项，您不应访问自定义函数中的文档对象模型（DOM）或使用依赖于 DOM 像 jQuery 这样的库。 在 Excel for Windows 中，自定义函数使用 JavaScript 运行时，自定义函数无法访问 DOM。</span><span class="sxs-lookup"><span data-stu-id="a9d1f-p108">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM. On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="a9d1f-136">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a9d1f-136">See also</span></span>

* [<span data-ttu-id="a9d1f-137">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="a9d1f-137">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a9d1f-138">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="a9d1f-138">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="a9d1f-139">自定义函数最佳做法</span><span class="sxs-lookup"><span data-stu-id="a9d1f-139">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a9d1f-140">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="a9d1f-140">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
