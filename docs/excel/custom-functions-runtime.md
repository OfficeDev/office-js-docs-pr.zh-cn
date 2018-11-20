---
ms.date: 10/17/2018
description: 了解开发使用新 JavaScript 运行时的 Excel 自定义函数时的关键方案。
title: Excel 自定义函数的运行时
ms.openlocfilehash: 7261eb214e97a2a05e08571a31ac9b449df14083
ms.sourcegitcommit: 161a0625646a8c2ebaf1773c6369ee7cc96aa07b
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/01/2018
ms.locfileid: "25891709"
---
# <a name="runtime-for-excel-custom-functions-preview"></a><span data-ttu-id="f77a8-103">Excel 自定义函数的运行时（预览）</span><span class="sxs-lookup"><span data-stu-id="f77a8-103">Runtime for Excel custom functions (preview)</span></span>

<span data-ttu-id="f77a8-104">自定义函数使用的是与外接程序的其他部分（例如任务窗格或其他 UI 元素）使用的运行时不同的新 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="f77a8-104">Custom functions use a new JavaScript runtime that differs from the runtime used by other parts of an add-in, such as the task pane or other UI elements.</span></span> <span data-ttu-id="f77a8-105">这种 JavaScript 运行时旨在优化自定义函数中的计算性能，并支持可用于在自定义函数中执行常见 Web 操作（例如请求外部数据或通过与服务器的持久连接交换数据）的新 API。</span><span class="sxs-lookup"><span data-stu-id="f77a8-105">This JavaScript runtime is designed to optimize performance of calculations in custom functions and exposes new APIs that you can use to perform common web-based actions within custom functions such as requesting external data or exchanging data over a persistent connection with a server.</span></span> <span data-ttu-id="f77a8-106">JavaScript 运行时还可提供对 `OfficeRuntime` 命名空间内的新 API 的访问，这些 API 可在自定义函数内或由外接程序的其他部分使用，用于存储数据或显示对话框。</span><span class="sxs-lookup"><span data-stu-id="f77a8-106">The JavaScript runtime also provides access to new APIs in the `OfficeRuntime` namespace that can be used within custom functions or by other parts of an add-in to store data or display a dialog box.</span></span> <span data-ttu-id="f77a8-107">本文介绍了如何在自定义函数内使用这些 API，还概述了在开发自定义函数时需要牢记的其他注意事项。</span><span class="sxs-lookup"><span data-stu-id="f77a8-107">This article describes how to use these APIs within custom functions and also outlines additional considerations to keep in mind as you develop custom functions.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a><span data-ttu-id="f77a8-108">请求外部数据</span><span class="sxs-lookup"><span data-stu-id="f77a8-108">Requesting external data</span></span>

<span data-ttu-id="f77a8-109">在自定义函数中，你可以使用 [Fetch](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API) 等 API 或使用 [XmlHttpRequest (XHR)](https://developer.mozilla.org/zh-CN/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。</span><span class="sxs-lookup"><span data-stu-id="f77a8-109">Within a custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/zh-CN/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="f77a8-110">在 JavaScript runtime 运行时内，XHR 通过要求[相同来源策略](https://developer.mozilla.org/zh-CN/docs/Web/Security/Same-origin_policy)和简单 [CORS](https://www.w3.org/TR/cors/) 来实施附加安全措施。</span><span class="sxs-lookup"><span data-stu-id="f77a8-110">Within the JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/zh-CN/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

### <a name="xhr-example"></a><span data-ttu-id="f77a8-111">XHR 示例</span><span class="sxs-lookup"><span data-stu-id="f77a8-111">XHR example</span></span>

<span data-ttu-id="f77a8-112">在下面的代码示例中，`getTemperature` 函数调用 `sendWebRequest` 函数，以基于温度计 ID 获取特定区域的温度。</span><span class="sxs-lookup"><span data-stu-id="f77a8-112">In the following code sample, the `getTemperature` function calls the `sendWebRequest` function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="f77a8-113">`sendWebRequest` 函数使用 XHR 来向可以提供相应数据的端点发出 `GET` 请求。</span><span class="sxs-lookup"><span data-stu-id="f77a8-113">The `sendWebRequest` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span> 

> [!NOTE] 
> <span data-ttu-id="f77a8-114">当使用提取或 XHR 时，将返回新的 JavaScript `Promise`。</span><span class="sxs-lookup"><span data-stu-id="f77a8-114">When using fetch or XHR, a new JavaScript `Promise` is returned.</span></span> <span data-ttu-id="f77a8-115">在 2018 年 9 月之前，必须指定 `OfficeExtension.Promise` 才能在 Office JavaScript API 中使用 promise，但现在可以直接使用 JavaScript `Promise`。</span><span class="sxs-lookup"><span data-stu-id="f77a8-115">Prior to September 2018, you had to specify `OfficeExtension.Promise` to use promises within the Office JavaScript API, but now you can simply use a JavaScript `Promise`.</span></span>

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

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="f77a8-116">通过 WebSocket 接收数据</span><span class="sxs-lookup"><span data-stu-id="f77a8-116">Receiving data via WebSockets</span></span>

<span data-ttu-id="f77a8-117">在自定义函数内，可使用 [WebSocket](https://developer.mozilla.org/zh-CN/docs/Web/API/WebSockets_API) 来通过与服务器的持久连接交换数据。</span><span class="sxs-lookup"><span data-stu-id="f77a8-117">Within a custom function, you can use [WebSockets](https://developer.mozilla.org/zh-CN/docs/Web/API/WebSockets_API) to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="f77a8-118">通过使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式地轮询服务器来获取数据。</span><span class="sxs-lookup"><span data-stu-id="f77a8-118">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="f77a8-119">WebSocket 示例</span><span class="sxs-lookup"><span data-stu-id="f77a8-119">WebSockets example</span></span>

<span data-ttu-id="f77a8-120">下面的代码示例建立了一个 `WebSocket` 连接，然后记录来自服务器的每一条传入消息。</span><span class="sxs-lookup"><span data-stu-id="f77a8-120">The following code sample establishes a `WebSocket` connection and then logs each incoming message from the server.</span></span> 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a><span data-ttu-id="f77a8-121">存储和访问数据</span><span class="sxs-lookup"><span data-stu-id="f77a8-121">Storing and accessing data</span></span>

<span data-ttu-id="f77a8-122">在自定义函数（或外接程序的任何其他部分）内，可以使用 `OfficeRuntime.AsyncStorage` 对象来存储和访问数据。</span><span class="sxs-lookup"><span data-stu-id="f77a8-122">Within a custom function (or within any other part of an add-in), you can store and access data by using the `OfficeRuntime.AsyncStorage` object.</span></span> <span data-ttu-id="f77a8-123">`AsyncStorage` 是一种未加密的持久键值存储系统，为无法在自定义函数内使用的 [localStorage](https://developer.mozilla.org/zh-CN/docs/Web/API/Window/localStorage) 提供了一种替代方案。</span><span class="sxs-lookup"><span data-stu-id="f77a8-123">`AsyncStorage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/zh-CN/docs/Web/API/Window/localStorage), which cannot be used within custom functions.</span></span> <span data-ttu-id="f77a8-124">一个外接程序可以使用 `AsyncStorage` 存储最多 10 MB 的数据。</span><span class="sxs-lookup"><span data-stu-id="f77a8-124">An add-in can store up to 10 MB of data using `AsyncStorage`.</span></span>

<span data-ttu-id="f77a8-125">`AsyncStorage` 旨在作为共享存储解决方案，这意味着外接程序的多个部分将能访问相同数据。</span><span class="sxs-lookup"><span data-stu-id="f77a8-125">`AsyncStorage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="f77a8-126">例如，用户身份验证令牌可能存储在 `AsyncStorage` 中，因为自定义函数和任务窗格等外接程序 UI 元素都有可能会访问该数据。</span><span class="sxs-lookup"><span data-stu-id="f77a8-126">For example, tokens for user authentication may be stored in `AsyncStorage` because it can be accessed by both a custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="f77a8-127">同样，如果两个外接程序共享同一个域（例如 www.contoso.com/addin1、www.contoso.com/addin2），则也可以通过 `AsyncStorage` 来回共享信息。</span><span class="sxs-lookup"><span data-stu-id="f77a8-127">Similarly, if two add-ins share the same domain (e.g. www.contoso.com/addin1, www.contoso.com/addin2), they are also permitted to share information back and forth through `AsyncStorage`.</span></span> <span data-ttu-id="f77a8-128">注意，具有不同子域的外接程序将具有不同的 `AsyncStorage` 实例（例如 subdomain.contoso.com/addin1、differentsubdomain.contoso.com/addin2）</span><span class="sxs-lookup"><span data-stu-id="f77a8-128">Note that add-ins which have different subdomains will have different instances of `AsyncStorage` (e.g. subdomain.contoso.com/addin1, differentsubdomain.contoso.com/addin2).</span></span> 

<span data-ttu-id="f77a8-129">由于 `AsyncStorage` 可能是共享的位置，因此一定要认识到，可能会存在替代键值对的情况。</span><span class="sxs-lookup"><span data-stu-id="f77a8-129">Because `AsyncStorage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="f77a8-130">`AsyncStorage` 对象支持以下方法：</span><span class="sxs-lookup"><span data-stu-id="f77a8-130">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - <span data-ttu-id="f77a8-131">`multiRemove`：你会注意到，我们没有实施可用于清除所有信息的方法（例如 `clear`）。</span><span class="sxs-lookup"><span data-stu-id="f77a8-131">`multiRemove`: You will note that there is no implementation of a method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="f77a8-132">相反，需要使用 `multiRemove` 来一次性删除多个条目。</span><span class="sxs-lookup"><span data-stu-id="f77a8-132">Instead, you should instead use `multiRemove` to remove multiple entries at a time.</span></span>

### <a name="asyncstorage-example"></a><span data-ttu-id="f77a8-133">AsyncStorage 示例</span><span class="sxs-lookup"><span data-stu-id="f77a8-133">AsyncStorage example</span></span> 

<span data-ttu-id="f77a8-134">下面的代码示例调用 `AsyncStorage.getItem` 函数来从存储器中检索值。</span><span class="sxs-lookup"><span data-stu-id="f77a8-134">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

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

## <a name="displaying-a-dialog-box"></a><span data-ttu-id="f77a8-135">显示对话框</span><span class="sxs-lookup"><span data-stu-id="f77a8-135">Displaying a dialog box</span></span>

<span data-ttu-id="f77a8-136">在自定义函数（或外接程序的任何其他部分）内，可以使用 `OfficeRuntime.displayWebDialogOptions` API 来显示对话框。</span><span class="sxs-lookup"><span data-stu-id="f77a8-136">Within a custom function (or within any other part of an add-in), you can use the `OfficeRuntime.displayWebDialogOptions` API to display a dialog box.</span></span> <span data-ttu-id="f77a8-137">此对话框 API 为可在任务窗格和附加程序命令内使用但无法在自定义函数中使用的[对话框 API](../develop/dialog-api-in-office-add-ins.md) 提供了一个替代方案。</span><span class="sxs-lookup"><span data-stu-id="f77a8-137">This dialog API provides an alternative to the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands, but not within custom functions.</span></span>

### <a name="dialog-api-example"></a><span data-ttu-id="f77a8-138">对话框 API 示例</span><span class="sxs-lookup"><span data-stu-id="f77a8-138">Dialog API example</span></span> 

<span data-ttu-id="f77a8-139">在下面的代码示例中，函数 `getTokenViaDialog` 使用对话框 API 的 `displayWebDialogOptions` 函数来显示对话框。</span><span class="sxs-lookup"><span data-stu-id="f77a8-139">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialogOptions` function to display a dialog box.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="f77a8-140">其他注意事项</span><span class="sxs-lookup"><span data-stu-id="f77a8-140">Additional considerations</span></span>

<span data-ttu-id="f77a8-141">为创建一个可在多个平台（Office 外接程序的关键租户之一）上运行的外接程序，请勿访问自定义函数中的文档对象模型 (DOM) 或使用 jQuery 等这类依赖于 DOM 的库。</span><span class="sxs-lookup"><span data-stu-id="f77a8-141">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="f77a8-142">在自定义函数会使用 JavaScript 运行时的 Windows 版 Excel 中，自定义函数无法访问 DOM。</span><span class="sxs-lookup"><span data-stu-id="f77a8-142">On Excel for Windows, where custom functions use the JavaScript runtime, custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="f77a8-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f77a8-143">See also</span></span>

* [<span data-ttu-id="f77a8-144">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="f77a8-144">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f77a8-145">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="f77a8-145">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f77a8-146">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="f77a8-146">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="f77a8-147">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="f77a8-147">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
