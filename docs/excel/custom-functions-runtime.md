---
ms.date: 09/20/2018
description: Excel 自定义函数使用新的 JavaScript 运行时，其不同于标准的加载项  WebView 控件运行时。
title: Excel 自定义函数运行时
ms.openlocfilehash: d31002096fccd682c0f2a23a8b43249af5d4df8f
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068812"
---
# <a name="runtime-for-excel-custom-functions"></a><span data-ttu-id="5cbbb-103">Excel 自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="5cbbb-103">Runtime for Excel custom functions</span></span>

<span data-ttu-id="5cbbb-104">自定义函数使用新的 JavaScript 运行时，其使用的是沙盒化的 JavaScript 引擎而不是 web 浏览器，由此扩展了 Excel 的功能。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-104">Custom functions extend Excel’s capabilities by using a new JavaScript runtime that uses a sandboxed JavaScript engine rather than a web browser.</span></span> <span data-ttu-id="5cbbb-105">因为自定义函数不需要呈现 UI 元素，新的 JavaScript 运行时为执行计算进行了优化，让你能够同时运行数千个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-105">Because custom functions do not need to render UI elements, the new JavaScript runtime is optimized for performing calculations, enabling you to run thousands of custom functions simultaneously.</span></span>

## <a name="key-facts-about-the-new-javascript-runtime"></a><span data-ttu-id="5cbbb-106">新的 JavaScript 运行时的有关要点</span><span class="sxs-lookup"><span data-stu-id="5cbbb-106">Key facts about the new JavaScript runtime</span></span> 

<span data-ttu-id="5cbbb-107">只有加载项中的自定义函数将会使用本文中介绍的新 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-107">Only custom functions within an add-in will use the new JavaScript runtime that's described in this article.</span></span> <span data-ttu-id="5cbbb-108">如果加载项包括其他组件，例如任务窗格和其他 UI 元素，除了自定义函数外，加载项的这些其他组件将继续运行在类似于浏览器的 WebView 运行时中。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-108">If an add-in includes other components such as task panes and other UI elements, in addition to custom functions, these other components of the add-in will continue to run in the browser-like WebView runtime.</span></span>  <span data-ttu-id="5cbbb-109">另外：</span><span class="sxs-lookup"><span data-stu-id="5cbbb-109">Additionally:</span></span> 

- <span data-ttu-id="5cbbb-110">JavaScript 运行时不提供对文档对象模型 (DOM) 的访问，也不支持类似于 jQuery 的依赖于 DOM 的库。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-110">The JavaScript runtime does not provide access to the Document Object Model (DOM) or support libraries like jQuery that rely on the DOM.</span></span>

- <span data-ttu-id="5cbbb-111">加载项的 JavaScript 文件中定义的自定义函数可以返回常规 JavaScript `Promise` 而不是返回 `OfficeExtension.Promise`。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-111">A custom function that's defined in an add-in's JavaScript file can return a regular JavaScript `Promise` instead of returning `OfficeExtension.Promise`.</span></span>  

- <span data-ttu-id="5cbbb-112">指定自定义函数元数据的 JSON 文件不需在**选项**中指定**同步**或**异步**。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-112">The JSON file that specifies custom function metatdata does not need to specify **sync** or **async** within **options**.</span></span>

## <a name="new-apis"></a><span data-ttu-id="5cbbb-113">新 API</span><span class="sxs-lookup"><span data-stu-id="5cbbb-113">New and updated APIs</span></span> 

<span data-ttu-id="5cbbb-114">自定义函数使用的 JavaScript 运行时有以下 API：</span><span class="sxs-lookup"><span data-stu-id="5cbbb-114">The JavaScript runtime that's used by custom functions has the following APIs:</span></span>

- [<span data-ttu-id="5cbbb-115">XHR</span><span class="sxs-lookup"><span data-stu-id="5cbbb-115">XHR</span></span>](#xhr)
- [<span data-ttu-id="5cbbb-116">Websocket</span><span class="sxs-lookup"><span data-stu-id="5cbbb-116">WebSockets</span></span>](#websockets)
- [<span data-ttu-id="5cbbb-117">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="5cbbb-117">AsyncStorage</span></span>](#asyncstorage)
- [<span data-ttu-id="5cbbb-118">Dialog API</span><span class="sxs-lookup"><span data-stu-id="5cbbb-118">Dialog API requirement sets</span></span>](#dialog-api)

### <a name="xhr"></a><span data-ttu-id="5cbbb-119">XHR</span><span class="sxs-lookup"><span data-stu-id="5cbbb-119">XHR</span></span>

<span data-ttu-id="5cbbb-120">XHR 代表 [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)，这是一个发出 HTTP 请求以便与服务器进行交互的标准 web API。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-120">XHR stands for [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span> <span data-ttu-id="5cbbb-121">在新的 JavaScript 运行时中，XHR 通过要求[同源策略](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)和简单 [CORS](https://www.w3.org/TR/cors/)实现附加安全措施。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-121">In the new JavaScript runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>  

<span data-ttu-id="5cbbb-122">在下面的代码示例中，`getTemperature()` 函数发出 web 请求，以获取基于温度计 ID 的特定区域温度。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-122">In the following code sample, the `getTemperature()` function sends a web request to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="5cbbb-123"> `sendWebRequest()` 函数使用 XHR 向可提供数据的端点发出 `GET` 请求。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-123">The `sendWebRequest()` function uses XHR to issue a `GET` request to an endpoint that can provide the data.</span></span>  

```js
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){ //sendWebRequest is defined later in this code sample
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

//Helper method that uses Office's implementation of XMLHttpRequest in the new JavaScript runtime for custom functions  
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

### <a name="websockets"></a><span data-ttu-id="5cbbb-124">Websocket</span><span class="sxs-lookup"><span data-stu-id="5cbbb-124">WebSockets</span></span>

<span data-ttu-id="5cbbb-125">[Websocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) 是在服务器与一个或多个客户端之间建立实时通信的网络协议。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-125">[WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) is a networking protocol that creates real-time communication between a server and one or more clients.</span></span> <span data-ttu-id="5cbbb-126">它通常用于聊天应用程序，因为它允许同时读取和写入文本。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-126">It is often used for chat applications because it allows you to read and write text simultaneously.</span></span>  

<span data-ttu-id="5cbbb-127">如下面的代码示例中所示，自定义函数可以使用 Websocket。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-127">As shown in the following code sample, custom functions can use WebSockets.</span></span> <span data-ttu-id="5cbbb-128">本示例中，WebSocket 记录其接收的每条消息。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-128">In this example, the WebSocket logs each message that it receives.</span></span>

```ts
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a><span data-ttu-id="5cbbb-129">AsyncStorage</span><span class="sxs-lookup"><span data-stu-id="5cbbb-129">AsyncStorage</span></span>

<span data-ttu-id="5cbbb-130">AsyncStorage 是可用于存储身份验证令牌的键值存储系统。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-130">AsyncStorage is a key-value storage system that can be used to store authentication tokens.</span></span> <span data-ttu-id="5cbbb-131">它特点是：</span><span class="sxs-lookup"><span data-stu-id="5cbbb-131">It is framework-agnostic.</span></span>

- <span data-ttu-id="5cbbb-132">持续</span><span class="sxs-lookup"><span data-stu-id="5cbbb-132">persistent</span></span>
- <span data-ttu-id="5cbbb-133">不加密</span><span class="sxs-lookup"><span data-stu-id="5cbbb-133">Unencrypted</span></span>
- <span data-ttu-id="5cbbb-134">异步</span><span class="sxs-lookup"><span data-stu-id="5cbbb-134">Asynchronous calls</span></span>

<span data-ttu-id="5cbbb-135">AsyncStorage 对于加载项的所有部件全局可用。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-135">AsyncStorage is globally available to all parts of your add-in.</span></span> <span data-ttu-id="5cbbb-136">对于自定义函数，`AsyncStorage` 作为全局对象公开。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-136">For custom functions, `AsyncStorage` is exposed as a global object.</span></span> <span data-ttu-id="5cbbb-137">（对于加载项的其他部件，如任务窗格和使用 WebView 运行时的其他元素，AsyncStorage 通过 `OfficeRuntime` 公开。）每一加载项都有自己的存储分区，默认大小为 5 MB。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-137">(For other parts of your add-in, such as task panes and other elements that use the WebView runtime, AsyncStorage is exposed through `OfficeRuntime`.) Each add-in has its own storage partition, with a default size of 5MB.</span></span> 

<span data-ttu-id="5cbbb-138">下面的方法在 `AsyncStorage` 对象上可用：</span><span class="sxs-lookup"><span data-stu-id="5cbbb-138">The following methods are available on the `AsyncStorage` object:</span></span>
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`
 
<span data-ttu-id="5cbbb-139">`mergeItem` 和 `multiMerge` 方法现时不受支持。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-139">At this time, the `mergeItem` and `multiMerge` methods are not supported.</span></span>

<span data-ttu-id="5cbbb-140">下面的代码示例调用 `AsyncStorage.getItem` 函数以从存储检索值。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-140">The following code sample calls the `AsyncStorage.getItem` function to retrieve a value from storage.</span></span>

```js
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
}
```

### <a name="dialog-api"></a><span data-ttu-id="5cbbb-141">Dialog API</span><span class="sxs-lookup"><span data-stu-id="5cbbb-141">Dialog API scenarios</span></span>

<span data-ttu-id="5cbbb-142">Dialog API 可打开一个提示用户登录的对话框。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-142">The Dialog API enables you to open a dialog box that prompts user sign-in.</span></span> <span data-ttu-id="5cbbb-143"> 可使用 Dialog API 要求通过Google 或 Facebook 等外部资源进行用户身份验证，此后用户才可以使用你的函数。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-143">You can use the Dialog API to require user authentication through an outside resource, such as Google or Facebook, before the user can use your function.</span></span>   

<span data-ttu-id="5cbbb-144">在下面的代码示例中，`getTokenViaDialog()` 方法使用 Dialog API 的 `displayWebDialog()` 方法打开一个对话框。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-144">In the following code sample, the `getTokenViaDialog()` method uses the Dialog API’s `displayWebDialog()` method to open a dialog box.</span></span>

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"
 
function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://myauthurl")
    .then(function (token) {
      
      // Use token to get stock price
      fetch("https://myservice.com/?token=token&ticker= + ticker")
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
        OfficeRuntime.displayWebDialog(url, {
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

> [!NOTE]
> <span data-ttu-id="5cbbb-145">本节中所述的 Dialog AP 是用于自定义函数的新 JavaScript 运行时的一部分，仅可在自定义的函数中使用。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-145">The Dialog API described in this section is part of the new JavaScript runtime for custom functions and can be used only within custom functions.</span></span> <span data-ttu-id="5cbbb-146">此 API 不同于可在任务窗格和加载项命令中使用的 [Dialog API](../develop/dialog-api-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="5cbbb-146">This API is different from the [Dialog API](../develop/dialog-api-in-office-add-ins.md) that can be used within task panes and add-in commands.</span></span>

## <a name="see-also"></a><span data-ttu-id="5cbbb-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5cbbb-147">See also</span></span>

* [<span data-ttu-id="5cbbb-148">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="5cbbb-148">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="5cbbb-149">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="5cbbb-149">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="5cbbb-150">自定义函数的最佳实践</span><span class="sxs-lookup"><span data-stu-id="5cbbb-150">Custom functions best practices</span></span>](custom-functions-best-practices.md)