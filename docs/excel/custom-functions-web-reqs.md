---
ms.date: 03/21/2019
description: 使用 Excel 中的自定义函数请求、流式处理和取消流式处理工作簿的外部数据
title: 使用自定义函数进行 Web 请求和其他数据处理（预览版）
localization_priority: Priority
ms.openlocfilehash: 9256e2aa87ec6d7b314314a1e4bc2b3793f1df5c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449706"
---
# <a name="receiving-and-handling-data-with-custom-functions"></a><span data-ttu-id="d2a88-103">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="d2a88-103">Receiving and handling data with custom functions</span></span>

<span data-ttu-id="d2a88-104">自定义函数增强 Excel 功能的方法之一是从工作簿以外的位置接收数据，例如 Web 或服务器（通过 WebSockets）。</span><span class="sxs-lookup"><span data-stu-id="d2a88-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="d2a88-105">自定义函数可以通过 XHR 请求数据和获取请求，也可以实时流式处理这些数据。</span><span class="sxs-lookup"><span data-stu-id="d2a88-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

<span data-ttu-id="d2a88-106">下面的文档说明了 Web 请求的一些示例，但是若要为自己构建流式处理函数，请尝试[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows)。</span><span class="sxs-lookup"><span data-stu-id="d2a88-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="d2a88-107">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="d2a88-107">Functions that return data from external sources</span></span>

<span data-ttu-id="d2a88-108">如果自定义函数从外部源（如 Web）检索数据，则必须：</span><span class="sxs-lookup"><span data-stu-id="d2a88-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="d2a88-109">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="d2a88-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="d2a88-110">使用回调函数解析带有最终值的 Promise。</span><span class="sxs-lookup"><span data-stu-id="d2a88-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="d2a88-111">你可以通过 API（如 [`Fetch`](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API)）或使用 `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/zh-CN/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。</span><span class="sxs-lookup"><span data-stu-id="d2a88-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/zh-CN/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="d2a88-112">在自定义函数运行时内，XHR 通过要求[相同来源策略](https://developer.mozilla.org/zh-CN/docs/Web/Security/Same-origin_policy)和简单 [CORS](https://www.w3.org/TR/cors/) 来实施附加安全措施。</span><span class="sxs-lookup"><span data-stu-id="d2a88-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/zh-CN/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="d2a88-113">请注意，简单的 CORS 实施不能使用 cookie，且仅支持简单的方法（GET、HEAD、POST）。</span><span class="sxs-lookup"><span data-stu-id="d2a88-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="d2a88-114">简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。</span><span class="sxs-lookup"><span data-stu-id="d2a88-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="d2a88-115">你还可以在简单 CORS 中使用内容类型标题，前提是内容类型为 `application/x-www-form-urlencoded`、`text/plain` 或 `multipart/form-data`。</span><span class="sxs-lookup"><span data-stu-id="d2a88-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="d2a88-116">XHR 示例</span><span class="sxs-lookup"><span data-stu-id="d2a88-116">XHR example</span></span>

<span data-ttu-id="d2a88-117">在下面的代码示例中，**getTemperature** 函数调用 sendWebRequest 函数，以基于温度计 ID 获取特定区域的温度。</span><span class="sxs-lookup"><span data-stu-id="d2a88-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="d2a88-118">sendWebRequest 函数使用 XHR 来向可以提供相应数据的端点发出 GET 请求。</span><span class="sxs-lookup"><span data-stu-id="d2a88-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```JavaScript
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

        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}

CustomFunctions.associate("GETTEMPERATURE", getTemperature);
```

<span data-ttu-id="d2a88-119">有关具有更多上下文的 XHR 请求的另一个示例，请参阅 [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github 存储库中[此文件](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js)中的 `getFile` 函数。</span><span class="sxs-lookup"><span data-stu-id="d2a88-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="d2a88-120">提取示例</span><span class="sxs-lookup"><span data-stu-id="d2a88-120">Fetch example</span></span>

<span data-ttu-id="d2a88-121">在以下代码示例中，stockPriceStream 函数使用股票代码符号来获取每 1000 毫秒的股票价格。</span><span class="sxs-lookup"><span data-stu-id="d2a88-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="d2a88-122">有关此示例的更多详细信息以及若要获取随附的 JSON，请参阅[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)。</span><span class="sxs-lookup"><span data-stu-id="d2a88-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span> 

```JavaScript
function stockPriceStream(ticker, handler) {
    var updateFrequency = 1000 /* milliseconds*/;
    var isPending = false;

    var timer = setInterval(function() {
        // If there is already a pending request, skip this iteration:
        if (isPending) {
            return;
        }

        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        isPending = true;

        fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                handler.setResult(parseFloat(text));
            })
            .catch(function(error) {
                handler.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    handler.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="d2a88-123">通过 WebSocket 接收数据</span><span class="sxs-lookup"><span data-stu-id="d2a88-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="d2a88-124">在自定义函数内，可使用 WebSocket 来通过与服务器的持久连接交换数据。</span><span class="sxs-lookup"><span data-stu-id="d2a88-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="d2a88-125">通过使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式地轮询服务器来获取数据。</span><span class="sxs-lookup"><span data-stu-id="d2a88-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="d2a88-126">WebSocket 示例</span><span class="sxs-lookup"><span data-stu-id="d2a88-126">WebSockets example</span></span>

<span data-ttu-id="d2a88-127">下面的代码示例建立了一个 WebSocket 连接，然后记录来自服务器的每一条传入消息。</span><span class="sxs-lookup"><span data-stu-id="d2a88-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a><span data-ttu-id="d2a88-128">流式处理函数</span><span class="sxs-lookup"><span data-stu-id="d2a88-128">Streaming functions</span></span>

<span data-ttu-id="d2a88-129">流式处理自定义函数使用户能够在不需要用户显式请求数据刷新的情况下，随着时间的推移向单元格重复输出数据。</span><span class="sxs-lookup"><span data-stu-id="d2a88-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="d2a88-130">以下代码示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="d2a88-130">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="d2a88-131">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="d2a88-131">Note the following about this code:</span></span>

- <span data-ttu-id="d2a88-132">Excel 使用 setResult 回调自动显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="d2a88-132">Excel displays each new value automatically using the setResult callback.</span></span>
- <span data-ttu-id="d2a88-133">当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数 handler。</span><span class="sxs-lookup"><span data-stu-id="d2a88-133">The second input parameter, handler, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="d2a88-134">onCanceled 回调定义取消函数时执行的函数。</span><span class="sxs-lookup"><span data-stu-id="d2a88-134">The onCanceled callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="d2a88-135">对于任何流式处理函数，都必须实现此类取消处理程序。</span><span class="sxs-lookup"><span data-stu-id="d2a88-135">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="d2a88-136">有关详细信息，请参阅[取消函数](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="d2a88-136">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```JavaScript
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}

CustomFunctions.associate("INCREMENTVALUE", incrementValue);
```

<span data-ttu-id="d2a88-137">在 JSON 元数据文件中为流式处理函数指定元数据时，必须在 options 对象中设置属性 "cancelable": true 和 "stream": true，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="d2a88-137">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties "cancelable": true and "stream": true within the options object, as shown in the following example.</span></span>

```JSON
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="d2a88-138">取消函数</span><span class="sxs-lookup"><span data-stu-id="d2a88-138">Canceling a function</span></span>

<span data-ttu-id="d2a88-139">在某些情况下，可能需要取消执行流式处理自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。</span><span class="sxs-lookup"><span data-stu-id="d2a88-139">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="d2a88-140">Excel 会在以下情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="d2a88-140">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="d2a88-141">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="d2a88-141">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="d2a88-142">函数的参数（输入）之一发生变化。</span><span class="sxs-lookup"><span data-stu-id="d2a88-142">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="d2a88-143">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="d2a88-143">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="d2a88-144">用户手动触发重新计算。</span><span class="sxs-lookup"><span data-stu-id="d2a88-144">When the user triggers recalculation manually.</span></span> <span data-ttu-id="d2a88-145">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="d2a88-145">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="d2a88-146">若要使函数可取消，请在函数代码中实施一个处理程序，以告诉它在取消时要执行的操作。</span><span class="sxs-lookup"><span data-stu-id="d2a88-146">To make a function cancelable, implement a handler in your function's code to tell it what to do when it is canceled.</span></span> <span data-ttu-id="d2a88-147">此外，指定描述函数的 JSON 元数据中 options 对象中的属性 `"cancelable": true`。</span><span class="sxs-lookup"><span data-stu-id="d2a88-147">Additionally, specify specify the property `"cancelable": true` within the options object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="d2a88-148">本文前一部分中的代码示例提供了这些方法的示例。</span><span class="sxs-lookup"><span data-stu-id="d2a88-148">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="see-also"></a><span data-ttu-id="d2a88-149">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d2a88-149">See also</span></span>

* [<span data-ttu-id="d2a88-150">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="d2a88-150">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="d2a88-151">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="d2a88-151">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="d2a88-152">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="d2a88-152">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="d2a88-153">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="d2a88-153">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="d2a88-154">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="d2a88-154">Custom functions changelog</span></span>](custom-functions-changelog.md)