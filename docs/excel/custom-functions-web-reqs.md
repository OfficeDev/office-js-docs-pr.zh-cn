---
ms.date: 05/07/2019
description: 使用 Excel 中的自定义函数请求、流式处理和取消流式处理工作簿的外部数据
title: 使用自定义函数接收和处理数据
localization_priority: Priority
ms.openlocfilehash: 61f4d0fdaea4277faedddbe075a587fb23842c08
ms.sourcegitcommit: 5b9c2b39dfe76cabd98bf28d5287d9718788e520
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/07/2019
ms.locfileid: "33659633"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="bbd27-103">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="bbd27-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="bbd27-104">自定义函数增强 Excel 功能的方法之一是从工作簿以外的位置接收数据，例如 Web 或服务器（通过 WebSockets）。</span><span class="sxs-lookup"><span data-stu-id="bbd27-104">One of the ways that custom functions enhance Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="bbd27-105">自定义函数可以通过 XHR 和 `fetch` 请求来请求数据，也可以实时流式处理这些数据。</span><span class="sxs-lookup"><span data-stu-id="bbd27-105">Custom functions can request data through XHR and fetch requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="bbd27-106">下面的文档说明了 Web 请求的一些示例，但是若要为自己构建流式处理函数，请尝试[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows)。</span><span class="sxs-lookup"><span data-stu-id="bbd27-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="bbd27-107">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="bbd27-107">Functions that return data from external sources</span></span>

<span data-ttu-id="bbd27-108">如果自定义函数从外部源（如 Web）检索数据，则必须：</span><span class="sxs-lookup"><span data-stu-id="bbd27-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="bbd27-109">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="bbd27-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="bbd27-110">使用回调函数解析带有最终值的 Promise。</span><span class="sxs-lookup"><span data-stu-id="bbd27-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="bbd27-111">你可以通过 API（如 [`Fetch`](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API)）或使用 `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/zh-CN/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。</span><span class="sxs-lookup"><span data-stu-id="bbd27-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/zh-CN/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="bbd27-112">在自定义函数运行时内，XHR 通过要求[相同来源策略](https://developer.mozilla.org/zh-CN/docs/Web/Security/Same-origin_policy)和简单 [CORS](https://www.w3.org/TR/cors/) 来实施附加安全措施。</span><span class="sxs-lookup"><span data-stu-id="bbd27-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/zh-CN/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="bbd27-113">请注意，简单的 CORS 实施不能使用 cookie，且仅支持简单的方法（GET、HEAD、POST）。</span><span class="sxs-lookup"><span data-stu-id="bbd27-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="bbd27-114">简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。</span><span class="sxs-lookup"><span data-stu-id="bbd27-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="bbd27-115">你还可以在简单 CORS 中使用内容类型标题，前提是内容类型为 `application/x-www-form-urlencoded`、`text/plain` 或 `multipart/form-data`。</span><span class="sxs-lookup"><span data-stu-id="bbd27-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="bbd27-116">XHR 示例</span><span class="sxs-lookup"><span data-stu-id="bbd27-116">XHR example</span></span>

<span data-ttu-id="bbd27-117">在下面的代码示例中，**getTemperature** 函数调用 sendWebRequest 函数，以基于温度计 ID 获取特定区域的温度。</span><span class="sxs-lookup"><span data-stu-id="bbd27-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="bbd27-118">sendWebRequest 函数使用 XHR 来向可以提供相应数据的端点发出 GET 请求。</span><span class="sxs-lookup"><span data-stu-id="bbd27-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

```js
/**
 * Receives a temperature from an online source.
 * @customfunction
 * @param {number} thermometerID Identification number of the thermometer.
 */
function getTemperature(thermometerID) {
  return new Promise(function(setResult) {
      sendWebRequest(thermometerID, function(data){
          storeLastTemperature(thermometerID, data.temperature);
          setResult(data.temperature);
      });
  });
}

// Helper method that uses Office's implementation of XMLHttpRequest in the JavaScript runtime for custom functions.  
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

<span data-ttu-id="bbd27-119">有关具有更多上下文的 XHR 请求的另一个示例，请参阅 [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github 存储库中[此文件](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js)中的 `getFile` 函数。</span><span class="sxs-lookup"><span data-stu-id="bbd27-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="bbd27-120">提取示例</span><span class="sxs-lookup"><span data-stu-id="bbd27-120">Fetch example</span></span>

<span data-ttu-id="bbd27-121">在以下代码示例中，`stockPriceStream` 函数使用股票代码符号来获取每 1000 毫秒的股票价格。</span><span class="sxs-lookup"><span data-stu-id="bbd27-121">In the following code sample, the stockPriceStream function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="bbd27-122">有关此示例的更多详细信息，请参阅[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)。</span><span class="sxs-lookup"><span data-stu-id="bbd27-122">For more details about this sample and to get the accompanying JSON, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span>

```js
/**
 * Streams a stock price.
 * @customfunction 
 * @param {string} ticker Stock ticker.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function stockPriceStream(ticker, invocation) {
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
                invocation.setResult(parseFloat(text));
            })
            .catch(function(error) {
                invocation.setResult(error);
            })
            .then(function() {
                isPending = false;
            });
    }, updateFrequency);

    invocation.onCanceled = () => {
        clearInterval(timer);
    };
}

CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
```

## <a name="receive-data-via-websockets"></a><span data-ttu-id="bbd27-123">通过 WebSocket 接收数据</span><span class="sxs-lookup"><span data-stu-id="bbd27-123">Receiving data via WebSockets</span></span>

<span data-ttu-id="bbd27-124">在自定义函数内，可使用 WebSocket 来通过与服务器的持久连接交换数据。</span><span class="sxs-lookup"><span data-stu-id="bbd27-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="bbd27-125">通过使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式地轮询服务器来获取数据。</span><span class="sxs-lookup"><span data-stu-id="bbd27-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="bbd27-126">WebSocket 示例</span><span class="sxs-lookup"><span data-stu-id="bbd27-126">WebSockets example</span></span>

<span data-ttu-id="bbd27-127">下面的代码示例建立了一个 WebSocket 连接，然后记录来自服务器的每一条传入消息。</span><span class="sxs-lookup"><span data-stu-id="bbd27-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="stream-and-cancel-functions"></a><span data-ttu-id="bbd27-128">流式处理和取消函数</span><span class="sxs-lookup"><span data-stu-id="bbd27-128">Stream and cancel functions</span></span>

<span data-ttu-id="bbd27-129">流式处理自定义函数使用户能够在不需要用户显式刷新数据的情况下，向重复更新的单元格输出数据。</span><span class="sxs-lookup"><span data-stu-id="bbd27-129">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span>

<span data-ttu-id="bbd27-130">可取消的自定义函数使用户能够取消执行流式处理自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。</span><span class="sxs-lookup"><span data-stu-id="bbd27-130">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span>

<span data-ttu-id="bbd27-131">若要将函数声明为流式传输或可取消，请使用 JSDOC 注释标记 `@stream` 或 `@cancelable`。</span><span class="sxs-lookup"><span data-stu-id="bbd27-131">To declare a function as streaming or cancelable, use the JSDOC comment tags `@stream` or `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="bbd27-132">使用调用参数</span><span class="sxs-lookup"><span data-stu-id="bbd27-132">Using an invocation parameter</span></span>

<span data-ttu-id="bbd27-133">默认情况下，`invocation` 参数是任何自定义函数的最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="bbd27-133">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="bbd27-134">`invocation` 参数提供与单元格相关的上下文（例如地址），并且还使你能够使用 `setResult` 和 `onCanceled` 方法。</span><span class="sxs-lookup"><span data-stu-id="bbd27-134">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="bbd27-135">这些方法可定义在函数流式传输 (`setResult`) 或被取消 (`onCanceled`) 时它所执行的操作。</span><span class="sxs-lookup"><span data-stu-id="bbd27-135">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="bbd27-136">如果使用 TypeScript，则调用处理程序需要为 `CustomFunctions.StreamingInvocation` 或 `CustomFunctions.CancelableInvocation` 类型。</span><span class="sxs-lookup"><span data-stu-id="bbd27-136">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="bbd27-137">流式传输和可取消函数示例</span><span class="sxs-lookup"><span data-stu-id="bbd27-137">Streaming and cancelable function example</span></span>
<span data-ttu-id="bbd27-138">以下代码示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="bbd27-138">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="bbd27-139">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="bbd27-139">Note the following about this code:</span></span>

- <span data-ttu-id="bbd27-140">Excel 使用 `setResult` 方法自动显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="bbd27-140">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="bbd27-141">当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数“invocation”。</span><span class="sxs-lookup"><span data-stu-id="bbd27-141">The second input parameter, , is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="bbd27-142">`onCanceled` 回调定义取消函数时执行的函数。</span><span class="sxs-lookup"><span data-stu-id="bbd27-142">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation Invocation parameter necessary for streaming functions.
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function(){
    clearInterval(timer);
    }
}
CustomFunctions.associate("INCREMENT", increment);
```

>[!NOTE]
> <span data-ttu-id="bbd27-143">Excel 会在以下情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="bbd27-143">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="bbd27-144">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="bbd27-144">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="bbd27-145">函数的参数（输入）之一发生变化。</span><span class="sxs-lookup"><span data-stu-id="bbd27-145">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="bbd27-146">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="bbd27-146">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="bbd27-147">用户手动触发重新计算。</span><span class="sxs-lookup"><span data-stu-id="bbd27-147">When the user triggers recalculation manually.</span></span> <span data-ttu-id="bbd27-148">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="bbd27-148">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="bbd27-149">后续步骤</span><span class="sxs-lookup"><span data-stu-id="bbd27-149">Next steps</span></span>

* <span data-ttu-id="bbd27-150">了解[您的函数可以使用的不同参数类型](custom-functions-parameter-options.md)。</span><span class="sxs-lookup"><span data-stu-id="bbd27-150">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="bbd27-151">发现如何[批处理多个 API 调用](custom-functions-batching.md)。</span><span class="sxs-lookup"><span data-stu-id="bbd27-151">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="bbd27-152">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bbd27-152">See also</span></span>

* [<span data-ttu-id="bbd27-153">函数中的可变值</span><span class="sxs-lookup"><span data-stu-id="bbd27-153">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="bbd27-154">创建自定义函数的 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="bbd27-154">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="bbd27-155">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="bbd27-155">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="bbd27-156">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="bbd27-156">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="bbd27-157">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="bbd27-157">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="bbd27-158">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="bbd27-158">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="bbd27-159">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="bbd27-159">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
