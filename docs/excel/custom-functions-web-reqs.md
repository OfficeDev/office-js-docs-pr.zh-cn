---
ms.date: 05/30/2019
description: 使用 Excel 中的自定义函数请求、流式处理和取消流式处理工作簿的外部数据
title: 使用自定义函数接收和处理数据
localization_priority: Priority
ms.openlocfilehash: add6a3bc91b28ff7dbd0f0b298ed8f38ed5dd1bc
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706142"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="cba86-103">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="cba86-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="cba86-104">自定义函数增强 Excel 功能的方法之一是从工作簿以外的位置接收数据，例如 Web 或服务器（通过 WebSockets）。</span><span class="sxs-lookup"><span data-stu-id="cba86-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="cba86-105">自定义函数可以通过 XHR 和 `fetch` 请求来请求数据，也可以实时流式处理这些数据。</span><span class="sxs-lookup"><span data-stu-id="cba86-105">Custom functions can request data through XHR and `fetch` requests as well as stream this data in real time.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cba86-106">下面的文档说明了 Web 请求的一些示例，但是若要为自己构建流式处理函数，请尝试[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows)。</span><span class="sxs-lookup"><span data-stu-id="cba86-106">The documentation below illustrates some samples of web requests, but to build a streaming function for yourself, try the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows).</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="cba86-107">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="cba86-107">Functions that return data from external sources</span></span>

<span data-ttu-id="cba86-108">如果自定义函数从外部源（如 Web）检索数据，则必须：</span><span class="sxs-lookup"><span data-stu-id="cba86-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="cba86-109">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="cba86-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="cba86-110">使用回调函数解析带有最终值的 Promise。</span><span class="sxs-lookup"><span data-stu-id="cba86-110">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="cba86-111">你可以通过 API（如 [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)）或使用 `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。</span><span class="sxs-lookup"><span data-stu-id="cba86-111">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="cba86-112">在自定义函数运行时内，XHR 通过要求[相同来源策略](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)和简单 [CORS](https://www.w3.org/TR/cors/) 来实施附加安全措施。</span><span class="sxs-lookup"><span data-stu-id="cba86-112">Within custom functions runtime, XHR implements additional security measures by requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="cba86-113">请注意，简单的 CORS 实施不能使用 cookie，且仅支持简单的方法（GET、HEAD、POST）。</span><span class="sxs-lookup"><span data-stu-id="cba86-113">Note that a simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="cba86-114">简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。</span><span class="sxs-lookup"><span data-stu-id="cba86-114">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="cba86-115">你还可以在简单 CORS 中使用内容类型标题，前提是内容类型为 `application/x-www-form-urlencoded`、`text/plain` 或 `multipart/form-data`。</span><span class="sxs-lookup"><span data-stu-id="cba86-115">You can also use a Content-Type header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="cba86-116">XHR 示例</span><span class="sxs-lookup"><span data-stu-id="cba86-116">XHR example</span></span>

<span data-ttu-id="cba86-117">在下面的代码示例中，**getTemperature** 函数调用 sendWebRequest 函数，以基于温度计 ID 获取特定区域的温度。</span><span class="sxs-lookup"><span data-stu-id="cba86-117">In the following code sample, the **getTemperature** function calls the sendWebRequest function to get the temperature of a particular area based on thermometer ID.</span></span> <span data-ttu-id="cba86-118">sendWebRequest 函数使用 XHR 来向可以提供相应数据的端点发出 GET 请求。</span><span class="sxs-lookup"><span data-stu-id="cba86-118">The sendWebRequest function uses XHR to issue a GET request to an endpoint that can provide the data.</span></span>

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

<span data-ttu-id="cba86-119">有关具有更多上下文的 XHR 请求的另一个示例，请参阅 [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github 存储库中[此文件](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js)中的 `getFile` 函数。</span><span class="sxs-lookup"><span data-stu-id="cba86-119">For another sample of an XHR request with more context, see the `getFile` function within [this file](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js) in the [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github repository.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="cba86-120">提取示例</span><span class="sxs-lookup"><span data-stu-id="cba86-120">Fetch example</span></span>

<span data-ttu-id="cba86-121">在以下代码示例中，`stockPriceStream` 函数使用股票代码符号来获取每 1000 毫秒的股票价格。</span><span class="sxs-lookup"><span data-stu-id="cba86-121">In the following code sample, the `stockPriceStream` function uses a stock ticker symbol to get the price of a stock every 1000 milliseconds.</span></span> <span data-ttu-id="cba86-122">有关此示例的更多详细信息，请参阅[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)。</span><span class="sxs-lookup"><span data-stu-id="cba86-122">For more details about this sample, see the [Custom functions tutorial](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function).</span></span>

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

## <a name="receive-data-via-websockets"></a><span data-ttu-id="cba86-123">通过 WebSocket 接收数据</span><span class="sxs-lookup"><span data-stu-id="cba86-123">Receive data via WebSockets</span></span>

<span data-ttu-id="cba86-124">在自定义函数内，可使用 WebSocket 来通过与服务器的持久连接交换数据。</span><span class="sxs-lookup"><span data-stu-id="cba86-124">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="cba86-125">通过使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式地轮询服务器来获取数据。</span><span class="sxs-lookup"><span data-stu-id="cba86-125">By using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="cba86-126">WebSocket 示例</span><span class="sxs-lookup"><span data-stu-id="cba86-126">WebSockets example</span></span>

<span data-ttu-id="cba86-127">下面的代码示例建立了一个 WebSocket 连接，然后记录来自服务器的每一条传入消息。</span><span class="sxs-lookup"><span data-stu-id="cba86-127">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="cba86-128">生成流式处理函数</span><span class="sxs-lookup"><span data-stu-id="cba86-128">Make a streaming function</span></span>

<span data-ttu-id="cba86-129">流式处理自定义函数使用户能够在不需要用户显式刷新数据的情况下，向重复更新的单元格输出数据。</span><span class="sxs-lookup"><span data-stu-id="cba86-129">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="cba86-130">这对于检查联机服务中的实时数据非常有用，如[自定义函数教程](/tutorials/excel-tutorial-create-custom-functions)中的函数。</span><span class="sxs-lookup"><span data-stu-id="cba86-130">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](/tutorials/excel-tutorial-create-custom-functions).</span></span>

<span data-ttu-id="cba86-131">若要声明函数，请使用 JSDoc 批注标记 `@stream`。</span><span class="sxs-lookup"><span data-stu-id="cba86-131">To declare a streaming function, use the JSDoc comment tag `@stream`.</span></span> <span data-ttu-id="cba86-132">若要提醒用户你的函数可能会根据新的信息重新提升，请考虑使用流或其他措辞，以在函数的名称或描述中说明此情况。</span><span class="sxs-lookup"><span data-stu-id="cba86-132">To alert users to the fact that your function may re-evaluate based on new information, consider putting stream or other wording to indicate this in the name or description of your function.</span></span>

<span data-ttu-id="cba86-133">以下示例显示了每秒按你指定的幅度提高给定数值的流式函数。</span><span class="sxs-lookup"><span data-stu-id="cba86-133">The following example shows a streaming function which increases a given number every second by an amount you specify.</span></span>

```JS
/**
 * Increments a value once a second.
 * @customfunction INC increment
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("INC", increment);
```

>[!NOTE]
> <span data-ttu-id="cba86-134">请注意，还有一类函数被称为可取消函数，它们与流式函数*无*关。</span><span class="sxs-lookup"><span data-stu-id="cba86-134">Note that there are also a category of functions called cancelable functions, which are *not* related to streaming functions.</span></span> <span data-ttu-id="cba86-135">以前版本的自定义函数需要在手写的 JSON 中声明 `"cancelable": true` 和 `"streaming": true`。</span><span class="sxs-lookup"><span data-stu-id="cba86-135">Previous versions of custom functions required you to declare `"cancelable": true` and `"streaming": true` in JSON written by hand.</span></span> <span data-ttu-id="cba86-136">引入自动生成的元数据之后，仅返回一个值的异步自定义函数可取消。</span><span class="sxs-lookup"><span data-stu-id="cba86-136">Since the introduction of autogenerated metadata, only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="cba86-137">可取消函数允许在请求中间终止 Web 请求，它使用 [`CancelableInvocation`](https://docs.microsoft.com/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation?view=office-js) 来决定取消时需要采取的操作。</span><span class="sxs-lookup"><span data-stu-id="cba86-137">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](https://docs.microsoft.com/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation?view=office-js) to decide what to do upon cancellation.</span></span> <span data-ttu-id="cba86-138">使用标记 `@cancelable` 声明可取消函数。</span><span class="sxs-lookup"><span data-stu-id="cba86-138">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="cba86-139">使用调用参数</span><span class="sxs-lookup"><span data-stu-id="cba86-139">Using an invocation parameter</span></span>

<span data-ttu-id="cba86-140">默认情况下，`invocation` 参数是任何自定义函数的最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="cba86-140">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="cba86-141">`invocation` 参数提供与单元格相关的上下文（例如地址），并且还使你能够使用 `setResult` 和 `onCanceled` 方法。</span><span class="sxs-lookup"><span data-stu-id="cba86-141">The `invocation` parameter gives context about the cell (such as its address) and also allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="cba86-142">这些方法可定义在函数流式传输 (`setResult`) 或被取消 (`onCanceled`) 时它所执行的操作。</span><span class="sxs-lookup"><span data-stu-id="cba86-142">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="cba86-143">如果使用 TypeScript，则调用处理程序需要为 `CustomFunctions.StreamingInvocation` 或 `CustomFunctions.CancelableInvocation` 类型。</span><span class="sxs-lookup"><span data-stu-id="cba86-143">If you're using TypeScript, the invocation handler needs to be of type `CustomFunctions.StreamingInvocation` or `CustomFunctions.CancelableInvocation`.</span></span>

### <a name="streaming-and-cancelable-function-example"></a><span data-ttu-id="cba86-144">流式传输和可取消函数示例</span><span class="sxs-lookup"><span data-stu-id="cba86-144">Streaming and cancelable function example</span></span>
<span data-ttu-id="cba86-145">以下代码示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="cba86-145">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="cba86-146">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="cba86-146">Note the following about this code:</span></span>

- <span data-ttu-id="cba86-147">Excel 使用 `setResult` 方法自动显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="cba86-147">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="cba86-148">当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数“invocation”。</span><span class="sxs-lookup"><span data-stu-id="cba86-148">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="cba86-149">`onCanceled` 回调定义取消函数时执行的函数。</span><span class="sxs-lookup"><span data-stu-id="cba86-149">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>

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
> <span data-ttu-id="cba86-150">Excel 会在以下情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="cba86-150">Excel cancels the execution of a function in the following situations:</span></span>
>
> - <span data-ttu-id="cba86-151">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="cba86-151">When the user edits or deletes a cell that references the function.</span></span>
> - <span data-ttu-id="cba86-152">函数的参数（输入）之一发生变化。</span><span class="sxs-lookup"><span data-stu-id="cba86-152">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="cba86-153">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="cba86-153">In this case, a new function call is triggered following the cancellation.</span></span>
> - <span data-ttu-id="cba86-154">用户手动触发重新计算。</span><span class="sxs-lookup"><span data-stu-id="cba86-154">When the user triggers recalculation manually.</span></span> <span data-ttu-id="cba86-155">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="cba86-155">In this case, a new function call is triggered following the cancellation.</span></span>

## <a name="next-steps"></a><span data-ttu-id="cba86-156">后续步骤</span><span class="sxs-lookup"><span data-stu-id="cba86-156">Next steps</span></span>

* <span data-ttu-id="cba86-157">了解[你的函数可以使用的不同参数类型](custom-functions-parameter-options.md)。</span><span class="sxs-lookup"><span data-stu-id="cba86-157">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
* <span data-ttu-id="cba86-158">发现如何[批处理多个 API 调用](custom-functions-batching.md)。</span><span class="sxs-lookup"><span data-stu-id="cba86-158">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="cba86-159">另请参阅</span><span class="sxs-lookup"><span data-stu-id="cba86-159">See also</span></span>

* [<span data-ttu-id="cba86-160">函数中的可变值</span><span class="sxs-lookup"><span data-stu-id="cba86-160">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="cba86-161">创建自定义函数的 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="cba86-161">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="cba86-162">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="cba86-162">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="cba86-163">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="cba86-163">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="cba86-164">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="cba86-164">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="cba86-165">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="cba86-165">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="cba86-166">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="cba86-166">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
