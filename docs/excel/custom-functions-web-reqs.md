---
ms.date: 04/29/2020
description: 使用 Excel 中的自定义函数请求、流式处理和取消流式处理工作簿的外部数据
title: 使用自定义函数接收和处理数据
localization_priority: Normal
ms.openlocfilehash: c53ad94c798f787447ab353201a245cd4f20d463
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610459"
---
# <a name="receive-and-handle-data-with-custom-functions"></a><span data-ttu-id="786e0-103">使用自定义函数接收和处理数据</span><span class="sxs-lookup"><span data-stu-id="786e0-103">Receive and handle data with custom functions</span></span>

<span data-ttu-id="786e0-104">自定义函数增强 Excel 功能的方法之一是从工作簿以外的位置接收数据，例如 Web 或服务器（通过 WebSockets）。</span><span class="sxs-lookup"><span data-stu-id="786e0-104">One of the ways that custom functions enhances Excel's power is by receiving data from locations other than the workbook, such as the web or a server (through WebSockets).</span></span> <span data-ttu-id="786e0-105">你可以通过 API（如 [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)）或使用 `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。</span><span class="sxs-lookup"><span data-stu-id="786e0-105">You can request external data through an API like [`Fetch`](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![自定义函数的 gif，可通过 API 对时间进行流式处理](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="786e0-107">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="786e0-107">Functions that return data from external sources</span></span>

<span data-ttu-id="786e0-108">如果自定义函数从外部源（如 Web）检索数据，则必须：</span><span class="sxs-lookup"><span data-stu-id="786e0-108">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="786e0-109">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="786e0-109">Return a JavaScript Promise to Excel.</span></span>
2. <span data-ttu-id="786e0-110">使用回调函数解析带有最终值的 Promise。</span><span class="sxs-lookup"><span data-stu-id="786e0-110">Resolve the Promise with the final value using the callback function.</span></span>

### <a name="fetch-example"></a><span data-ttu-id="786e0-111">Fetch 示例</span><span class="sxs-lookup"><span data-stu-id="786e0-111">Fetch example</span></span>

<span data-ttu-id="786e0-112">在下面的代码示例中， `webRequest` 函数将进入假设的 Contoso "Space Of 人数" API，用于跟踪当前国际空间站上的用户数。</span><span class="sxs-lookup"><span data-stu-id="786e0-112">In the following code sample, the `webRequest` function reaches out to the hypothetical Contoso "Number of People in Space" API, which tracks the number of people currently on the International Space Station.</span></span> <span data-ttu-id="786e0-113">该函数返回一个 JavaScript Promise 并使用 fetch 从 API 请求信息。</span><span class="sxs-lookup"><span data-stu-id="786e0-113">The function returns a JavaScript Promise and uses fetch to request information from the API.</span></span> <span data-ttu-id="786e0-114">生成的数据被转换成 JSON，而 `names` 属性则被转换成一个字符串，用于解析 Promise。</span><span class="sxs-lookup"><span data-stu-id="786e0-114">The resulting data is transformed into JSON and the `names` property is converted into a string, which is used to resolve the Promise.</span></span>

<span data-ttu-id="786e0-115">在开发自己的函数时，可能需要在相应 Web 请求没有及时完成时执行某个操作，或者需要考虑[批处理多个 API 请求](./custom-functions-batching.md)。</span><span class="sxs-lookup"><span data-stu-id="786e0-115">When developing your own functions, you may want to perform an action if the web request does not complete in a timely manner or consider [batching up multiple API requests](./custom-functions-batching.md).</span></span>

```JS
/**
 * Requests the names of the people currently on the International Space Station from a hypothetical API.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace";
  return new Promise(function (resolve, reject) {
    fetch(url)
      .then(function (response){
        return response.json();
        }
      )
      .then(function (json) {
        resolve(JSON.stringify(json.names));
      })
  })
}
```

>[!NOTE]
><span data-ttu-id="786e0-116">使用 `Fetch` 可以避免嵌套回调，在某些情况下可能优于 XHR。</span><span class="sxs-lookup"><span data-stu-id="786e0-116">Using `Fetch` avoids nested callbacks and may be preferable to XHR in some cases.</span></span>

### <a name="xhr-example"></a><span data-ttu-id="786e0-117">XHR 示例</span><span class="sxs-lookup"><span data-stu-id="786e0-117">XHR example</span></span>

<span data-ttu-id="786e0-118">在下面的代码示例中， `getStarCount` 函数将调用 GITHUB API，以发现给定用户存储库中指定的星数。</span><span class="sxs-lookup"><span data-stu-id="786e0-118">In the following code sample, the `getStarCount` function calls the Github API to discover the amount of stars given to a particular user's repository.</span></span> <span data-ttu-id="786e0-119">这是一个可返回 JavaScript Promise 的异步函数。</span><span class="sxs-lookup"><span data-stu-id="786e0-119">This is an asynchronous function which returns a JavaScript Promise.</span></span> <span data-ttu-id="786e0-120">当从 Web 调用中获取数据时，系统将对 Promise 进行解析，以将数据返回到单元格。</span><span class="sxs-lookup"><span data-stu-id="786e0-120">When data is obtained from the web call, the Promise is resolved which returns the data to the cell.</span></span>

```TS
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(userName: string, repoName: string) {

  const url = "https://api.github.com/repos/" + userName + "/" + repoName;

  let xhttp = new XMLHttpRequest();

  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        resolve(JSON.parse(xhttp.responseText).watchers_count);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
```

## <a name="make-a-streaming-function"></a><span data-ttu-id="786e0-121">生成流式处理函数</span><span class="sxs-lookup"><span data-stu-id="786e0-121">Make a streaming function</span></span>

<span data-ttu-id="786e0-122">流式处理自定义函数使用户能够在不需要用户显式刷新数据的情况下，向重复更新的单元格输出数据。</span><span class="sxs-lookup"><span data-stu-id="786e0-122">Streaming custom functions enable you to output data to cells that updates repeatedly, without requiring a user to explicitly refresh anything.</span></span> <span data-ttu-id="786e0-123">这对于检查联机服务中的实时数据非常有用，如[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)中的函数。</span><span class="sxs-lookup"><span data-stu-id="786e0-123">This can be useful to check live data from a service online, like the function in [the custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

<span data-ttu-id="786e0-124">若要声明流式处理函数，您可以使用以下任一方法：</span><span class="sxs-lookup"><span data-stu-id="786e0-124">To declare a streaming function, you can use either:</span></span>

- <span data-ttu-id="786e0-125">`@streaming`标记。</span><span class="sxs-lookup"><span data-stu-id="786e0-125">The `@streaming` tag.</span></span>
- <span data-ttu-id="786e0-126">`CustomFunctions.StreamingInvocation`调用参数。</span><span class="sxs-lookup"><span data-stu-id="786e0-126">The `CustomFunctions.StreamingInvocation` invocation parameter.</span></span>

<span data-ttu-id="786e0-127">以下代码示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="786e0-127">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="786e0-128">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="786e0-128">Note the following about this code:</span></span>

- <span data-ttu-id="786e0-129">Excel 使用 `setResult` 方法自动显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="786e0-129">Excel displays each new value automatically using the `setResult` method.</span></span>
- <span data-ttu-id="786e0-130">当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数“invocation”。</span><span class="sxs-lookup"><span data-stu-id="786e0-130">The second input parameter, invocation, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>
- <span data-ttu-id="786e0-131">`onCanceled` 回调定义取消函数时执行的函数。</span><span class="sxs-lookup"><span data-stu-id="786e0-131">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span>
- <span data-ttu-id="786e0-132">流式处理不一定与发出 Web 请求有关：在本例中，该函数不会发出 Web 请求，但仍以设置的时间间隔获取数据，因此需要使用流式处理 `invocation` 参数。</span><span class="sxs-lookup"><span data-stu-id="786e0-132">Streaming isn't necessarily tied to making a web request: in this case, the function isn't making a web request but is still getting data at set intervals, so it requires the use of the streaming `invocation` parameter.</span></span>

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
```

## <a name="canceling-a-function"></a><span data-ttu-id="786e0-133">取消函数</span><span class="sxs-lookup"><span data-stu-id="786e0-133">Canceling a function</span></span>

<span data-ttu-id="786e0-134">Excel 会在以下情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="786e0-134">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="786e0-135">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="786e0-135">When the user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="786e0-136">函数的参数（输入）之一发生变化。</span><span class="sxs-lookup"><span data-stu-id="786e0-136">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="786e0-137">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="786e0-137">In this case, a new function call is triggered following the cancellation.</span></span>
- <span data-ttu-id="786e0-138">用户手动触发重新计算。</span><span class="sxs-lookup"><span data-stu-id="786e0-138">When the user triggers recalculation manually.</span></span> <span data-ttu-id="786e0-139">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="786e0-139">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="786e0-140">你还可以考虑设置默认流式处理值，以在发出请求但你处于脱机状态时处理案例。</span><span class="sxs-lookup"><span data-stu-id="786e0-140">You can also consider setting a default streaming value to handle cases when a request is made but you are offline.</span></span>

<span data-ttu-id="786e0-141">请注意，还有一类函数被称为可取消函数，它们与流式处理函数_无_关。</span><span class="sxs-lookup"><span data-stu-id="786e0-141">Note that there are also a category of functions called cancelable functions, which are _not_ related to streaming functions.</span></span> <span data-ttu-id="786e0-142">仅可取消可返回一个值的异步自定义函数。</span><span class="sxs-lookup"><span data-stu-id="786e0-142">Only asynchronous custom functions which return one value are cancelable.</span></span> <span data-ttu-id="786e0-143">可取消函数允许在请求中间终止 Web 请求，它使用 [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) 来决定取消时需要采取的操作。</span><span class="sxs-lookup"><span data-stu-id="786e0-143">Cancelable functions allow a web request to be terminated in the middle of a request, using a [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) to decide what to do upon cancellation.</span></span> <span data-ttu-id="786e0-144">使用标记 `@cancelable` 声明可取消函数。</span><span class="sxs-lookup"><span data-stu-id="786e0-144">Declare a cancelable function using the tag `@cancelable`.</span></span>

### <a name="using-an-invocation-parameter"></a><span data-ttu-id="786e0-145">使用调用参数</span><span class="sxs-lookup"><span data-stu-id="786e0-145">Using an invocation parameter</span></span>

<span data-ttu-id="786e0-146">默认情况下，`invocation` 参数是任何自定义函数的最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="786e0-146">The `invocation` parameter is the last parameter of any custom function by default.</span></span> <span data-ttu-id="786e0-147">`invocation`参数提供有关单元格（如其地址和内容）的上下文，并允许您使用 `setResult` 和 `onCanceled` 方法。</span><span class="sxs-lookup"><span data-stu-id="786e0-147">The `invocation` parameter gives context about the cell (such as its address and contents) and allows you to use `setResult` and `onCanceled` methods.</span></span> <span data-ttu-id="786e0-148">这些方法可定义在函数流式传输 (`setResult`) 或被取消 (`onCanceled`) 时它所执行的操作。</span><span class="sxs-lookup"><span data-stu-id="786e0-148">These methods define what a function does when the function streams (`setResult`) or is canceled (`onCanceled`).</span></span>

<span data-ttu-id="786e0-149">如果使用的是 TypeScript，则调用处理程序必须为类型 [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) 或 [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) 。</span><span class="sxs-lookup"><span data-stu-id="786e0-149">If you're using TypeScript, the invocation handler needs to be of type [`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation) or[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation).</span></span>

## <a name="receiving-data-via-websockets"></a><span data-ttu-id="786e0-150">通过 WebSocket 接收数据</span><span class="sxs-lookup"><span data-stu-id="786e0-150">Receiving data via WebSockets</span></span>

<span data-ttu-id="786e0-151">在自定义函数内，可使用 WebSocket 来通过与服务器的持久连接交换数据。</span><span class="sxs-lookup"><span data-stu-id="786e0-151">Within a custom function, you can use WebSockets to exchange data over a persistent connection with a server.</span></span> <span data-ttu-id="786e0-152">使用 Websocket 时，您的自定义函数可以打开与服务器的连接，然后在发生特定事件时自动从服务器接收邮件，而无需显式轮询服务器以获取数据。</span><span class="sxs-lookup"><span data-stu-id="786e0-152">Using WebSockets, your custom function can open a connection with a server and then automatically receive messages from the server when certain events occur, without having to explicitly poll the server for data.</span></span>

### <a name="websockets-example"></a><span data-ttu-id="786e0-153">WebSocket 示例</span><span class="sxs-lookup"><span data-stu-id="786e0-153">WebSockets example</span></span>

<span data-ttu-id="786e0-154">下面的代码示例建立了一个 WebSocket 连接，然后记录来自服务器的每一条传入消息。</span><span class="sxs-lookup"><span data-stu-id="786e0-154">The following code sample establishes a WebSocket connection and then logs each incoming message from the server.</span></span>

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="next-steps"></a><span data-ttu-id="786e0-155">后续步骤</span><span class="sxs-lookup"><span data-stu-id="786e0-155">Next steps</span></span>

- <span data-ttu-id="786e0-156">了解[你的函数可以使用的不同参数类型](custom-functions-parameter-options.md)。</span><span class="sxs-lookup"><span data-stu-id="786e0-156">Learn about [different parameter types your functions can use](custom-functions-parameter-options.md).</span></span>
- <span data-ttu-id="786e0-157">发现如何[批处理多个 API 调用](custom-functions-batching.md)。</span><span class="sxs-lookup"><span data-stu-id="786e0-157">Discover how to [batch multiple API calls](custom-functions-batching.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="786e0-158">另请参阅</span><span class="sxs-lookup"><span data-stu-id="786e0-158">See also</span></span>

- [<span data-ttu-id="786e0-159">函数中的可变值</span><span class="sxs-lookup"><span data-stu-id="786e0-159">Volatile values in functions</span></span>](custom-functions-volatile.md)
- [<span data-ttu-id="786e0-160">创建自定义函数的 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="786e0-160">Create JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="786e0-161">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="786e0-161">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="786e0-162">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="786e0-162">Create custom functions in Excel</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="786e0-163">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="786e0-163">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
