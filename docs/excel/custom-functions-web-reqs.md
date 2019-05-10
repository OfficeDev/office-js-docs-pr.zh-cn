---
ms.date: 05/03/2019
description: 使用 Excel 中的自定义函数请求、流式处理和取消流式处理工作簿的外部数据
title: 使用自定义函数接收和处理数据
localization_priority: Priority
ms.openlocfilehash: 2f70bd5cd5d8e645b47f2bc97dcec3e8bbacef55
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2019
ms.locfileid: "33627983"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>使用自定义函数接收和处理数据

自定义函数增强 Excel 功能的方法之一是从工作簿以外的位置接收数据，例如 Web 或服务器（通过 WebSockets）。 自定义函数可以通过 XHR 和 `fetch` 请求来请求数据，也可以实时流式处理这些数据。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

下面的文档说明了 Web 请求的一些示例，但是若要为自己构建流式处理函数，请尝试[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows)。

## <a name="functions-that-return-data-from-external-sources"></a>从外部源返回数据的函数

如果自定义函数从外部源（如 Web）检索数据，则必须：

1. 将 JavaScript Promise 返回到 Excel。
2. 使用回调函数解析带有最终值的 Promise。

你可以通过 API（如 [`Fetch`](https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API)）或使用 `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/zh-CN/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。

在自定义函数运行时内，XHR 通过要求[相同来源策略](https://developer.mozilla.org/zh-CN/docs/Web/Security/Same-origin_policy)和简单 [CORS](https://www.w3.org/TR/cors/) 来实施附加安全措施。

请注意，简单的 CORS 实施不能使用 cookie，且仅支持简单的方法（GET、HEAD、POST）。 简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。 你还可以在简单 CORS 中使用内容类型标题，前提是内容类型为 `application/x-www-form-urlencoded`、`text/plain` 或 `multipart/form-data`。

### <a name="xhr-example"></a>XHR 示例

在下面的代码示例中，**getTemperature** 函数调用 sendWebRequest 函数，以基于温度计 ID 获取特定区域的温度。 sendWebRequest 函数使用 XHR 来向可以提供相应数据的端点发出 GET 请求。

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

有关具有更多上下文的 XHR 请求的另一个示例，请参阅 [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github 存储库中[此文件](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js)中的 `getFile` 函数。

### <a name="fetch-example"></a>提取示例

在以下代码示例中，`stockPriceStream` 函数使用股票代码符号来获取每 1000 毫秒的股票价格。 有关此示例的更多详细信息，请参阅[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)。

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

## <a name="receive-data-via-websockets"></a>通过 WebSocket 接收数据

在自定义函数内，可使用 WebSocket 来通过与服务器的持久连接交换数据。 通过使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式地轮询服务器来获取数据。

### <a name="websockets-example"></a>WebSocket 示例

下面的代码示例建立了一个 WebSocket 连接，然后记录来自服务器的每一条传入消息。

```js
let ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Received: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="stream-and-cancel-functions"></a>流式处理和取消函数

流式处理自定义函数使用户能够在不需要用户显式刷新数据的情况下，向重复更新的单元格输出数据。

可取消的自定义函数使用户能够取消执行流式处理自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。

若要将函数声明为流式传输或可取消，请使用 JSDOC 注释标记 `@stream` 或 `@cancelable`。

### <a name="using-an-invocation-parameter"></a>使用调用参数

默认情况下，`invocation` 参数是任何自定义函数的最后一个参数。 `invocation` 参数提供与单元格相关的上下文（例如地址），并且还使你能够使用 `setResult` 和 `onCanceled` 方法。 这些方法可定义在函数流式传输 (`setResult`) 或被取消 (`onCanceled`) 时它所执行的操作。

如果使用 TypeScript，则调用处理程序需要为 `CustomFunctions.StreamingInvocation` 或 `CustomFunctions.CancelableInvocation` 类型。

### <a name="streaming-and-cancelable-function-example"></a>流式传输和可取消函数示例
以下代码示例是一个自定义函数，它每秒向结果添加一个数字。 关于此代码，请注意以下几点：

- Excel 使用 `setResult` 方法自动显示每个新值。
- 当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数“invocation”。
- `onCanceled` 回调定义取消函数时执行的函数。

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
> Excel 会在以下情况下取消函数的执行：
>
> - 用户编辑或删除引用函数的单元格。
> - 函数的参数（输入）之一发生变化。 在这种情况下，取消之后还会触发新的函数调用。
> - 用户手动触发重新计算。 在这种情况下，取消之后还会触发新的函数调用。

## <a name="next-steps"></a>后续步骤
了解[函数可使用的不同参数类型](custom-functions-parameter-options.md) 探索如何[批处理多个 API 调用](custom-functions-batching.md)。

## <a name="see-also"></a>另请参阅

* [函数中的可变值](custom-functions-volatile.md)
* [创建自定义函数的 JSON 元数据](custom-functions-json-autogeneration.md)
* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)