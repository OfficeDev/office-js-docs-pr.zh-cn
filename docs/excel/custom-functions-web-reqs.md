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
# <a name="receiving-and-handling-data-with-custom-functions"></a>使用自定义函数接收和处理数据

自定义函数增强 Excel 功能的方法之一是从工作簿以外的位置接收数据，例如 Web 或服务器（通过 WebSockets）。 自定义函数可以通过 XHR 请求数据和获取请求，也可以实时流式处理这些数据。

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

有关具有更多上下文的 XHR 请求的另一个示例，请参阅 [Office-Add-in-JavaScript-FileDownload](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload) Github 存储库中[此文件](https://github.com/OfficeDev/Office-Add-in-JavaScript-FileDownload/blob/master/FileDownloadSampleWeb/Home.js)中的 `getFile` 函数。

### <a name="fetch-example"></a>提取示例

在以下代码示例中，stockPriceStream 函数使用股票代码符号来获取每 1000 毫秒的股票价格。 有关此示例的更多详细信息以及若要获取随附的 JSON，请参阅[自定义函数教程](https://docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions?tabs=excel-windows#create-a-streaming-asynchronous-custom-function)。 

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

## <a name="receiving-data-via-websockets"></a>通过 WebSocket 接收数据

在自定义函数内，可使用 WebSocket 来通过与服务器的持久连接交换数据。 通过使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式地轮询服务器来获取数据。

### <a name="websockets-example"></a>WebSocket 示例

下面的代码示例建立了一个 WebSocket 连接，然后记录来自服务器的每一条传入消息。

```JavaScript
var ws = new WebSocket('wss://bundles.office.com');

ws.onmessage(message) {
    console.log(`Recieved: ${message}`);
}

ws.onerror(error){
    console.err(`Failed: ${error}`);
}
```

## <a name="streaming-functions"></a>流式处理函数

流式处理自定义函数使用户能够在不需要用户显式请求数据刷新的情况下，随着时间的推移向单元格重复输出数据。 以下代码示例是一个自定义函数，它每秒向结果添加一个数字。 关于此代码，请注意以下几点：

- Excel 使用 setResult 回调自动显示每个新值。
- 当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数 handler。
- onCanceled 回调定义取消函数时执行的函数。 对于任何流式处理函数，都必须实现此类取消处理程序。 有关详细信息，请参阅[取消函数](#canceling-a-function)。

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

在 JSON 元数据文件中为流式处理函数指定元数据时，必须在 options 对象中设置属性 "cancelable": true 和 "stream": true，如以下示例所示。

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

## <a name="canceling-a-function"></a>取消函数

在某些情况下，可能需要取消执行流式处理自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。 Excel 会在以下情况下取消函数的执行：

- 用户编辑或删除引用函数的单元格。
- 函数的参数（输入）之一发生变化。 在这种情况下，取消之后还会触发新的函数调用。
- 用户手动触发重新计算。 在这种情况下，取消之后还会触发新的函数调用。

若要使函数可取消，请在函数代码中实施一个处理程序，以告诉它在取消时要执行的操作。 此外，指定描述函数的 JSON 元数据中 options 对象中的属性 `"cancelable": true`。 本文前一部分中的代码示例提供了这些方法的示例。

## <a name="see-also"></a>另请参阅

* [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数的运行时](custom-functions-runtime.md)
* [自定义函数最佳实践](custom-functions-best-practices.md)
* [自定义函数更改日志](custom-functions-changelog.md)
