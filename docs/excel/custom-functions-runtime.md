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
# <a name="runtime-for-excel-custom-functions-preview"></a>Excel 自定义函数的运行时（预览）

自定义函数使用新的JavaScript运行时，该运行时与加载项的其他部分（例如任务窗格或其他UI元素）使用的运行时不同。此JavaScript运行时旨在优化自定义函数中的计算性能，并公开可用于在自定义函数中执行常见的基于Web的操作（如请求外部数据或通过与服务器的持久连接以交换数据）的新API。JavaScript运行时还提供对 `OfficeRuntime`  命名空间中的新 API 的访问，这些 API 可以在自定义函数中使用，也可以由加载项的其他部分用于存储数据或显示对话框。 本文介绍如何在自定义函数中使用这些 API，并概述在开发自定义函数时要记住的其他注意事项。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="requesting-external-data"></a>请求外部数据

在自定义函数中，您可以使用 [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) 等 API 或使用 [XmlHttpRequest  (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)  （发布 HTTP 请求以与服务器交互的一个标准 Web API）来请求外部数据。在 JavaScript  运行时中，XHR 通过要求 [同源策略](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) 和简单的 [CORS](https://www.w3.org/TR/cors/) 来实现其他安全措施。  

### <a name="xhr-example"></a>XHR 示例

在下面的代码示例中， `getTemperature` 函数调用 `sendWebRequest` 函数以根据温度计 ID 获取特定区域的温度。 `sendWebRequest` 函数使用 XHR 向可提供数据的端点发出 `GET` 请求。 

> [!NOTE] 
> 当使用 fetch 或 XHR 时，会返回一个新的 JavaScript `Promise` 。在2018年9月之前，您必须指定 `OfficeExtension.Promise` 在 Office  JavaScript API 中使用 promises，但现在您只需使用 JavaScript `Promise` 即可。

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

## <a name="receiving-data-via-websockets"></a>通过 WebSockets 接收数据

在自定义函数中，您可以使用 [WebSockets](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) 通过与服务器的持久连接来交换数据。 通过使用 WebSockets，您的自定义函数可以打开与服务器的连接，然后在发生特定事件时自动从服务器接收消息，而无需显式轮询服务器以获取数据。

### <a name="websockets-example"></a>WebSockets 示例

下面的代码示例建立了一个 `WebSocket` 连接，然后记录来自服务器的每一封传入的邮件。 

```typescript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>存储和访问数据

在自定义函数内（或在加载项的任何其他部分内），您可以使用 `OfficeRuntime.AsyncStorage` 对象来存储和访问数据。 `AsyncStorage` 是一个持久的，未加密的键值存储系统，它提供了 [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) 的替代方法，它不能在自定义函数中使用。加载项可以使用 `AsyncStorage` 来存储多达 10 MB 的数据。

下面的方法在 `AsyncStorage` 对象上可用：
 
 - `getItem`
 - `setItem`
 - `removeItem`
 - `clear`
 - `getAllKeys`
 - `flushGetRequests`
 - `multiGet`
 - `multiSet`
 - `multiRemove`

### <a name="asyncstorage-example"></a>AsyncStorage 示例 

下面的代码示例调用 `AsyncStorage.getItem` 函数以从存储检索值。

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

## <a name="displaying-a-dialog-box"></a>显示对话框

在自定义函数内（或在加载项的任何其他部分内），您可以使用 `OfficeRuntime.displayWebDialogOptions`  API 以显示对话框。 此对话框 API 提供了 [Dialog API](../develop/dialog-api-in-office-add-ins.md) 的替代方法，可以在任务窗格和加载项命令中使用，但不能在自定义函数中使用。

### <a name="dialog-api-example"></a>Dialog API 示例 

在下面的代码示例中，该函数 `getTokenViaDialog` 使用 Dialog API 的 `displayWebDialogOptions` 函数显示一个对话框。

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

## <a name="additional-considerations"></a>其他注意事项

为了创建一个将在多个平台（Office 加载项的关键租户之一）上运行的加载项，您不应访问自定义函数中的文档对象模型（DOM）或使用依赖于 DOM 像 jQuery 这样的库。 在 Excel for Windows 中，自定义函数使用 JavaScript 运行时，自定义函数无法访问 DOM。

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [自定义函数元数据](custom-functions-json.md)
* [自定义函数最佳做法](custom-functions-best-practices.md)
* [Excel 自定义函数教程](excel-tutorial-custom-functions.md)
