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
# <a name="runtime-for-excel-custom-functions"></a>Excel 自定义函数运行时

自定义函数使用新的 JavaScript 运行时，其使用的是沙盒化的 JavaScript 引擎而不是 web 浏览器，由此扩展了 Excel 的功能。 因为自定义函数不需要呈现 UI 元素，新的 JavaScript 运行时为执行计算进行了优化，让你能够同时运行数千个自定义函数。

## <a name="key-facts-about-the-new-javascript-runtime"></a>新的 JavaScript 运行时的有关要点 

只有加载项中的自定义函数将会使用本文中介绍的新 JavaScript 运行时。 如果加载项包括其他组件，例如任务窗格和其他 UI 元素，除了自定义函数外，加载项的这些其他组件将继续运行在类似于浏览器的 WebView 运行时中。  另外： 

- JavaScript 运行时不提供对文档对象模型 (DOM) 的访问，也不支持类似于 jQuery 的依赖于 DOM 的库。

- 加载项的 JavaScript 文件中定义的自定义函数可以返回常规 JavaScript `Promise` 而不是返回 `OfficeExtension.Promise`。  

- 指定自定义函数元数据的 JSON 文件不需在**选项**中指定**同步**或**异步**。

## <a name="new-apis"></a>新 API 

自定义函数使用的 JavaScript 运行时有以下 API：

- [XHR](#xhr)
- [Websocket](#websockets)
- [AsyncStorage](#asyncstorage)
- [Dialog API](#dialog-api)

### <a name="xhr"></a>XHR

XHR 代表 [XmlHttpRequest](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)，这是一个发出 HTTP 请求以便与服务器进行交互的标准 web API。 在新的 JavaScript 运行时中，XHR 通过要求[同源策略](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)和简单 [CORS](https://www.w3.org/TR/cors/)实现附加安全措施。  

在下面的代码示例中，`getTemperature()` 函数发出 web 请求，以获取基于温度计 ID 的特定区域温度。  `sendWebRequest()` 函数使用 XHR 向可提供数据的端点发出 `GET` 请求。  

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

### <a name="websockets"></a>Websocket

[Websocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) 是在服务器与一个或多个客户端之间建立实时通信的网络协议。 它通常用于聊天应用程序，因为它允许同时读取和写入文本。  

如下面的代码示例中所示，自定义函数可以使用 Websocket。 本示例中，WebSocket 记录其接收的每条消息。

```ts
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = (message) => {
    console.log(`Received: ${message}`);
};
ws.onerror = (error) => {
    console.err(`Failed: ${error}`);
}
```

### <a name="asyncstorage"></a>AsyncStorage

AsyncStorage 是可用于存储身份验证令牌的键值存储系统。 它特点是：

- 持续
- 不加密
- 异步

AsyncStorage 对于加载项的所有部件全局可用。 对于自定义函数，`AsyncStorage` 作为全局对象公开。 （对于加载项的其他部件，如任务窗格和使用 WebView 运行时的其他元素，AsyncStorage 通过 `OfficeRuntime` 公开。）每一加载项都有自己的存储分区，默认大小为 5 MB。 

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
 
`mergeItem` 和 `multiMerge` 方法现时不受支持。

下面的代码示例调用 `AsyncStorage.getItem` 函数以从存储检索值。

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

### <a name="dialog-api"></a>Dialog API

Dialog API 可打开一个提示用户登录的对话框。  可使用 Dialog API 要求通过Google 或 Facebook 等外部资源进行用户身份验证，此后用户才可以使用你的函数。   

在下面的代码示例中，`getTokenViaDialog()` 方法使用 Dialog API 的 `displayWebDialog()` 方法打开一个对话框。

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
> 本节中所述的 Dialog AP 是用于自定义函数的新 JavaScript 运行时的一部分，仅可在自定义的函数中使用。 此 API 不同于可在任务窗格和加载项命令中使用的 [Dialog API](../develop/dialog-api-in-office-add-ins.md)。

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [自定义函数元数据](custom-functions-json.md)
* [自定义函数的最佳实践](custom-functions-best-practices.md)