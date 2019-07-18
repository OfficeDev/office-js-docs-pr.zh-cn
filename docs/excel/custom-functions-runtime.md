---
ms.date: 05/08/2019
description: 了解开发使用新 JavaScript 运行时的 Excel 自定义函数时的关键方案。
title: Excel 自定义函数的运行时
localization_priority: Normal
ms.openlocfilehash: e0246170bc80ec63705031cb32a36b5033d42f3a
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771388"
---
# <a name="runtime-for-excel-custom-functions"></a>Excel 自定义函数的运行时

自定义函数使用的是与外接程序的其他部分（例如任务窗格或其他 UI 元素）使用的运行时不同的新 JavaScript 运行时。 这种 JavaScript 运行时旨在优化自定义函数中的计算性能，并支持可用于在自定义函数中执行常见 Web 操作（例如请求外部数据或通过与服务器的持久连接交换数据）的新 API。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

JavaScript 运行时还可提供对 `OfficeRuntime` 命名空间内的新 API 的访问，这些 API 可在自定义函数内或由外接程序的其他部分使用，用于存储数据或显示对话框。 本文介绍了如何在自定义函数内使用这些 API，还概述了在开发自定义函数时需要牢记的其他注意事项。

## <a name="requesting-external-data"></a>请求外部数据

在自定义函数中，你可以使用 [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) 等 API 或使用 [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。

在自定义函数使用的 JavaScript 运行时中, XHR 通过要求使用[相同的源策略](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)和简单的[CORS](https://www.w3.org/TR/cors/)来实现其他安全措施。

请注意，简单的 CORS 实施不能使用 cookie，且仅支持简单的方法（GET、HEAD、POST）。 简单的 CORS 接受字段名称为 `Accept`、`Accept-Language`、`Content-Language` 的简单标题。 您还可以使用简单`Content-Type` CORS 中的标头, 只要内容类型为`application/x-www-form-urlencoded`、 `text/plain`或。 `multipart/form-data`

### <a name="xhr-example"></a>XHR 示例

在下面的代码示例中，`getTemperature` 函数调用 `sendWebRequest` 函数，以基于温度计 ID 获取特定区域的温度。 `sendWebRequest` 函数使用 XHR 来向可以提供相应数据的端点发出 `GET` 请求。

> [!NOTE] 
> 当使用提取或 XHR 时，将返回新的 JavaScript `Promise`。 在 2018 年 9 月之前，必须指定 `OfficeExtension.Promise` 才能在 Office JavaScript API 中使用 promise，但现在可以直接使用 JavaScript `Promise`。

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
        
        //set Content-Type to application/text. Application/json is not currently supported with Simple CORS
        xhttp.setRequestHeader("Content-Type", "application/text");
        xhttp.open("GET", "https://contoso.com/temperature/" + thermometerID), true)
        xhttp.send();  
    }
}
```

## <a name="receiving-data-via-websockets"></a>通过 WebSocket 接收数据

在自定义函数内，可使用 [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) 来通过与服务器的持久连接交换数据。 通过使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式地轮询服务器来获取数据。

### <a name="websockets-example"></a>WebSocket 示例

下面的代码示例建立了一个 `WebSocket` 连接，然后记录来自服务器的每一条传入消息。

```JavaScript
const ws = new WebSocket('wss://bundles.office.com');
ws.onmessage = function (message) {
    console.log(`Received: ${message}`);
}
ws.onerror = function (error) {
    console.err(`Failed: ${error}`);
}
```

## <a name="storing-and-accessing-data"></a>存储和访问数据

在自定义函数（或外接程序的任何其他部分）内，可以使用 `OfficeRuntime.storage` 对象来存储和访问数据。 `Storage` 是一种未加密的持久键值存储系统，为无法在自定义函数内使用的 [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage) 提供了一种替代方案。 `Storage`每个域提供 10 MB 的数据。 域可由多个加载项共享。

`Storage` 旨在作为共享存储解决方案，这意味着外接程序的多个部分将能访问相同数据。 例如，用户身份验证令牌可能存储在 `storage` 中，因为自定义函数和任务窗格等外接程序 UI 元素都有可能会访问该数据。 同样，如果两个外接程序共享同一个域（例如 www.contoso.com/addin1、www.contoso.com/addin2），则也可以通过 `storage` 来回共享信息。 注意，具有不同子域的外接程序将具有不同的 `storage` 实例（例如 subdomain.contoso.com/addin1、differentsubdomain.contoso.com/addin2）

由于 `storage` 可能是共享的位置，因此一定要认识到，可能会存在替代键值对的情况。

`storage` 对象支持以下方法：

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

.[!NOTE]
> 没有用于清除所有信息的方法 (例如`clear`)。 相反，需要使用 `removeItems` 来一次性删除多个条目。

### <a name="officeruntimestorage-example"></a>OfficeRuntime 示例

下面的代码示例调用`OfficeRuntime.storage.setItem`函数, 以将键和值设置为`storage`。

```JavaScript
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a>其他注意事项

为创建一个可在多个平台（Office 外接程序的关键租户之一）上运行的外接程序，请勿访问自定义函数中的文档对象模型 (DOM) 或使用 jQuery 等这类依赖于 DOM 的库。 在 Windows 的 Excel 中, 自定义函数使用 JavaScript 运行时, 自定义函数无法访问 DOM。

## <a name="next-steps"></a>后续步骤
了解如何[使用自定义函数执行 web 请求](custom-functions-web-reqs.md)。

## <a name="see-also"></a>另请参阅

* [在 Excel 中创建自定义函数](custom-functions-overview.md)
* [自定义函数体系结构](custom-functions-architecture.md)
* [在自定义函数中显示对话框](custom-functions-dialog.md)
* [自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
