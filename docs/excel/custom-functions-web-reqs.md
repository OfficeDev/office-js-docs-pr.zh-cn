---
ms.date: 05/02/2022
description: 使用 Excel 中的自定义函数请求、流式传输和取消将外部数据流式传输到工作簿。
title: 使用自定义函数接收和处理数据
ms.localizationpriority: medium
ms.openlocfilehash: fbe319e79d4cded5fe4b37ce5a654e633996f22a
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958543"
---
# <a name="receive-and-handle-data-with-custom-functions"></a>使用自定义函数接收和处理数据

自定义函数增强 Excel 功能的方法之一是从工作簿以外的位置接收数据，例如 Web 或服务器 (通过 [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API)) 。 你可以通过 API（如 [`Fetch`](https://developer.mozilla.org/docs/Web/API/Fetch_API)）或使用 `XmlHttpRequest` [(XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)（一种发出与服务器交互的 HTTP 请求的标准 Web API）来请求外部数据。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

![从 API 流式传输时间的自定义函数的 GIF。](../images/custom-functions-web-api.gif)

## <a name="functions-that-return-data-from-external-sources"></a>从外部源返回数据的函数

如果自定义函数从外部源（如 Web）检索数据，则必须：

1. 将 [JavaScript `Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 返回到 Excel。
2. `Promise`使用回调函数解析最终值。

### <a name="fetch-example"></a>Fetch 示例

在下面的代码示例中，该 `webRequest` 函数可接触到一个假设的外部 API，该 API 跟踪国际空间站上当前的人数。 该函数返回 JavaScript `Promise` ，并用于 `fetch` 从假设 API 请求信息。 生成的数据将转换为 JSON， `names` 并将属性转换为字符串，用于解析承诺。

在开发自己的函数时，可能需要在相应 Web 请求没有及时完成时执行某个操作，或者需要考虑[批处理多个 API 请求](custom-functions-batching.md)。

```JS
/**
 * Requests the names of the people currently on the International Space Station.
 * Note: This function requests data from a hypothetical URL. In practice, replace the URL with a data source for your scenario.
 * @customfunction
 */
function webRequest() {
  let url = "https://www.contoso.com/NumberOfPeopleInSpace"; // This is a hypothetical URL.
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

> [!NOTE]
> 使用 `fetch` 可以避免嵌套回调，在某些情况下可能优于 XHR。

### <a name="xhr-example"></a>XHR 示例

在下面的代码示例中，该 `getStarCount` 函数调用 Github API 来发现给特定用户存储库的星数。 这是一个返回 JavaScript `Promise`的异步函数。 从 Web 调用获取数据时，会解析承诺，以便将数据返回到单元格。

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

## <a name="make-a-streaming-function"></a>生成流式处理函数

流式处理自定义函数使用户能够在不需要用户显式刷新数据的情况下，向重复更新的单元格输出数据。 这对于检查联机服务中的实时数据非常有用，如[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)中的函数。

若要声明流式处理函数，可以使用以下两个选项之一。

- 标记 `@streaming` 。
- `CustomFunctions.StreamingInvocation`调用参数。

以下代码示例是一个自定义函数，它每秒向结果添加一个数字。 对于此代码，请注意以下事项。

- Excel 使用 `setResult` 方法自动显示每个新值。
- 当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数 `invocation`。
- 回 `onCanceled` 调定义在取消函数时运行的函数。
- 流式处理不一定与发出 Web 请求相关。 在这种情况下，函数不会发出 Web 请求，但仍按设置的时间间隔获取数据，因此需要使用流式处理 `invocation` 参数。

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

## <a name="cancel-a-function"></a>取消函数

Excel 在以下情况下取消执行函数。

- 用户编辑或删除引用函数的单元格。
- 函数的参数（输入）之一发生变化。 在这种情况下，取消之后还会触发新的函数调用。
- 用户手动触发重新计算。 在这种情况下，取消之后还会触发新的函数调用。

你还可以考虑设置默认流式处理值，以在发出请求但你处于脱机状态时处理案例。

> [!NOTE]
> 还有一类称为可取消函数的函数，这些函数 _与_ 流式处理函数无关。 只有返回一个值的异步自定义函数才可取消。 可取消函数允许在请求中间终止 Web 请求，它使用 [`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation) 来决定取消时需要采取的操作。 使用标记 `@cancelable` 声明可取消函数。

### <a name="use-an-invocation-parameter"></a>使用调用参数

默认情况下，`invocation` 参数是任何自定义函数的最后一个参数。 该 `invocation` 参数提供有关单元格 (的上下文，例如其地址和内容) ，并允许你使用 `setResult` 该方法和 `onCanceled` 事件来定义函数在流式传输 () `setResult` 或在)  (`onCanceled` 取消时的作用。

如果使用的是 TypeScript，则调用处理程序的类型或[`CancelableInvocation`](/javascript/api/custom-functions-runtime/customfunctions.cancelableinvocation)类型[`CustomFunctions.StreamingInvocation`](/javascript/api/custom-functions-runtime/customfunctions.streaminginvocation)。

## <a name="receiving-data-via-websockets"></a>通过 WebSocket 接收数据

在自定义函数内，可使用 [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) 来通过与服务器的持久连接交换数据。 使用 WebSocket，自定义函数可以打开与服务器的连接，然后在发生某些事件时自动从服务器接收消息，而无需显式轮询服务器中的数据。

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

## <a name="next-steps"></a>后续步骤

- 了解[你的函数可以使用的不同参数类型](custom-functions-parameter-options.md)。
- 发现如何[批处理多个 API 调用](custom-functions-batching.md)。

## <a name="see-also"></a>另请参阅

- [函数中的可变值](custom-functions-volatile.md)
- [创建自定义函数的 JSON 元数据](custom-functions-json-autogeneration.md)
- [为自定义函数手动创建 JSON 元数据](custom-functions-json.md)
- [在 Excel 中创建自定义函数](custom-functions-overview.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
