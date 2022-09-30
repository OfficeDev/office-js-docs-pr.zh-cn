---
ms.date: 09/09/2022
description: 将自定义函数集体进行批处理，以减少对远程服务的网络调用。
title: 对远程服务的自定义函数调用进行批处理
ms.localizationpriority: medium
ms.openlocfilehash: f779351789350bbc591b1b5d7a975ff9f70cda26
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234919"
---
# <a name="batch-custom-function-calls-for-a-remote-service"></a>远程服务的 Batch 自定义函数调用

如果自定义函数调用远程服务，可以使用批处理模式来减少对远程服务的网络调用次数。 为了减少网络往返，你可以将所有调用批处理为对 Web 服务的单个调用。 当重新计算电子表格时，此方法非常合适。

例如，如果有人在电子表格的 100 个单元格中使用了自定义函数，然后重新计算电子表格，则自定义函数将运行 100 次并进行 100 次网络调用。 通过使用批处理模式，可以将这些调用组合起来，在单次网络调用中完成总共 100 次计算。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="view-the-completed-sample"></a>查看已完成的示例

若要查看已完成的示例，请按照本文操作，并将代码示例粘贴到自己的项目中。 例如，若要为 TypeScript 创建新的自定义函数项目，请使用 [Office 外接程序的 Yeoman 生成器](../develop/yeoman-generator-overview.md)，然后将本文中的所有代码添加到项目。 运行代码并试用。

或者，下载或查看 [自定义函数批处理模式下的](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching)完整示例项目。 如果要在进一步阅读之前查看完整代码，请查看[脚本文件](https://github.com/OfficeDev/Office-Add-in-samples/blob/main/Excel-custom-functions/Batching/src/functions/functions.js)。

## <a name="create-the-batching-pattern-in-this-article"></a>创建本文中所述的批处理模式

若要为自定义函数设置批处理，需要编写三个主要的代码部分。

1. [一个推送操作](#add-the-_pushoperation-function)，用于在每次 Excel 调用自定义函数时向一批调用添加新操作。
2. 一个 [函数，用于在](#make-the-remote-request) 批处理准备就绪时发出远程请求。
3. [服务器代码响应批处理请求](#process-the-batch-call-on-the-remote-service)，计算所有操作结果，并返回值。

在以下部分中，你将了解如何一次构造一个示例的代码。 建议使用 [适用于 Office 外接程序生成器的 Yeoman 生成器](../develop/yeoman-generator-overview.md) 创建全新的自定义函数项目。 若要创建新项目，请参阅 [开始开发 Excel 自定义函数](../quickstarts/excel-custom-functions-quickstart.md)。 可以使用 TypeScript 或 JavaScript。

## <a name="batch-each-call-to-your-custom-function"></a>批处理对自定义函数的每次调用

自定义函数通过调用远程服务来执行运算并计算其所需的结果。 这为它们提供了一种将每个请求的运算存储到批处理中的方法。 稍后，你将看到如何创建 `_pushOperation` 函数来批处理这些运算。 首先，看看下面的代码示例，以了解如何从自定义函数调用 `_pushOperation`。

在下面的代码中，自定义函数执行除法，但实际计算依赖于远程服务。 它调用 `_pushOperation`，从而将该运算与其他运算一起批处理到远程服务。 它将该运算命名为“div2”。 你可以为运算使用任何所需的命名方案，只要远程服务也使用相同的方案即可（稍后将对远程服务方面进行详细介绍）。 此外，还将传递远程服务运行该运算所需的参数。

### <a name="add-the-div2-custom-function"></a>添加 div2 自定义函数

将以下代码添加到 **functions.js** 或 **functions.ts** 文件 (，具体取决于使用 JavaScript 还是 TypeScript) 。

```javascript
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend, divisor) {
  return _pushOperation("div2", [dividend, divisor]);
}
```

### <a name="add-global-variables-for-tracking-batch-requests"></a>添加用于跟踪批处理请求的全局变量

接下来，将两个全局变量添加到 **functions.js** 或 **functions.ts** 文件。 `_isBatchedRequestScheduled` 稍后对远程服务的批处理调用计时非常重要。

```javascript
let _batch = [];
let _isBatchedRequestScheduled = false;
```

### <a name="add-the-_pushoperation-function"></a>添加函数`_pushOperation`

当 Excel 调用自定义函数时，需要将操作推送到批处理数组。 以下 **_pushOperation函数** 代码演示如何从自定义函数添加新操作。 它会创建一个新的批处理条目，创建一个新的承诺来解决或拒绝相应运算，并将该条目推送到批处理数组中。

此段代码还会检查是否对批处理进行了安排。 在本例中，将每个批处理安排为每 100 毫秒运行一次。 可以根据需要调整此值。 值越大，发送到远程服务的批处理越大，用户查看结果的等待时间越长。 较低的值倾向于向远程服务发送更多的批处理，但可为用户提供较快的响应时间。

该函数创建一个 **invocationEntry** 对象，该对象包含要运行的操作的字符串名称。 例如，如果有两个分别名为 `multiply` 和 `divide` 的自定义函数，则可以在批处理条目中将它们作为运算名称重复使用。 `args` 保存从 Excel 传递到自定义函数的参数。 最后， `resolve` 或 `reject` 方法存储保存远程服务返回的信息的承诺。

将以下代码添加到 **functions.js** 或 **functions.ts** 文件。

```javascript
// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
  // Create an entry for your custom function.
  console.log("pushOperation");
  const invocationEntry = {
    operation: op, // e.g., sum
    args: args,
    resolve: undefined,
    reject: undefined,
  };

  // Create a unique promise for this invocation,
  // and save its resolve and reject functions into the invocation entry.
  const promise = new Promise((resolve, reject) => {
    invocationEntry.resolve = resolve;
    invocationEntry.reject = reject;
  });

  // Push the invocation entry into the next batch.
  _batch.push(invocationEntry);

  // If a remote request hasn't been scheduled yet,
  // schedule it after a certain timeout, e.g., 100 ms.
  if (!_isBatchedRequestScheduled) {
    console.log("schedule remote request");
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a>发出远程请求

`_makeRemoteRequest` 函数的目的是将一批运算传递给远程服务，然后将结果返回给每个自定义函数。 它首先创建批处理数组的副本。 这样，来自 Excel 的并发自定义函数调用便可以立即在新数组中开始批处理。 然后将副本转换为不包含承诺信息的较简单的数组。 将这些承诺传递给远程服务是没有意义的，因为它们不会发生作用。 `_makeRemoteRequest` 将根据远程服务返回的内容拒绝或解决每个承诺。

将以下代码添加到 **functions.js** 或 **functions.ts** 文件。

```javascript
// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  try{
  console.log("makeRemoteRequest");
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });
  console.log("makeRemoteRequest2");
  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      console.log("responseBatch in fetchFromRemoteService");
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
          console.log("rejecting promise");
        } else {
          console.log("fulfilling promise");
          console.log(response);

          batchCopy[index].resolve(response.result);
        }
      });
    });
    console.log("makeRemoteRequest3");
  } catch (error) {
    console.log("error name:" + error.name);
    console.log("error message:" + error.message);
    console.log(error);
  }
}
```

### <a name="modify-_makeremoterequest-for-your-own-solution"></a>根据自己的解决方案修改 `_makeRemoteRequest`

`_makeRemoteRequest` 函数调用 `_fetchFromRemoteService`，正如稍后将会看到的，后者只是一个表示远程服务的模拟。 这使得研究和运行本文中的代码更加容易。 但是，如果要将此代码用于实际远程服务，应进行以下更改。

- 决定如何通过网络将批处理运算序列化。 例如，你可能希望将数组放入 JSON 主体中。
- 你不需要调用 `_fetchFromRemoteService`，而是需要对传递批量运算的远程服务进行实际的网络调用。

## <a name="process-the-batch-call-on-the-remote-service"></a>处理远程服务上的批处理调用

最后一步是处理远程服务中的批处理调用。 下面的代码示例展示了 `_fetchFromRemoteService` 函数。 此函数会解包每个运算，执行指定的运算，并返回结果。 出于学习目的，在本文中，`_fetchFromRemoteService` 函数适用于在 Web 加载项中运行并模拟远程服务。 可以将此代码添加到 **functions.js** 或 **functions.ts** 文件，以便可以研究并运行本文中的所有代码，而无需设置实际的远程服务。

将以下代码添加到 **functions.js** 或 **functions.ts** 文件。

```javascript
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a promise that may contain a batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
async function _fetchFromRemoteService(requestBatch) {
  // Simulate a slow network request to the server.
  console.log("_fetchFromRemoteService");
  await pause(1000);
  console.log("postpause");
  return requestBatch.map((request) => {
    console.log("requestBatch server side");
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myResult = args[0] * args[1];
        console.log(myResult);
        return {
          result: myResult
        };
      } else {
        return {
          error: `Operation not supported: ${operation}`
        };
      }
    } catch (error) {
      return {
        error: `Operation failed: ${operation}`
      };
    }
  });
}

function pause(ms) {
  console.log("pause");
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-_fetchfromremoteservice-for-your-live-remote-service"></a>根据自己的实时远程服务修改 `_fetchFromRemoteService`

若要修改要 `_fetchFromRemoteService` 在实时远程服务中运行的函数，请进行以下更改。

- 根据服务器平台（Node.js 或其他平台），将客户端网络调用映射到此函数。
- 删除作为模拟的一部分来模拟网络延迟的 `pause` 函数。
- 修改函数声明，以便在出于网络目的更改传递的参数时使用该参数。 例如，它可能是要处理的批量运算的 JSON 主体，而不是数组。
- 修改函数以执行运算（或调用执行运算的函数）。
- 应用相应的身份验证机制。 确保只有正确的调用者才可以访问该函数。
- 将代码放入远程服务。

## <a name="next-steps"></a>后续步骤

了解可以在自定义函数中使用的[各种参数](custom-functions-parameter-options.md)。 或者查阅[通过自定义函数进行 Web 调用](custom-functions-web-reqs.md)后面的基础知识。

## <a name="see-also"></a>另请参阅

- [函数中的可变值](custom-functions-volatile.md)
- [在 Excel 中创建自定义函数](custom-functions-overview.md)
- [Excel 自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)
