---
ms.date: 06/17/2019
description: 将自定义函数集体进行批处理，以减少对远程服务的网络调用。
title: 对远程服务的自定义函数调用进行批处理
localization_priority: Priority
ms.openlocfilehash: 2e01c981dd71a4b6eebf0e191302ba2f8f71ef2a
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059837"
---
# <a name="batching-custom-function-calls-for-a-remote-service"></a><span data-ttu-id="dea08-103">对远程服务的自定义函数调用进行批处理</span><span class="sxs-lookup"><span data-stu-id="dea08-103">Batching custom function calls for a remote service</span></span>

<span data-ttu-id="dea08-104">如果自定义函数调用远程服务，可以使用批处理模式来减少对远程服务的网络调用次数。</span><span class="sxs-lookup"><span data-stu-id="dea08-104">If your custom functions call a remote service you can use a batching pattern to reduce the number of network calls to the remote service.</span></span> <span data-ttu-id="dea08-105">为了减少网络往返，你可以将所有调用批处理为对 Web 服务的单个调用。</span><span class="sxs-lookup"><span data-stu-id="dea08-105">To reduce network round trips you batch all the calls into a single call to the web service.</span></span> <span data-ttu-id="dea08-106">当重新计算电子表格时，此方法非常合适。</span><span class="sxs-lookup"><span data-stu-id="dea08-106">This is ideal when the spreadsheet is recalculated.</span></span>

<span data-ttu-id="dea08-107">例如，如果有人在电子表格的 100 个单元格中使用了自定义函数，然后重新计算电子表格，则自定义函数将运行 100 次并进行 100 次网络调用。</span><span class="sxs-lookup"><span data-stu-id="dea08-107">For example, if someone used your custom function in 100 cells in a spreadsheet, and then recalculated the spreadsheet, your custom function would run 100 times and make 100 network calls.</span></span> <span data-ttu-id="dea08-108">通过使用批处理模式，可以将这些调用组合起来，在单次网络调用中完成总共 100 次计算。</span><span class="sxs-lookup"><span data-stu-id="dea08-108">By using a batching pattern, the calls can be combined to make all 100 calculations in a single network call.</span></span>

## <a name="view-the-completed-sample"></a><span data-ttu-id="dea08-109">查看已完成的示例</span><span class="sxs-lookup"><span data-stu-id="dea08-109">View the completed sample</span></span>

<span data-ttu-id="dea08-110">你可以按照本文操作，将代码示例粘贴到自己的项目中。</span><span class="sxs-lookup"><span data-stu-id="dea08-110">You can follow this article and paste the code examples into your own project.</span></span> <span data-ttu-id="dea08-111">例如，可以使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)为 TypeScript 创建一个新的自定义函数项目，然后将本文中的所有代码添加到该项目中。</span><span class="sxs-lookup"><span data-stu-id="dea08-111">For example, you can use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create a new custom function project for TypeScript, then add all the code from this article to the project.</span></span> <span data-ttu-id="dea08-112">然后，可以运行代码并尝试执行。</span><span class="sxs-lookup"><span data-stu-id="dea08-112">You can then run the code and try it out.</span></span>

<span data-ttu-id="dea08-113">此外，还可以在[自定义函数批处理模式](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching)处下载或查看完整的示例项目。</span><span class="sxs-lookup"><span data-stu-id="dea08-113">Also, you can download or view the complete sample project at [Custom function batching pattern](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/Batching).</span></span> <span data-ttu-id="dea08-114">如果要在进一步阅读之前查看完整代码，请查看[脚本文件](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts)。</span><span class="sxs-lookup"><span data-stu-id="dea08-114">If you want to view the code in whole before reading any further, take a look at the [script file](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Excel-custom-functions/Batching/src/functions/functions.ts).</span></span>

## <a name="create-the-batching-pattern-in-this-article"></a><span data-ttu-id="dea08-115">创建本文中所述的批处理模式</span><span class="sxs-lookup"><span data-stu-id="dea08-115">Create the batching pattern in this article</span></span>

<span data-ttu-id="dea08-116">若要为自定义函数设置批处理，需要编写三个主要的代码部分。</span><span class="sxs-lookup"><span data-stu-id="dea08-116">To set up batching for your custom functions you'll need to write three main sections of code.</span></span>

1. <span data-ttu-id="dea08-117">用以在 Excel 每次调用自定义函数时向一批调用添加新运算的推送运算。</span><span class="sxs-lookup"><span data-stu-id="dea08-117">A push operation to add a new operation to the batch of calls each time Excel calls your custom function.</span></span>
2. <span data-ttu-id="dea08-118">用以在批处理就绪时发出远程请求的函数。</span><span class="sxs-lookup"><span data-stu-id="dea08-118">A function to make the remote request when the batch is ready.</span></span>
3. <span data-ttu-id="dea08-119">用以响应批处理请求、计算所有运算结果并返回值的服务器代码。</span><span class="sxs-lookup"><span data-stu-id="dea08-119">Server code to respond to the batch request, calculate all of the operation results, and return the values.</span></span>

<span data-ttu-id="dea08-120">以下部分将为你展示如何构造代码（每次一个示例）。</span><span class="sxs-lookup"><span data-stu-id="dea08-120">In the following sections you will be shown how to construct the code one example at a time.</span></span> <span data-ttu-id="dea08-121">你将把各个代码示例添加到 **functions.ts** 文件中。</span><span class="sxs-lookup"><span data-stu-id="dea08-121">You'll add each code example to your **functions.ts** file.</span></span> <span data-ttu-id="dea08-122">建议使用 yo office 生成器创建全新的自定义函数项目。</span><span class="sxs-lookup"><span data-stu-id="dea08-122">It's recommended you create a brand new custom functions project using the Yo Office generator.</span></span> <span data-ttu-id="dea08-123">若要创建新项目，请参阅[开始开发 Excel 自定义函数](../quickstarts/excel-custom-functions-quickstart.md)并使用 TypeScript，而不是 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="dea08-123">To create a new project see [Get started developing Excel custom functions](../quickstarts/excel-custom-functions-quickstart.md) and use TypeScript instead of JavaScript.</span></span>

## <a name="batch-each-call-to-your-custom-function"></a><span data-ttu-id="dea08-124">批处理对自定义函数的每次调用</span><span class="sxs-lookup"><span data-stu-id="dea08-124">Batch each call to your custom function</span></span>

<span data-ttu-id="dea08-125">自定义函数通过调用远程服务来执行运算并计算其所需的结果。</span><span class="sxs-lookup"><span data-stu-id="dea08-125">Your custom functions work by calling a remote service to perform the operation and calculate the result they need.</span></span> <span data-ttu-id="dea08-126">这为它们提供了一种将每个请求的运算存储到批处理中的方法。</span><span class="sxs-lookup"><span data-stu-id="dea08-126">This provides a way for them to store each requested operation into a batch.</span></span> <span data-ttu-id="dea08-127">稍后，你将看到如何创建 `_pushOperation` 函数来批处理这些运算。</span><span class="sxs-lookup"><span data-stu-id="dea08-127">Later you'll see how to create a `_pushOperation` function to batch the operations.</span></span> <span data-ttu-id="dea08-128">首先，看看下面的代码示例，以了解如何从自定义函数调用 `_pushOperation`。</span><span class="sxs-lookup"><span data-stu-id="dea08-128">First, take a look at the following code example to see how to call `_pushOperation` from your custom function.</span></span>

<span data-ttu-id="dea08-129">在下面的代码中，自定义函数执行除法，但实际计算依赖于远程服务。</span><span class="sxs-lookup"><span data-stu-id="dea08-129">In the following code, the custom function performs division but relies on a remote service to do the actual calculation.</span></span> <span data-ttu-id="dea08-130">它调用 `_pushOperation`，从而将该运算与其他运算一起批处理到远程服务。</span><span class="sxs-lookup"><span data-stu-id="dea08-130">It calls `_pushOperation` to batch the operation along with other operations to the remote service.</span></span> <span data-ttu-id="dea08-131">它将该运算命名为“div2”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dea08-131">It names the operation **div2**.</span></span> <span data-ttu-id="dea08-132">你可以为运算使用任何所需的命名方案，只要远程服务也使用相同的方案即可（稍后将对远程服务方面进行详细介绍）。</span><span class="sxs-lookup"><span data-stu-id="dea08-132">You can use any naming scheme you want for operations as long as the remote service is also using the same scheme (more on the remote service later).</span></span> <span data-ttu-id="dea08-133">此外，还将传递远程服务运行该运算所需的参数。</span><span class="sxs-lookup"><span data-stu-id="dea08-133">Also, the arguments the remote service will need to run the operation are passed.</span></span>

### <a name="add-the-div2-custom-function-to-functionsts"></a><span data-ttu-id="dea08-134">将 div2 自定义函数添加到 functions.ts</span><span class="sxs-lookup"><span data-stu-id="dea08-134">Add the div2 custom function to functions.ts</span></span>

```typescript
/**
 * @CustomFunction
 * Divides two numbers using batching
 * @param dividend The number being divided
 * @param divisor The number the dividend is divided by
 * @returns The result of dividing the two numbers
 */
function div2(dividend: number, divisor: number) {
  return _pushOperation(
    "div2",
    [dividend, divisor]
  );
}

CustomFunctions.associate("DIV2", div2);
```

<span data-ttu-id="dea08-135">接下来，你将定义批处理数组，该数组将存储要在一个网络调用中传递的所有运算。</span><span class="sxs-lookup"><span data-stu-id="dea08-135">Next, you will define the batch array which will store all operations to be passed in one network call.</span></span> <span data-ttu-id="dea08-136">以下代码展示了如何定义描述数组中每个批处理条目的接口。</span><span class="sxs-lookup"><span data-stu-id="dea08-136">The following code shows how to define an interface describing each batch entry in the array.</span></span> <span data-ttu-id="dea08-137">接口定义了一个运算，是要运行的运算的字符串名称。</span><span class="sxs-lookup"><span data-stu-id="dea08-137">The interface defines an operation, which is a string name of which operation to run.</span></span> <span data-ttu-id="dea08-138">例如，如果有两个分别名为 `multiply` 和 `divide` 的自定义函数，则可以在批处理条目中将它们作为运算名称重复使用。</span><span class="sxs-lookup"><span data-stu-id="dea08-138">For example, if you had two custom functions named `multiply` and `divide`, you could reuse those as the operation names in your batch entries.</span></span> <span data-ttu-id="dea08-139">`args` 将保留从 Excel 传递到自定义函数的参数。</span><span class="sxs-lookup"><span data-stu-id="dea08-139">`args` will hold the arguments that were passed to your custom function from Excel.</span></span> <span data-ttu-id="dea08-140">最后，`resolve` 或 `reject` 将存储一个承诺，其中存有远程服务返回的信息。</span><span class="sxs-lookup"><span data-stu-id="dea08-140">And finally, `resolve` or `reject` will store a promise holding the information the remote service returns.</span></span>

```typescript
interface IBatchEntry {
  operation: string;
  args: any[];
  resolve: (data: any) => void;
  reject: (error: Error) => void;
}
```

<span data-ttu-id="dea08-141">接下来，创建使用上一个接口的批处理数组。</span><span class="sxs-lookup"><span data-stu-id="dea08-141">Next, create the batch array that uses the previous interface.</span></span> <span data-ttu-id="dea08-142">若要跟踪是否已安排某个批处理，请创建一个 `_isBatchedRequestSchedule` 变量。</span><span class="sxs-lookup"><span data-stu-id="dea08-142">To track if a batch is scheduled or not, create an `_isBatchedRequestSchedule` variable.</span></span> <span data-ttu-id="dea08-143">这一点在稍后对远程服务的批处理调用进行计时时很重要。</span><span class="sxs-lookup"><span data-stu-id="dea08-143">This will be important later for timing batch calls to the remote service.</span></span>

```typescript
const _batch: IBatchEntry[] = [];
let _isBatchedRequestScheduled = false;
```

<span data-ttu-id="dea08-144">最后，当 Excel 调用自定义函数时，你需要将该运算推送到批处理数组中。</span><span class="sxs-lookup"><span data-stu-id="dea08-144">Finally when Excel calls your custom function, you need to push the operation into the batch array.</span></span> <span data-ttu-id="dea08-145">以下代码展示了如何从自定义函数添加新运算。</span><span class="sxs-lookup"><span data-stu-id="dea08-145">The following code shows how to add a new operation from a custom function.</span></span> <span data-ttu-id="dea08-146">它会创建一个新的批处理条目，创建一个新的承诺来解决或拒绝相应运算，并将该条目推送到批处理数组中。</span><span class="sxs-lookup"><span data-stu-id="dea08-146">It creates a new batch entry, creates a new promise to resolve or reject the operation, and pushes the entry into the batch array.</span></span>

<span data-ttu-id="dea08-147">此段代码还会检查是否对批处理进行了安排。</span><span class="sxs-lookup"><span data-stu-id="dea08-147">This code also checks to see if a batch is scheduled.</span></span> <span data-ttu-id="dea08-148">在本例中，将每个批处理安排为每 100 毫秒运行一次。</span><span class="sxs-lookup"><span data-stu-id="dea08-148">In this example, each batch is scheduled to run every 100ms.</span></span> <span data-ttu-id="dea08-149">可以根据需要调整此值。</span><span class="sxs-lookup"><span data-stu-id="dea08-149">You can adjust this value as needed.</span></span> <span data-ttu-id="dea08-150">值越大，发送到远程服务的批处理越大，用户查看结果的等待时间越长。</span><span class="sxs-lookup"><span data-stu-id="dea08-150">Higher values result in bigger batches being sent to the remote service, and a longer wait time for the user to see results.</span></span> <span data-ttu-id="dea08-151">较低的值倾向于向远程服务发送更多的批处理，但可为用户提供较快的响应时间。</span><span class="sxs-lookup"><span data-stu-id="dea08-151">Lower values tend to send more batches to the remote service, but with a quick response time for users.</span></span>

### <a name="add-the-pushoperation-function-to-functionsts"></a><span data-ttu-id="dea08-152">将 `_pushOperation` 函数添加到 functions.ts</span><span class="sxs-lookup"><span data-stu-id="dea08-152">Add the `_pushOperation` function to functions.ts</span></span>

```typescript
function _pushOperation(op: string, args: any[]) {
  // Create an entry for your custom function.
  const invocationEntry: IBatchEntry = {
    operation: op, // e.g. sum
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
  // schedule it after a certain timeout, e.g. 100 ms.
  if (!_isBatchedRequestScheduled) {
    _isBatchedRequestScheduled = true;
    setTimeout(_makeRemoteRequest, 100);
  }

  // Return the promise for this invocation.
  return promise;
}
```

## <a name="make-the-remote-request"></a><span data-ttu-id="dea08-153">发出远程请求</span><span class="sxs-lookup"><span data-stu-id="dea08-153">Make the remote request</span></span>

<span data-ttu-id="dea08-154">`_makeRemoteRequest` 函数的目的是将一批运算传递给远程服务，然后将结果返回给每个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="dea08-154">The purpose of the `_makeRemoteRequest` function is to pass the batch of operations to the remote service, and then return the results to each custom function.</span></span> <span data-ttu-id="dea08-155">它首先创建批处理数组的副本。</span><span class="sxs-lookup"><span data-stu-id="dea08-155">It first creates a copy of the batch array.</span></span> <span data-ttu-id="dea08-156">这样，来自 Excel 的并发自定义函数调用便可以立即在新数组中开始批处理。</span><span class="sxs-lookup"><span data-stu-id="dea08-156">This allows concurrent custom function calls from Excel to immediately begin batching in a new array.</span></span> <span data-ttu-id="dea08-157">然后将副本转换为不包含承诺信息的较简单的数组。</span><span class="sxs-lookup"><span data-stu-id="dea08-157">The copy is then turned into a simpler array that does not contain the promise information.</span></span> <span data-ttu-id="dea08-158">将这些承诺传递给远程服务是没有意义的，因为它们不会发生作用。</span><span class="sxs-lookup"><span data-stu-id="dea08-158">It wouldn't make sense to pass the promises to a remote service since they would not work.</span></span> <span data-ttu-id="dea08-159">`_makeRemoteRequest` 将根据远程服务返回的内容拒绝或解决每个承诺。</span><span class="sxs-lookup"><span data-stu-id="dea08-159">The `_makeRemoteRequest` will either reject or resolve each promise based on what the remote service returns.</span></span>

### <a name="add-the-following-makeremoterequest-method-to-functionsts"></a><span data-ttu-id="dea08-160">将以下 `_makeRemoteRequest` 方法添加到 functions.ts</span><span class="sxs-lookup"><span data-stu-id="dea08-160">Add the following `_makeRemoteRequest` method to functions.ts</span></span>

```typescript
function _makeRemoteRequest() {
  // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
  // Note the use of "splice" rather than "slice", which will modify the original _batch array
  // to empty it out.
  const batchCopy = _batch.splice(0, _batch.length);
  _isBatchedRequestScheduled = false;

  // Build a simpler request batch that only contains the arguments for each invocation.
  const requestBatch = batchCopy.map((item) => {
    return { operation: item.operation, args: item.args };
  });

  // Make the remote request.
  _fetchFromRemoteService(requestBatch)
    .then((responseBatch) => {
      // Match each value from the response batch to its corresponding invocation entry from the request batch,
      // and resolve the invocation promise with its corresponding response value.
      responseBatch.forEach((response, index) => {
        if (response.error) {
          batchCopy[index].reject(new Error(response.error));
        } else {
          console.log(response);
          batchCopy[index].resolve(response.result);
        }
      });
    });
}
```

### <a name="modify-makeremoterequest-for-your-own-solution"></a><span data-ttu-id="dea08-161">根据自己的解决方案修改 `_makeRemoteRequest`</span><span class="sxs-lookup"><span data-stu-id="dea08-161">Modify `_makeRemoteRequest` for your own solution</span></span>

<span data-ttu-id="dea08-162">`_makeRemoteRequest` 函数调用 `_fetchFromRemoteService`，正如稍后将会看到的，后者只是一个表示远程服务的模拟。</span><span class="sxs-lookup"><span data-stu-id="dea08-162">The `_makeRemoteRequest` function calls `_fetchFromRemoteService` which, as you'll see later, is just a mock representing the remote service.</span></span> <span data-ttu-id="dea08-163">这使得研究和运行本文中的代码更加容易。</span><span class="sxs-lookup"><span data-stu-id="dea08-163">This makes it easier to study and run the code in this article.</span></span> <span data-ttu-id="dea08-164">但是，如果要将此代码用于实际的远程服务，则应进行以下更改：</span><span class="sxs-lookup"><span data-stu-id="dea08-164">But when you want to use this code for an actual remote service you should make the following changes:</span></span>

- <span data-ttu-id="dea08-165">决定如何通过网络将批处理运算序列化。</span><span class="sxs-lookup"><span data-stu-id="dea08-165">Decide how to serialize the batch operations over the network.</span></span> <span data-ttu-id="dea08-166">例如，你可能希望将数组放入 JSON 主体中。</span><span class="sxs-lookup"><span data-stu-id="dea08-166">For example, you may want to put the array into a JSON body.</span></span>
- <span data-ttu-id="dea08-167">你不需要调用 `_fetchFromRemoteService`，而是需要对传递批量运算的远程服务进行实际的网络调用。</span><span class="sxs-lookup"><span data-stu-id="dea08-167">Instead of calling `_fetchFromRemoteService` you need to make the actual network call to the remote service passing the batch of operations.</span></span>

## <a name="process-the-batch-call-on-the-remote-service"></a><span data-ttu-id="dea08-168">处理远程服务上的批处理调用</span><span class="sxs-lookup"><span data-stu-id="dea08-168">Process the batch call on the remote service</span></span>

<span data-ttu-id="dea08-169">最后一步是处理远程服务中的批处理调用。</span><span class="sxs-lookup"><span data-stu-id="dea08-169">The last step is to handle the batch call in the remote service.</span></span> <span data-ttu-id="dea08-170">下面的代码示例展示了 `_fetchFromRemoteService` 函数。</span><span class="sxs-lookup"><span data-stu-id="dea08-170">The following code sample shows the `_fetchFromRemoteService` function.</span></span> <span data-ttu-id="dea08-171">此函数会解包每个运算，执行指定的运算，并返回结果。</span><span class="sxs-lookup"><span data-stu-id="dea08-171">This function unpacks each operation, performs the specified operation, and returns the results.</span></span> <span data-ttu-id="dea08-172">出于学习目的，在本文中，`_fetchFromRemoteService` 函数适用于在 Web 加载项中运行并模拟远程服务。</span><span class="sxs-lookup"><span data-stu-id="dea08-172">For learning purposes in this article, the `_fetchFromRemoteService` function is designed to run in your web add-in and mock a remote service.</span></span> <span data-ttu-id="dea08-173">你可以将此代码添加到 **functions.ts** 文件中，这样就可以研究和运行本文中的所有代码，而无需设置实际的远程服务。</span><span class="sxs-lookup"><span data-stu-id="dea08-173">You can add this code to your **functions.ts** file so that you can study and run all the code in this article without having to set up an actual remote service.</span></span>

### <a name="add-the-following-fetchfromremoteservice-function-to-functionsts"></a><span data-ttu-id="dea08-174">将以下 `_fetchFromRemoteService` 函数添加到 functions.ts</span><span class="sxs-lookup"><span data-stu-id="dea08-174">Add the following `_fetchFromRemoteService` function to functions.ts</span></span>

```typescript
async function _fetchFromRemoteService(
  requestBatch: Array<{ operation: string, args: any[] }>
): Promise<IServerResponse[]> {
  // Simulate a slow network request to the server;
  await pause(1000);

  return requestBatch.map((request): IServerResponse => {
    const { operation, args } = request;

    try {
      if (operation === "div2") {
        // Divide the first argument by the second argument.
        return {
          result: args[0] / args[1]
        };
      } else if (operation === "mul2") {
        // Multiply the arguments for the given entry.
        const myresult = args[0] * args[1];
        console.log(myresult);
        return {
          result: myresult
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

function pause(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
```

### <a name="modify-fetchfromremoteservice-for-your-live-remote-service"></a><span data-ttu-id="dea08-175">根据自己的实时远程服务修改 `_fetchFromRemoteService`</span><span class="sxs-lookup"><span data-stu-id="dea08-175">Modify `_fetchFromRemoteService` for your live remote service</span></span>

<span data-ttu-id="dea08-176">若要修改 `_fetchFromRemoteService` 以便在实时远程服务中运行，请进行以下更改：</span><span class="sxs-lookup"><span data-stu-id="dea08-176">To modify the `_fetchFromRemoteService` function to run in your live remote service, make the following changes:</span></span>

- <span data-ttu-id="dea08-177">根据服务器平台（Node.js 或其他平台），将客户端网络调用映射到此函数。</span><span class="sxs-lookup"><span data-stu-id="dea08-177">Depending on your server platform (Node.js or others) map the client network call to this function.</span></span>
- <span data-ttu-id="dea08-178">删除作为模拟的一部分来模拟网络延迟的 `pause` 函数。</span><span class="sxs-lookup"><span data-stu-id="dea08-178">Remove the `pause` function which simulates network latency as part of the mock.</span></span>
- <span data-ttu-id="dea08-179">修改函数声明，以便在出于网络目的更改传递的参数时使用该参数。</span><span class="sxs-lookup"><span data-stu-id="dea08-179">Modify the function declaration to work with the parameter passed if the parameter is changed for network purposes.</span></span> <span data-ttu-id="dea08-180">例如，它可能是要处理的批量运算的 JSON 主体，而不是数组。</span><span class="sxs-lookup"><span data-stu-id="dea08-180">For example, instead of an array, it may be a JSON body of batched operations to process.</span></span>
- <span data-ttu-id="dea08-181">修改函数以执行运算（或调用执行运算的函数）。</span><span class="sxs-lookup"><span data-stu-id="dea08-181">Modify the function to perform the operations (or call functions that do the operations).</span></span>
- <span data-ttu-id="dea08-182">应用相应的身份验证机制。</span><span class="sxs-lookup"><span data-stu-id="dea08-182">Apply an appropriate authentication mechanism.</span></span> <span data-ttu-id="dea08-183">确保只有正确的调用者才可以访问该函数。</span><span class="sxs-lookup"><span data-stu-id="dea08-183">Ensure that only the correct callers can access the function.</span></span>
- <span data-ttu-id="dea08-184">将代码放入远程服务。</span><span class="sxs-lookup"><span data-stu-id="dea08-184">Place the code in the remote service.</span></span>

## <a name="next-steps"></a><span data-ttu-id="dea08-185">后续步骤</span><span class="sxs-lookup"><span data-stu-id="dea08-185">Next steps</span></span>
<span data-ttu-id="dea08-186">了解可以在自定义函数中使用的[各种参数](custom-functions-parameter-options.md)。</span><span class="sxs-lookup"><span data-stu-id="dea08-186">Learn about [the various parameters](custom-functions-parameter-options.md) you can use in your custom functions.</span></span> <span data-ttu-id="dea08-187">或者查阅[通过自定义函数进行 Web 调用](custom-functions-web-reqs.md)后面的基础知识。</span><span class="sxs-lookup"><span data-stu-id="dea08-187">Or review the basics behind making [a web call through a custom function](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="dea08-188">另请参阅</span><span class="sxs-lookup"><span data-stu-id="dea08-188">See also</span></span>

* [<span data-ttu-id="dea08-189">函数中的可变值</span><span class="sxs-lookup"><span data-stu-id="dea08-189">Volatile values in functions</span></span>](custom-functions-volatile.md)
* [<span data-ttu-id="dea08-190">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="dea08-190">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="dea08-191">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="dea08-191">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="dea08-192">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="dea08-192">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
