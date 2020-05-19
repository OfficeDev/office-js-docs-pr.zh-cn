---
ms.date: 05/06/2020
description: 在 Excel 中定义自定义函数的 JSON 元数据，并将您的函数 id 和 name 属性相关联。
title: Excel 中自定义函数的元数据
localization_priority: Normal
ms.openlocfilehash: 848a65a0eda7b8cfd6a28df16b44dbbfc207c7b9
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275992"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="28209-103">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="28209-103">Custom functions metadata</span></span>

<span data-ttu-id="28209-104">如 "[自定义函数概述](custom-functions-overview.md)" 一文中所述，自定义函数项目必须包括 JSON 元数据文件和脚本（JavaScript 或 TypeScript）文件才能注册函数，使其可供使用。</span><span class="sxs-lookup"><span data-stu-id="28209-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="28209-105">自定义函数在用户首次运行外接程序且在所有工作簿中对同一用户可用时注册。</span><span class="sxs-lookup"><span data-stu-id="28209-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="28209-106">我们建议在可能的情况（而不是创建自己的 JSON 文件）中使用 JSON 自动生成。</span><span class="sxs-lookup"><span data-stu-id="28209-106">We recommend using JSON auto-generation when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="28209-107">自动生成不易出现用户错误， `yo office` 搭建文件已包含此文件。</span><span class="sxs-lookup"><span data-stu-id="28209-107">Auto-generation is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="28209-108">有关 JSDoc 注释 JSON 文件生成的过程的详细信息，请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="28209-108">For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="28209-109">不过，您可以从头开始创建自定义函数项目;它要求您执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="28209-109">However, you can make a custom functions project from scratch; it requires that you:</span></span>

- <span data-ttu-id="28209-110">编写 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="28209-110">Write your JSON file.</span></span>
- <span data-ttu-id="28209-111">检查您的清单文件是否已连接到您的 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="28209-111">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="28209-112">`id`在脚本文件中关联函数和 `name` 属性，以便注册您的函数</span><span class="sxs-lookup"><span data-stu-id="28209-112">Associate your functions' `id` and `name` properties in the script file in order to register your functions</span></span>

<span data-ttu-id="28209-113">下图说明了使用 `yo office` 搭建文件和从草稿写入 JSON 之间的差异。</span><span class="sxs-lookup"><span data-stu-id="28209-113">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>
<span data-ttu-id="28209-114">![使用 Yo 办公室和编写自己的 JSON 的差异的图像](../images/custom-functions-json.png)</span><span class="sxs-lookup"><span data-stu-id="28209-114">![Image of differences between using Yo Office and writing your own JSON](../images/custom-functions-json.png)</span></span>

> [!NOTE]
> <span data-ttu-id="28209-115">请记住，如果不使用生成器，请将清单连接到您创建的 JSON 文件，并通过 `<Resources>` XML 清单文件中的节 `yo office` 。</span><span class="sxs-lookup"><span data-stu-id="28209-115">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="28209-116">创作元数据并连接到清单</span><span class="sxs-lookup"><span data-stu-id="28209-116">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="28209-117">在项目中创建一个 JSON 文件，并提供有关函数中的函数的所有详细信息，如函数的参数。</span><span class="sxs-lookup"><span data-stu-id="28209-117">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="28209-118">有关函数属性的完整列表，请参阅[以下元数据示例](#json-metadata-example)和[元数据参考](#metadata-reference)。</span><span class="sxs-lookup"><span data-stu-id="28209-118">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="28209-119">请确保您的 XML 清单文件引用了节中的 JSON 文件 `<Resources>` ，如下例所示。</span><span class="sxs-lookup"><span data-stu-id="28209-119">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a><span data-ttu-id="28209-120">JSON 元数据示例</span><span class="sxs-lookup"><span data-stu-id="28209-120">JSON metadata example</span></span>

<span data-ttu-id="28209-121">以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="28209-121">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="28209-122">此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="28209-122">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "first",
          "description": "first number to add",
          "type": "number",
          "dimensionality": "scalar"
        },
        {
          "name": "second",
          "description": "second number to add",
          "type": "number",
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE",
      "description": "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST",
      "description": "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> <span data-ttu-id="28209-123">[OfficeDev/Excel 自定义函数](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)GitHub 存储库的提交历史记录中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="28209-123">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="28209-124">随着项目已调整为自动生成 JSON，手写 JSON 的完整示例仅在项目的早期版本中可用。</span><span class="sxs-lookup"><span data-stu-id="28209-124">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="28209-125">元数据参考</span><span class="sxs-lookup"><span data-stu-id="28209-125">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="28209-126">functions</span><span class="sxs-lookup"><span data-stu-id="28209-126">functions</span></span>

<span data-ttu-id="28209-127">`functions` 属性是自定义函数对象的一个数组。</span><span class="sxs-lookup"><span data-stu-id="28209-127">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="28209-128">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="28209-128">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="28209-129">属性</span><span class="sxs-lookup"><span data-stu-id="28209-129">Property</span></span>      | <span data-ttu-id="28209-130">数据类型</span><span class="sxs-lookup"><span data-stu-id="28209-130">Data type</span></span> | <span data-ttu-id="28209-131">必需</span><span class="sxs-lookup"><span data-stu-id="28209-131">Required</span></span> | <span data-ttu-id="28209-132">说明</span><span class="sxs-lookup"><span data-stu-id="28209-132">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="28209-133">string</span><span class="sxs-lookup"><span data-stu-id="28209-133">string</span></span>    | <span data-ttu-id="28209-134">否</span><span class="sxs-lookup"><span data-stu-id="28209-134">No</span></span>       | <span data-ttu-id="28209-135">最终用户在 Excel 中看到的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="28209-135">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="28209-136">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="28209-136">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="28209-137">string</span><span class="sxs-lookup"><span data-stu-id="28209-137">string</span></span>    | <span data-ttu-id="28209-138">否</span><span class="sxs-lookup"><span data-stu-id="28209-138">No</span></span>       | <span data-ttu-id="28209-139">提供有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="28209-139">URL that provides information about the function.</span></span> <span data-ttu-id="28209-140">（它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。</span><span class="sxs-lookup"><span data-stu-id="28209-140">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="28209-141">string</span><span class="sxs-lookup"><span data-stu-id="28209-141">string</span></span>    | <span data-ttu-id="28209-142">是</span><span class="sxs-lookup"><span data-stu-id="28209-142">Yes</span></span>      | <span data-ttu-id="28209-143">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="28209-143">A unique ID for the function.</span></span> <span data-ttu-id="28209-144">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="28209-144">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="28209-145">string</span><span class="sxs-lookup"><span data-stu-id="28209-145">string</span></span>    | <span data-ttu-id="28209-146">是</span><span class="sxs-lookup"><span data-stu-id="28209-146">Yes</span></span>      | <span data-ttu-id="28209-147">最终用户在 Excel 中看到的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="28209-147">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="28209-148">在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="28209-148">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="28209-149">object</span><span class="sxs-lookup"><span data-stu-id="28209-149">object</span></span>    | <span data-ttu-id="28209-150">否</span><span class="sxs-lookup"><span data-stu-id="28209-150">No</span></span>       | <span data-ttu-id="28209-151">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="28209-151">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="28209-152">有关详细信息，请参阅[选项](#options)。</span><span class="sxs-lookup"><span data-stu-id="28209-152">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="28209-153">array</span><span class="sxs-lookup"><span data-stu-id="28209-153">array</span></span>     | <span data-ttu-id="28209-154">是</span><span class="sxs-lookup"><span data-stu-id="28209-154">Yes</span></span>      | <span data-ttu-id="28209-155">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="28209-155">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="28209-156">有关详细信息，请参阅[参数](#parameters)。</span><span class="sxs-lookup"><span data-stu-id="28209-156">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="28209-157">object</span><span class="sxs-lookup"><span data-stu-id="28209-157">object</span></span>    | <span data-ttu-id="28209-158">是</span><span class="sxs-lookup"><span data-stu-id="28209-158">Yes</span></span>      | <span data-ttu-id="28209-159">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="28209-159">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="28209-160">有关详细信息，请参阅[结果](#result)。</span><span class="sxs-lookup"><span data-stu-id="28209-160">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="28209-161">options</span><span class="sxs-lookup"><span data-stu-id="28209-161">options</span></span>

<span data-ttu-id="28209-162">`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="28209-162">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="28209-163">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="28209-163">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="28209-164">属性</span><span class="sxs-lookup"><span data-stu-id="28209-164">Property</span></span>          | <span data-ttu-id="28209-165">数据类型</span><span class="sxs-lookup"><span data-stu-id="28209-165">Data type</span></span> | <span data-ttu-id="28209-166">必需</span><span class="sxs-lookup"><span data-stu-id="28209-166">Required</span></span>                               | <span data-ttu-id="28209-167">说明</span><span class="sxs-lookup"><span data-stu-id="28209-167">Description</span></span>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | <span data-ttu-id="28209-168">boolean</span><span class="sxs-lookup"><span data-stu-id="28209-168">boolean</span></span>   | <span data-ttu-id="28209-169">否</span><span class="sxs-lookup"><span data-stu-id="28209-169">No</span></span><br/><br/><span data-ttu-id="28209-170">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="28209-170">Default value is `false`.</span></span>  | <span data-ttu-id="28209-171">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="28209-171">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="28209-172">可取消函数通常仅用于返回单个结果的异步函数，并需要处理对数据请求的取消操作。</span><span class="sxs-lookup"><span data-stu-id="28209-172">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="28209-173">函数不能同时为流式处理和可取消。</span><span class="sxs-lookup"><span data-stu-id="28209-173">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="28209-174">有关详细信息，请参阅[Make a 流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)结尾附近的注释。</span><span class="sxs-lookup"><span data-stu-id="28209-174">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="28209-175">boolean</span><span class="sxs-lookup"><span data-stu-id="28209-175">boolean</span></span>   | <span data-ttu-id="28209-176">否</span><span class="sxs-lookup"><span data-stu-id="28209-176">No</span></span> <br/><br/><span data-ttu-id="28209-177">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="28209-177">Default value is `false`.</span></span> | <span data-ttu-id="28209-178">如果为 `true` ，则自定义函数可以访问调用自定义函数的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="28209-178">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="28209-179">若要获取调用自定义函数的单元格的地址，请在自定义函数中使用 context。</span><span class="sxs-lookup"><span data-stu-id="28209-179">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="28209-180">不能将自定义函数同时设置为流式处理和 requiresAddress。</span><span class="sxs-lookup"><span data-stu-id="28209-180">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="28209-181">使用此选项时，"调用" 参数必须是在 options 中传递的最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="28209-181">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span>                                              |
| `stream`          | <span data-ttu-id="28209-182">boolean</span><span class="sxs-lookup"><span data-stu-id="28209-182">boolean</span></span>   | <span data-ttu-id="28209-183">否</span><span class="sxs-lookup"><span data-stu-id="28209-183">No</span></span><br/><br/><span data-ttu-id="28209-184">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="28209-184">Default value is `false`.</span></span>  | <span data-ttu-id="28209-185">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="28209-185">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="28209-186">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="28209-186">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="28209-187">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="28209-187">The function should have no `return` statement.</span></span> <span data-ttu-id="28209-188">相反，结果值将作为 `StreamingInvocation.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="28209-188">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="28209-189">有关详细信息，请参阅[流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)。</span><span class="sxs-lookup"><span data-stu-id="28209-189">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>                                                                                                                                                                |
| `volatile`        | <span data-ttu-id="28209-190">boolean</span><span class="sxs-lookup"><span data-stu-id="28209-190">boolean</span></span>   | <span data-ttu-id="28209-191">否</span><span class="sxs-lookup"><span data-stu-id="28209-191">No</span></span> <br/><br/><span data-ttu-id="28209-192">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="28209-192">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="28209-193">如果 `true` 为，则函数会在 Excel 重新计算时重新计算，而不是仅在公式的依赖值发生更改时进行重新计算。</span><span class="sxs-lookup"><span data-stu-id="28209-193">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="28209-194">函数不能同时为流式处理和可变。</span><span class="sxs-lookup"><span data-stu-id="28209-194">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="28209-195">如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。</span><span class="sxs-lookup"><span data-stu-id="28209-195">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span>                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a><span data-ttu-id="28209-196">参数</span><span class="sxs-lookup"><span data-stu-id="28209-196">parameters</span></span>

<span data-ttu-id="28209-197">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="28209-197">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="28209-198">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="28209-198">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="28209-199">属性</span><span class="sxs-lookup"><span data-stu-id="28209-199">Property</span></span>  |  <span data-ttu-id="28209-200">数据类型</span><span class="sxs-lookup"><span data-stu-id="28209-200">Data type</span></span>  |  <span data-ttu-id="28209-201">必需</span><span class="sxs-lookup"><span data-stu-id="28209-201">Required</span></span>  |  <span data-ttu-id="28209-202">说明</span><span class="sxs-lookup"><span data-stu-id="28209-202">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="28209-203">string</span><span class="sxs-lookup"><span data-stu-id="28209-203">string</span></span>  |  <span data-ttu-id="28209-204">否</span><span class="sxs-lookup"><span data-stu-id="28209-204">No</span></span> |  <span data-ttu-id="28209-205">参数的说明。</span><span class="sxs-lookup"><span data-stu-id="28209-205">A description of the parameter.</span></span> <span data-ttu-id="28209-206">这将显示在 Excel 的 IntelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="28209-206">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="28209-207">string</span><span class="sxs-lookup"><span data-stu-id="28209-207">string</span></span>  |  <span data-ttu-id="28209-208">否</span><span class="sxs-lookup"><span data-stu-id="28209-208">No</span></span>  |  <span data-ttu-id="28209-209">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="28209-209">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="28209-210">string</span><span class="sxs-lookup"><span data-stu-id="28209-210">string</span></span>  |  <span data-ttu-id="28209-211">是</span><span class="sxs-lookup"><span data-stu-id="28209-211">Yes</span></span>  |  <span data-ttu-id="28209-212">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="28209-212">The name of the parameter.</span></span> <span data-ttu-id="28209-213">此名称显示在 Excel 的 IntelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="28209-213">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="28209-214">string</span><span class="sxs-lookup"><span data-stu-id="28209-214">string</span></span>  |  <span data-ttu-id="28209-215">否</span><span class="sxs-lookup"><span data-stu-id="28209-215">No</span></span>  |  <span data-ttu-id="28209-216">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="28209-216">The data type of the parameter.</span></span> <span data-ttu-id="28209-217">可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="28209-217">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="28209-218">如果未指定此属性，则数据类型默认为 **any**。</span><span class="sxs-lookup"><span data-stu-id="28209-218">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="28209-219">boolean</span><span class="sxs-lookup"><span data-stu-id="28209-219">boolean</span></span> | <span data-ttu-id="28209-220">否</span><span class="sxs-lookup"><span data-stu-id="28209-220">No</span></span> | <span data-ttu-id="28209-221">如果为 `true`，则参数是可选的。</span><span class="sxs-lookup"><span data-stu-id="28209-221">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="28209-222">boolean</span><span class="sxs-lookup"><span data-stu-id="28209-222">boolean</span></span> | <span data-ttu-id="28209-223">否</span><span class="sxs-lookup"><span data-stu-id="28209-223">No</span></span> | <span data-ttu-id="28209-224">如果 `true` 为，则参数将从指定的数组中填充。</span><span class="sxs-lookup"><span data-stu-id="28209-224">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="28209-225">请注意，根据定义，所有重复参数均被视为可选参数。</span><span class="sxs-lookup"><span data-stu-id="28209-225">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="28209-226">结果</span><span class="sxs-lookup"><span data-stu-id="28209-226">result</span></span>

<span data-ttu-id="28209-227">`result` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="28209-227">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="28209-228">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="28209-228">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="28209-229">属性</span><span class="sxs-lookup"><span data-stu-id="28209-229">Property</span></span>         | <span data-ttu-id="28209-230">数据类型</span><span class="sxs-lookup"><span data-stu-id="28209-230">Data type</span></span> | <span data-ttu-id="28209-231">必需</span><span class="sxs-lookup"><span data-stu-id="28209-231">Required</span></span> | <span data-ttu-id="28209-232">说明</span><span class="sxs-lookup"><span data-stu-id="28209-232">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="28209-233">string</span><span class="sxs-lookup"><span data-stu-id="28209-233">string</span></span>    | <span data-ttu-id="28209-234">否</span><span class="sxs-lookup"><span data-stu-id="28209-234">No</span></span>       | <span data-ttu-id="28209-235">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="28209-235">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="28209-236">将函数名称与 JSON 元数据相关联</span><span class="sxs-lookup"><span data-stu-id="28209-236">Associating function names with JSON metadata</span></span>

<span data-ttu-id="28209-237">若要使函数正常工作，需要将函数的 `id` 属性与 JavaScript 实现相关联。</span><span class="sxs-lookup"><span data-stu-id="28209-237">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="28209-238">请确保存在关联，否则将无法注册该函数，也无法在 Excel 中使用它。</span><span class="sxs-lookup"><span data-stu-id="28209-238">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="28209-239">下面的代码示例演示如何使用方法进行关联 `CustomFunctions.associate()` 。</span><span class="sxs-lookup"><span data-stu-id="28209-239">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="28209-240">该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。</span><span class="sxs-lookup"><span data-stu-id="28209-240">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="28209-241">下面的 JSON 显示了与上一个自定义函数 JavaScript 代码相关联的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="28209-241">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

<span data-ttu-id="28209-242">在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。</span><span class="sxs-lookup"><span data-stu-id="28209-242">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="28209-243">在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="28209-243">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="28209-244">在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。</span><span class="sxs-lookup"><span data-stu-id="28209-244">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="28209-245">也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。</span><span class="sxs-lookup"><span data-stu-id="28209-245">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="28209-246">在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。</span><span class="sxs-lookup"><span data-stu-id="28209-246">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="28209-247">你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="28209-247">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="28209-248">在 JavaScript 文件中，使用每个函数的后面指定自定义函数关联 `CustomFunctions.associate` 。</span><span class="sxs-lookup"><span data-stu-id="28209-248">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="28209-249">以下示例显示了与此 JavaScript 代码示例中定义的函数相对应的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="28209-249">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="28209-250">`id`和 `name` 属性值以大写形式表示，这是描述自定义函数的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="28209-250">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="28209-251">仅当您手动准备自己的 JSON 文件，而不是使用自动生成时，才需要添加此 JSON。</span><span class="sxs-lookup"><span data-stu-id="28209-251">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="28209-252">有关自动生成的详细信息，请参阅[CREATE JSON metadata for custom 函数](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="28209-252">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      ...
    },
    {
      "id": "INCREMENT",
      "name": "INCREMENT",
      ...
    }
  ]
}
```

## <a name="next-steps"></a><span data-ttu-id="28209-253">后续步骤</span><span class="sxs-lookup"><span data-stu-id="28209-253">Next steps</span></span>

<span data-ttu-id="28209-254">了解[有关命名函数](custom-functions-naming.md)或了解如何使用前面所述的手写 JSON 方法对[函数进行本地化](custom-functions-localize.md)的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="28209-254">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="28209-255">另请参阅</span><span class="sxs-lookup"><span data-stu-id="28209-255">See also</span></span>

- [<span data-ttu-id="28209-256">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="28209-256">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="28209-257">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="28209-257">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="28209-258">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="28209-258">Create custom functions in Excel</span></span>](custom-functions-overview.md)
