---
ms.date: 12/22/2020
description: 定义自定义函数的 JSON 元数据Excel并关联函数 ID 和名称属性。
title: 手动为自定义函数创建 JSON Excel
localization_priority: Normal
ms.openlocfilehash: 514eacba5045d160eb6f3d4823adbd8c2f45292a
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075899"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a><span data-ttu-id="bbd4c-103">手动为自定义函数创建 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="bbd4c-103">Manually create JSON metadata for custom functions</span></span>

<span data-ttu-id="bbd4c-104">如自定义函数概述[](custom-functions-overview.md)文章中所述，自定义函数项目必须同时包括 JSON 元数据文件和脚本 (JavaScript 或 TypeScript) 文件以注册函数，使其可供使用。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="bbd4c-105">当用户首次运行外接程序时以及之后，自定义函数将注册到所有工作簿中的同一用户。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="bbd4c-106">我们建议尽可能使用 JSON 自动生成，而不是创建自己的 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="bbd4c-107">自动生成不易出现用户错误，并且基架文件 `yo office` 已包含此错误。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="bbd4c-108">有关 JSDoc 标记和 JSON 自动生成过程的信息，请参阅自动生成 [自定义函数的 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="bbd4c-109">但是，你可以从头开始创建自定义函数项目。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-109">However, you can make a custom functions project from scratch.</span></span> <span data-ttu-id="bbd4c-110">此过程需要您：</span><span class="sxs-lookup"><span data-stu-id="bbd4c-110">This process requires you to:</span></span>

- <span data-ttu-id="bbd4c-111">编写 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-111">Write your JSON file.</span></span>
- <span data-ttu-id="bbd4c-112">检查清单文件是否连接到 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-112">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="bbd4c-113">在脚本文件中 `id` 关联 `name` 函数和属性，以便注册函数。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-113">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="bbd4c-114">下图说明了使用基架文件和从头开始编写 `yo office` JSON 的区别。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-114">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![使用 Yo 方法与编写Office JSON 之间的差异的图像。](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="bbd4c-116">请记住，如果不使用生成器，请通过 XML 清单文件的 部分将清单连接到你创建的 JSON `<Resources>` `yo office` 文件。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-116">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="bbd4c-117">创作元数据并连接到清单</span><span class="sxs-lookup"><span data-stu-id="bbd4c-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="bbd4c-118">在项目中创建 JSON 文件，并提供其中函数的所有详细信息，例如函数的参数。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-118">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="bbd4c-119">有关[函数属性](#json-metadata-example)[的完整列表](#metadata-reference)，请参阅以下元数据示例和元数据引用。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="bbd4c-120">确保 XML 清单文件引用 部分中的 JSON 文件， `<Resources>` 类似于以下示例。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-120">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="bbd4c-121">JSON 元数据示例</span><span class="sxs-lookup"><span data-stu-id="bbd4c-121">JSON metadata example</span></span>

<span data-ttu-id="bbd4c-122">以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="bbd4c-123">此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="bbd4c-124">[OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json)中提供了完整的 JSON GitHub存储库的提交历史记录。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="bbd4c-125">由于项目已调整为自动生成 JSON，因此手写 JSON 的完整示例仅在项目的早期版本中可用。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="bbd4c-126">元数据参考</span><span class="sxs-lookup"><span data-stu-id="bbd4c-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="bbd4c-127">functions</span><span class="sxs-lookup"><span data-stu-id="bbd4c-127">functions</span></span>

<span data-ttu-id="bbd4c-128">`functions` 属性是自定义函数对象的一个数组。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="bbd4c-129">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="bbd4c-130">属性</span><span class="sxs-lookup"><span data-stu-id="bbd4c-130">Property</span></span>      | <span data-ttu-id="bbd4c-131">数据类型</span><span class="sxs-lookup"><span data-stu-id="bbd4c-131">Data type</span></span> | <span data-ttu-id="bbd4c-132">必需</span><span class="sxs-lookup"><span data-stu-id="bbd4c-132">Required</span></span> | <span data-ttu-id="bbd4c-133">说明</span><span class="sxs-lookup"><span data-stu-id="bbd4c-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="bbd4c-134">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-134">string</span></span>    | <span data-ttu-id="bbd4c-135">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-135">No</span></span>       | <span data-ttu-id="bbd4c-136">最终用户在 Excel 中看到的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="bbd4c-137">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="bbd4c-138">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-138">string</span></span>    | <span data-ttu-id="bbd4c-139">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-139">No</span></span>       | <span data-ttu-id="bbd4c-140">提供有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-140">URL that provides information about the function.</span></span> <span data-ttu-id="bbd4c-141">（它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="bbd4c-142">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-142">string</span></span>    | <span data-ttu-id="bbd4c-143">是</span><span class="sxs-lookup"><span data-stu-id="bbd4c-143">Yes</span></span>      | <span data-ttu-id="bbd4c-144">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-144">A unique ID for the function.</span></span> <span data-ttu-id="bbd4c-145">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="bbd4c-146">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-146">string</span></span>    | <span data-ttu-id="bbd4c-147">是</span><span class="sxs-lookup"><span data-stu-id="bbd4c-147">Yes</span></span>      | <span data-ttu-id="bbd4c-148">最终用户在 Excel 中看到的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="bbd4c-149">在Excel中，此函数名称以 XML 清单文件中指定的自定义函数命名空间作为前缀。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-149">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="bbd4c-150">object</span><span class="sxs-lookup"><span data-stu-id="bbd4c-150">object</span></span>    | <span data-ttu-id="bbd4c-151">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-151">No</span></span>       | <span data-ttu-id="bbd4c-152">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="bbd4c-153">有关详细信息，请参阅[选项](#options)。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="bbd4c-154">array</span><span class="sxs-lookup"><span data-stu-id="bbd4c-154">array</span></span>     | <span data-ttu-id="bbd4c-155">是</span><span class="sxs-lookup"><span data-stu-id="bbd4c-155">Yes</span></span>      | <span data-ttu-id="bbd4c-156">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="bbd4c-157">有关详细信息 [，](#parameters) 请参阅参数。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="bbd4c-158">object</span><span class="sxs-lookup"><span data-stu-id="bbd4c-158">object</span></span>    | <span data-ttu-id="bbd4c-159">是</span><span class="sxs-lookup"><span data-stu-id="bbd4c-159">Yes</span></span>      | <span data-ttu-id="bbd4c-160">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="bbd4c-161">有关详细信息，请参阅[结果](#result)。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="bbd4c-162">options</span><span class="sxs-lookup"><span data-stu-id="bbd4c-162">options</span></span>

<span data-ttu-id="bbd4c-163">`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="bbd4c-164">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="bbd4c-165">属性</span><span class="sxs-lookup"><span data-stu-id="bbd4c-165">Property</span></span>          | <span data-ttu-id="bbd4c-166">数据类型</span><span class="sxs-lookup"><span data-stu-id="bbd4c-166">Data type</span></span> | <span data-ttu-id="bbd4c-167">必需</span><span class="sxs-lookup"><span data-stu-id="bbd4c-167">Required</span></span>                               | <span data-ttu-id="bbd4c-168">说明</span><span class="sxs-lookup"><span data-stu-id="bbd4c-168">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="bbd4c-169">boolean</span><span class="sxs-lookup"><span data-stu-id="bbd4c-169">boolean</span></span>   | <span data-ttu-id="bbd4c-170">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-170">No</span></span><br/><br/><span data-ttu-id="bbd4c-171">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-171">Default value is `false`.</span></span>  | <span data-ttu-id="bbd4c-172">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="bbd4c-173">可取消函数通常仅用于返回单个结果并需要处理数据请求取消的异步函数。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="bbd4c-174">函数不能同时使用 和 `stream` `cancelable` 属性。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-174">A function can't use both the `stream` and `cancelable` properties.</span></span> |
| `requiresAddress` | <span data-ttu-id="bbd4c-175">boolean</span><span class="sxs-lookup"><span data-stu-id="bbd4c-175">boolean</span></span>   | <span data-ttu-id="bbd4c-176">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-176">No</span></span> <br/><br/><span data-ttu-id="bbd4c-177">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-177">Default value is `false`.</span></span> | <span data-ttu-id="bbd4c-178">如果 `true` 为 ，则自定义函数可以访问调用它的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-178">If `true`, your custom function can access the address of the cell that invoked it.</span></span> <span data-ttu-id="bbd4c-179">`address`调用参数的[属性](custom-functions-parameter-options.md#invocation-parameter)包含调用自定义函数的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-179">The `address` property of the [invocation parameter](custom-functions-parameter-options.md#invocation-parameter) contains the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="bbd4c-180">函数不能同时使用 和 `stream` `requiresAddress` 属性。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-180">A function can't use both the `stream` and `requiresAddress` properties.</span></span> |
| `requiresParameterAddresses` | <span data-ttu-id="bbd4c-181">boolean</span><span class="sxs-lookup"><span data-stu-id="bbd4c-181">boolean</span></span>   | <span data-ttu-id="bbd4c-182">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-182">No</span></span> <br/><br/><span data-ttu-id="bbd4c-183">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-183">Default value is `false`.</span></span> | <span data-ttu-id="bbd4c-184">如果 `true` 为 ，则自定义函数可以访问函数的输入参数的地址。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-184">If `true`, your custom function can access the addresses of the function's input parameters.</span></span> <span data-ttu-id="bbd4c-185">此属性必须与结果对象的 属性结合使用， `dimensionality` 并且[](#result) `dimensionality` 必须设置为 `matrix` 。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-185">This property must be used in combination with the `dimensionality` property of the [result](#result) object, and `dimensionality` must be set to `matrix`.</span></span> <span data-ttu-id="bbd4c-186">有关详细信息 [，请参阅检测参数](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) 的地址。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-186">See [Detect the address of a parameter](custom-functions-parameter-options.md#detect-the-address-of-a-parameter) for more information.</span></span> |
| `stream`          | <span data-ttu-id="bbd4c-187">boolean</span><span class="sxs-lookup"><span data-stu-id="bbd4c-187">boolean</span></span>   | <span data-ttu-id="bbd4c-188">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-188">No</span></span><br/><br/><span data-ttu-id="bbd4c-189">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-189">Default value is `false`.</span></span>  | <span data-ttu-id="bbd4c-190">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-190">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="bbd4c-191">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-191">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="bbd4c-192">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-192">The function should have no `return` statement.</span></span> <span data-ttu-id="bbd4c-193">相反，结果值将作为 `StreamingInvocation.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-193">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="bbd4c-194">有关详细信息，请参阅 [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function)。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-194">For more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="bbd4c-195">boolean</span><span class="sxs-lookup"><span data-stu-id="bbd4c-195">boolean</span></span>   | <span data-ttu-id="bbd4c-196">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-196">No</span></span> <br/><br/><span data-ttu-id="bbd4c-197">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-197">Default value is `false`.</span></span> | <span data-ttu-id="bbd4c-198">如果为 ，则函数每次Excel重新计算，而不是仅在公式的从属值发生更改 `true` 时重新计算。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-198">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="bbd4c-199">函数不能同时使用 和 `stream` `volatile` 属性。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-199">A function can't use both the `stream` and `volatile` properties.</span></span> <span data-ttu-id="bbd4c-200">如果 `stream` 和 `volatile` 属性都设置为 `true` ，则可变属性将被忽略。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-200">If the `stream` and `volatile` properties are both set to `true`, the volatile property will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="bbd4c-201">参数</span><span class="sxs-lookup"><span data-stu-id="bbd4c-201">parameters</span></span>

<span data-ttu-id="bbd4c-202">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-202">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="bbd4c-203">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-203">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="bbd4c-204">属性</span><span class="sxs-lookup"><span data-stu-id="bbd4c-204">Property</span></span>  |  <span data-ttu-id="bbd4c-205">数据类型</span><span class="sxs-lookup"><span data-stu-id="bbd4c-205">Data type</span></span>  |  <span data-ttu-id="bbd4c-206">必需</span><span class="sxs-lookup"><span data-stu-id="bbd4c-206">Required</span></span>  |  <span data-ttu-id="bbd4c-207">说明</span><span class="sxs-lookup"><span data-stu-id="bbd4c-207">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="bbd4c-208">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-208">string</span></span>  |  <span data-ttu-id="bbd4c-209">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-209">No</span></span> |  <span data-ttu-id="bbd4c-210">参数的说明。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-210">A description of the parameter.</span></span> <span data-ttu-id="bbd4c-211">这将显示在Excel中IntelliSense。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-211">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="bbd4c-212">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-212">string</span></span>  |  <span data-ttu-id="bbd4c-213">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-213">No</span></span>  |  <span data-ttu-id="bbd4c-214">必须是 (一个非数组值) 或 (`scalar` `matrix` 二维数组) 。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-214">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="bbd4c-215">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-215">string</span></span>  |  <span data-ttu-id="bbd4c-216">是</span><span class="sxs-lookup"><span data-stu-id="bbd4c-216">Yes</span></span>  |  <span data-ttu-id="bbd4c-217">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-217">The name of the parameter.</span></span> <span data-ttu-id="bbd4c-218">此名称显示在Excel中IntelliSense。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-218">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="bbd4c-219">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-219">string</span></span>  |  <span data-ttu-id="bbd4c-220">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-220">No</span></span>  |  <span data-ttu-id="bbd4c-221">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-221">The data type of the parameter.</span></span> <span data-ttu-id="bbd4c-222">可以是 `boolean` 、 、 或 ，它允许您使用前三种类型 `number` `string` `any` 中的任意一种。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-222">Can be `boolean`, `number`, `string`, or `any`, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="bbd4c-223">如果未指定此属性，则数据类型默认值 `any` 。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-223">If this property is not specified, the data type defaults to `any`.</span></span> |
|  `optional`  | <span data-ttu-id="bbd4c-224">boolean</span><span class="sxs-lookup"><span data-stu-id="bbd4c-224">boolean</span></span> | <span data-ttu-id="bbd4c-225">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-225">No</span></span> | <span data-ttu-id="bbd4c-226">如果为 `true`，则参数是可选的。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-226">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="bbd4c-227">boolean</span><span class="sxs-lookup"><span data-stu-id="bbd4c-227">boolean</span></span> | <span data-ttu-id="bbd4c-228">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-228">No</span></span> | <span data-ttu-id="bbd4c-229">如果 `true` 为 ，参数从指定数组填充。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-229">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="bbd4c-230">请注意，根据定义，函数的所有重复参数都被视为可选参数。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-230">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="bbd4c-231">结果</span><span class="sxs-lookup"><span data-stu-id="bbd4c-231">result</span></span>

<span data-ttu-id="bbd4c-232">`result` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-232">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="bbd4c-233">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-233">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="bbd4c-234">属性</span><span class="sxs-lookup"><span data-stu-id="bbd4c-234">Property</span></span>         | <span data-ttu-id="bbd4c-235">数据类型</span><span class="sxs-lookup"><span data-stu-id="bbd4c-235">Data type</span></span> | <span data-ttu-id="bbd4c-236">必需</span><span class="sxs-lookup"><span data-stu-id="bbd4c-236">Required</span></span> | <span data-ttu-id="bbd4c-237">说明</span><span class="sxs-lookup"><span data-stu-id="bbd4c-237">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="bbd4c-238">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-238">string</span></span>    | <span data-ttu-id="bbd4c-239">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-239">No</span></span>       | <span data-ttu-id="bbd4c-240">必须是 (一个非数组值) 或 (`scalar` `matrix` 二维数组) 。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-240">Must be either `scalar` (a non-array value) or `matrix` (a 2-dimensional array).</span></span> |
| `type` | <span data-ttu-id="bbd4c-241">string</span><span class="sxs-lookup"><span data-stu-id="bbd4c-241">string</span></span>    | <span data-ttu-id="bbd4c-242">否</span><span class="sxs-lookup"><span data-stu-id="bbd4c-242">No</span></span>       | <span data-ttu-id="bbd4c-243">结果数据类型。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-243">The data type of the result.</span></span> <span data-ttu-id="bbd4c-244">可以是 `boolean` `number` 、、或 (，这允许你使用前 `string` `any` 三种类型中的任意) 。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-244">Can be `boolean`, `number`, `string`, or `any` (which allows you to use of any of the previous three types).</span></span> <span data-ttu-id="bbd4c-245">如果未指定此属性，则数据类型默认值 `any` 。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-245">If this property is not specified, the data type defaults to `any`.</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="bbd4c-246">将函数名称与 JSON 元数据相关联</span><span class="sxs-lookup"><span data-stu-id="bbd4c-246">Associating function names with JSON metadata</span></span>

<span data-ttu-id="bbd4c-247">若要使函数正常工作，需要将函数的 属性 `id` 与 JavaScript 实现关联。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-247">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="bbd4c-248">请确保存在关联，否则函数将不会注册并且不可在 Excel。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-248">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="bbd4c-249">下面的代码示例演示如何使用 方法进行 `CustomFunctions.associate()` 关联。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-249">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="bbd4c-250">该示例定义了自定义函数 `add`，并将其与 JSON 元数据文件中的对象关联，其中 `id` 属性的值为 **ADD**。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-250">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="bbd4c-251">以下 JSON 显示与以前的自定义函数 JavaScript 代码关联的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-251">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="bbd4c-252">在 JavaScript 文件中创建自定义函数和在 JSON 元数据文件中指定相应信息时，请记住以下最佳实践。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-252">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="bbd4c-253">在 JSON 元数据文件中，确保每个 `id` 属性的值仅包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-253">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="bbd4c-254">在 JSON 元数据文件中，确保每个 `id` 属性的值在该文件范围内是唯一的。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-254">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="bbd4c-255">也就是说，元数据文件中不应存在具有相同 `id` 值的两个函数对象。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-255">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="bbd4c-256">在将 JSON 元数据文件中的 `id` 属性的值与相应的 JavaScript 函数名称关联后，请勿再更改该值。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-256">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="bbd4c-257">你可以通过更新 JSON 元数据文件中的 `name` 属性来更改最终用户在 Excel 中看到的函数名称，但绝不能更改已确定的 `id` 属性的值。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-257">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="bbd4c-258">在 JavaScript 文件中，在每个函数后使用 `CustomFunctions.associate` 指定自定义函数关联。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-258">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="bbd4c-259">以下示例显示对应于前面 JavaScript 代码示例中定义的函数的 JSON 元数据。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-259">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="bbd4c-260">`id`和 `name` 属性值为大写，这是描述自定义函数时的最佳操作。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-260">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="bbd4c-261">只有在手动准备自己的 JSON 文件而不是使用自动生成时，才需要添加此 JSON。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-261">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="bbd4c-262">有关自动生成的信息，请参阅自动生成 [自定义函数的 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-262">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="bbd4c-263">后续步骤</span><span class="sxs-lookup"><span data-stu-id="bbd4c-263">Next steps</span></span>

<span data-ttu-id="bbd4c-264">了解[命名函数的](custom-functions-naming.md)最佳实践，或了解如何使用前面描述的手写[](custom-functions-localize.md)JSON 方法本地化函数。</span><span class="sxs-lookup"><span data-stu-id="bbd4c-264">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="bbd4c-265">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bbd4c-265">See also</span></span>

- [<span data-ttu-id="bbd4c-266">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="bbd4c-266">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="bbd4c-267">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="bbd4c-267">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="bbd4c-268">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="bbd4c-268">Create custom functions in Excel</span></span>](custom-functions-overview.md)
