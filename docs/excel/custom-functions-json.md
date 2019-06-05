---
ms.date: 05/30/2019
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中自定义函数的元数据
localization_priority: Normal
ms.openlocfilehash: e51e4e8ee89eb1f345ee0c564e9b2ff8119806b2
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706121"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="0dec4-103">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="0dec4-103">Custom functions metadata</span></span>

<span data-ttu-id="0dec4-104">在 Excel 加载项中定义[自定义函数](custom-functions-overview.md)时, 加载项项目包含 JSON 元数据文件, 该文件提供了 Excel 注册自定义函数并使其可供最终用户使用的信息。</span><span class="sxs-lookup"><span data-stu-id="0dec4-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0dec4-105">此文件的生成方式为:</span><span class="sxs-lookup"><span data-stu-id="0dec4-105">This file is generated either:</span></span>

- <span data-ttu-id="0dec4-106">您, 在手写 JSON 文件中</span><span class="sxs-lookup"><span data-stu-id="0dec4-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="0dec4-107">从您在函数开头输入的 JSDoc 注释</span><span class="sxs-lookup"><span data-stu-id="0dec4-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="0dec4-108">自定义函数在用户首次运行外接程序且在所有工作簿中对同一用户可用时注册。</span><span class="sxs-lookup"><span data-stu-id="0dec4-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="0dec4-109">本文介绍了 JSON 元数据文件的格式, 假定您正在手动编写元数据文件。</span><span class="sxs-lookup"><span data-stu-id="0dec4-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="0dec4-110">有关 JSDoc 注释 JSON 文件生成的信息, 请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="0dec4-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="0dec4-111">有关为启用自定义函数必须在加载项项目中包含的其他文件的信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="0dec4-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="0dec4-112">托管 JSON 文件的服务器上的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)，以便自定义函数在 Excel Online 中正常工作。</span><span class="sxs-lookup"><span data-stu-id="0dec4-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="0dec4-113">示例元数据</span><span class="sxs-lookup"><span data-stu-id="0dec4-113">Example metadata</span></span>

<span data-ttu-id="0dec4-114">以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="0dec4-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="0dec4-115">此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="0dec4-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
      "description":  "Count up from zero",
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
      "description":  "Get the second highest number from a range",
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
> <span data-ttu-id="0dec4-116">在 [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub 存储库中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="0dec4-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="0dec4-117">functions</span><span class="sxs-lookup"><span data-stu-id="0dec4-117">functions</span></span> 

<span data-ttu-id="0dec4-118">`functions` 属性是自定义函数对象的一个数组。</span><span class="sxs-lookup"><span data-stu-id="0dec4-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="0dec4-119">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0dec4-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0dec4-120">属性</span><span class="sxs-lookup"><span data-stu-id="0dec4-120">Property</span></span>  |  <span data-ttu-id="0dec4-121">数据类型</span><span class="sxs-lookup"><span data-stu-id="0dec4-121">Data type</span></span>  |  <span data-ttu-id="0dec4-122">必需</span><span class="sxs-lookup"><span data-stu-id="0dec4-122">Required</span></span>  |  <span data-ttu-id="0dec4-123">说明</span><span class="sxs-lookup"><span data-stu-id="0dec4-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0dec4-124">string</span><span class="sxs-lookup"><span data-stu-id="0dec4-124">string</span></span>  |  <span data-ttu-id="0dec4-125">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-125">No</span></span>  |  <span data-ttu-id="0dec4-126">最终用户在 Excel 中看到的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="0dec4-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="0dec4-127">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="0dec4-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="0dec4-128">字符串</span><span class="sxs-lookup"><span data-stu-id="0dec4-128">string</span></span>  |   <span data-ttu-id="0dec4-129">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-129">No</span></span>  |  <span data-ttu-id="0dec4-130">提供有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="0dec4-130">URL that provides information about the function.</span></span> <span data-ttu-id="0dec4-131">（它显示在任务窗格中。）例如，`http://contoso.com/help/convertcelsiustofahrenheit.html`。</span><span class="sxs-lookup"><span data-stu-id="0dec4-131">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="0dec4-132">string</span><span class="sxs-lookup"><span data-stu-id="0dec4-132">string</span></span> | <span data-ttu-id="0dec4-133">是</span><span class="sxs-lookup"><span data-stu-id="0dec4-133">Yes</span></span> | <span data-ttu-id="0dec4-134">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="0dec4-134">A unique ID for the function.</span></span> <span data-ttu-id="0dec4-135">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="0dec4-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="0dec4-136">string</span><span class="sxs-lookup"><span data-stu-id="0dec4-136">string</span></span>  |  <span data-ttu-id="0dec4-137">是</span><span class="sxs-lookup"><span data-stu-id="0dec4-137">Yes</span></span>  |  <span data-ttu-id="0dec4-138">最终用户在 Excel 中看到的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="0dec4-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="0dec4-139">在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="0dec4-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="0dec4-140">object</span><span class="sxs-lookup"><span data-stu-id="0dec4-140">object</span></span>  |  <span data-ttu-id="0dec4-141">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-141">No</span></span>  |  <span data-ttu-id="0dec4-142">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="0dec4-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0dec4-143">有关详细信息，请参阅[选项](#options)。</span><span class="sxs-lookup"><span data-stu-id="0dec4-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="0dec4-144">array</span><span class="sxs-lookup"><span data-stu-id="0dec4-144">array</span></span>  |  <span data-ttu-id="0dec4-145">是</span><span class="sxs-lookup"><span data-stu-id="0dec4-145">Yes</span></span>  |  <span data-ttu-id="0dec4-146">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="0dec4-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="0dec4-147">有关详细信息，请参阅[参数](#parameters)。</span><span class="sxs-lookup"><span data-stu-id="0dec4-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="0dec4-148">object</span><span class="sxs-lookup"><span data-stu-id="0dec4-148">object</span></span>  |  <span data-ttu-id="0dec4-149">是</span><span class="sxs-lookup"><span data-stu-id="0dec4-149">Yes</span></span>  |  <span data-ttu-id="0dec4-150">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="0dec4-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0dec4-151">有关详细信息，请参阅[结果](#result)。</span><span class="sxs-lookup"><span data-stu-id="0dec4-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="0dec4-152">options</span><span class="sxs-lookup"><span data-stu-id="0dec4-152">options</span></span>

<span data-ttu-id="0dec4-153">`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="0dec4-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0dec4-154">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0dec4-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="0dec4-155">属性</span><span class="sxs-lookup"><span data-stu-id="0dec4-155">Property</span></span>  |  <span data-ttu-id="0dec4-156">数据类型</span><span class="sxs-lookup"><span data-stu-id="0dec4-156">Data type</span></span>  |  <span data-ttu-id="0dec4-157">必需</span><span class="sxs-lookup"><span data-stu-id="0dec4-157">Required</span></span>  |  <span data-ttu-id="0dec4-158">说明</span><span class="sxs-lookup"><span data-stu-id="0dec4-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="0dec4-159">boolean</span><span class="sxs-lookup"><span data-stu-id="0dec4-159">boolean</span></span>  |  <span data-ttu-id="0dec4-160">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-160">No</span></span><br/><br/><span data-ttu-id="0dec4-161">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="0dec4-161">Default value is `false`.</span></span>  |  <span data-ttu-id="0dec4-162">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `CancelableInvocation` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="0dec4-162">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="0dec4-163">可取消函数通常仅用于返回单个结果的异步函数, 并需要处理对数据请求的取消操作。</span><span class="sxs-lookup"><span data-stu-id="0dec4-163">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="0dec4-164">函数不能同时为流式处理和可取消。</span><span class="sxs-lookup"><span data-stu-id="0dec4-164">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="0dec4-165">有关详细信息, 请参阅[Make a 流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)结尾附近的注释。</span><span class="sxs-lookup"><span data-stu-id="0dec4-165">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="0dec4-166">boolean</span><span class="sxs-lookup"><span data-stu-id="0dec4-166">boolean</span></span> | <span data-ttu-id="0dec4-167">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-167">No</span></span> <br/><br/><span data-ttu-id="0dec4-168">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="0dec4-168">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="0dec4-169">如果为 true, 则自定义函数可以访问调用自定义函数的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="0dec4-169">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="0dec4-170">若要获取调用自定义函数的单元格的地址, 请在自定义函数中使用 context。</span><span class="sxs-lookup"><span data-stu-id="0dec4-170">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="0dec4-171">有关详细信息，请参阅[确定调用自定义函数的单元格](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function)。</span><span class="sxs-lookup"><span data-stu-id="0dec4-171">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="0dec4-172">不能将自定义函数同时设置为流式处理和 requiresAddress。</span><span class="sxs-lookup"><span data-stu-id="0dec4-172">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="0dec4-173">使用此选项时, "调用" 参数必须是在 options 中传递的最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="0dec4-173">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="0dec4-174">boolean</span><span class="sxs-lookup"><span data-stu-id="0dec4-174">boolean</span></span>  |  <span data-ttu-id="0dec4-175">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-175">No</span></span><br/><br/><span data-ttu-id="0dec4-176">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="0dec4-176">Default value is `false`.</span></span>  |  <span data-ttu-id="0dec4-177">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="0dec4-177">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="0dec4-178">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="0dec4-178">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="0dec4-179">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="0dec4-179">The function should have no `return` statement.</span></span> <span data-ttu-id="0dec4-180">相反，结果值将作为 `StreamingInvocation.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="0dec4-180">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="0dec4-181">有关详细信息，请参阅[流式处理函数](custom-functions-web-reqs.md#make-a-streaming-function)。</span><span class="sxs-lookup"><span data-stu-id="0dec4-181">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="0dec4-182">boolean</span><span class="sxs-lookup"><span data-stu-id="0dec4-182">boolean</span></span> | <span data-ttu-id="0dec4-183">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-183">No</span></span> <br/><br/><span data-ttu-id="0dec4-184">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="0dec4-184">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="0dec4-185">如果为 `true`，则该函数会在每次 Excel 重新计算时（而不是仅当公式的从属值发生更改时）进行重新计算。</span><span class="sxs-lookup"><span data-stu-id="0dec4-185">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="0dec4-186">函数不能同时为流式处理和可变。</span><span class="sxs-lookup"><span data-stu-id="0dec4-186">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="0dec4-187">如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。</span><span class="sxs-lookup"><span data-stu-id="0dec4-187">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="0dec4-188">参数</span><span class="sxs-lookup"><span data-stu-id="0dec4-188">parameters</span></span>

<span data-ttu-id="0dec4-189">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="0dec4-189">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="0dec4-190">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0dec4-190">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0dec4-191">属性</span><span class="sxs-lookup"><span data-stu-id="0dec4-191">Property</span></span>  |  <span data-ttu-id="0dec4-192">数据类型</span><span class="sxs-lookup"><span data-stu-id="0dec4-192">Data type</span></span>  |  <span data-ttu-id="0dec4-193">必需</span><span class="sxs-lookup"><span data-stu-id="0dec4-193">Required</span></span>  |  <span data-ttu-id="0dec4-194">说明</span><span class="sxs-lookup"><span data-stu-id="0dec4-194">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0dec4-195">string</span><span class="sxs-lookup"><span data-stu-id="0dec4-195">string</span></span>  |  <span data-ttu-id="0dec4-196">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-196">No</span></span> |  <span data-ttu-id="0dec4-197">参数的说明。</span><span class="sxs-lookup"><span data-stu-id="0dec4-197">A description of the parameter.</span></span> <span data-ttu-id="0dec4-198">这显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="0dec4-198">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="0dec4-199">字符串</span><span class="sxs-lookup"><span data-stu-id="0dec4-199">string</span></span>  |  <span data-ttu-id="0dec4-200">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-200">No</span></span>  |  <span data-ttu-id="0dec4-201">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="0dec4-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="0dec4-202">string</span><span class="sxs-lookup"><span data-stu-id="0dec4-202">string</span></span>  |  <span data-ttu-id="0dec4-203">是</span><span class="sxs-lookup"><span data-stu-id="0dec4-203">Yes</span></span>  |  <span data-ttu-id="0dec4-204">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="0dec4-204">The name of the parameter.</span></span> <span data-ttu-id="0dec4-205">此名称显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="0dec4-205">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="0dec4-206">string</span><span class="sxs-lookup"><span data-stu-id="0dec4-206">string</span></span>  |  <span data-ttu-id="0dec4-207">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-207">No</span></span>  |  <span data-ttu-id="0dec4-208">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="0dec4-208">The data type of the parameter.</span></span> <span data-ttu-id="0dec4-209">可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="0dec4-209">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="0dec4-210">如果未指定此属性，则数据类型默认为 **any**。</span><span class="sxs-lookup"><span data-stu-id="0dec4-210">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="0dec4-211">boolean</span><span class="sxs-lookup"><span data-stu-id="0dec4-211">boolean</span></span> | <span data-ttu-id="0dec4-212">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-212">No</span></span> | <span data-ttu-id="0dec4-213">如果为 `true`，则参数是可选的。</span><span class="sxs-lookup"><span data-stu-id="0dec4-213">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="0dec4-214">结果</span><span class="sxs-lookup"><span data-stu-id="0dec4-214">result</span></span>

<span data-ttu-id="0dec4-215">`result` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="0dec4-215">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0dec4-216">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0dec4-216">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="0dec4-217">属性</span><span class="sxs-lookup"><span data-stu-id="0dec4-217">Property</span></span>  |  <span data-ttu-id="0dec4-218">数据类型</span><span class="sxs-lookup"><span data-stu-id="0dec4-218">Data type</span></span>  |  <span data-ttu-id="0dec4-219">必需</span><span class="sxs-lookup"><span data-stu-id="0dec4-219">Required</span></span>  |  <span data-ttu-id="0dec4-220">说明</span><span class="sxs-lookup"><span data-stu-id="0dec4-220">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="0dec4-221">string</span><span class="sxs-lookup"><span data-stu-id="0dec4-221">string</span></span>  |  <span data-ttu-id="0dec4-222">否</span><span class="sxs-lookup"><span data-stu-id="0dec4-222">No</span></span>  |  <span data-ttu-id="0dec4-223">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="0dec4-223">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="0dec4-224">后续步骤</span><span class="sxs-lookup"><span data-stu-id="0dec4-224">Next steps</span></span>
<span data-ttu-id="0dec4-225">了解[有关命名函数](custom-functions-naming.md)或了解如何使用前面所述的手写 JSON 方法对[函数进行本地化](custom-functions-localize.md)的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="0dec4-225">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="0dec4-226">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0dec4-226">See also</span></span>

* [<span data-ttu-id="0dec4-227">为自定义函数自动生成 JSON 元数据</span><span class="sxs-lookup"><span data-stu-id="0dec4-227">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="0dec4-228">自定义函数参数选项</span><span class="sxs-lookup"><span data-stu-id="0dec4-228">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="0dec4-229">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="0dec4-229">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0dec4-230">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="0dec4-230">Create custom functions in Excel</span></span>](custom-functions-overview.md)