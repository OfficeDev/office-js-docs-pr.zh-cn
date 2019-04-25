---
ms.date: 03/29/2019
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中的自定义函数的元数据（预览）
localization_priority: Normal
ms.openlocfilehash: 3703699348e99fd076fe0e3affac88038e3aaf59
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448203"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="bad80-103">自定义函数元数据（预览）</span><span class="sxs-lookup"><span data-stu-id="bad80-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="bad80-104">在 Excel 加载项中定义[自定义函数](custom-functions-overview.md)时, 加载项项目包含 JSON 元数据文件, 该文件提供了 Excel 注册自定义函数并使其可供最终用户使用的信息。</span><span class="sxs-lookup"><span data-stu-id="bad80-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="bad80-105">此文件的生成方式为:</span><span class="sxs-lookup"><span data-stu-id="bad80-105">This file is generated either:</span></span>

- <span data-ttu-id="bad80-106">您, 在手写 JSON 文件中</span><span class="sxs-lookup"><span data-stu-id="bad80-106">by you, in a handwritten JSON file</span></span>
- <span data-ttu-id="bad80-107">从您在函数开头输入的 JSDoc 注释</span><span class="sxs-lookup"><span data-stu-id="bad80-107">from the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="bad80-108">自定义函数在用户首次运行外接程序且在所有工作簿中对同一用户可用时注册。</span><span class="sxs-lookup"><span data-stu-id="bad80-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="bad80-109">本文介绍了 JSON 元数据文件的格式, 假定您正在手动编写元数据文件。</span><span class="sxs-lookup"><span data-stu-id="bad80-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="bad80-110">有关 JSDoc 注释 JSON 文件生成的信息, 请参阅[为自定义函数生成 JSON 元数据](custom-functions-json-autogeneration.md)。</span><span class="sxs-lookup"><span data-stu-id="bad80-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="bad80-111">有关为启用自定义函数必须在加载项项目中包含的其他文件的信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="bad80-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> <span data-ttu-id="bad80-112">托管 JSON 文件的服务器上的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)，以便自定义函数在 Excel Online 中正常工作。</span><span class="sxs-lookup"><span data-stu-id="bad80-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="bad80-113">示例元数据</span><span class="sxs-lookup"><span data-stu-id="bad80-113">Example metadata</span></span>

<span data-ttu-id="bad80-114">以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="bad80-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="bad80-115">此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="bad80-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
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
        "type": "number",
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
> <span data-ttu-id="bad80-116">在 [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub 存储库中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="bad80-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="bad80-117">functions</span><span class="sxs-lookup"><span data-stu-id="bad80-117">functions</span></span> 

<span data-ttu-id="bad80-118">`functions` 属性是自定义函数对象的一个数组。</span><span class="sxs-lookup"><span data-stu-id="bad80-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="bad80-119">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bad80-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="bad80-120">属性</span><span class="sxs-lookup"><span data-stu-id="bad80-120">Property</span></span>  |  <span data-ttu-id="bad80-121">数据类型</span><span class="sxs-lookup"><span data-stu-id="bad80-121">Data type</span></span>  |  <span data-ttu-id="bad80-122">必需</span><span class="sxs-lookup"><span data-stu-id="bad80-122">Required</span></span>  |  <span data-ttu-id="bad80-123">说明</span><span class="sxs-lookup"><span data-stu-id="bad80-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="bad80-124">string</span><span class="sxs-lookup"><span data-stu-id="bad80-124">string</span></span>  |  <span data-ttu-id="bad80-125">否</span><span class="sxs-lookup"><span data-stu-id="bad80-125">No</span></span>  |  <span data-ttu-id="bad80-126">最终用户在 Excel 中看到的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="bad80-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="bad80-127">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="bad80-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="bad80-128">string</span><span class="sxs-lookup"><span data-stu-id="bad80-128">string</span></span>  |   <span data-ttu-id="bad80-129">否</span><span class="sxs-lookup"><span data-stu-id="bad80-129">No</span></span>  |  <span data-ttu-id="bad80-130">提供有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="bad80-130">URL that provides information about the function.</span></span> <span data-ttu-id="bad80-131">（它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。</span><span class="sxs-lookup"><span data-stu-id="bad80-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="bad80-132">string</span><span class="sxs-lookup"><span data-stu-id="bad80-132">string</span></span> | <span data-ttu-id="bad80-133">是</span><span class="sxs-lookup"><span data-stu-id="bad80-133">Yes</span></span> | <span data-ttu-id="bad80-134">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="bad80-134">A unique ID for the function.</span></span> <span data-ttu-id="bad80-135">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="bad80-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="bad80-136">string</span><span class="sxs-lookup"><span data-stu-id="bad80-136">string</span></span>  |  <span data-ttu-id="bad80-137">是</span><span class="sxs-lookup"><span data-stu-id="bad80-137">Yes</span></span>  |  <span data-ttu-id="bad80-138">最终用户在 Excel 中看到的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="bad80-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="bad80-139">在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="bad80-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="bad80-140">object</span><span class="sxs-lookup"><span data-stu-id="bad80-140">object</span></span>  |  <span data-ttu-id="bad80-141">否</span><span class="sxs-lookup"><span data-stu-id="bad80-141">No</span></span>  |  <span data-ttu-id="bad80-142">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="bad80-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="bad80-143">有关详细信息，请参阅[选项](#options)。</span><span class="sxs-lookup"><span data-stu-id="bad80-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="bad80-144">array</span><span class="sxs-lookup"><span data-stu-id="bad80-144">array</span></span>  |  <span data-ttu-id="bad80-145">是</span><span class="sxs-lookup"><span data-stu-id="bad80-145">Yes</span></span>  |  <span data-ttu-id="bad80-146">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="bad80-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="bad80-147">有关详细信息，请参阅[参数](#parameters)。</span><span class="sxs-lookup"><span data-stu-id="bad80-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="bad80-148">object</span><span class="sxs-lookup"><span data-stu-id="bad80-148">object</span></span>  |  <span data-ttu-id="bad80-149">是</span><span class="sxs-lookup"><span data-stu-id="bad80-149">Yes</span></span>  |  <span data-ttu-id="bad80-150">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="bad80-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="bad80-151">有关详细信息，请参阅[结果](#result)。</span><span class="sxs-lookup"><span data-stu-id="bad80-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="bad80-152">options</span><span class="sxs-lookup"><span data-stu-id="bad80-152">options</span></span>

<span data-ttu-id="bad80-153">`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="bad80-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="bad80-154">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bad80-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="bad80-155">属性</span><span class="sxs-lookup"><span data-stu-id="bad80-155">Property</span></span>  |  <span data-ttu-id="bad80-156">数据类型</span><span class="sxs-lookup"><span data-stu-id="bad80-156">Data type</span></span>  |  <span data-ttu-id="bad80-157">必需</span><span class="sxs-lookup"><span data-stu-id="bad80-157">Required</span></span>  |  <span data-ttu-id="bad80-158">说明</span><span class="sxs-lookup"><span data-stu-id="bad80-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="bad80-159">boolean</span><span class="sxs-lookup"><span data-stu-id="bad80-159">boolean</span></span>  |  <span data-ttu-id="bad80-160">否</span><span class="sxs-lookup"><span data-stu-id="bad80-160">No</span></span><br/><br/><span data-ttu-id="bad80-161">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bad80-161">Default value is `false`.</span></span>  |  <span data-ttu-id="bad80-162">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="bad80-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="bad80-163">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="bad80-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="bad80-164">（请***不要***在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="bad80-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="bad80-165">在函数正文中，必须将处理程序分配给 `caller.onCanceled` 成员。</span><span class="sxs-lookup"><span data-stu-id="bad80-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="bad80-166">有关详细信息，请参阅[取消函数](custom-functions-web-reqs.md#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="bad80-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="bad80-167">boolean</span><span class="sxs-lookup"><span data-stu-id="bad80-167">boolean</span></span> | <span data-ttu-id="bad80-168">否</span><span class="sxs-lookup"><span data-stu-id="bad80-168">No</span></span> <br/><br/><span data-ttu-id="bad80-169">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bad80-169">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="bad80-170">如果为 true, 则自定义函数可以访问调用自定义函数的单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="bad80-170">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="bad80-171">若要获取调用自定义函数的单元格的地址, 请在自定义函数中使用 context。</span><span class="sxs-lookup"><span data-stu-id="bad80-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="bad80-172">有关详细信息，请参阅[确定调用自定义函数的单元格](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function)。</span><span class="sxs-lookup"><span data-stu-id="bad80-172">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="bad80-173">不能将自定义函数同时设置为流式处理和 requiresAddress。</span><span class="sxs-lookup"><span data-stu-id="bad80-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="bad80-174">使用此选项时, "invocationContext" 参数必须是在 options 中传递的最后一个参数。</span><span class="sxs-lookup"><span data-stu-id="bad80-174">When using this option, the 'invocationContext' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="bad80-175">boolean</span><span class="sxs-lookup"><span data-stu-id="bad80-175">boolean</span></span>  |  <span data-ttu-id="bad80-176">否</span><span class="sxs-lookup"><span data-stu-id="bad80-176">No</span></span><br/><br/><span data-ttu-id="bad80-177">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bad80-177">Default value is `false`.</span></span>  |  <span data-ttu-id="bad80-178">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="bad80-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="bad80-179">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="bad80-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="bad80-180">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="bad80-180">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="bad80-181">（请***不要***在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="bad80-181">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="bad80-182">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="bad80-182">The function should have no `return` statement.</span></span> <span data-ttu-id="bad80-183">相反，结果值将作为 `caller.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="bad80-183">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="bad80-184">有关详细信息，请参阅[流式处理函数](custom-functions-web-reqs.md#streaming-functions)。</span><span class="sxs-lookup"><span data-stu-id="bad80-184">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="bad80-185">boolean</span><span class="sxs-lookup"><span data-stu-id="bad80-185">boolean</span></span> | <span data-ttu-id="bad80-186">否</span><span class="sxs-lookup"><span data-stu-id="bad80-186">No</span></span> <br/><br/><span data-ttu-id="bad80-187">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="bad80-187">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="bad80-188">如果为 `true`，则该函数会在每次 Excel 重新计算时（而不是仅当公式的从属值发生更改时）进行重新计算。</span><span class="sxs-lookup"><span data-stu-id="bad80-188">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="bad80-189">函数不能同时为流式处理和可变。</span><span class="sxs-lookup"><span data-stu-id="bad80-189">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="bad80-190">如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。</span><span class="sxs-lookup"><span data-stu-id="bad80-190">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="bad80-191">参数</span><span class="sxs-lookup"><span data-stu-id="bad80-191">parameters</span></span>

<span data-ttu-id="bad80-192">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="bad80-192">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="bad80-193">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bad80-193">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="bad80-194">属性</span><span class="sxs-lookup"><span data-stu-id="bad80-194">Property</span></span>  |  <span data-ttu-id="bad80-195">数据类型</span><span class="sxs-lookup"><span data-stu-id="bad80-195">Data type</span></span>  |  <span data-ttu-id="bad80-196">必需</span><span class="sxs-lookup"><span data-stu-id="bad80-196">Required</span></span>  |  <span data-ttu-id="bad80-197">说明</span><span class="sxs-lookup"><span data-stu-id="bad80-197">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="bad80-198">string</span><span class="sxs-lookup"><span data-stu-id="bad80-198">string</span></span>  |  <span data-ttu-id="bad80-199">否</span><span class="sxs-lookup"><span data-stu-id="bad80-199">No</span></span> |  <span data-ttu-id="bad80-200">参数的说明。</span><span class="sxs-lookup"><span data-stu-id="bad80-200">A description of the parameter.</span></span> <span data-ttu-id="bad80-201">这显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="bad80-201">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="bad80-202">string</span><span class="sxs-lookup"><span data-stu-id="bad80-202">string</span></span>  |  <span data-ttu-id="bad80-203">否</span><span class="sxs-lookup"><span data-stu-id="bad80-203">No</span></span>  |  <span data-ttu-id="bad80-204">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="bad80-204">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="bad80-205">string</span><span class="sxs-lookup"><span data-stu-id="bad80-205">string</span></span>  |  <span data-ttu-id="bad80-206">是</span><span class="sxs-lookup"><span data-stu-id="bad80-206">Yes</span></span>  |  <span data-ttu-id="bad80-207">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="bad80-207">The name of the parameter.</span></span> <span data-ttu-id="bad80-208">此名称显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="bad80-208">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="bad80-209">string</span><span class="sxs-lookup"><span data-stu-id="bad80-209">string</span></span>  |  <span data-ttu-id="bad80-210">否</span><span class="sxs-lookup"><span data-stu-id="bad80-210">No</span></span>  |  <span data-ttu-id="bad80-211">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="bad80-211">The data type of the parameter.</span></span> <span data-ttu-id="bad80-212">可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="bad80-212">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="bad80-213">如果未指定此属性，则数据类型默认为 **any**。</span><span class="sxs-lookup"><span data-stu-id="bad80-213">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="bad80-214">boolean</span><span class="sxs-lookup"><span data-stu-id="bad80-214">boolean</span></span> | <span data-ttu-id="bad80-215">否</span><span class="sxs-lookup"><span data-stu-id="bad80-215">No</span></span> | <span data-ttu-id="bad80-216">如果为 `true`，则参数是可选的。</span><span class="sxs-lookup"><span data-stu-id="bad80-216">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="bad80-217">如果可选参数的 `type` 属性未指定或设置为 `any`，则可能会发现 IDE 中的 lint 错误以及当将函数输入到 Excel 的单元格中时未显示可选参数等问题。</span><span class="sxs-lookup"><span data-stu-id="bad80-217">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="bad80-218">预计将于 2018 年 12 月有所改变。</span><span class="sxs-lookup"><span data-stu-id="bad80-218">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="bad80-219">结果</span><span class="sxs-lookup"><span data-stu-id="bad80-219">result</span></span>

<span data-ttu-id="bad80-220">`result` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="bad80-220">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="bad80-221">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="bad80-221">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="bad80-222">属性</span><span class="sxs-lookup"><span data-stu-id="bad80-222">Property</span></span>  |  <span data-ttu-id="bad80-223">数据类型</span><span class="sxs-lookup"><span data-stu-id="bad80-223">Data type</span></span>  |  <span data-ttu-id="bad80-224">必需</span><span class="sxs-lookup"><span data-stu-id="bad80-224">Required</span></span>  |  <span data-ttu-id="bad80-225">说明</span><span class="sxs-lookup"><span data-stu-id="bad80-225">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="bad80-226">string</span><span class="sxs-lookup"><span data-stu-id="bad80-226">string</span></span>  |  <span data-ttu-id="bad80-227">否</span><span class="sxs-lookup"><span data-stu-id="bad80-227">No</span></span>  |  <span data-ttu-id="bad80-228">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="bad80-228">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="bad80-229">string</span><span class="sxs-lookup"><span data-stu-id="bad80-229">string</span></span>  |  <span data-ttu-id="bad80-230">是</span><span class="sxs-lookup"><span data-stu-id="bad80-230">Yes</span></span>  |  <span data-ttu-id="bad80-231">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="bad80-231">The data type of the parameter.</span></span> <span data-ttu-id="bad80-232">必须是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="bad80-232">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="bad80-233">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bad80-233">See also</span></span>

* [<span data-ttu-id="bad80-234">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="bad80-234">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="bad80-235">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="bad80-235">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="bad80-236">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="bad80-236">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="bad80-237">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="bad80-237">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="bad80-238">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="bad80-238">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
