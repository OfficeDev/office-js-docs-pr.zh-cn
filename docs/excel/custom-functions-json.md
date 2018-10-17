---
ms.date: 09/27/2018
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中的自定义函数的元数据
ms.openlocfilehash: b7c7f26d56309f43758b9b13025ebaad661aeaed
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579868"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="de495-103">自定义函数元数据 （预览）</span><span class="sxs-lookup"><span data-stu-id="de495-103">Custom functions metadata</span></span>

<span data-ttu-id="de495-p101">在 Excel 加载项内定义[自定义函数](custom-functions-overview.md)时，加载项项目必须包含一个 JSON 元数据文件，它提供 Excel 需要用来注册自定义函数并使其为最终用户可用的信息。本文介绍了 JSON 元数据文件的格式。</span><span class="sxs-lookup"><span data-stu-id="de495-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="de495-106">有关必须包含在加载项项目中以启用自定义函数的其他文件的信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="de495-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="de495-107">元数据示例</span><span class="sxs-lookup"><span data-stu-id="de495-107">Example metadata</span></span>

<span data-ttu-id="de495-p102">下面的示例显示用于定义自定义函数的加载项的 JSON 元数据文件的内容。下面示例中的各节提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="de495-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="de495-110">[OfficeDev/Excel-Custom-Functions ](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)GitHub 存储库中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="de495-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="de495-111">functions</span><span class="sxs-lookup"><span data-stu-id="de495-111">functions</span></span> 

<span data-ttu-id="de495-p103">`functions` 属性是自定义函数对象的数组。下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="de495-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="de495-114">属性</span><span class="sxs-lookup"><span data-stu-id="de495-114">Property</span></span>  |  <span data-ttu-id="de495-115">数据类型</span><span class="sxs-lookup"><span data-stu-id="de495-115">Data type</span></span>  |  <span data-ttu-id="de495-116">是否必需</span><span class="sxs-lookup"><span data-stu-id="de495-116">Required</span></span>  |  <span data-ttu-id="de495-117">说明</span><span class="sxs-lookup"><span data-stu-id="de495-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="de495-118">String</span><span class="sxs-lookup"><span data-stu-id="de495-118">string</span></span>  |  <span data-ttu-id="de495-119">否</span><span class="sxs-lookup"><span data-stu-id="de495-119">No</span></span>  |  <span data-ttu-id="de495-p104">最终用户在 Excel 中看到的函数的说明。例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="de495-p104">The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="de495-122">String</span><span class="sxs-lookup"><span data-stu-id="de495-122">string</span></span>  |   <span data-ttu-id="de495-123">否</span><span class="sxs-lookup"><span data-stu-id="de495-123">No</span></span>  |  <span data-ttu-id="de495-p105">提供有关函数的信息的 URL。（它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。</span><span class="sxs-lookup"><span data-stu-id="de495-p105">URL that provides information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="de495-126">String</span><span class="sxs-lookup"><span data-stu-id="de495-126">string</span></span> | <span data-ttu-id="de495-127">是</span><span class="sxs-lookup"><span data-stu-id="de495-127">Yes</span></span> | <span data-ttu-id="de495-128">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="de495-128">A unique ID for the function.</span></span> <span data-ttu-id="de495-129">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="de495-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="de495-130">String</span><span class="sxs-lookup"><span data-stu-id="de495-130">string</span></span>  |  <span data-ttu-id="de495-131">是</span><span class="sxs-lookup"><span data-stu-id="de495-131">Yes</span></span>  |  <span data-ttu-id="de495-p107">最终用户在 Excel 中看到的函数的名称。在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="de495-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="de495-134">object</span><span class="sxs-lookup"><span data-stu-id="de495-134">object</span></span>  |  <span data-ttu-id="de495-135">否</span><span class="sxs-lookup"><span data-stu-id="de495-135">No</span></span>  |  <span data-ttu-id="de495-p108">使你可以自定义 Excel 执行函数的方式和时间等的某些方面。有关详细信息，请参阅[选项对象](#options-object)。</span><span class="sxs-lookup"><span data-stu-id="de495-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="de495-138">数组</span><span class="sxs-lookup"><span data-stu-id="de495-138">array</span></span>  |  <span data-ttu-id="de495-139">是</span><span class="sxs-lookup"><span data-stu-id="de495-139">Yes</span></span>  |  <span data-ttu-id="de495-p109">定义函数的输入参数的数组。有关详细信息，请参阅[参数数组](#parameters-array)。</span><span class="sxs-lookup"><span data-stu-id="de495-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="de495-142">object</span><span class="sxs-lookup"><span data-stu-id="de495-142">object</span></span>  |  <span data-ttu-id="de495-143">是</span><span class="sxs-lookup"><span data-stu-id="de495-143">Yes</span></span>  |  <span data-ttu-id="de495-p110">定义函数返回的信息类型的对象。有关详细信息，请参阅[结果对象](#result-object)。</span><span class="sxs-lookup"><span data-stu-id="de495-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="de495-146">options</span><span class="sxs-lookup"><span data-stu-id="de495-146">options</span></span>

<span data-ttu-id="de495-p111">`options` 对象使你可以自定义 Excel 执行函数的方式和时间等的某些方面。下表列出 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="de495-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="de495-149">属性</span><span class="sxs-lookup"><span data-stu-id="de495-149">Property</span></span>  |  <span data-ttu-id="de495-150">数据类型</span><span class="sxs-lookup"><span data-stu-id="de495-150">Data type</span></span>  |  <span data-ttu-id="de495-151">是否必需</span><span class="sxs-lookup"><span data-stu-id="de495-151">Required</span></span>  |  <span data-ttu-id="de495-152">说明</span><span class="sxs-lookup"><span data-stu-id="de495-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="de495-153">boolean</span><span class="sxs-lookup"><span data-stu-id="de495-153">boolean</span></span>  |  <span data-ttu-id="de495-154">否</span><span class="sxs-lookup"><span data-stu-id="de495-154">No</span></span><br/><br/><span data-ttu-id="de495-155">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="de495-155">Default value is 4.</span></span>  |  <span data-ttu-id="de495-p112">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。如果您使用此选项，Excel 将使用其他 `caller` 参数调用 JavaScript 函数。（请***不要***在 `parameters` 属性中注册此参数）。在函数的正文中，必须将处理程序分配给 `caller.onCanceled` 成员。有关详细信息，请参阅[取消函数](custom-functions-overview.md#canceling-a-function) 。</span><span class="sxs-lookup"><span data-stu-id="de495-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="de495-161">boolean</span><span class="sxs-lookup"><span data-stu-id="de495-161">boolean</span></span>  |  <span data-ttu-id="de495-162">否</span><span class="sxs-lookup"><span data-stu-id="de495-162">No</span></span><br/><br/><span data-ttu-id="de495-163">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="de495-163">Default value is 4.</span></span>  |  <span data-ttu-id="de495-164">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="de495-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="de495-165">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="de495-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="de495-166">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="de495-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="de495-167">（请***不要***在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="de495-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="de495-168">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="de495-168">The function should have no `return` statement.</span></span> <span data-ttu-id="de495-169">相反，结果值将作为 `caller.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="de495-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="de495-170">有关详细信息，请参阅[流式函数](custom-functions-overview.md#streaming-functions)。</span><span class="sxs-lookup"><span data-stu-id="de495-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="de495-171">parameters</span><span class="sxs-lookup"><span data-stu-id="de495-171">parameters</span></span>

<span data-ttu-id="de495-p114">`parameters` 属性是参数对象的数组。下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="de495-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="de495-174">属性</span><span class="sxs-lookup"><span data-stu-id="de495-174">Property</span></span>  |  <span data-ttu-id="de495-175">数据类型</span><span class="sxs-lookup"><span data-stu-id="de495-175">Data type</span></span>  |  <span data-ttu-id="de495-176">是否必需</span><span class="sxs-lookup"><span data-stu-id="de495-176">Required</span></span>  |  <span data-ttu-id="de495-177">说明</span><span class="sxs-lookup"><span data-stu-id="de495-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="de495-178">String</span><span class="sxs-lookup"><span data-stu-id="de495-178">string</span></span>  |  <span data-ttu-id="de495-179">否</span><span class="sxs-lookup"><span data-stu-id="de495-179">No</span></span> |  <span data-ttu-id="de495-180">参数的描述。</span><span class="sxs-lookup"><span data-stu-id="de495-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="de495-181">String</span><span class="sxs-lookup"><span data-stu-id="de495-181">string</span></span>  |  <span data-ttu-id="de495-182">否</span><span class="sxs-lookup"><span data-stu-id="de495-182">No</span></span>  |  <span data-ttu-id="de495-183">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="de495-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="de495-184">String</span><span class="sxs-lookup"><span data-stu-id="de495-184">string</span></span>  |  <span data-ttu-id="de495-185">是</span><span class="sxs-lookup"><span data-stu-id="de495-185">Yes</span></span>  |  <span data-ttu-id="de495-p115">参数的名称。此名称显示在 Excel 的 IntelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="de495-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="de495-188">String</span><span class="sxs-lookup"><span data-stu-id="de495-188">string</span></span>  |  <span data-ttu-id="de495-189">否</span><span class="sxs-lookup"><span data-stu-id="de495-189">No</span></span>  |  <span data-ttu-id="de495-p116">参数的数据类型。必须是 **boolean**、 **number** 或 **string**。</span><span class="sxs-lookup"><span data-stu-id="de495-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="de495-192">result</span><span class="sxs-lookup"><span data-stu-id="de495-192">result</span></span>

<span data-ttu-id="de495-p117">`results` 对象定义函数返回的信息类型。下表列出 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="de495-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="de495-195">属性</span><span class="sxs-lookup"><span data-stu-id="de495-195">Property</span></span>  |  <span data-ttu-id="de495-196">数据类型</span><span class="sxs-lookup"><span data-stu-id="de495-196">Data type</span></span>  |  <span data-ttu-id="de495-197">是否必需</span><span class="sxs-lookup"><span data-stu-id="de495-197">Required</span></span>  |  <span data-ttu-id="de495-198">说明</span><span class="sxs-lookup"><span data-stu-id="de495-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="de495-199">String</span><span class="sxs-lookup"><span data-stu-id="de495-199">string</span></span>  |  <span data-ttu-id="de495-200">否</span><span class="sxs-lookup"><span data-stu-id="de495-200">No</span></span>  |  <span data-ttu-id="de495-201">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="de495-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="de495-202">String</span><span class="sxs-lookup"><span data-stu-id="de495-202">string</span></span>  |  <span data-ttu-id="de495-203">是</span><span class="sxs-lookup"><span data-stu-id="de495-203">Yes</span></span>  |  <span data-ttu-id="de495-p118">参数的数据类型。必须是 **boolean**、**number** 或 **string**。</span><span class="sxs-lookup"><span data-stu-id="de495-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="de495-206">另请参阅</span><span class="sxs-lookup"><span data-stu-id="de495-206">See also</span></span>

* [<span data-ttu-id="de495-207">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="de495-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="de495-208">Excel 自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="de495-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="de495-209">自定义函数最佳做法</span><span class="sxs-lookup"><span data-stu-id="de495-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="de495-210">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="de495-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)