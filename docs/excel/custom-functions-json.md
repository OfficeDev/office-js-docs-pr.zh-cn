---
ms.date: 11/26/2018
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中的自定义函数的元数据
ms.openlocfilehash: 4bdf27173c5e912aa3eba3c8661ba45dd8b453cb
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724856"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="0d293-103">自定义函数元数据（预览）</span><span class="sxs-lookup"><span data-stu-id="0d293-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="0d293-104">在 Excel 加载项中定义[自定义函数](custom-functions-overview.md)时，加载项项目必须包含 JSON 元数据文件，该文件提供 Excel 注册自定义函数并使其可供最终用户使用所需的信息。</span><span class="sxs-lookup"><span data-stu-id="0d293-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="0d293-105">本文介绍了 JSON 元数据文件的格式。</span><span class="sxs-lookup"><span data-stu-id="0d293-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="0d293-106">有关为启用自定义函数必须在加载项项目中包含的其他文件的信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="0d293-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="0d293-107">示例元数据</span><span class="sxs-lookup"><span data-stu-id="0d293-107">Example metadata</span></span>

<span data-ttu-id="0d293-108">以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="0d293-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="0d293-109">此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="0d293-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="0d293-110">在 [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub 存储库中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="0d293-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="0d293-111">functions</span><span class="sxs-lookup"><span data-stu-id="0d293-111">functions</span></span> 

<span data-ttu-id="0d293-112">`functions` 属性是自定义函数对象的一个数组。</span><span class="sxs-lookup"><span data-stu-id="0d293-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="0d293-113">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0d293-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0d293-114">属性</span><span class="sxs-lookup"><span data-stu-id="0d293-114">Property</span></span>  |  <span data-ttu-id="0d293-115">数据类型</span><span class="sxs-lookup"><span data-stu-id="0d293-115">Data type</span></span>  |  <span data-ttu-id="0d293-116">是否必需</span><span class="sxs-lookup"><span data-stu-id="0d293-116">Required</span></span>  |  <span data-ttu-id="0d293-117">说明</span><span class="sxs-lookup"><span data-stu-id="0d293-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0d293-118">string</span><span class="sxs-lookup"><span data-stu-id="0d293-118">string</span></span>  |  <span data-ttu-id="0d293-119">否</span><span class="sxs-lookup"><span data-stu-id="0d293-119">No</span></span>  |  <span data-ttu-id="0d293-120">最终用户在 Excel 中看到的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="0d293-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="0d293-121">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="0d293-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="0d293-122">string</span><span class="sxs-lookup"><span data-stu-id="0d293-122">string</span></span>  |   <span data-ttu-id="0d293-123">否</span><span class="sxs-lookup"><span data-stu-id="0d293-123">No</span></span>  |  <span data-ttu-id="0d293-124">提供有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="0d293-124">URL that provides information about the function.</span></span> <span data-ttu-id="0d293-125">（它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。</span><span class="sxs-lookup"><span data-stu-id="0d293-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="0d293-126">string</span><span class="sxs-lookup"><span data-stu-id="0d293-126">string</span></span> | <span data-ttu-id="0d293-127">是</span><span class="sxs-lookup"><span data-stu-id="0d293-127">Yes</span></span> | <span data-ttu-id="0d293-128">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="0d293-128">A unique ID for the function.</span></span> <span data-ttu-id="0d293-129">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="0d293-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="0d293-130">string</span><span class="sxs-lookup"><span data-stu-id="0d293-130">string</span></span>  |  <span data-ttu-id="0d293-131">是</span><span class="sxs-lookup"><span data-stu-id="0d293-131">Yes</span></span>  |  <span data-ttu-id="0d293-132">最终用户在 Excel 中看到的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="0d293-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="0d293-133">在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="0d293-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="0d293-134">object</span><span class="sxs-lookup"><span data-stu-id="0d293-134">object</span></span>  |  <span data-ttu-id="0d293-135">否</span><span class="sxs-lookup"><span data-stu-id="0d293-135">No</span></span>  |  <span data-ttu-id="0d293-136">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="0d293-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0d293-137">有关详细信息，请参阅[选项](#options)。</span><span class="sxs-lookup"><span data-stu-id="0d293-137">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="0d293-138">array</span><span class="sxs-lookup"><span data-stu-id="0d293-138">array</span></span>  |  <span data-ttu-id="0d293-139">是</span><span class="sxs-lookup"><span data-stu-id="0d293-139">Yes</span></span>  |  <span data-ttu-id="0d293-140">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="0d293-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="0d293-141">有关详细信息，请参阅[参数](#parameters)。</span><span class="sxs-lookup"><span data-stu-id="0d293-141">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="0d293-142">object</span><span class="sxs-lookup"><span data-stu-id="0d293-142">object</span></span>  |  <span data-ttu-id="0d293-143">是</span><span class="sxs-lookup"><span data-stu-id="0d293-143">Yes</span></span>  |  <span data-ttu-id="0d293-144">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="0d293-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0d293-145">有关详细信息，请参阅[结果](#result)。</span><span class="sxs-lookup"><span data-stu-id="0d293-145">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="0d293-146">options</span><span class="sxs-lookup"><span data-stu-id="0d293-146">options</span></span>

<span data-ttu-id="0d293-147">`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="0d293-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0d293-148">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0d293-148">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="0d293-149">属性</span><span class="sxs-lookup"><span data-stu-id="0d293-149">Property</span></span>  |  <span data-ttu-id="0d293-150">数据类型</span><span class="sxs-lookup"><span data-stu-id="0d293-150">Data type</span></span>  |  <span data-ttu-id="0d293-151">是否必需</span><span class="sxs-lookup"><span data-stu-id="0d293-151">Required</span></span>  |  <span data-ttu-id="0d293-152">说明</span><span class="sxs-lookup"><span data-stu-id="0d293-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="0d293-153">boolean</span><span class="sxs-lookup"><span data-stu-id="0d293-153">boolean</span></span>  |  <span data-ttu-id="0d293-154">否</span><span class="sxs-lookup"><span data-stu-id="0d293-154">No</span></span><br/><br/><span data-ttu-id="0d293-155">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="0d293-155">Default value is `false`.</span></span>  |  <span data-ttu-id="0d293-156">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="0d293-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="0d293-157">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="0d293-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="0d293-158">（请***不要***在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="0d293-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="0d293-159">在函数正文中，必须将处理程序分配给 `caller.onCanceled` 成员。</span><span class="sxs-lookup"><span data-stu-id="0d293-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="0d293-160">有关详细信息，请参阅[取消函数](custom-functions-overview.md#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="0d293-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="0d293-161">boolean</span><span class="sxs-lookup"><span data-stu-id="0d293-161">boolean</span></span>  |  <span data-ttu-id="0d293-162">否</span><span class="sxs-lookup"><span data-stu-id="0d293-162">No</span></span><br/><br/><span data-ttu-id="0d293-163">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="0d293-163">Default value is `false`.</span></span>  |  <span data-ttu-id="0d293-164">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="0d293-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="0d293-165">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="0d293-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="0d293-166">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="0d293-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="0d293-167">（请***不要***在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="0d293-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="0d293-168">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="0d293-168">The function should have no `return` statement.</span></span> <span data-ttu-id="0d293-169">相反，结果值将作为 `caller.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="0d293-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="0d293-170">有关详细信息，请参阅[流式处理函数](custom-functions-overview.md#streaming-functions)。</span><span class="sxs-lookup"><span data-stu-id="0d293-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="0d293-171">boolean</span><span class="sxs-lookup"><span data-stu-id="0d293-171">boolean</span></span> | <span data-ttu-id="0d293-172">否</span><span class="sxs-lookup"><span data-stu-id="0d293-172">No</span></span> <br/><br/><span data-ttu-id="0d293-173">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="0d293-173">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="0d293-174">如果为 `true`，则该函数会在每次 Excel 重新计算时（而不是仅当公式的从属值发生更改时）进行重新计算。</span><span class="sxs-lookup"><span data-stu-id="0d293-174">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="0d293-175">函数不能同时为流式处理和可变。</span><span class="sxs-lookup"><span data-stu-id="0d293-175">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="0d293-176">如果 `stream` 和 `volatile` 属性同时设置为 `true`，则将忽略可变选项。</span><span class="sxs-lookup"><span data-stu-id="0d293-176">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="0d293-177">参数</span><span class="sxs-lookup"><span data-stu-id="0d293-177">parameters</span></span>

<span data-ttu-id="0d293-178">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="0d293-178">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="0d293-179">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0d293-179">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0d293-180">属性</span><span class="sxs-lookup"><span data-stu-id="0d293-180">Property</span></span>  |  <span data-ttu-id="0d293-181">数据类型</span><span class="sxs-lookup"><span data-stu-id="0d293-181">Data type</span></span>  |  <span data-ttu-id="0d293-182">是否必需</span><span class="sxs-lookup"><span data-stu-id="0d293-182">Required</span></span>  |  <span data-ttu-id="0d293-183">说明</span><span class="sxs-lookup"><span data-stu-id="0d293-183">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0d293-184">string</span><span class="sxs-lookup"><span data-stu-id="0d293-184">string</span></span>  |  <span data-ttu-id="0d293-185">否</span><span class="sxs-lookup"><span data-stu-id="0d293-185">No</span></span> |  <span data-ttu-id="0d293-186">参数的说明。</span><span class="sxs-lookup"><span data-stu-id="0d293-186">A description of the parameter.</span></span> <span data-ttu-id="0d293-187">这显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="0d293-187">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="0d293-188">string</span><span class="sxs-lookup"><span data-stu-id="0d293-188">string</span></span>  |  <span data-ttu-id="0d293-189">否</span><span class="sxs-lookup"><span data-stu-id="0d293-189">No</span></span>  |  <span data-ttu-id="0d293-190">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="0d293-190">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="0d293-191">string</span><span class="sxs-lookup"><span data-stu-id="0d293-191">string</span></span>  |  <span data-ttu-id="0d293-192">是</span><span class="sxs-lookup"><span data-stu-id="0d293-192">Yes</span></span>  |  <span data-ttu-id="0d293-193">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="0d293-193">The name of the parameter.</span></span> <span data-ttu-id="0d293-194">此名称显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="0d293-194">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="0d293-195">string</span><span class="sxs-lookup"><span data-stu-id="0d293-195">string</span></span>  |  <span data-ttu-id="0d293-196">否</span><span class="sxs-lookup"><span data-stu-id="0d293-196">No</span></span>  |  <span data-ttu-id="0d293-197">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="0d293-197">The data type of the parameter.</span></span> <span data-ttu-id="0d293-198">可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="0d293-198">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="0d293-199">如果未指定此属性，则数据类型默认为 **any**。</span><span class="sxs-lookup"><span data-stu-id="0d293-199">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="0d293-200">boolean</span><span class="sxs-lookup"><span data-stu-id="0d293-200">boolean</span></span> | <span data-ttu-id="0d293-201">否</span><span class="sxs-lookup"><span data-stu-id="0d293-201">No</span></span> | <span data-ttu-id="0d293-202">如果为 `true`，则参数是可选的。</span><span class="sxs-lookup"><span data-stu-id="0d293-202">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="0d293-203">如果可选参数的 `type` 属性未指定或设置为 `any`，则可能会发现 IDE 中的 lint 错误以及当将函数输入到 Excel 的单元格中时未显示可选参数等问题。</span><span class="sxs-lookup"><span data-stu-id="0d293-203">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="0d293-204">预计将于 2018 年 12 月有所改变。</span><span class="sxs-lookup"><span data-stu-id="0d293-204">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="0d293-205">结果</span><span class="sxs-lookup"><span data-stu-id="0d293-205">result</span></span>

<span data-ttu-id="0d293-206">`result` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="0d293-206">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0d293-207">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="0d293-207">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="0d293-208">属性</span><span class="sxs-lookup"><span data-stu-id="0d293-208">Property</span></span>  |  <span data-ttu-id="0d293-209">数据类型</span><span class="sxs-lookup"><span data-stu-id="0d293-209">Data type</span></span>  |  <span data-ttu-id="0d293-210">是否必需</span><span class="sxs-lookup"><span data-stu-id="0d293-210">Required</span></span>  |  <span data-ttu-id="0d293-211">说明</span><span class="sxs-lookup"><span data-stu-id="0d293-211">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="0d293-212">string</span><span class="sxs-lookup"><span data-stu-id="0d293-212">string</span></span>  |  <span data-ttu-id="0d293-213">否</span><span class="sxs-lookup"><span data-stu-id="0d293-213">No</span></span>  |  <span data-ttu-id="0d293-214">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="0d293-214">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="0d293-215">string</span><span class="sxs-lookup"><span data-stu-id="0d293-215">string</span></span>  |  <span data-ttu-id="0d293-216">是</span><span class="sxs-lookup"><span data-stu-id="0d293-216">Yes</span></span>  |  <span data-ttu-id="0d293-217">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="0d293-217">The data type of the parameter.</span></span> <span data-ttu-id="0d293-218">必须是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="0d293-218">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="0d293-219">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0d293-219">See also</span></span>

* [<span data-ttu-id="0d293-220">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="0d293-220">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="0d293-221">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="0d293-221">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0d293-222">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="0d293-222">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0d293-223">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="0d293-223">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
