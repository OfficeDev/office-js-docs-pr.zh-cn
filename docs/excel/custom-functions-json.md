---
ms.date: 10/17/2018
description: 在 Excel 中定义自定义函数的元数据。
title: Excel 中的自定义函数的元数据
ms.openlocfilehash: 0c77474188a2deefd23a73bb64e87569bb1fa52a
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298542"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="a6c91-103">自定义函数元数据（预览）</span><span class="sxs-lookup"><span data-stu-id="a6c91-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="a6c91-104">在 Excel 加载项中定义[自定义函数](custom-functions-overview.md)时，加载项项目必须包含 JSON 元数据文件，该文件提供 Excel 注册自定义函数并使其可供最终用户使用所需的信息。</span><span class="sxs-lookup"><span data-stu-id="a6c91-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="a6c91-105">本文介绍了 JSON 元数据文件的格式。</span><span class="sxs-lookup"><span data-stu-id="a6c91-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="a6c91-106">有关为启用自定义函数必须在加载项项目中包含的其他文件的信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="a6c91-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="a6c91-107">示例元数据</span><span class="sxs-lookup"><span data-stu-id="a6c91-107">Example metadata</span></span>

<span data-ttu-id="a6c91-108">以下示例介绍了定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="a6c91-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="a6c91-109">此示例后面的部分提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="a6c91-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="a6c91-110">在 [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub 存储库中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="a6c91-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="a6c91-111">functions</span><span class="sxs-lookup"><span data-stu-id="a6c91-111">functions</span></span> 

<span data-ttu-id="a6c91-112">`functions` 属性是自定义函数对象的一个数组。</span><span class="sxs-lookup"><span data-stu-id="a6c91-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="a6c91-113">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="a6c91-113">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="a6c91-114">属性</span><span class="sxs-lookup"><span data-stu-id="a6c91-114">Property</span></span>  |  <span data-ttu-id="a6c91-115">数据类型</span><span class="sxs-lookup"><span data-stu-id="a6c91-115">Data type</span></span>  |  <span data-ttu-id="a6c91-116">是否必需</span><span class="sxs-lookup"><span data-stu-id="a6c91-116">Required</span></span>  |  <span data-ttu-id="a6c91-117">说明</span><span class="sxs-lookup"><span data-stu-id="a6c91-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="a6c91-118">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-118">string</span></span>  |  <span data-ttu-id="a6c91-119">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-119">No</span></span>  |  <span data-ttu-id="a6c91-120">最终用户在 Excel 中看到的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="a6c91-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="a6c91-121">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="a6c91-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="a6c91-122">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-122">string</span></span>  |   <span data-ttu-id="a6c91-123">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-123">No</span></span>  |  <span data-ttu-id="a6c91-124">提供有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="a6c91-124">URL that provides information about the function.</span></span> <span data-ttu-id="a6c91-125">（它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。</span><span class="sxs-lookup"><span data-stu-id="a6c91-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="a6c91-126">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-126">string</span></span> | <span data-ttu-id="a6c91-127">是</span><span class="sxs-lookup"><span data-stu-id="a6c91-127">Yes</span></span> | <span data-ttu-id="a6c91-128">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="a6c91-128">A unique ID for the group.</span></span> <span data-ttu-id="a6c91-129">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="a6c91-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="a6c91-130">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-130">string</span></span>  |  <span data-ttu-id="a6c91-131">是</span><span class="sxs-lookup"><span data-stu-id="a6c91-131">Yes</span></span>  |  <span data-ttu-id="a6c91-132">最终用户在 Excel 中看到的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="a6c91-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="a6c91-133">在 Excel 中，此函数名称将以 XML 清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="a6c91-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="a6c91-134">object</span><span class="sxs-lookup"><span data-stu-id="a6c91-134">object</span></span>  |  <span data-ttu-id="a6c91-135">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-135">No</span></span>  |  <span data-ttu-id="a6c91-136">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="a6c91-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="a6c91-137">有关详细信息，请参阅[选项对象](#options-object)。</span><span class="sxs-lookup"><span data-stu-id="a6c91-137">See object load [options](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="a6c91-138">array</span><span class="sxs-lookup"><span data-stu-id="a6c91-138">array</span></span>  |  <span data-ttu-id="a6c91-139">是</span><span class="sxs-lookup"><span data-stu-id="a6c91-139">Yes</span></span>  |  <span data-ttu-id="a6c91-140">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="a6c91-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="a6c91-141">有关详细信息，请参阅[参数数组](#parameters-array)。</span><span class="sxs-lookup"><span data-stu-id="a6c91-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="a6c91-142">object</span><span class="sxs-lookup"><span data-stu-id="a6c91-142">object</span></span>  |  <span data-ttu-id="a6c91-143">是</span><span class="sxs-lookup"><span data-stu-id="a6c91-143">Yes</span></span>  |  <span data-ttu-id="a6c91-144">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="a6c91-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="a6c91-145">有关详细信息，请参阅[结果对象](#result-object)。</span><span class="sxs-lookup"><span data-stu-id="a6c91-145">See object load [options](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="a6c91-146">options</span><span class="sxs-lookup"><span data-stu-id="a6c91-146">options</span></span>

<span data-ttu-id="a6c91-147">`options` 对象使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="a6c91-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="a6c91-148">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="a6c91-148">The following table lists the properties of the</span></span>

|  <span data-ttu-id="a6c91-149">属性</span><span class="sxs-lookup"><span data-stu-id="a6c91-149">Property</span></span>  |  <span data-ttu-id="a6c91-150">数据类型</span><span class="sxs-lookup"><span data-stu-id="a6c91-150">Data type</span></span>  |  <span data-ttu-id="a6c91-151">是否必需</span><span class="sxs-lookup"><span data-stu-id="a6c91-151">Required</span></span>  |  <span data-ttu-id="a6c91-152">说明</span><span class="sxs-lookup"><span data-stu-id="a6c91-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="a6c91-153">boolean</span><span class="sxs-lookup"><span data-stu-id="a6c91-153">boolean</span></span>  |  <span data-ttu-id="a6c91-154">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-154">No</span></span><br/><br/><span data-ttu-id="a6c91-155">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="a6c91-155">Default value is `false`.</span></span>  |  <span data-ttu-id="a6c91-156">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="a6c91-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="a6c91-157">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="a6c91-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="a6c91-158">（请***不要***在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="a6c91-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="a6c91-159">在函数正文中，必须将处理程序分配给 `caller.onCanceled` 成员。</span><span class="sxs-lookup"><span data-stu-id="a6c91-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="a6c91-160">有关详细信息，请参阅[取消函数](custom-functions-overview.md#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="a6c91-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="a6c91-161">boolean</span><span class="sxs-lookup"><span data-stu-id="a6c91-161">boolean</span></span>  |  <span data-ttu-id="a6c91-162">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-162">No</span></span><br/><br/><span data-ttu-id="a6c91-163">默认值为 `false`。</span><span class="sxs-lookup"><span data-stu-id="a6c91-163">Default value is `false`.</span></span>  |  <span data-ttu-id="a6c91-164">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="a6c91-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="a6c91-165">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="a6c91-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="a6c91-166">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="a6c91-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="a6c91-167">（请***不要***在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="a6c91-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="a6c91-168">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="a6c91-168">The function should have no `return` statement.</span></span> <span data-ttu-id="a6c91-169">相反，结果值将作为 `caller.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="a6c91-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="a6c91-170">有关详细信息，请参阅[流式处理函数](custom-functions-overview.md#streaming-functions)。</span><span class="sxs-lookup"><span data-stu-id="a6c91-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="a6c91-171">parameters</span><span class="sxs-lookup"><span data-stu-id="a6c91-171">parameters</span></span>

<span data-ttu-id="a6c91-172">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="a6c91-172">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="a6c91-173">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="a6c91-173">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="a6c91-174">属性</span><span class="sxs-lookup"><span data-stu-id="a6c91-174">Property</span></span>  |  <span data-ttu-id="a6c91-175">数据类型</span><span class="sxs-lookup"><span data-stu-id="a6c91-175">Data type</span></span>  |  <span data-ttu-id="a6c91-176">是否必需</span><span class="sxs-lookup"><span data-stu-id="a6c91-176">Required</span></span>  |  <span data-ttu-id="a6c91-177">说明</span><span class="sxs-lookup"><span data-stu-id="a6c91-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="a6c91-178">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-178">string</span></span>  |  <span data-ttu-id="a6c91-179">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-179">No</span></span> |  <span data-ttu-id="a6c91-180">参数的说明。</span><span class="sxs-lookup"><span data-stu-id="a6c91-180">A description of the value.</span></span> <span data-ttu-id="a6c91-181">这显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="a6c91-181">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="a6c91-182">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-182">string</span></span>  |  <span data-ttu-id="a6c91-183">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-183">No</span></span>  |  <span data-ttu-id="a6c91-184">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="a6c91-184">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="a6c91-185">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-185">string</span></span>  |  <span data-ttu-id="a6c91-186">是</span><span class="sxs-lookup"><span data-stu-id="a6c91-186">Yes</span></span>  |  <span data-ttu-id="a6c91-187">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="a6c91-187">The name of the parameter.</span></span> <span data-ttu-id="a6c91-188">此名称显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="a6c91-188">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="a6c91-189">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-189">string</span></span>  |  <span data-ttu-id="a6c91-190">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-190">No</span></span>  |  <span data-ttu-id="a6c91-191">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="a6c91-191">The System data type of the parameter.</span></span> <span data-ttu-id="a6c91-192">可以是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="a6c91-192">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="a6c91-193">如果未指定此属性，则数据类型默认为 **any**。</span><span class="sxs-lookup"><span data-stu-id="a6c91-193">If this property is not specified, the data type defaults to **any**.</span></span> |

## <a name="result"></a><span data-ttu-id="a6c91-194">result</span><span class="sxs-lookup"><span data-stu-id="a6c91-194">result</span></span>

<span data-ttu-id="a6c91-195">`result` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="a6c91-195">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="a6c91-196">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="a6c91-196">The following table lists the properties of the</span></span>

|  <span data-ttu-id="a6c91-197">属性</span><span class="sxs-lookup"><span data-stu-id="a6c91-197">Property</span></span>  |  <span data-ttu-id="a6c91-198">数据类型</span><span class="sxs-lookup"><span data-stu-id="a6c91-198">Data type</span></span>  |  <span data-ttu-id="a6c91-199">是否必需</span><span class="sxs-lookup"><span data-stu-id="a6c91-199">Required</span></span>  |  <span data-ttu-id="a6c91-200">说明</span><span class="sxs-lookup"><span data-stu-id="a6c91-200">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="a6c91-201">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-201">string</span></span>  |  <span data-ttu-id="a6c91-202">否</span><span class="sxs-lookup"><span data-stu-id="a6c91-202">No</span></span>  |  <span data-ttu-id="a6c91-203">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="a6c91-203">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="a6c91-204">string</span><span class="sxs-lookup"><span data-stu-id="a6c91-204">string</span></span>  |  <span data-ttu-id="a6c91-205">是</span><span class="sxs-lookup"><span data-stu-id="a6c91-205">Yes</span></span>  |  <span data-ttu-id="a6c91-206">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="a6c91-206">The System data type of the parameter.</span></span> <span data-ttu-id="a6c91-207">必须是 **boolean**、**number**、**string** 或 **any**，允许使用前三种类型中的任何一种。</span><span class="sxs-lookup"><span data-stu-id="a6c91-207">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="a6c91-208">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a6c91-208">See also</span></span>

* [<span data-ttu-id="a6c91-209">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="a6c91-209">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="a6c91-210">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="a6c91-210">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="a6c91-211">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="a6c91-211">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="a6c91-212">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="a6c91-212">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
