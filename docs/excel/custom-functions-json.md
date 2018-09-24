---
ms.date: 09/20/2018
description: 在 Excel 中定义自定义函数的元数据。
title: 在 Excel 中的自定义函数的元数据
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062142"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="6e6c5-103">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="6e6c5-103">Custom functions metadata</span></span>

<span data-ttu-id="6e6c5-104">在 Excel 加载项内定义[自定义函数](custom-functions-overview.md)时，加载项项目必须包含一个 JSON 元数据文件，它提供 Excel 需要用来注册自定义函数并使其为最终用户可用的信息。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="6e6c5-105">本文介绍了 JSON 元数据文件的格式。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-105">This article describes the format of the JSON file with examples.</span></span>

> [!NOTE]
> <span data-ttu-id="6e6c5-106">有关其他文件的信息，你必须在加载项项目中加入其他文件才能启用自定义函数，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md#learn-the-basics)。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md#learn-the-basics).</span></span>

## <a name="example-metadata"></a><span data-ttu-id="6e6c5-107">元数据示例</span><span class="sxs-lookup"><span data-stu-id="6e6c5-107">Example metadata</span></span>

<span data-ttu-id="6e6c5-108">下面的示例显示用于定义自定义函数的加载项的 JSON 元数据文件的内容。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="6e6c5-109">下面示例中的各节提供了有关此 JSON 示例中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
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
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
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
> <span data-ttu-id="6e6c5-110">[OfficeDev/Excel-Custom-Functions GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)中提供了完整的示例 JSON 文件。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="6e6c5-111">函数</span><span class="sxs-lookup"><span data-stu-id="6e6c5-111">functions</span></span> 

<span data-ttu-id="6e6c5-112"> `functions` 属性是自定义函数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="6e6c5-113">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-113">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="6e6c5-114">属性</span><span class="sxs-lookup"><span data-stu-id="6e6c5-114">Property</span></span>  |  <span data-ttu-id="6e6c5-115">数据类型</span><span class="sxs-lookup"><span data-stu-id="6e6c5-115">Data type</span></span>  |  <span data-ttu-id="6e6c5-116">必需</span><span class="sxs-lookup"><span data-stu-id="6e6c5-116">Required</span></span>  |  <span data-ttu-id="6e6c5-117">说明</span><span class="sxs-lookup"><span data-stu-id="6e6c5-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="6e6c5-118">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-118">string</span></span>  |  <span data-ttu-id="6e6c5-119">No</span><span class="sxs-lookup"><span data-stu-id="6e6c5-119">No</span></span>  |  <span data-ttu-id="6e6c5-120">Excel UI 中显示的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-120">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="6e6c5-121">例如，**将摄氏度值转换为华氏度**。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-121">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="6e6c5-122">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-122">string</span></span>  |   <span data-ttu-id="6e6c5-123">No</span><span class="sxs-lookup"><span data-stu-id="6e6c5-123">No</span></span>  |  <span data-ttu-id="6e6c5-124">用户可在其中获取有关函数的信息的 URL。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-124">URL where your users can get help about the function.</span></span> <span data-ttu-id="6e6c5-125">（它显示在任务窗格中。）例如，**http://contoso.com/help/convertcelsiustofahrenheit.html**。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-125">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span> |
| `id`     | <span data-ttu-id="6e6c5-126">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-126">string</span></span> | <span data-ttu-id="6e6c5-127">Yes</span><span class="sxs-lookup"><span data-stu-id="6e6c5-127">Yes</span></span> | <span data-ttu-id="6e6c5-128">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-128">A unique ID for the group.</span></span> <span data-ttu-id="6e6c5-129">设置之后，不应更改此 ID。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="6e6c5-130">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-130">string</span></span>  |  <span data-ttu-id="6e6c5-131">Yes</span><span class="sxs-lookup"><span data-stu-id="6e6c5-131">Yes</span></span>  |  <span data-ttu-id="6e6c5-132">用户选择函数时出现在 Excel 用户界面中的函数名称（以命名空间为前缀）。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-132">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="6e6c5-133">它不必与 JavaScript 中定义的函数名称相同。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-133">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="6e6c5-134">object</span><span class="sxs-lookup"><span data-stu-id="6e6c5-134">object</span></span>  |  <span data-ttu-id="6e6c5-135">No</span><span class="sxs-lookup"><span data-stu-id="6e6c5-135">No</span></span>  |  <span data-ttu-id="6e6c5-136">使你可以自定义 Excel 执行函数的方式和时间等的某些方面。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="6e6c5-137">有关详细信息，请参阅[选项对象](#options-object)。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="6e6c5-138">array</span><span class="sxs-lookup"><span data-stu-id="6e6c5-138">array</span></span>  |  <span data-ttu-id="6e6c5-139">Yes</span><span class="sxs-lookup"><span data-stu-id="6e6c5-139">Yes</span></span>  |  <span data-ttu-id="6e6c5-140">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="6e6c5-141">有关详细信息，请参阅[参数数组](#parameters-array)。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="6e6c5-142">object</span><span class="sxs-lookup"><span data-stu-id="6e6c5-142">object</span></span>  |  <span data-ttu-id="6e6c5-143">Yes</span><span class="sxs-lookup"><span data-stu-id="6e6c5-143">Yes</span></span>  |  <span data-ttu-id="6e6c5-144">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="6e6c5-145">有关详细信息，请参阅[结果对象](#result-object)。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="6e6c5-146">选项</span><span class="sxs-lookup"><span data-stu-id="6e6c5-146">options</span></span>

<span data-ttu-id="6e6c5-147">`options` 对象使你可以自定义 Excel 执行函数的方式和时间等的某些方面。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="6e6c5-148">下表列出了 `options` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-148">The following table lists the properties of the</span></span>

|  <span data-ttu-id="6e6c5-149">属性</span><span class="sxs-lookup"><span data-stu-id="6e6c5-149">Property</span></span>  |  <span data-ttu-id="6e6c5-150">数据类型</span><span class="sxs-lookup"><span data-stu-id="6e6c5-150">Data type</span></span>  |  <span data-ttu-id="6e6c5-151">必需</span><span class="sxs-lookup"><span data-stu-id="6e6c5-151">Required</span></span>  |  <span data-ttu-id="6e6c5-152">说明</span><span class="sxs-lookup"><span data-stu-id="6e6c5-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="6e6c5-153">boolean</span><span class="sxs-lookup"><span data-stu-id="6e6c5-153">boolean</span></span>  |  <span data-ttu-id="6e6c5-154">否，默认值为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-154">No, default is `false`.</span></span>  |  <span data-ttu-id="6e6c5-155">如果为 `true`，每当用户执行的操作会取消函数时，Excel 将调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-155">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="6e6c5-156">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-156">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="6e6c5-157">（请***不要*** 在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-157">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="6e6c5-158">在函数的主体中，必须将处理程序分配给 `caller.onCanceled` 成员。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-158">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="6e6c5-159">有关详细信息，请参阅[取消函数](custom-functions-overview.md#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-159">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="6e6c5-160">boolean</span><span class="sxs-lookup"><span data-stu-id="6e6c5-160">boolean</span></span>  |  <span data-ttu-id="6e6c5-161">否，默认值为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-161">No, default is `false`.</span></span>  |  <span data-ttu-id="6e6c5-162">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-162">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="6e6c5-163">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-163">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="6e6c5-164">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-164">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="6e6c5-165">（请***不要*** 在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-165">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="6e6c5-166">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-166">The function should have no `return` statement.</span></span> <span data-ttu-id="6e6c5-167">相反，结果值将作为 `caller.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-167">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="6e6c5-168">有关详细信息，请参阅[流式函数](custom-functions-overview.md#streamed-functions)。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-168">For more information, see [Excel functions by category](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="6e6c5-169">参数</span><span class="sxs-lookup"><span data-stu-id="6e6c5-169">parameters</span></span>

<span data-ttu-id="6e6c5-170">`parameters` 属性是参数对象的数组。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-170">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="6e6c5-171">下表列出了每个对象的属性。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-171">The following table lists the properties of the SP.FieldRatingScale object.</span></span>

|  <span data-ttu-id="6e6c5-172">属性</span><span class="sxs-lookup"><span data-stu-id="6e6c5-172">Property</span></span>  |  <span data-ttu-id="6e6c5-173">数据类型</span><span class="sxs-lookup"><span data-stu-id="6e6c5-173">Data type</span></span>  |  <span data-ttu-id="6e6c5-174">必需</span><span class="sxs-lookup"><span data-stu-id="6e6c5-174">Required</span></span>  |  <span data-ttu-id="6e6c5-175">说明</span><span class="sxs-lookup"><span data-stu-id="6e6c5-175">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="6e6c5-176">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-176">string</span></span>  |  <span data-ttu-id="6e6c5-177">No</span><span class="sxs-lookup"><span data-stu-id="6e6c5-177">No</span></span> |  <span data-ttu-id="6e6c5-178">参数的描述。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-178">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="6e6c5-179">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-179">string</span></span>  |  <span data-ttu-id="6e6c5-180">No</span><span class="sxs-lookup"><span data-stu-id="6e6c5-180">No</span></span>  |  <span data-ttu-id="6e6c5-181">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-181">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="6e6c5-182">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-182">string</span></span>  |  <span data-ttu-id="6e6c5-183">Yes</span><span class="sxs-lookup"><span data-stu-id="6e6c5-183">Yes</span></span>  |  <span data-ttu-id="6e6c5-184">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-184">The name of the parameter.</span></span> <span data-ttu-id="6e6c5-185">此名称显示在 Excel 的 intelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-185">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="6e6c5-186">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-186">string</span></span>  |  <span data-ttu-id="6e6c5-187">No</span><span class="sxs-lookup"><span data-stu-id="6e6c5-187">No</span></span>  |  <span data-ttu-id="6e6c5-188">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-188">The data type of the parameter.</span></span> <span data-ttu-id="6e6c5-189">必须是**布尔值**、**数字**或**字符串**。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-189">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="6e6c5-190">结果</span><span class="sxs-lookup"><span data-stu-id="6e6c5-190">result</span></span>

<span data-ttu-id="6e6c5-191">`results` 对象定义函数返回的信息类型。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-191">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="6e6c5-192">下表列出了 `result` 对象的属性。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-192">The following table lists the properties of the</span></span>

|  <span data-ttu-id="6e6c5-193">属性</span><span class="sxs-lookup"><span data-stu-id="6e6c5-193">Property</span></span>  |  <span data-ttu-id="6e6c5-194">数据类型</span><span class="sxs-lookup"><span data-stu-id="6e6c5-194">Data type</span></span>  |  <span data-ttu-id="6e6c5-195">必需</span><span class="sxs-lookup"><span data-stu-id="6e6c5-195">Required</span></span>  |  <span data-ttu-id="6e6c5-196">说明</span><span class="sxs-lookup"><span data-stu-id="6e6c5-196">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="6e6c5-197">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-197">string</span></span>  |  <span data-ttu-id="6e6c5-198">No</span><span class="sxs-lookup"><span data-stu-id="6e6c5-198">No</span></span>  |  <span data-ttu-id="6e6c5-199">必须是**标量**（非数组值）或**矩阵**（二维数组）。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-199">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="6e6c5-200">string</span><span class="sxs-lookup"><span data-stu-id="6e6c5-200">string</span></span>  |  <span data-ttu-id="6e6c5-201">Yes</span><span class="sxs-lookup"><span data-stu-id="6e6c5-201">Yes</span></span>  |  <span data-ttu-id="6e6c5-202">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-202">The data type of the parameter.</span></span> <span data-ttu-id="6e6c5-203">必须是**布尔值**、**数字**或**字符串**。</span><span class="sxs-lookup"><span data-stu-id="6e6c5-203">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="6e6c5-204">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6e6c5-204">See also</span></span>

* [<span data-ttu-id="6e6c5-205">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="6e6c5-205">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="6e6c5-206">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="6e6c5-206">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="6e6c5-207">自定义函数的最佳做法</span><span class="sxs-lookup"><span data-stu-id="6e6c5-207">Custom functions best practices</span></span>](custom-functions-best-practices.md)