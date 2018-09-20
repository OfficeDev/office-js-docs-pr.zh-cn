# <a name="custom-function-metadata"></a><span data-ttu-id="87ad5-101">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="87ad5-101">Custom function metadata</span></span>

<span data-ttu-id="87ad5-102">如果在 Excel 加载项中包括[自定义函数](custom-functions-overview.md)，你必须托管一个 JSON 文件，其中包含有关函数的元数据（此外，还要托管包含函数的 JavaScript文件，以及充当 JavaScript 文件父项的无用户界面的 HTML 文件）。</span><span class="sxs-lookup"><span data-stu-id="87ad5-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="87ad5-103">本文使用示例描述了 JSON 文件的格式。</span><span class="sxs-lookup"><span data-stu-id="87ad5-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="87ad5-104">[此处](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json)提供了一个完整的示例 JSON文件。</span><span class="sxs-lookup"><span data-stu-id="87ad5-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="87ad5-105">函数数组</span><span class="sxs-lookup"><span data-stu-id="87ad5-105">Functions array</span></span>

<span data-ttu-id="87ad5-106">元数据是包含单一 `functions` 属性的 JSON 对象，其值是一个对象数组。</span><span class="sxs-lookup"><span data-stu-id="87ad5-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="87ad5-107">其中的每个对象都代表一个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="87ad5-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="87ad5-108">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="87ad5-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="87ad5-109">属性</span><span class="sxs-lookup"><span data-stu-id="87ad5-109">Property</span></span>  |  <span data-ttu-id="87ad5-110">数据类型</span><span class="sxs-lookup"><span data-stu-id="87ad5-110">Data Type</span></span>  |  <span data-ttu-id="87ad5-111">是否必需？</span><span class="sxs-lookup"><span data-stu-id="87ad5-111">Required?</span></span>  |  <span data-ttu-id="87ad5-112">说明</span><span class="sxs-lookup"><span data-stu-id="87ad5-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="87ad5-113">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-113">string</span></span>  |  <span data-ttu-id="87ad5-114">否</span><span class="sxs-lookup"><span data-stu-id="87ad5-114">No</span></span>  |  <span data-ttu-id="87ad5-115">Excel 用户界面中显示的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="87ad5-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="87ad5-116">例如，“将摄氏度值转换为华氏度”。</span><span class="sxs-lookup"><span data-stu-id="87ad5-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="87ad5-117">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-117">string</span></span>  |   <span data-ttu-id="87ad5-118">否</span><span class="sxs-lookup"><span data-stu-id="87ad5-118">No</span></span>  |  <span data-ttu-id="87ad5-119">用户可在其中获得函数相关帮助的 URL。</span><span class="sxs-lookup"><span data-stu-id="87ad5-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="87ad5-120">（它显示在任务窗格中。）例如，“http://contoso.com/help/convertcelsiustofahrenheit.html”</span><span class="sxs-lookup"><span data-stu-id="87ad5-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="87ad5-121">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-121">string</span></span>  |  <span data-ttu-id="87ad5-122">是</span><span class="sxs-lookup"><span data-stu-id="87ad5-122">Yes</span></span>  |  <span data-ttu-id="87ad5-123">用户选择函数时出现在 Excel 用户界面中的函数名称（以命名空间为前缀）。</span><span class="sxs-lookup"><span data-stu-id="87ad5-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="87ad5-124">它应该与 JavaScript 中定义的函数名称相同。</span><span class="sxs-lookup"><span data-stu-id="87ad5-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="87ad5-125">对象</span><span class="sxs-lookup"><span data-stu-id="87ad5-125">object</span></span>  |  <span data-ttu-id="87ad5-126">否</span><span class="sxs-lookup"><span data-stu-id="87ad5-126">No</span></span>  |  <span data-ttu-id="87ad5-127">配置 Excel 处理函数的方式。</span><span class="sxs-lookup"><span data-stu-id="87ad5-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="87ad5-128">有关详细信息，请参阅[选项对象](#options-object)。</span><span class="sxs-lookup"><span data-stu-id="87ad5-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="87ad5-129">数组</span><span class="sxs-lookup"><span data-stu-id="87ad5-129">array</span></span>  |  <span data-ttu-id="87ad5-130">是</span><span class="sxs-lookup"><span data-stu-id="87ad5-130">Yes</span></span>  |  <span data-ttu-id="87ad5-131">有关函数参数的元数据。</span><span class="sxs-lookup"><span data-stu-id="87ad5-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="87ad5-132">有关详细信息，请参阅[参数数组](#parameters-array)。</span><span class="sxs-lookup"><span data-stu-id="87ad5-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="87ad5-133">对象</span><span class="sxs-lookup"><span data-stu-id="87ad5-133">object</span></span>  |  <span data-ttu-id="87ad5-134">是</span><span class="sxs-lookup"><span data-stu-id="87ad5-134">Yes</span></span>  |  <span data-ttu-id="87ad5-135">有关函数返回的值的元数据。</span><span class="sxs-lookup"><span data-stu-id="87ad5-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="87ad5-136">有关详细信息，请参阅[结果对象](#result-object)。</span><span class="sxs-lookup"><span data-stu-id="87ad5-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="87ad5-137">Options 对象</span><span class="sxs-lookup"><span data-stu-id="87ad5-137">Options object</span></span>

<span data-ttu-id="87ad5-138">对象配置 Excel 处理函数的 方式。`options`</span><span class="sxs-lookup"><span data-stu-id="87ad5-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="87ad5-139">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="87ad5-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="87ad5-140">属性</span><span class="sxs-lookup"><span data-stu-id="87ad5-140">Property</span></span>  |  <span data-ttu-id="87ad5-141">数据类型</span><span class="sxs-lookup"><span data-stu-id="87ad5-141">Data Type</span></span>  |  <span data-ttu-id="87ad5-142">是否必需？</span><span class="sxs-lookup"><span data-stu-id="87ad5-142">Required?</span></span>  |  <span data-ttu-id="87ad5-143">说明</span><span class="sxs-lookup"><span data-stu-id="87ad5-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="87ad5-144">布尔值</span><span class="sxs-lookup"><span data-stu-id="87ad5-144">boolean</span></span>  |  <span data-ttu-id="87ad5-145">否，默认值为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="87ad5-145">No, default is `false`.</span></span>  |  <span data-ttu-id="87ad5-p110">如果为 `true`，则每次用户执行具有取消函数效果的操作时，Excel 都会调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。如果您使用此选项，Excel 将使用其他 `caller` 参数调用 JavaScript 函数。（请***不要***在 `parameters` 属性中注册此参数）。在函数的正文中， 必须将处理程序分配给 `caller.onCanceled` 成员。</span><span class="sxs-lookup"><span data-stu-id="87ad5-p110">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. Note,  and  cannot both be .</span></span>|
|  `stream`  |  <span data-ttu-id="87ad5-150">布尔值</span><span class="sxs-lookup"><span data-stu-id="87ad5-150">boolean</span></span>  |  <span data-ttu-id="87ad5-151">否，默认值为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="87ad5-151">No, default is `false`.</span></span>  |  <span data-ttu-id="87ad5-152">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="87ad5-152">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="87ad5-153">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="87ad5-153">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="87ad5-154">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="87ad5-154">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="87ad5-155">（请***不要*** 在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="87ad5-155">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="87ad5-156">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="87ad5-156">The function should have no `return` statement.</span></span> <span data-ttu-id="87ad5-157">相反，结果值将作为 `caller.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="87ad5-157">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span>|

## <a name="parameters-array"></a><span data-ttu-id="87ad5-158">参数数组</span><span class="sxs-lookup"><span data-stu-id="87ad5-158">Parameters array</span></span>

<span data-ttu-id="87ad5-159">属性是一个对象数组。`parameters`</span><span class="sxs-lookup"><span data-stu-id="87ad5-159">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="87ad5-160">其中每个对象代表一个参数。</span><span class="sxs-lookup"><span data-stu-id="87ad5-160">Each of these objects represents a parameter.</span></span> <span data-ttu-id="87ad5-161">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="87ad5-161">The following table contains its properties:</span></span>

|  <span data-ttu-id="87ad5-162">属性</span><span class="sxs-lookup"><span data-stu-id="87ad5-162">Property</span></span>  |  <span data-ttu-id="87ad5-163">数据类型</span><span class="sxs-lookup"><span data-stu-id="87ad5-163">Data Type</span></span>  |  <span data-ttu-id="87ad5-164">是否必需？</span><span class="sxs-lookup"><span data-stu-id="87ad5-164">Required?</span></span>  |  <span data-ttu-id="87ad5-165">说明</span><span class="sxs-lookup"><span data-stu-id="87ad5-165">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="87ad5-166">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-166">string</span></span>  |  <span data-ttu-id="87ad5-167">否</span><span class="sxs-lookup"><span data-stu-id="87ad5-167">No</span></span> |  <span data-ttu-id="87ad5-168">参数的描述。</span><span class="sxs-lookup"><span data-stu-id="87ad5-168">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="87ad5-169">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-169">string</span></span>  |  <span data-ttu-id="87ad5-170">是</span><span class="sxs-lookup"><span data-stu-id="87ad5-170">Yes</span></span>  |  <span data-ttu-id="87ad5-171">必须是“标量”（即非数组值）或“矩阵”（即一系列行数组）。</span><span class="sxs-lookup"><span data-stu-id="87ad5-171">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="87ad5-172">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-172">string</span></span>  |  <span data-ttu-id="87ad5-173">是</span><span class="sxs-lookup"><span data-stu-id="87ad5-173">Yes</span></span>  |  <span data-ttu-id="87ad5-174">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="87ad5-174">The name of the parameter.</span></span> <span data-ttu-id="87ad5-175">此名称显示在 Excel 的 IntelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="87ad5-175">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="87ad5-176">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-176">string</span></span>  |  <span data-ttu-id="87ad5-177">是</span><span class="sxs-lookup"><span data-stu-id="87ad5-177">Yes</span></span>  |  <span data-ttu-id="87ad5-178">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="87ad5-178">The data type of the parameter.</span></span> <span data-ttu-id="87ad5-179">必须为“布尔值”、“数字”或“字符串”。</span><span class="sxs-lookup"><span data-stu-id="87ad5-179">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="87ad5-180">结果对象</span><span class="sxs-lookup"><span data-stu-id="87ad5-180">Result object</span></span>

<span data-ttu-id="87ad5-181">属性提供有关函数返回的值的元数据。`results`</span><span class="sxs-lookup"><span data-stu-id="87ad5-181">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="87ad5-182">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="87ad5-182">The following table contains its properties:</span></span>

|  <span data-ttu-id="87ad5-183">属性</span><span class="sxs-lookup"><span data-stu-id="87ad5-183">Property</span></span>  |  <span data-ttu-id="87ad5-184">数据类型</span><span class="sxs-lookup"><span data-stu-id="87ad5-184">Data Type</span></span>  |  <span data-ttu-id="87ad5-185">是否必需？</span><span class="sxs-lookup"><span data-stu-id="87ad5-185">Required?</span></span>  |  <span data-ttu-id="87ad5-186">说明</span><span class="sxs-lookup"><span data-stu-id="87ad5-186">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="87ad5-187">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-187">string</span></span>  |  <span data-ttu-id="87ad5-188">否</span><span class="sxs-lookup"><span data-stu-id="87ad5-188">No</span></span>  |  <span data-ttu-id="87ad5-189">必须是“标量”（即非数组值）或“矩阵”（即一系列行数组）。</span><span class="sxs-lookup"><span data-stu-id="87ad5-189">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="87ad5-190">字符串</span><span class="sxs-lookup"><span data-stu-id="87ad5-190">string</span></span>  |  <span data-ttu-id="87ad5-191">是</span><span class="sxs-lookup"><span data-stu-id="87ad5-191">Yes</span></span>  |  <span data-ttu-id="87ad5-192">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="87ad5-192">The data type of the parameter.</span></span> <span data-ttu-id="87ad5-193">必须为“布尔值”、“数字”或“字符串”。</span><span class="sxs-lookup"><span data-stu-id="87ad5-193">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="87ad5-194">示例</span><span class="sxs-lookup"><span data-stu-id="87ad5-194">Example</span></span>

<span data-ttu-id="87ad5-195">以下 JSON 代码是自定义函数元数据的示例。</span><span class="sxs-lookup"><span data-stu-id="87ad5-195">The following JSON code is an example of a metadata file for custom functions.</span></span>

```json
{
    "functions": [
        {
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
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
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

## <a name="see-also"></a><span data-ttu-id="87ad5-196">另请参阅</span><span class="sxs-lookup"><span data-stu-id="87ad5-196">See also</span></span>
[<span data-ttu-id="87ad5-197">自定义函数</span><span class="sxs-lookup"><span data-stu-id="87ad5-197">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="87ad5-198">有关数组公式的指导和示例</span><span class="sxs-lookup"><span data-stu-id="87ad5-198">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
