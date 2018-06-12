# <a name="custom-function-metadata"></a><span data-ttu-id="ddfd6-101">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="ddfd6-101">Custom function metadata</span></span>

<span data-ttu-id="ddfd6-102">如果在 Excel 加载项中包括[自定义函数](custom-functions-overview.md)，你必须托管一个 JSON 文件，其中包含有关函数的元数据（此外，还要托管包含函数的 JavaScript文件，以及充当 JavaScript 文件父项的无用户界面的 HTML 文件）。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="ddfd6-103">本文使用示例描述了 JSON 文件的格式。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="ddfd6-104">[此处](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json)提供了一个完整的示例 JSON文件。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="ddfd6-105">函数数组</span><span class="sxs-lookup"><span data-stu-id="ddfd6-105">Functions array</span></span>

<span data-ttu-id="ddfd6-106">元数据是包含单一 `functions` 属性的 JSON 对象，其值是一个对象数组。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="ddfd6-107">其中的每个对象都代表一个自定义函数。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="ddfd6-108">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="ddfd6-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="ddfd6-109">属性</span><span class="sxs-lookup"><span data-stu-id="ddfd6-109">Property</span></span>  |  <span data-ttu-id="ddfd6-110">数据类型</span><span class="sxs-lookup"><span data-stu-id="ddfd6-110">Data Type</span></span>  |  <span data-ttu-id="ddfd6-111">是否必需？</span><span class="sxs-lookup"><span data-stu-id="ddfd6-111">Required?</span></span>  |  <span data-ttu-id="ddfd6-112">说明</span><span class="sxs-lookup"><span data-stu-id="ddfd6-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="ddfd6-113">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-113">string</span></span>  |  <span data-ttu-id="ddfd6-114">否</span><span class="sxs-lookup"><span data-stu-id="ddfd6-114">No</span></span>  |  <span data-ttu-id="ddfd6-115">Excel 用户界面中显示的函数的说明。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="ddfd6-116">例如，“将摄氏度值转换为华氏度”。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="ddfd6-117">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-117">string</span></span>  |   <span data-ttu-id="ddfd6-118">否</span><span class="sxs-lookup"><span data-stu-id="ddfd6-118">No</span></span>  |  <span data-ttu-id="ddfd6-119">用户可在其中获得函数相关帮助的 URL。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="ddfd6-120">（它显示在任务窗格中。）例如，“http://contoso.com/help/convertcelsiustofahrenheit.html”</span><span class="sxs-lookup"><span data-stu-id="ddfd6-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="ddfd6-121">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-121">string</span></span>  |  <span data-ttu-id="ddfd6-122">是</span><span class="sxs-lookup"><span data-stu-id="ddfd6-122">Yes</span></span>  |  <span data-ttu-id="ddfd6-123">用户选择函数时出现在 Excel 用户界面中的函数名称（以命名空间为前缀）。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="ddfd6-124">它应该与 JavaScript 中定义的函数名称相同。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="ddfd6-125">对象</span><span class="sxs-lookup"><span data-stu-id="ddfd6-125">object</span></span>  |  <span data-ttu-id="ddfd6-126">否</span><span class="sxs-lookup"><span data-stu-id="ddfd6-126">No</span></span>  |  <span data-ttu-id="ddfd6-127">配置 Excel 处理函数的方式。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="ddfd6-128">有关详细信息，请参阅[选项对象](#options-object)。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="ddfd6-129">数组</span><span class="sxs-lookup"><span data-stu-id="ddfd6-129">array</span></span>  |  <span data-ttu-id="ddfd6-130">是</span><span class="sxs-lookup"><span data-stu-id="ddfd6-130">Yes</span></span>  |  <span data-ttu-id="ddfd6-131">有关函数参数的元数据。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="ddfd6-132">有关详细信息，请参阅[参数数组](#parameters-array)。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="ddfd6-133">对象</span><span class="sxs-lookup"><span data-stu-id="ddfd6-133">object</span></span>  |  <span data-ttu-id="ddfd6-134">是</span><span class="sxs-lookup"><span data-stu-id="ddfd6-134">Yes</span></span>  |  <span data-ttu-id="ddfd6-135">有关函数返回的值的元数据。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="ddfd6-136">有关详细信息，请参阅[结果对象](#result-object)。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="ddfd6-137">Options 对象</span><span class="sxs-lookup"><span data-stu-id="ddfd6-137">Options object</span></span>

<span data-ttu-id="ddfd6-138">对象配置 Excel 处理函数的 方式。`options`</span><span class="sxs-lookup"><span data-stu-id="ddfd6-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="ddfd6-139">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="ddfd6-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="ddfd6-140">属性</span><span class="sxs-lookup"><span data-stu-id="ddfd6-140">Property</span></span>  |  <span data-ttu-id="ddfd6-141">数据类型</span><span class="sxs-lookup"><span data-stu-id="ddfd6-141">Data Type</span></span>  |  <span data-ttu-id="ddfd6-142">是否必需？</span><span class="sxs-lookup"><span data-stu-id="ddfd6-142">Required?</span></span>  |  <span data-ttu-id="ddfd6-143">说明</span><span class="sxs-lookup"><span data-stu-id="ddfd6-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="ddfd6-144">布尔值</span><span class="sxs-lookup"><span data-stu-id="ddfd6-144">boolean</span></span>  |  <span data-ttu-id="ddfd6-145">否，默认值为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-145">No, default is `false`.</span></span>  |  <span data-ttu-id="ddfd6-146">如果为 `true`，每当用户执行的操作会取消函数时，Excel 将调用 `onCanceled` 处理程序；例如，手动触发重新计算或编辑函数引用的单元格。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-146">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="ddfd6-147">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-147">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="ddfd6-148">（请***不要*** 在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-148">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="ddfd6-149">在函数的主体中，必须将处理程序分配给 `caller.onCanceled` 成员。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-149">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="ddfd6-150">注意，`cancelable` 和 `sync` 不能同时为 `true`。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-150">Note, `cancelable` and `sync` cannot both be `true`.</span></span>  |
|  `stream`  |  <span data-ttu-id="ddfd6-151">布尔值</span><span class="sxs-lookup"><span data-stu-id="ddfd6-151">boolean</span></span>  |  <span data-ttu-id="ddfd6-152">否，默认值为 `false` 。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-152">No, default is `false`.</span></span>  |  <span data-ttu-id="ddfd6-153">如果为 `true`，即使只调用一次，该函数也可能会重复输出到单元格。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-153">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="ddfd6-154">此选项对于快速变化的数据源（如股票价格）非常有用。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-154">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="ddfd6-155">如果使用此选项，Excel 将使用额外的 `caller` 参数调用 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-155">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="ddfd6-156">（请***不要*** 在 `parameters` 属性中注册此参数）。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-156">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="ddfd6-157">函数不应存在 `return` 语句。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-157">The function should have no `return` statement.</span></span> <span data-ttu-id="ddfd6-158">相反，结果值将作为 `caller.setResult` 回调方法的参数传递。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-158">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="ddfd6-159">注意，`stream` 和 `sync` 不能同时为 `true`。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-159">Note, `stream` and `sync` may not both be `true`.</span></span>|
|  `sync`  |  <span data-ttu-id="ddfd6-160">布尔值</span><span class="sxs-lookup"><span data-stu-id="ddfd6-160">boolean</span></span>  |  <span data-ttu-id="ddfd6-161">否，默认值为 `false`</span><span class="sxs-lookup"><span data-stu-id="ddfd6-161">No, default is `false`</span></span>  |  <span data-ttu-id="ddfd6-162">如果为 `true`，函数会同步运行，并且它必须返回一个值。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-162">If `true`, the function runs synchronously and it must return a value.</span></span> <span data-ttu-id="ddfd6-163">如果为 `false`，则函数将异步运行，并且它必须返回 `OfficeExtension.Promise` 对象。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-163">If `false`, the function runs asynchronously and it must return a `OfficeExtension.Promise` object.</span></span> <span data-ttu-id="ddfd6-164">注意，如果 `sync`  或 `true`  为 `cancelable` ，则 `stream`   不能为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-164">Note, `sync`  may not be `true` if either `cancelable` or `stream` are `true`.</span></span>  |

## <a name="parameters-array"></a><span data-ttu-id="ddfd6-165">参数数组</span><span class="sxs-lookup"><span data-stu-id="ddfd6-165">Parameters array</span></span>

<span data-ttu-id="ddfd6-166">属性是一个对象数组。`parameters`</span><span class="sxs-lookup"><span data-stu-id="ddfd6-166">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="ddfd6-167">其中每个对象代表一个参数。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-167">Each of these objects represents a parameter.</span></span> <span data-ttu-id="ddfd6-168">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="ddfd6-168">The following table contains its properties:</span></span>

|  <span data-ttu-id="ddfd6-169">属性</span><span class="sxs-lookup"><span data-stu-id="ddfd6-169">Property</span></span>  |  <span data-ttu-id="ddfd6-170">数据类型</span><span class="sxs-lookup"><span data-stu-id="ddfd6-170">Data Type</span></span>  |  <span data-ttu-id="ddfd6-171">是否必需？</span><span class="sxs-lookup"><span data-stu-id="ddfd6-171">Required?</span></span>  |  <span data-ttu-id="ddfd6-172">说明</span><span class="sxs-lookup"><span data-stu-id="ddfd6-172">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="ddfd6-173">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-173">string</span></span>  |  <span data-ttu-id="ddfd6-174">否</span><span class="sxs-lookup"><span data-stu-id="ddfd6-174">No</span></span> |  <span data-ttu-id="ddfd6-175">参数的描述。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-175">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="ddfd6-176">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-176">string</span></span>  |  <span data-ttu-id="ddfd6-177">是</span><span class="sxs-lookup"><span data-stu-id="ddfd6-177">Yes</span></span>  |  <span data-ttu-id="ddfd6-178">必须是“标量”（即非数组值）或“矩阵”（即一系列行数组）。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-178">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="ddfd6-179">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-179">string</span></span>  |  <span data-ttu-id="ddfd6-180">是</span><span class="sxs-lookup"><span data-stu-id="ddfd6-180">Yes</span></span>  |  <span data-ttu-id="ddfd6-181">参数的名称。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-181">The name of the parameter.</span></span> <span data-ttu-id="ddfd6-182">此名称显示在 Excel 的 IntelliSense 中。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-182">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="ddfd6-183">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-183">string</span></span>  |  <span data-ttu-id="ddfd6-184">是</span><span class="sxs-lookup"><span data-stu-id="ddfd6-184">Yes</span></span>  |  <span data-ttu-id="ddfd6-185">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-185">The data type of the parameter.</span></span> <span data-ttu-id="ddfd6-186">必须为“布尔值”、“数字”或“字符串”。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-186">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="ddfd6-187">结果对象</span><span class="sxs-lookup"><span data-stu-id="ddfd6-187">Result object</span></span>

<span data-ttu-id="ddfd6-188">属性提供有关函数返回的值的元数据。`results`</span><span class="sxs-lookup"><span data-stu-id="ddfd6-188">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="ddfd6-189">下表包含其属性：</span><span class="sxs-lookup"><span data-stu-id="ddfd6-189">The following table contains its properties:</span></span>

|  <span data-ttu-id="ddfd6-190">属性</span><span class="sxs-lookup"><span data-stu-id="ddfd6-190">Property</span></span>  |  <span data-ttu-id="ddfd6-191">数据类型</span><span class="sxs-lookup"><span data-stu-id="ddfd6-191">Data Type</span></span>  |  <span data-ttu-id="ddfd6-192">是否必需？</span><span class="sxs-lookup"><span data-stu-id="ddfd6-192">Required?</span></span>  |  <span data-ttu-id="ddfd6-193">说明</span><span class="sxs-lookup"><span data-stu-id="ddfd6-193">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="ddfd6-194">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-194">string</span></span>  |  <span data-ttu-id="ddfd6-195">否</span><span class="sxs-lookup"><span data-stu-id="ddfd6-195">No</span></span>  |  <span data-ttu-id="ddfd6-196">必须是“标量”（即非数组值）或“矩阵”（即一系列行数组）。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-196">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="ddfd6-197">字符串</span><span class="sxs-lookup"><span data-stu-id="ddfd6-197">string</span></span>  |  <span data-ttu-id="ddfd6-198">是</span><span class="sxs-lookup"><span data-stu-id="ddfd6-198">Yes</span></span>  |  <span data-ttu-id="ddfd6-199">参数的数据类型。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-199">The data type of the parameter.</span></span> <span data-ttu-id="ddfd6-200">必须为“布尔值”、“数字”或“字符串”。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-200">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="ddfd6-201">示例</span><span class="sxs-lookup"><span data-stu-id="ddfd6-201">Example</span></span>

<span data-ttu-id="ddfd6-202">以下 JSON 代码是自定义函数元数据的示例。</span><span class="sxs-lookup"><span data-stu-id="ddfd6-202">The following JSON code is an example of a metadata file for custom functions.</span></span>

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
            ],
            "options": {
                "sync": true
            }
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
            ],
            "options": {
                "sync": false
            }
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
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": [],
            "options": {
                "sync": true
            }
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
                "sync": false,
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
            ],
            "options": {
                "sync": true
            }
        }
    ]
}

```

## <a name="see-also"></a><span data-ttu-id="ddfd6-203">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ddfd6-203">See also</span></span>
[<span data-ttu-id="ddfd6-204">自定义函数</span><span class="sxs-lookup"><span data-stu-id="ddfd6-204">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="ddfd6-205">有关数组公式的指导和示例</span><span class="sxs-lookup"><span data-stu-id="ddfd6-205">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
