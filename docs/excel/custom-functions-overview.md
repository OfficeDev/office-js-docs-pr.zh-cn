# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="da977-101">在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="da977-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="da977-102">借助自定义函数（类似于用户定义的函数 [UDF]），开发人员可以使用加载项向 Excel 添加任何 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="da977-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="da977-103">然后，用户可以像使用 Excel 中的其他本地函数（例如 `=SUM()`）一样访问自定义函数。</span><span class="sxs-lookup"><span data-stu-id="da977-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="da977-104">本文介绍了如何在 Excel 中创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="da977-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="da977-105">下图显示了最终用户如何将自定义函数插入到单元格中。</span><span class="sxs-lookup"><span data-stu-id="da977-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="da977-106">将 42 添加到一对数字的函数。</span><span class="sxs-lookup"><span data-stu-id="da977-106">Here’s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="da977-107">以下是相同自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="da977-107">Here’s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="da977-108">自定义函数现可在 Windows、Mac 和 Excel Online 的开发人员预览版中使用。</span><span class="sxs-lookup"><span data-stu-id="da977-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="da977-109">若要试用，请按照以下步骤操作：</span><span class="sxs-lookup"><span data-stu-id="da977-109">Follow these steps to try them:</span></span>

1. <span data-ttu-id="da977-110">安装 Office（Windows 的内部版本 9325 或 Mac 上的内部版本 13.329）并加入 [Office 预览体验成员](https://products.office.com/office-insider)计划。</span><span class="sxs-lookup"><span data-stu-id="da977-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="da977-111">（请注意，仅仅获取最新版本是不够的；在加入预览体验成员计划之前，任何版本的功能都将禁用）</span><span class="sxs-lookup"><span data-stu-id="da977-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2. <span data-ttu-id="da977-112">使用 [Yo Office](https://github.com/OfficeDev/generator-office) 创建 Excel 自定义函数的加载项项目，并按照 [project README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) 中的说明在 Excel 中启动加载项，在代码中进行更改并调试。</span><span class="sxs-lookup"><span data-stu-id="da977-112">Create an Excel Custom Functions Add-in project using [Yo Office](https://github.com/OfficeDev/generator-office), and follow the instructions in the [project README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to start the add-in in Excel, make changes in the code, and debug.</span></span>
3. <span data-ttu-id="da977-113">在任意单元格中键入“`=CONTOSO.ADD42(1,2)`”，再按 **Enter** 运行自定义函数。</span><span class="sxs-lookup"><span data-stu-id="da977-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="da977-114">请参阅本文末尾的**已知问题**部分，其中包括自定义函数的当前限制，该部分将随时间进行更新。</span><span class="sxs-lookup"><span data-stu-id="da977-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="da977-115">学习基础知识</span><span class="sxs-lookup"><span data-stu-id="da977-115">Learn the basics</span></span>

<span data-ttu-id="da977-116">在克隆的示例存储库中，将看到以下文件：</span><span class="sxs-lookup"><span data-stu-id="da977-116">In the cloned sample repo, you’ll see the following files:</span></span>

- <span data-ttu-id="da977-117">**./src/customfunctions.js**，其中包含自定义函数代码（请参阅上面 `ADD42` 函数的简单代码示例）。</span><span class="sxs-lookup"><span data-stu-id="da977-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="da977-p105">**./config/customfunctions.json**，其中包含将自定义函数告诉 Excel 的注册 JSON。注册使您的自定义函数在用户键入于单元格时显示在可用的函数列表中。</span><span class="sxs-lookup"><span data-stu-id="da977-p105">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function. Registration makes your custom functions appear in the list of available functions displayed when a user types in a cell.</span></span>
- <span data-ttu-id="da977-p106">**./index.html**，它提供 JS 文件的&lt;脚本&gt;引用。此文件不在 Excel 中显示 UI。</span><span class="sxs-lookup"><span data-stu-id="da977-p106">**customfunctions.html**, which provides a &lt;Script&gt; reference to the JS file. This file does not display UI in Excel.</span></span>
- <span data-ttu-id="da977-122">**./manifest.xml**，它将 HTML、JavaScript 和 JSON 文件的的位置告诉 Excel；还为与该加载项一起安装的所有自定义函数指定一个名称空间。</span><span class="sxs-lookup"><span data-stu-id="da977-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="da977-123">JSON 文件 (./config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="da977-123">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="da977-124">customfunctions.json中的以下代码相同的 `ADD42` 功能指定元数据。</span><span class="sxs-lookup"><span data-stu-id="da977-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="da977-125">JSON文件的详细参考信息（包括本示例中未使用的选项）位于 [自定义函数注册JSON](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="da977-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](custom-functions-json.md).</span></span>

<span data-ttu-id="da977-126">请注意，对于这个例子：</span><span class="sxs-lookup"><span data-stu-id="da977-126">Note that for this example:</span></span>

- <span data-ttu-id="da977-127">只有一个自定义函数，所以只有 `functions` 阵列的一个成员。</span><span class="sxs-lookup"><span data-stu-id="da977-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="da977-128">该 `name` 属性定义了函数名称。</span><span class="sxs-lookup"><span data-stu-id="da977-128">The `name` property defines the function name.</span></span> <span data-ttu-id="da977-129">正如您在前面的动画gif中看到的，名称空间（`CONTOSO`）预先添加到Excel自动完成菜单中的函数名称。</span><span class="sxs-lookup"><span data-stu-id="da977-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="da977-130">此前缀在加载项清单中定义，如下所述。</span><span class="sxs-lookup"><span data-stu-id="da977-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="da977-131">前缀和函数名使用句点分隔，按照惯例，前缀和函数名都是大写。</span><span class="sxs-lookup"><span data-stu-id="da977-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="da977-132">要使用自定义函数，用户键入名称空间，后跟该函数的名称（`ADD42`）进入一个单元格，在这种情况下 `=CONTOSO.ADD42`。</span><span class="sxs-lookup"><span data-stu-id="da977-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="da977-133">前缀将用作公司或加载项的标识符。</span><span class="sxs-lookup"><span data-stu-id="da977-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="da977-134">`description` 将在 Excel 的自动完成菜单中显示。</span><span class="sxs-lookup"><span data-stu-id="da977-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="da977-135">当用户针对某个函数请求帮助时，Excel 将打开任务窗格并显示位于 `helpUrl` 所指定 URL 的网页。</span><span class="sxs-lookup"><span data-stu-id="da977-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="da977-136">该 `result` 属性指定函数返回给 Excel 之信息的类型。</span><span class="sxs-lookup"><span data-stu-id="da977-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="da977-137">该 `type` 子属性可以 `"string"`， `"number"`， 或 `"boolean"`。</span><span class="sxs-lookup"><span data-stu-id="da977-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="da977-138">该 `dimensionality` 属性可以 `scalar` 或 `matrix` （指定 `type`值的二维数组。）</span><span class="sxs-lookup"><span data-stu-id="da977-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="da977-139">该 `parameters` 数组 *按顺序*指定了传递给函数的每个参数中的数据类型。</span><span class="sxs-lookup"><span data-stu-id="da977-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="da977-140">该 `name` 和 `description` 在Excel智能感知中使用子属性。</span><span class="sxs-lookup"><span data-stu-id="da977-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="da977-141">该 `type` 和 `dimensionality` 子属性与上述 `result` 属性之子属性相同。</span><span class="sxs-lookup"><span data-stu-id="da977-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="da977-142">该 `options` 属性使您可以自定义Excel执行功能之方式和时间的某些方面。</span><span class="sxs-lookup"><span data-stu-id="da977-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="da977-143">本文后面有关于这些选项的更多信息。</span><span class="sxs-lookup"><span data-stu-id="da977-143">There is more information about these options later in this article.</span></span>

 ```js
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "name": "ADD42", 
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
            "options": {
                "sync": true
            }
        }
    ]
}
```

> [!NOTE]
> <span data-ttu-id="da977-144">自定义函数是在用户第一次运行加载项时注册的。</span><span class="sxs-lookup"><span data-stu-id="da977-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="da977-145">之后，对于同一用户，在所有工作簿中都可以使用它们（不仅是最初加载项运行的那个。）</span><span class="sxs-lookup"><span data-stu-id="da977-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="da977-146">您的JSON文件的服务器设置必须具有 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) 启用以使自定义函数在Excel Online中正常工作。</span><span class="sxs-lookup"><span data-stu-id="da977-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-manifestxml"></a><span data-ttu-id="da977-147">清单文件 (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="da977-147">Manifest file (manifest.xml)</span></span>


<span data-ttu-id="da977-148">以下是您在加载项的清单中包含的 `<ExtensionPoint>` 和 `<Resources>` 标记的例子，它们使 Excel 能够运行您的函数。</span><span class="sxs-lookup"><span data-stu-id="da977-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="da977-149">请注意有关此标记的以下事实：</span><span class="sxs-lookup"><span data-stu-id="da977-149">Note the following facts about this markup:</span></span>

- <span data-ttu-id="da977-150">该 `<Script>` 元素及其相应的资源ID指定JavaScript文件在您的函数中的位置。</span><span class="sxs-lookup"><span data-stu-id="da977-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="da977-151">该 `<Page>` 元素及其相应的资源ID指定加载项之HTML页面的位置。</span><span class="sxs-lookup"><span data-stu-id="da977-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="da977-152">HTML页面包含一个 `<Script>` 加载JavaScript文件的标签（customfunctions.js）。</span><span class="sxs-lookup"><span data-stu-id="da977-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="da977-153">HTML 页面是一个隐藏页面，始终不会在 UI 中显示。</span><span class="sxs-lookup"><span data-stu-id="da977-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="da977-154">该 `<Metadata>` 元素及其相应的资源ID指定JSON文件的位置。</span><span class="sxs-lookup"><span data-stu-id="da977-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="da977-155">一个 `<Namespace>` 元素及其相应的资源ID指定加载项中所有自定义函数的前缀。</span><span class="sxs-lookup"><span data-stu-id="da977-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="residjs" />
                    </Script>
                    <Page>
                        <SourceLocation resid="residhtml"/>
                    </Page>
                    <Metadata>
                        <SourceLocation resid="residjson" />
                    </Metadata>
                    <Namespace resid="residNS" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
            <bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
            <bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="residNS" DefaultValue="CONTOSO" />
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>

```

## <a name="initializing-custom-functions"></a><span data-ttu-id="da977-156">初始化自定义函数</span><span class="sxs-lookup"><span data-stu-id="da977-156">Initializing custom functions</span></span>

<span data-ttu-id="da977-157">您的代码在使用之前必须初始化自定义函数功能。</span><span class="sxs-lookup"><span data-stu-id="da977-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="da977-158">你可以在一个 &lt;脚本&gt; 在HTML文件（customfunctions.html）中的标记或JavaScript文件（customfunctions.js）的顶部。</span><span class="sxs-lookup"><span data-stu-id="da977-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="da977-159">在预览自定义函数期间，您可以选择两种初始化语法。</span><span class="sxs-lookup"><span data-stu-id="da977-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="da977-160">回购库中的HTML文件使用以下语法：</span><span class="sxs-lookup"><span data-stu-id="da977-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="da977-161">您还可以使用以下语法：</span><span class="sxs-lookup"><span data-stu-id="da977-161">You can also use the following syntax:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="handling-errors"></a><span data-ttu-id="da977-162">处理错误</span><span class="sxs-lookup"><span data-stu-id="da977-162">Handling errors</span></span>
<span data-ttu-id="da977-163">自定义的函数的错误处理与 [Excel JavaScript API 的错误处理整体类同](./excel-add-ins-error-handling.md)。</span><span class="sxs-lookup"><span data-stu-id="da977-163">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](./excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="da977-164">一般情况下，将使用 `.catch` 来处理错误。</span><span class="sxs-lookup"><span data-stu-id="da977-164">Generally, you will use `.catch` to handle errors.</span></span> <span data-ttu-id="da977-165">下面的代码是 `.catch` 的例子。</span><span class="sxs-lookup"><span data-stu-id="da977-165">The code below gives an example of `.catch`.</span></span> 

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="da977-166">同步和异步函数</span><span class="sxs-lookup"><span data-stu-id="da977-166">Synchronous and asynchronous functions</span></span>

<span data-ttu-id="da977-167">上面的 `ADD42` 功能是关于Excel同步的（通过设置在JSON文件中的 `"sync": true` 选项来指定）。</span><span class="sxs-lookup"><span data-stu-id="da977-167">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="da977-168">同步函数提供了快速的性能，因为它们与Excel运行的过程相同，并且在多线程计算过程中它们并行运行。</span><span class="sxs-lookup"><span data-stu-id="da977-168">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="da977-169">另一方面，如果您的自定义函数从Web中检索数据，则它必须相对于Excel异步。</span><span class="sxs-lookup"><span data-stu-id="da977-169">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="da977-170">异步函数必须：</span><span class="sxs-lookup"><span data-stu-id="da977-170">Asynchronous functions must:</span></span>

1. <span data-ttu-id="da977-171">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="da977-171">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="da977-172">使用回调函数，用最终值解析Promise。</span><span class="sxs-lookup"><span data-stu-id="da977-172">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="da977-173">下面的代码显示用于检索温度计温度的自定义异步函数示例。</span><span class="sxs-lookup"><span data-stu-id="da977-173">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="da977-174">注意 `sendWebRequest` 是一个假设的功能，这里没有指定，它使用XHR来调用温度网络服务。</span><span class="sxs-lookup"><span data-stu-id="da977-174">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="da977-175">异步函数显示在Excel等待最终结果时，单元格中出现的一个 `GETTING_DATA` 暂时错误。</span><span class="sxs-lookup"><span data-stu-id="da977-175">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="da977-176">用户可以在等待结果时与电子表格的其余部分进行正常交互。</span><span class="sxs-lookup"><span data-stu-id="da977-176">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="da977-177">自定义函数默认是异步的。</span><span class="sxs-lookup"><span data-stu-id="da977-177">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="da977-178">要将功能指定为同步，设置 `"sync": true` 选项在注册JSON文件中自定义函数的 `options` 属性中。</span><span class="sxs-lookup"><span data-stu-id="da977-178">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="da977-179">流式处理函数</span><span class="sxs-lookup"><span data-stu-id="da977-179">Streamed functions</span></span>

<span data-ttu-id="da977-180">异步功能可以流式处理。</span><span class="sxs-lookup"><span data-stu-id="da977-180">An asynchronous function can be streamed.</span></span> <span data-ttu-id="da977-181">借助流式处理自定义函数，可以随时间推移将数据重复输出到单元格，而无需等待 Excel 或用户请求重新计算。</span><span class="sxs-lookup"><span data-stu-id="da977-181">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="da977-182">以下示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="da977-182">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="da977-183">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="da977-183">Note the following about this code:</span></span>

- <span data-ttu-id="da977-184">Excel会自动使用 `setResult` 回调来显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="da977-184">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="da977-185">始终不会在注册代码中指定最后的 `caller` 参数，且当用户输入此函数时，该参数不会在 Excel 用户的自动完成菜单中显示。</span><span class="sxs-lookup"><span data-stu-id="da977-185">For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="da977-186">它是包含`setResult` 回调函数的对象，用于将数据从函数传递到 Excel，以更新单元格值。</span><span class="sxs-lookup"><span data-stu-id="da977-186">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="da977-187">为了让Excel通过 `setResult` 功能在 `caller` 对象，您必须通过设置 `"stream": true` 选项在注册JSON文件中自定义函数的 `options` 属性里，来声明在函数注册期间支持流式处理。</span><span class="sxs-lookup"><span data-stu-id="da977-187">In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="da977-188">取消</span><span class="sxs-lookup"><span data-stu-id="da977-188">Cancellation</span></span>

<span data-ttu-id="da977-189">可以取消流式处理函数和异步函数。</span><span class="sxs-lookup"><span data-stu-id="da977-189">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="da977-190">对于减少带宽消耗、工作内存和 CPU 负载，取消函数调用非常重要。</span><span class="sxs-lookup"><span data-stu-id="da977-190">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="da977-191">Excel 在以下情况下取消函数调用：</span><span class="sxs-lookup"><span data-stu-id="da977-191">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="da977-192">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="da977-192">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="da977-193">函数的参数（输入）之一发生变化。</span><span class="sxs-lookup"><span data-stu-id="da977-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="da977-194">在这种情况下，除了取消之外，还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="da977-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="da977-p125">用户手动触发重新计算。与上述情况一样，除了取消之外，还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="da977-p125">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="da977-197">您 *必须* 为每个流式传输功能实施取消处理程序。</span><span class="sxs-lookup"><span data-stu-id="da977-197">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="da977-198">异步、非流式功能可取消，也可不取消；由你决定。</span><span class="sxs-lookup"><span data-stu-id="da977-198">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="da977-199">同步功能无法取消。</span><span class="sxs-lookup"><span data-stu-id="da977-199">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="da977-200">要使功能可取消，请设置选项 `"cancelable": true` 在注册JSON文件中自定义函数的 `options` 属性里面。</span><span class="sxs-lookup"><span data-stu-id="da977-200">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="da977-201">下面的代码展示了已实现取消的上一个示例。</span><span class="sxs-lookup"><span data-stu-id="da977-201">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="da977-202">在此代码中，`caller` 对象包含一个为每个可取消的自定义函数定义的 `onCanceled` 函数。</span><span class="sxs-lookup"><span data-stu-id="da977-202">In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="da977-203">保存和共享状态</span><span class="sxs-lookup"><span data-stu-id="da977-203">Saving and sharing state</span></span>

<span data-ttu-id="da977-204">异步自定义函数可以将数据保存在全局 JavaScript 变量中。</span><span class="sxs-lookup"><span data-stu-id="da977-204">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="da977-205">在后续调用中，自定义函数可以使用保存在这些变量中的值。</span><span class="sxs-lookup"><span data-stu-id="da977-205">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="da977-206">当用户将相同的自定义函数添加到多个单元格时，保存状态很有用，因为该函数的所有实例都可以共享该状态。</span><span class="sxs-lookup"><span data-stu-id="da977-206">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="da977-207">例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。</span><span class="sxs-lookup"><span data-stu-id="da977-207">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="da977-208">下面的代码演示之前的温度流式处理函数的实施过程，该函数将状态保存在全局作用域。</span><span class="sxs-lookup"><span data-stu-id="da977-208">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="da977-209">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="da977-209">Note the following about this code:</span></span>

- <span data-ttu-id="da977-210">`refreshTemperature` 是一个流式处理函数，它会在每一秒内读取特定温度计的温度。</span><span class="sxs-lookup"><span data-stu-id="da977-210">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="da977-211">新的温度保存在 `savedTemperatures` 变量，但不直接更新单元格值。</span><span class="sxs-lookup"><span data-stu-id="da977-211">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="da977-212">它不应该直接从工作表单元格中调用， *所以它没有在JSON文件中注册*。</span><span class="sxs-lookup"><span data-stu-id="da977-212">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="da977-213">`streamTemperature` 每秒钟更新单元格中显示的温度值并使用 `savedTemperatures` 变量作为其数据源。</span><span class="sxs-lookup"><span data-stu-id="da977-213">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="da977-214">它必须在JSON文件中注册，并用所有大写字母命名， `STREAMTEMPERATURE`。</span><span class="sxs-lookup"><span data-stu-id="da977-214">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="da977-215">用户可以从 Excel UI 的多个单元格中调用 `streamTemperature`。</span><span class="sxs-lookup"><span data-stu-id="da977-215">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="da977-216">每次调用都从相同的 `savedTemperatures` 变量读取数据。</span><span class="sxs-lookup"><span data-stu-id="da977-216">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
         setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
     }
     getNextTemperature();
}

function refreshTemperature(thermometerID){
     sendWebRequest(thermometerID, function(data){
         savedTemperatures[thermometerID] = data.temperature;
     });
     setTimeout(function(){
         refreshTemperature(thermometerID);
     }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

> [!NOTE]
> <span data-ttu-id="da977-217">同步功能（通过在JSON文件中设置 `"sync": true` 选项指定）不能共享状态，因为Excel在多线程计算过程中将它们并行化。</span><span class="sxs-lookup"><span data-stu-id="da977-217">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="da977-218">只有异步函数可以共享状态，因为加载项的同步函数在每个进程中共享相同的JavaScript上下文。</span><span class="sxs-lookup"><span data-stu-id="da977-218">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="da977-219">使用数据区域</span><span class="sxs-lookup"><span data-stu-id="da977-219">Working with ranges of data</span></span>

<span data-ttu-id="da977-220">自定义函数可以将数据区域用作参数，或者可以从自定义函数返回数据区域。</span><span class="sxs-lookup"><span data-stu-id="da977-220">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="da977-221">例如，假设您的函数返回 Excel 中存储的一系列数字的第二个最大值。</span><span class="sxs-lookup"><span data-stu-id="da977-221">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="da977-222">下面的函数需要使用参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 参数类型。</span><span class="sxs-lookup"><span data-stu-id="da977-222">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="da977-223">请注意，在此函数的注册JSON中，您可以设置参数的 `type` 属性给 `matrix`。</span><span class="sxs-lookup"><span data-stu-id="da977-223">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){ 
     var highest = values[0][0], secondHighest = values[0][0];
     for(var i = 0; i < values.length; i++){
         for(var j = 1; j < values[i].length; j++){
             if(values[i][j] >= highest){
                 secondHighest = highest;
                 highest = values[i][j];
             }
             else if(values[i][j] >= secondHighest){
                 secondHighest = values[i][j];
             }
         }
     }
     return secondHighest;
 }
```

<span data-ttu-id="da977-224">如您所见，范围在JavaScript中以行矩阵的矩阵（如2维矩阵）处理。</span><span class="sxs-lookup"><span data-stu-id="da977-224">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="da977-225">已知问题</span><span class="sxs-lookup"><span data-stu-id="da977-225">Known issues</span></span>

- <span data-ttu-id="da977-226">Excel 暂未使用帮助 URL 和参数说明。</span><span class="sxs-lookup"><span data-stu-id="da977-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="da977-227">自定义功能目前不适用于移动客户的Excel。</span><span class="sxs-lookup"><span data-stu-id="da977-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="da977-228">目前，加载项依赖隐藏的浏览器进程来运行异步自定义函数。</span><span class="sxs-lookup"><span data-stu-id="da977-228">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="da977-229">将来，JavaScript 将直接在某些平台上运行，以确保自定义函数运行速度更快并占用更少的内存。</span><span class="sxs-lookup"><span data-stu-id="da977-229">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="da977-230">此外，大多数平台将不再需要清单的 `<Page>` 元素所引用的 HTML 页面，因为 Excel 将直接运行 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="da977-230">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="da977-231">若要为这一更改做准备，请确保自定义函数未使用网页 DOM。</span><span class="sxs-lookup"><span data-stu-id="da977-231">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="da977-232">使用GET或POST，用于访问网络所支持的主机APIs将会是 [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) 和 [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) 。</span><span class="sxs-lookup"><span data-stu-id="da977-232">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="da977-233">易失性函数（当电子表格中不相关数据发生变化时自动重新计算的函数）尚不受支持。</span><span class="sxs-lookup"><span data-stu-id="da977-233">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="da977-234">调试仅适用于Excel for Windows上的异步功能。</span><span class="sxs-lookup"><span data-stu-id="da977-234">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="da977-235">尚未启用通过Office 365管理门户和AppSource进行的部署。</span><span class="sxs-lookup"><span data-stu-id="da977-235">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="da977-236">Excel Online中的自定义功能，可能会在一段时间无活动后，在进程期间停止工作。</span><span class="sxs-lookup"><span data-stu-id="da977-236">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="da977-237">刷新浏览器页面（F5）并重新输入自定义函数以恢复该功能。</span><span class="sxs-lookup"><span data-stu-id="da977-237">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="da977-238">更改日志</span><span class="sxs-lookup"><span data-stu-id="da977-238">Changelog</span></span>

- <span data-ttu-id="da977-239">**2017 年 11 月 7 日**：发布了自定义函数（预览）和示例</span><span class="sxs-lookup"><span data-stu-id="da977-239">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="da977-240">**2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题</span><span class="sxs-lookup"><span data-stu-id="da977-240">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="da977-241">**2017 年 11 月 28 日**：发布了对取消异步函数的支持（需要对流式处理函数进行相应更改）</span><span class="sxs-lookup"><span data-stu-id="da977-241">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="da977-242">**2018年5月7日**：提供对Mac、Excel Online和进程运行中的同步功能的支持</span><span class="sxs-lookup"><span data-stu-id="da977-242">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
