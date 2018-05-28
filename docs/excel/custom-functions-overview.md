# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="88962-101">? Excel ????????????</span><span class="sxs-lookup"><span data-stu-id="88962-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="88962-102">?????????????????? [UDF]??????????????? Excel ???? JavaScript ???</span><span class="sxs-lookup"><span data-stu-id="88962-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="88962-103">?????????? Excel ???????????????????? `=SUM()`??</span><span class="sxs-lookup"><span data-stu-id="88962-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="88962-104">???????? Excel ?????????</span><span class="sxs-lookup"><span data-stu-id="88962-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="88962-105">?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="88962-106">? 42 ???????????</span><span class="sxs-lookup"><span data-stu-id="88962-106">Here?s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="88962-107">??????????????</span><span class="sxs-lookup"><span data-stu-id="88962-107">Here?s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="88962-108">???????? Windows?Mac ? Excel Online ????????????</span><span class="sxs-lookup"><span data-stu-id="88962-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="88962-109">???????????????</span><span class="sxs-lookup"><span data-stu-id="88962-109">Follow these steps to try them:</span></span>

1.  <span data-ttu-id="88962-110">?? Office?Windows ????? 9325 ? Mac ?????? 13.329???? [Office ??????](https://products.office.com/en-us/office-insider)???</span><span class="sxs-lookup"><span data-stu-id="88962-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/en-us/office-insider) program.</span></span> <span data-ttu-id="88962-111">????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2.  <span data-ttu-id="88962-112">?? [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) ??????? README.md ????? Excel ??????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-112">Clone the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repo and follow the instructions in the README.md to start the add-in in Excel, make changes in the code, and debug.</span></span>
3.  <span data-ttu-id="88962-113">????????? `=CONTOSO.ADD42(1,2)`??? **Enter** ????????</span><span class="sxs-lookup"><span data-stu-id="88962-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="88962-114">????????**????**??????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="88962-115">??????</span><span class="sxs-lookup"><span data-stu-id="88962-115">Learn the basics</span></span>

<span data-ttu-id="88962-116">???????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-116">In the cloned sample repo, you?ll see the following files:</span></span>

- <span data-ttu-id="88962-117">**customfunctions.js**????????????????????????? `ADD42` ????</span><span class="sxs-lookup"><span data-stu-id="88962-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="88962-118">**customfunctions.json**???????JSON???Excel????????</span><span class="sxs-lookup"><span data-stu-id="88962-118">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function.</span></span> <span data-ttu-id="88962-119">????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-119">Registration makes your custom functions appear in the list of available functions displayed when users type in cells.</span></span>
- <span data-ttu-id="88962-120">**customfunctions.html**??????? &lt;??&gt; ??JS???</span><span class="sxs-lookup"><span data-stu-id="88962-120">customfunctions.html, which provides a Script reference to customfunctions.js.</span></span> <span data-ttu-id="88962-121">????? Excel ??? UI?</span><span class="sxs-lookup"><span data-stu-id="88962-121">This file does not display UI in Excel.</span></span>
- <span data-ttu-id="88962-122">**customfunctions.xml**????Excel HTML?JavaScript?JSON?????;????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-customfunctionsjson"></a><span data-ttu-id="88962-123">JSON???customfunctions.json?</span><span class="sxs-lookup"><span data-stu-id="88962-123">JSON file (customfunctions.json)</span></span>

<span data-ttu-id="88962-124">customfunctions.json????????? `ADD42` ????????</span><span class="sxs-lookup"><span data-stu-id="88962-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="88962-125">JSON????????????????????????? [???????JSON](https://dev.office.com/reference/add-ins/custom-functions-json)?</span><span class="sxs-lookup"><span data-stu-id="88962-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](https://dev.office.com/reference/add-ins/custom-functions-json).</span></span>

<span data-ttu-id="88962-126">???????????</span><span class="sxs-lookup"><span data-stu-id="88962-126">Note that for this example:</span></span>

- <span data-ttu-id="88962-127">?????????????? `functions` ????????</span><span class="sxs-lookup"><span data-stu-id="88962-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="88962-128">? `name` ??????????</span><span class="sxs-lookup"><span data-stu-id="88962-128">The `name` property defines the function name.</span></span> <span data-ttu-id="88962-129">?????????gif??????????`CONTOSO`??????Excel?????????????</span><span class="sxs-lookup"><span data-stu-id="88962-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="88962-130">??????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="88962-131">?????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="88962-132">???????????????????????????`ADD42`??????????????? `=CONTOSO.ADD42`?</span><span class="sxs-lookup"><span data-stu-id="88962-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="88962-133">?????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="88962-134">`description` ?? Excel ???????????</span><span class="sxs-lookup"><span data-stu-id="88962-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="88962-135">???????????????Excel ?????????????`helpUrl`?? URL ????</span><span class="sxs-lookup"><span data-stu-id="88962-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="88962-136">? `result` ?????????Excel???????</span><span class="sxs-lookup"><span data-stu-id="88962-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="88962-137">? `type` ????? `"string"`? `"number"`? ? `"boolean"`?</span><span class="sxs-lookup"><span data-stu-id="88962-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="88962-138">? `dimensionality` ???? `scalar` ? `matrix` ??? `type`????????</span><span class="sxs-lookup"><span data-stu-id="88962-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="88962-139">? `parameters` ?? *???*????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="88962-140">? `name` ? `description` ?Excel???????????</span><span class="sxs-lookup"><span data-stu-id="88962-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="88962-141">? `type` ? `dimensionality` ?????? `result` ?????????</span><span class="sxs-lookup"><span data-stu-id="88962-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="88962-142">? `options` ?????????Excel????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="88962-143">?????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-143">There is more information about these options later in this article.</span></span>

 ```js
{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
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
> <span data-ttu-id="88962-144">??????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="88962-145">????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="88962-146">??JSON???????????? [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) ??????????Excel Online??????</span><span class="sxs-lookup"><span data-stu-id="88962-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-customfunctionsxml"></a><span data-ttu-id="88962-147">?????customfunctions.xml?</span><span class="sxs-lookup"><span data-stu-id="88962-147">Manifest file (customfunctions.xml)</span></span>


<span data-ttu-id="88962-148">??????? `<ExtensionPoint>` ? `<Resources>` ???????????????Excel?????????</span><span class="sxs-lookup"><span data-stu-id="88962-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="88962-149">??????????????</span><span class="sxs-lookup"><span data-stu-id="88962-149">Note the following facts about this markup:</span></span>

- <span data-ttu-id="88962-150">? `<Script>` ?????????ID??JavaScript????????????</span><span class="sxs-lookup"><span data-stu-id="88962-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="88962-151">? `<Page>` ?????????ID??????HTML??????</span><span class="sxs-lookup"><span data-stu-id="88962-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="88962-152">HTML?????? `<Script>` ??JavaScript??????customfunctions.js??</span><span class="sxs-lookup"><span data-stu-id="88962-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="88962-153">HTML ??????????????? UI ????</span><span class="sxs-lookup"><span data-stu-id="88962-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="88962-154">? `<Metadata>` ?????????ID??JSON??????</span><span class="sxs-lookup"><span data-stu-id="88962-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="88962-155">?? `<Namespace>` ?????????ID?????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


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

## <a name="initializing-custom-functions"></a><span data-ttu-id="88962-156">????????</span><span class="sxs-lookup"><span data-stu-id="88962-156">Initializing custom functions</span></span>

<span data-ttu-id="88962-157">??????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="88962-158">?????? &lt;??&gt; ?HTML???customfunctions.html??????JavaScript???customfunctions.js?????</span><span class="sxs-lookup"><span data-stu-id="88962-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="88962-159">????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="88962-160">?????HTML?????????</span><span class="sxs-lookup"><span data-stu-id="88962-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="88962-161">???????????</span><span class="sxs-lookup"><span data-stu-id="88962-161">You can also use the following syntax:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="88962-162">???????</span><span class="sxs-lookup"><span data-stu-id="88962-162">Synchronous and asynchronous functions</span></span>

<span data-ttu-id="88962-163">??? `ADD42` ?????Excel?????????JSON???? `"sync": true` ???????</span><span class="sxs-lookup"><span data-stu-id="88962-163">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="88962-164">??????????????????Excel??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-164">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="88962-165">???????????????Web?????????????Excel???</span><span class="sxs-lookup"><span data-stu-id="88962-165">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="88962-166">???????</span><span class="sxs-lookup"><span data-stu-id="88962-166">Asynchronous functions must:</span></span>

1. <span data-ttu-id="88962-167">? JavaScript ????? Excel?</span><span class="sxs-lookup"><span data-stu-id="88962-167">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="88962-168">?????????????Promise?</span><span class="sxs-lookup"><span data-stu-id="88962-168">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="88962-169">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-169">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="88962-170">?? `sendWebRequest` ???????????????????XHR??????????</span><span class="sxs-lookup"><span data-stu-id="88962-170">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="88962-171">???????Excel????????????????? `GETTING_DATA` ?????</span><span class="sxs-lookup"><span data-stu-id="88962-171">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="88962-172">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-172">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="88962-173">????????????</span><span class="sxs-lookup"><span data-stu-id="88962-173">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="88962-174">???????????? `"sync": true` ?????JSON????????? `options` ????</span><span class="sxs-lookup"><span data-stu-id="88962-174">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="88962-175">??????</span><span class="sxs-lookup"><span data-stu-id="88962-175">Streamed functions</span></span>

<span data-ttu-id="88962-176">???????????</span><span class="sxs-lookup"><span data-stu-id="88962-176">An asynchronous function can be streamed.</span></span> <span data-ttu-id="88962-177">??????????????????????????????????? Excel ??????????</span><span class="sxs-lookup"><span data-stu-id="88962-177">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="88962-178">??????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="88962-179">??????????????
</span><span class="sxs-lookup"><span data-stu-id="88962-179">Note the following about this code:</span></span>

- <span data-ttu-id="88962-180">Excel????? `setResult` ??????????</span><span class="sxs-lookup"><span data-stu-id="88962-180">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="88962-181">???????????????? `caller`?????????????????? Excel ?????????????</span><span class="sxs-lookup"><span data-stu-id="88962-181">For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="88962-182">????`setResult` ??????????????????? Excel?????????</span><span class="sxs-lookup"><span data-stu-id="88962-182">It?s an object that contains a `setResult` callback function that?s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="88962-183">???Excel?? `setResult` ??? `caller` ?????????? `"stream": true` ?????JSON????????? `options` ?????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-183">In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="88962-184">??</span><span class="sxs-lookup"><span data-stu-id="88962-184">Cancellation</span></span>

<span data-ttu-id="88962-185">????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-185">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="88962-186">?????????????? CPU ??????????????</span><span class="sxs-lookup"><span data-stu-id="88962-186">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="88962-187">Excel ?????????????</span><span class="sxs-lookup"><span data-stu-id="88962-187">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="88962-188">????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-188">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="88962-189">????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-189">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="88962-190">?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-190">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="88962-p124">?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-p124">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="88962-193">? *??* ??????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-193">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="88962-194">???????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-194">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="88962-195">?????????</span><span class="sxs-lookup"><span data-stu-id="88962-195">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="88962-196">????????????? `"cancelable": true` ???JSON????????? `options` ?????</span><span class="sxs-lookup"><span data-stu-id="88962-196">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="88962-197">????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-197">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="88962-198">??????`caller` ????????????????????? `onCanceled` ???</span><span class="sxs-lookup"><span data-stu-id="88962-198">In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="88962-199">???????</span><span class="sxs-lookup"><span data-stu-id="88962-199">Saving and sharing state</span></span>

<span data-ttu-id="88962-200">????????????????? JavaScript ????</span><span class="sxs-lookup"><span data-stu-id="88962-200">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="88962-201">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-201">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="88962-202">?????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-202">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="88962-203">??????????? Web ??????????????????? Web ???</span><span class="sxs-lookup"><span data-stu-id="88962-203">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="88962-204">????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-204">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="88962-205">??????????????
</span><span class="sxs-lookup"><span data-stu-id="88962-205">Note the following about this code:</span></span>

- <span data-ttu-id="88962-206">`refreshTemperature` ???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-206">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="88962-207">??????? `savedTemperatures` ??????????????</span><span class="sxs-lookup"><span data-stu-id="88962-207">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="88962-208">????????????????? *??????JSON?????*?</span><span class="sxs-lookup"><span data-stu-id="88962-208">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="88962-209">`streamTemperature` ?????????????????? `savedTemperatures` ?????????</span><span class="sxs-lookup"><span data-stu-id="88962-209">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="88962-210">????JSON????????????????? `STREAMTEMPERATURE`?</span><span class="sxs-lookup"><span data-stu-id="88962-210">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="88962-211">????? Excel UI ????????? `streamTemperature`?</span><span class="sxs-lookup"><span data-stu-id="88962-211">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="88962-212">????????? `savedTemperatures` ???????</span><span class="sxs-lookup"><span data-stu-id="88962-212">Each call reads data from the same `savedTemperatures` variable.</span></span>

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
> <span data-ttu-id="88962-213">????????JSON????? `"sync": true` ??????????????Excel????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-213">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="88962-214">??????????????????????????????????JavaScript????</span><span class="sxs-lookup"><span data-stu-id="88962-214">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="88962-215">??????</span><span class="sxs-lookup"><span data-stu-id="88962-215">Working with ranges of data</span></span>

<span data-ttu-id="88962-216">??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-216">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="88962-217">???????????Excel???????????????</span><span class="sxs-lookup"><span data-stu-id="88962-217">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="88962-218">??????????? `values`?? `Excel.CustomFunctionDimensionality.matrix` ?????</span><span class="sxs-lookup"><span data-stu-id="88962-218">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="88962-219">???????????JSON?????????? `type` ??? `matrix`?</span><span class="sxs-lookup"><span data-stu-id="88962-219">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

<span data-ttu-id="88962-220">????????JavaScript??????????2???????</span><span class="sxs-lookup"><span data-stu-id="88962-220">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="88962-221">????</span><span class="sxs-lookup"><span data-stu-id="88962-221">Known issues</span></span>

- <span data-ttu-id="88962-222">Excel ?????? URL ??????</span><span class="sxs-lookup"><span data-stu-id="88962-222">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="88962-223">????????????????Excel?</span><span class="sxs-lookup"><span data-stu-id="88962-223">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="88962-224">???????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-224">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="88962-225">???JavaScript ???????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-225">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="88962-226">???????????????? `<Page>` ?????? HTML ????? Excel ????? JavaScript?</span><span class="sxs-lookup"><span data-stu-id="88962-226">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won?t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="88962-227">???????????????????????? DOM?</span><span class="sxs-lookup"><span data-stu-id="88962-227">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="88962-228">??GET?POST?????????????APIs??? [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) ? [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) ?</span><span class="sxs-lookup"><span data-stu-id="88962-228">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="88962-229">??????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-229">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="88962-230">??????Excel for Windows???????</span><span class="sxs-lookup"><span data-stu-id="88962-230">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="88962-231">??????Office 365?????AppSource??????</span><span class="sxs-lookup"><span data-stu-id="88962-231">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="88962-232">Excel Online???????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-232">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="88962-233">????????F5??????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-233">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="88962-234">????</span><span class="sxs-lookup"><span data-stu-id="88962-234">Changelog</span></span>

- <span data-ttu-id="88962-235">**2017 ? 11 ? 7 ?**????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-235">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="88962-236">**2017 ? 11 ? 20 ?**?????????? 8801 ??????????????</span><span class="sxs-lookup"><span data-stu-id="88962-236">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="88962-237">**2017?11?28?**?????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="88962-237">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="88962-238">**2018?5?7?**????Mac?Excel Online??????????????</span><span class="sxs-lookup"><span data-stu-id="88962-238">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
