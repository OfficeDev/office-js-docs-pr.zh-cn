---
ms.date: 09/20/2018
description: 在 Excel 中使用 JavaScript 创建自定义的函数。
title: 在 Excel 中创建自定义函数（预览）
ms.openlocfilehash: 295152ca14cf56293d51b8b0512b729373841208
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062128"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="87eff-103">在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="87eff-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="87eff-104">自定义函数使开发人员可以通过在 JavaScript 中定义这些函数作为外接程序的一部分，将新函数添加到 Excel。</span><span class="sxs-lookup"><span data-stu-id="87eff-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="87eff-105">然后，用户可以像使用 Excel 中的其他本机函数（例如 `SUM()`）一样访问自定义函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-105">Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`).</span></span> <span data-ttu-id="87eff-106">本文介绍了如何在 Excel 中创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-106">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="87eff-107">下图显示了最终用户将插入 Excel 工作表的单元格的自定义的函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="87eff-108"> `CONTOSO.ADD42` 自定义函数用于将 42 添加到用户指定为函数的输入参数的一对数字中。</span><span class="sxs-lookup"><span data-stu-id="87eff-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="87eff-109">下面的代码定义 `ADD42` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-109">The following code defines the `ADD42` custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="87eff-110">自定义函数现可在 Windows、Mac 和 Excel Online 的开发人员预览版中使用。</span><span class="sxs-lookup"><span data-stu-id="87eff-110">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="87eff-111">若要试用它们，请完成以下步骤：</span><span class="sxs-lookup"><span data-stu-id="87eff-111">To try them, complete these steps:</span></span>

1. <span data-ttu-id="87eff-112">安装 Office（以 10827 为基础的 Windows 或 以 13.329 为基础的 Mac）并加入 [Office 预览体验计划](https://products.office.com/office-insider) 程序。</span><span class="sxs-lookup"><span data-stu-id="87eff-112">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="87eff-113">您必须加入 Office 预览体验计划才能访问自定义的函数；目前，除非您是 office 预览体验计划程序的成员，否则在所有 office 生成中都会禁用自定义函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-113">You must join the Office Insider program in order to have access to custom functions; currently, custom functions are disabled across all Office builds unless you are a member of the Office Insider program.</span></span>

2. <span data-ttu-id="87eff-114">使用 [Yo Office](https://github.com/OfficeDev/generator-office) 创建 Excel 自定义函数外接程序项目，然后按照 [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) 中使用项目的说明。</span><span class="sxs-lookup"><span data-stu-id="87eff-114">Use [Yo Office](https://github.com/OfficeDev/generator-office) to create an Excel Custom Functions add-in project, and then follow the instructions in the [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to use the project.</span></span>

3. <span data-ttu-id="87eff-115">`=CONTOSO.ADD42(1,2)`  在Excel 工作表的任意单元格键入，并按 **Enter** 运行自定义的函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-115">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

> [!NOTE]
> <span data-ttu-id="87eff-116">本文后面的 [已知问题](#known-issues) 一节指定了自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="87eff-116">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="87eff-117">学习基础知识</span><span class="sxs-lookup"><span data-stu-id="87eff-117">Learn the basics</span></span>

<span data-ttu-id="87eff-118">在您使用 [Yo Office](https://github.com/OfficeDev/generator-office)创建的自定义函数项目中，您将看到以下文件：</span><span class="sxs-lookup"><span data-stu-id="87eff-118">In the custom functions project that you've created using [Yo Office](https://github.com/OfficeDev/generator-office), you’ll see the following files:</span></span>

| <span data-ttu-id="87eff-119">文件</span><span class="sxs-lookup"><span data-stu-id="87eff-119">File</span></span> | <span data-ttu-id="87eff-120">文件格式</span><span class="sxs-lookup"><span data-stu-id="87eff-120">File format</span></span> | <span data-ttu-id="87eff-121">说明</span><span class="sxs-lookup"><span data-stu-id="87eff-121">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="87eff-122">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="87eff-122">**./src/customfunctions.js**</span></span> | <span data-ttu-id="87eff-123">JavaScript</span><span class="sxs-lookup"><span data-stu-id="87eff-123">JavaScript</span></span> | <span data-ttu-id="87eff-124">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="87eff-124">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="87eff-125">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="87eff-125">**./config/customfunctions.json**</span></span> | <span data-ttu-id="87eff-126">JSON</span><span class="sxs-lookup"><span data-stu-id="87eff-126">JSON</span></span> | <span data-ttu-id="87eff-127">包含描述自定义函数的元数据，并使 Excel 能够注册自定义函数以使其可供最终用户使用。</span><span class="sxs-lookup"><span data-stu-id="87eff-127">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="87eff-128">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="87eff-128">**./index.html**</span></span> | <span data-ttu-id="87eff-129">HTML</span><span class="sxs-lookup"><span data-stu-id="87eff-129">HTML</span></span> | <span data-ttu-id="87eff-130">提供 &lt;脚本&gt; 定义自定义函数的 JavaScript 文件的引用。</span><span class="sxs-lookup"><span data-stu-id="87eff-130">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="87eff-131">**Manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="87eff-131">**Manifest.xml**</span></span> | <span data-ttu-id="87eff-132">XML</span><span class="sxs-lookup"><span data-stu-id="87eff-132">XML</span></span> | <span data-ttu-id="87eff-133">此表中指定外接程序中所有自定义函数的命名空间，以及前面列出的JavaScript、 JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="87eff-133">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

### <a name="manifest-file-manifestxml"></a><span data-ttu-id="87eff-134">清单文件（./manifest.xml）</span><span class="sxs-lookup"><span data-stu-id="87eff-134">Manifest file (manifest.xml)</span></span>

<span data-ttu-id="87eff-135">定义自定义函数的外接程序的 XML 清单文件指定外接程序和 JavaScript、 JSON 和 HTML 文件的位置中的所有自定义函数的命名空间。</span><span class="sxs-lookup"><span data-stu-id="87eff-135">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="87eff-136">以下 XML 标记显示 了`<ExtensionPoint>` 和 `<Resources>` 元素的示例，您必须在外接程序的清单中包含该实例，才能使 Excel 能够运行自定义函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-136">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="87eff-137">Excel 中的函数由 XML 清单文件中指定的命名空间预置。</span><span class="sxs-lookup"><span data-stu-id="87eff-137">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="87eff-138">函数的命名空间出现在函数名之前并由句点分隔。</span><span class="sxs-lookup"><span data-stu-id="87eff-138">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="87eff-139">例如，若要`ADD42()`在 Excel 工作表的单元格中调用函数，则需要键入 `=CONTOSO.ADD42`，因为 CONTOSO 是命名空间并且`ADD42`是 JSON 文件中指定的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="87eff-139">For example, to call the function `ADD42()` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="87eff-140">该命名空间旨在用作公司或加载项的标识符。</span><span class="sxs-lookup"><span data-stu-id="87eff-140">The prefix is intended to be used as an identifier for your add-in.</span></span> 

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="87eff-141">JSON 文件（./config/customfunctions.json）</span><span class="sxs-lookup"><span data-stu-id="87eff-141">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="87eff-142">自定义函数元数据文件提供 Excel 要求注册自定义函数并使其可供最终用户使用的信息。</span><span class="sxs-lookup"><span data-stu-id="87eff-142">A custom functions metadata file provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="87eff-143">自定义函数是在用户第一次运行加载项时注册的。</span><span class="sxs-lookup"><span data-stu-id="87eff-143">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="87eff-144">之后，所有工作簿中的同一用户都可以使用它们 （即，不仅在加载项最初运行的工作簿中。）</span><span class="sxs-lookup"><span data-stu-id="87eff-144">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="87eff-145">您的 JSON 文件的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)才能使自定义函数在 Excel Online 中正常工作。</span><span class="sxs-lookup"><span data-stu-id="87eff-145">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="87eff-146">下面的代码 **customfunctions.json** 指定的元数据 `ADD42` 是以前本文中所述的函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-146">The following code in **customfunctions.json** specifies the metadata for the `ADD42` function that was described previously in this article.</span></span> <span data-ttu-id="87eff-147">此元数据定义的函数名称、说明、返回值、输入的参数等。</span><span class="sxs-lookup"><span data-stu-id="87eff-147">This metadata defines the function's name, description, return value, input parameters, and more.</span></span> <span data-ttu-id="87eff-148">下表提供了此代码示例有关的 JSON 对象中的各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="87eff-148">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span>

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
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
        }
    ]
}
```

<span data-ttu-id="87eff-149">下表列出了通常存在于 JSON 元数据文件的属性。</span><span class="sxs-lookup"><span data-stu-id="87eff-149">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="87eff-150">有关 JSON 元数据文件的详细信息，包括上一示例中未使用的选项，请参阅 [自定义函数元数据](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="87eff-150">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="87eff-151">属性</span><span class="sxs-lookup"><span data-stu-id="87eff-151">Property</span></span>  | <span data-ttu-id="87eff-152">说明</span><span class="sxs-lookup"><span data-stu-id="87eff-152">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="87eff-153">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="87eff-153">A unique ID for the group.</span></span> <span data-ttu-id="87eff-154">设置之后，不应更改此 ID。</span><span class="sxs-lookup"><span data-stu-id="87eff-154">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="87eff-155">当用户在单元格中键入公式时，自动完成菜单中显示函数的名称。</span><span class="sxs-lookup"><span data-stu-id="87eff-155">Name of the function that is shown in the autocomplete menu as a user types a formula within a cell.</span></span> <span data-ttu-id="87eff-156">在自动完成菜单中，此值将由自定义函数的命名空间中的 XML 清单文件指定作为前缀。</span><span class="sxs-lookup"><span data-stu-id="87eff-156">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="87eff-157">当用户请求帮助显示的页面的 Url。</span><span class="sxs-lookup"><span data-stu-id="87eff-157">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="87eff-158">介绍函数的用途。</span><span class="sxs-lookup"><span data-stu-id="87eff-158">Describes what the function does.</span></span> <span data-ttu-id="87eff-159">当函数是 Excel 中自动完成菜单中的选定项时，此值将显示为工具提示。</span><span class="sxs-lookup"><span data-stu-id="87eff-159">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="87eff-160">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="87eff-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="87eff-161"> `type` 子属性的值可以是 *\*字符串*\*、*\*数字*\*或 *\*布尔值*\*。</span><span class="sxs-lookup"><span data-stu-id="87eff-161">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="87eff-162"> `dimensionality` 子属性的值可以是 *\*标量*\*或 *\*矩阵** （指定值的一个二维数组 `type\`）。</span><span class="sxs-lookup"><span data-stu-id="87eff-162">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="87eff-163">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="87eff-163">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="87eff-164">该 `name` 和 `description`在 Excel intelliSense 中使用的子属性。</span><span class="sxs-lookup"><span data-stu-id="87eff-164">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="87eff-165"> `type` 和 `dimensionality\`子属性与此表中`result\`前面描述的对象的子属性相同。</span><span class="sxs-lookup"><span data-stu-id="87eff-165">The `type` and `dimensionality` child properties are identical to the child properties of the `result` object that is described previously in this table.</span></span> |
| `options` | <span data-ttu-id="87eff-166">使你可以自定义 Excel 执行函数的方式和时间等的某些方面。</span><span class="sxs-lookup"><span data-stu-id="87eff-166">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="87eff-167">有关如何使用此属性的详细信息，请参阅本文后面的 [Streamed 函数](#streamed-functions)和 [取消](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="87eff-167">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="87eff-168">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="87eff-168">Functions that return data from external sources</span></span>

<span data-ttu-id="87eff-169">如果自定义的函数从 web 等外部源检索数据，它必须：</span><span class="sxs-lookup"><span data-stu-id="87eff-169">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="87eff-170">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="87eff-170">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="87eff-171">使用回调函数，用最终值解析 Promise。</span><span class="sxs-lookup"><span data-stu-id="87eff-171">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="87eff-172">自定义功能显示 `#GETTING_DATA` Excel 等待最终结果的同时在单元格中的临时结果。</span><span class="sxs-lookup"><span data-stu-id="87eff-172">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="87eff-173">用户可以在等待结果时与电子表格的其余部分进行正常交互。</span><span class="sxs-lookup"><span data-stu-id="87eff-173">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="87eff-174">在下面的代码示例中，`getTemperature()` 自定义函数检索温度计当前温度。</span><span class="sxs-lookup"><span data-stu-id="87eff-174">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="87eff-175">注意，这`sendWebRequest` 是一个假设的函数（这里没有指定）它使用 XHR 来调用温度 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="87eff-175">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="87eff-176">流式函数</span><span class="sxs-lookup"><span data-stu-id="87eff-176">Streamed functions</span></span>

<span data-ttu-id="87eff-177">流式自定义函数使您能够在一段时间内重复地将数据输出到单元格，而无需用户明确请求重新计算。</span><span class="sxs-lookup"><span data-stu-id="87eff-177">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="87eff-178">以下示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="87eff-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="87eff-179">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="87eff-179">Note the following about this code:</span></span>

- <span data-ttu-id="87eff-180">Excel会自动使用 `setResult` 回调来显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="87eff-180">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="87eff-181">最后一个参数，`handler`，永远不会在注册代码中指定，当 Excel 用户输入该函数时，它不会显示在自动完成菜单中。</span><span class="sxs-lookup"><span data-stu-id="87eff-181">For streamed functions, the final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="87eff-182">它是包含`setResult` 回调函数的对象，用于将数据从函数传递到 Excel，以更新单元格值。</span><span class="sxs-lookup"><span data-stu-id="87eff-182">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>

- <span data-ttu-id="87eff-183">为了让Excel 在对象中中传递`setResult`函数`handler`，您必须通过设置选项`"stream": true` 在`options`注册 JSON 元数据文件中自定义函数的属性里，来声明在函数注册期间支持流。</span><span class="sxs-lookup"><span data-stu-id="87eff-183">In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="87eff-184">取消函数</span><span class="sxs-lookup"><span data-stu-id="87eff-184">Canceling a function</span></span>

<span data-ttu-id="87eff-185">在某些情况下，您可能需要取消执行流的自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。</span><span class="sxs-lookup"><span data-stu-id="87eff-185">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="87eff-186">Excel 在下列情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="87eff-186">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="87eff-187">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="87eff-187">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="87eff-188">当函数的一个参数（输入）发生变化时。</span><span class="sxs-lookup"><span data-stu-id="87eff-188">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="87eff-189">在这种情况下，取消后触发一个新函数调用。</span><span class="sxs-lookup"><span data-stu-id="87eff-189">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="87eff-190">用户手动触发重新计算。</span><span class="sxs-lookup"><span data-stu-id="87eff-190">The user triggers recalculation manually.</span></span> <span data-ttu-id="87eff-191">在这种情况下，取消后触发一个新函数调用。</span><span class="sxs-lookup"><span data-stu-id="87eff-191">In this case, a new function call is triggered in addition to the cancelation.</span></span>

> [!NOTE]
> <span data-ttu-id="87eff-192">必须为每个流式函数实现一个取消处理程序。</span><span class="sxs-lookup"><span data-stu-id="87eff-192">You must implement a cancellation handler for every streaming function.</span></span>

<span data-ttu-id="87eff-193">若要使函数取消，请 `"cancelable": true` 在`options`注册 JSON 文件中自定义函数的属性里面设置选项。</span><span class="sxs-lookup"><span data-stu-id="87eff-193">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="87eff-194">下面的代码显示以前描述的相同`incrementValue` 函数，但这一次实现了一个取消处理程序。</span><span class="sxs-lookup"><span data-stu-id="87eff-194">The following code shows the same `incrementValue` function that was described previously, but this time with a cancellation handler implemented.</span></span> <span data-ttu-id="87eff-195">本示例中， `clearInterval()` 将运行时 `incrementValue` 取消函数。</span><span class="sxs-lookup"><span data-stu-id="87eff-195">In this example, `clearInterval()` will run when the `incrementValue` function is canceled.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="87eff-196">保存和共享状态</span><span class="sxs-lookup"><span data-stu-id="87eff-196">Saving and sharing state</span></span>

<span data-ttu-id="87eff-197">自定义函数可以将数据保存在全局 JavaScript 变量中。</span><span class="sxs-lookup"><span data-stu-id="87eff-197">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="87eff-198">在后续调用中，自定义函数可以使用保存在这些变量中的值。</span><span class="sxs-lookup"><span data-stu-id="87eff-198">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="87eff-199">当用户将相同的自定义函数添加到多个单元格时，保存状态很有用，因为该函数的所有实例都可以共享该状态。</span><span class="sxs-lookup"><span data-stu-id="87eff-199">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="87eff-200">例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。</span><span class="sxs-lookup"><span data-stu-id="87eff-200">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="87eff-201">下面的代码示例演示了以前在全局上保存状态的温度流函数的实现。</span><span class="sxs-lookup"><span data-stu-id="87eff-201">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="87eff-202">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="87eff-202">Note the following about this code:</span></span>

- <span data-ttu-id="87eff-203">`refreshTemperature` 是一个流式处理函数，它会在每一秒内读取特定温度计的温度。</span><span class="sxs-lookup"><span data-stu-id="87eff-203">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="87eff-204">新的温度保存在 `savedTemperatures` 变量，但不直接更新单元格值。</span><span class="sxs-lookup"><span data-stu-id="87eff-204">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="87eff-205">它不应该直接从工作表单元格中调用， *所以它没有在JSON文件中注册*。</span><span class="sxs-lookup"><span data-stu-id="87eff-205">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="87eff-206">`streamTemperature` 每秒钟更新单元格中显示的温度值并使用 `savedTemperatures` 变量作为其数据源。</span><span class="sxs-lookup"><span data-stu-id="87eff-206">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="87eff-207">它必须在JSON文件中注册，并用所有大写字母命名， `STREAMTEMPERATURE`。</span><span class="sxs-lookup"><span data-stu-id="87eff-207">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="87eff-208">用户可以`streamTemperature` 从 Excel UI 的多个单元格中调用 。</span><span class="sxs-lookup"><span data-stu-id="87eff-208">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="87eff-209">每次调用都从相同的 `savedTemperatures` 变量读取数据。</span><span class="sxs-lookup"><span data-stu-id="87eff-209">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="87eff-210">使用数据范围</span><span class="sxs-lookup"><span data-stu-id="87eff-210">Working with ranges of data</span></span>

<span data-ttu-id="87eff-211">您自定义的函数可能接受范围的数据作为输入参数，或它可能返回的数据范围。</span><span class="sxs-lookup"><span data-stu-id="87eff-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="87eff-212">JavaScript 中，数据范围表示为一个二维数组。</span><span class="sxs-lookup"><span data-stu-id="87eff-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="87eff-213">例如，假设函数从 Excel 中存储的一系列数字中返回第二个最大值。</span><span class="sxs-lookup"><span data-stu-id="87eff-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="87eff-214">下面的函数接受参数 `values`，这是类型 `Excel.CustomFunctionDimensionality.matrix`。</span><span class="sxs-lookup"><span data-stu-id="87eff-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="87eff-215">请注意，在该函数的 JSON 元数据中，您可以将该参数的`type` 属性`matrix` 设置为。</span><span class="sxs-lookup"><span data-stu-id="87eff-215">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){
     let highest = values[0][0], secondHighest = values[0][0];
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

## <a name="handling-errors"></a><span data-ttu-id="87eff-216">处理错误</span><span class="sxs-lookup"><span data-stu-id="87eff-216">Handling errors</span></span>

<span data-ttu-id="87eff-217">在生成定义自定义函数的加载项时，请务必包含错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="87eff-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="87eff-218">自定义的函数的错误处理与 [Excel JavaScript API 的错误处理整体类同](excel-add-ins-error-handling.md)。</span><span class="sxs-lookup"><span data-stu-id="87eff-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="87eff-219">在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="87eff-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    //this delivers a section of lorem ipsum from the jsonplaceholder API
    let url = "https://jsonplaceholder.typicode.com/comments/" + x;

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

## <a name="known-issues"></a><span data-ttu-id="87eff-220">已知问题</span><span class="sxs-lookup"><span data-stu-id="87eff-220">Known issues</span></span>

- <span data-ttu-id="87eff-221">不支持可变函数（每当电子表格中不相关的数据更改时自动重新计算）。</span><span class="sxs-lookup"><span data-stu-id="87eff-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="87eff-222">自定义函数目前不适用于移动客户的 Excel。</span><span class="sxs-lookup"><span data-stu-id="87eff-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="87eff-223">不支持可变函数（每当电子表格中不相关的数据更改时自动重新计算）。</span><span class="sxs-lookup"><span data-stu-id="87eff-223">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="87eff-224">尚未启用通过 Office 365 Admin Portal 和 AppSource 进行的部署。</span><span class="sxs-lookup"><span data-stu-id="87eff-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="87eff-225">Excel Online中的自定义功能，可能会在一段时间无活动后，在进程期间停止工作。</span><span class="sxs-lookup"><span data-stu-id="87eff-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="87eff-226">刷新浏览器页面（F5）并重新输入自定义函数以恢复该功能。</span><span class="sxs-lookup"><span data-stu-id="87eff-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="87eff-227">如果您有多个加载项在 Excel for Windows 上运行，您可能会看到 **#GETTING_DATA**临时结果单元格内的工作表。</span><span class="sxs-lookup"><span data-stu-id="87eff-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="87eff-228">关闭 Excel 的所有窗口，并重新启动 Excel。</span><span class="sxs-lookup"><span data-stu-id="87eff-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="87eff-229">将来可能会提供专门用于自定义函数的调试工具。</span><span class="sxs-lookup"><span data-stu-id="87eff-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="87eff-230">同时，您可以在 Excel Online 使用 F12 开发人员工具调试。</span><span class="sxs-lookup"><span data-stu-id="87eff-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="87eff-231">请参阅 [自定义函数最佳实践](custom-functions-best-practices.md)中的详细信息。</span><span class="sxs-lookup"><span data-stu-id="87eff-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="87eff-232">更改日志</span><span class="sxs-lookup"><span data-stu-id="87eff-232">Changelog</span></span>

- <span data-ttu-id="87eff-233">**2017 年 11 月 7 日**：发布了\*自定义函数预览和示例</span><span class="sxs-lookup"><span data-stu-id="87eff-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="87eff-234">**2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题</span><span class="sxs-lookup"><span data-stu-id="87eff-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="87eff-235">**2017 年 11 月 28 日**：发布了\*对取消异步函数的支持（需要对流式函数进行相应更改）</span><span class="sxs-lookup"><span data-stu-id="87eff-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="87eff-236">**2018 年 5 月 7 日**：发布了\*对 Mac、Excel Online 和在进程中运行的同步函数的支持</span><span class="sxs-lookup"><span data-stu-id="87eff-236">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="87eff-237">**2018 年 9 月 20 日，** 发布了支持自定义函数 JavaScript 的运行时。</span><span class="sxs-lookup"><span data-stu-id="87eff-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="87eff-238">有关详细信息，请参阅 [Excel 自定义函数的运行时](custom-functions-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="87eff-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="87eff-239">\* 至 Office 预览体验计划渠道</span><span class="sxs-lookup"><span data-stu-id="87eff-239">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="87eff-240">另请参阅</span><span class="sxs-lookup"><span data-stu-id="87eff-240">See also</span></span>

* [<span data-ttu-id="87eff-241">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="87eff-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="87eff-242">Excel 自定义函数运行运行时</span><span class="sxs-lookup"><span data-stu-id="87eff-242">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="87eff-243">自定义函数的最佳做法</span><span class="sxs-lookup"><span data-stu-id="87eff-243">Custom functions best practices</span></span>](custom-functions-best-practices.md)