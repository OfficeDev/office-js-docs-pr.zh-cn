---
ms.date: 10/17/2018
description: 在 Excel 中使用 JavaScript 创建自定义函数。
title: 在 Excel 中创建自定义函数（预览）
ms.openlocfilehash: 8383b5f6d568a1ce2da036fbacfb90404bbe8297
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298549"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="18919-103">在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="18919-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="18919-104">开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="18919-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="18919-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="18919-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="18919-106">本文介绍了如何在 Excel 中创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="18919-106">This article explains how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="18919-107">下图演示最终用户将自定义函数插入到 Excel 工作表单元格的过程。</span><span class="sxs-lookup"><span data-stu-id="18919-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="18919-108">`CONTOSO.ADD42` 自定义函数旨在向用户指定作为函数输入参数的数字对添加 42。</span><span class="sxs-lookup"><span data-stu-id="18919-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="18919-109">以下代码定义 `ADD42` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="18919-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="18919-110">本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="18919-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="18919-111">自定义函数加载项项目的组件</span><span class="sxs-lookup"><span data-stu-id="18919-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="18919-112">如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，将在生成器创建的项目中看到以下文件：</span><span class="sxs-lookup"><span data-stu-id="18919-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="18919-113">文件</span><span class="sxs-lookup"><span data-stu-id="18919-113">File</span></span> | <span data-ttu-id="18919-114">文件格式</span><span class="sxs-lookup"><span data-stu-id="18919-114">File format</span></span> | <span data-ttu-id="18919-115">说明</span><span class="sxs-lookup"><span data-stu-id="18919-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="18919-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="18919-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="18919-117">或</span><span class="sxs-lookup"><span data-stu-id="18919-117">or</span></span><br/><span data-ttu-id="18919-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="18919-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="18919-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="18919-119">JavaScript</span></span><br/><span data-ttu-id="18919-120">或</span><span class="sxs-lookup"><span data-stu-id="18919-120">or</span></span><br/><span data-ttu-id="18919-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="18919-121">TypeScript</span></span> | <span data-ttu-id="18919-122">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="18919-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="18919-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="18919-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="18919-124">JSON</span><span class="sxs-lookup"><span data-stu-id="18919-124">JSON</span></span> | <span data-ttu-id="18919-125">包含描述自定义函数的元数据，使 Excel 能够注册自定义函数，并使其可供最终用户使用。</span><span class="sxs-lookup"><span data-stu-id="18919-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="18919-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="18919-126">**./index.html**</span></span> | <span data-ttu-id="18919-127">HTML</span><span class="sxs-lookup"><span data-stu-id="18919-127">HTML</span></span> | <span data-ttu-id="18919-128">提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。</span><span class="sxs-lookup"><span data-stu-id="18919-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="18919-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="18919-129">**Manifest.xml**</span></span> | <span data-ttu-id="18919-130">XML</span><span class="sxs-lookup"><span data-stu-id="18919-130">XML</span></span> | <span data-ttu-id="18919-131">指定加载项中所有自定义函数的命名空间以及此表中前面列出的 JavaScript、JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="18919-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="18919-132">下列部分将提供有关这些文件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="18919-132">The following sections provide more information about those changes.</span></span>

### <a name="script-file"></a><span data-ttu-id="18919-133">脚本文件</span><span class="sxs-lookup"><span data-stu-id="18919-133">Script file</span></span> 

<span data-ttu-id="18919-134">脚本文件（Yo Office 生成器创建的项目中的 **./src/customfunctions.js** 或 **./src/customfunctions.ts**）包含定义自定义函数并将自定义函数名称映射到 [JSON 元数据文件](#json-metadata-file)中的对象的代码。</span><span class="sxs-lookup"><span data-stu-id="18919-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="18919-135">例如，以下代码定义自定义函数 `add` 和 `increment`，然后指定这两个函数的映射信息。</span><span class="sxs-lookup"><span data-stu-id="18919-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="18919-136">将 `add` 函数映射到 JSON 元数据文件中的对象，其中 `id` 属性的值为 **ADD**，将 `increment` 函数映射到元数据文件中的对象，其中 `id` 属性的值为 **INCREMENT**。</span><span class="sxs-lookup"><span data-stu-id="18919-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="18919-137">有关将脚本文件中的函数名称映射到 JSON 元数据文件中的对象的更多信息，请参阅[自定义函数最佳实践](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="18919-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

```js
function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a><span data-ttu-id="18919-138">JSON 元数据文件</span><span class="sxs-lookup"><span data-stu-id="18919-138">JSON metadata file</span></span> 

<span data-ttu-id="18919-139">自定义函数元数据文件（Yo Office 生成器创建的项目中的 **./config/customfunctions.json**）提供 Excel 注册自定义函数并使其可供最终用户使用所需的信息。</span><span class="sxs-lookup"><span data-stu-id="18919-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="18919-140">自定义函数在用户首次运行加载项时注册。</span><span class="sxs-lookup"><span data-stu-id="18919-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="18919-141">之后，它们可在所有工作簿（即，不仅仅是在加载项初始运行的工作簿）中供同一用户使用。</span><span class="sxs-lookup"><span data-stu-id="18919-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="18919-142">托管 JSON 文件的服务器上的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)，以便自定义函数在 Excel Online 中正常工作。</span><span class="sxs-lookup"><span data-stu-id="18919-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="18919-143">**customfunctions.json** 中的以下代码指定上述 `add` 函数和 `increment` 函数的元数据。</span><span class="sxs-lookup"><span data-stu-id="18919-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="18919-144">此代码示例后面的表提供了有关此 JSON 对象中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="18919-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="18919-145">有关在 JSON 元数据文件中指定 `id` 和 `name` 属性值的详细信息，请参阅[自定义函数最佳实践](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="18919-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com",
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
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

<span data-ttu-id="18919-146">下表列出了 JSON 元数据文件中的常见属性。</span><span class="sxs-lookup"><span data-stu-id="18919-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="18919-147">有关 JSON 元数据文件的更多详细信息，请参阅[自定义函数元数据](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="18919-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="18919-148">属性</span><span class="sxs-lookup"><span data-stu-id="18919-148">Property</span></span>  | <span data-ttu-id="18919-149">说明</span><span class="sxs-lookup"><span data-stu-id="18919-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="18919-150">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="18919-150">A unique ID for the group.</span></span> <span data-ttu-id="18919-151">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="18919-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="18919-152">最终用户在 Excel 中看到的函数名称。</span><span class="sxs-lookup"><span data-stu-id="18919-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="18919-153">在 Excel 中，此函数名称将以 [XML 清单文件](#manifest-file)中指定的自定义函数命名空间作为前缀。</span><span class="sxs-lookup"><span data-stu-id="18919-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="18919-154">当用户请求帮助时显示的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="18919-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="18919-155">说明函数的功能。</span><span class="sxs-lookup"><span data-stu-id="18919-155">Describes what the function does.</span></span> <span data-ttu-id="18919-156">当函数是 Excel 自动完成菜单中的选中项时，此值将作为工具提示显示。</span><span class="sxs-lookup"><span data-stu-id="18919-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="18919-157">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="18919-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="18919-158">有关此对象的详细信息，请参阅[结果](custom-functions-json.md#result)。</span><span class="sxs-lookup"><span data-stu-id="18919-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="18919-159">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="18919-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="18919-160">有关此对象的详细信息，请参阅[参数](custom-functions-json.md#parameters)。</span><span class="sxs-lookup"><span data-stu-id="18919-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="18919-161">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="18919-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="18919-162">有关如何使用此属性的详细信息，请参阅本文后面的[流式处理函数](#streaming-functions)和[取消函数](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="18919-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="18919-163">清单文件</span><span class="sxs-lookup"><span data-stu-id="18919-163">Manifest file</span></span>

<span data-ttu-id="18919-164">定义自定义函数的加载项的 XML 清单文件（Yo Office 生成器创建的项目中的 **./manifest.xml**）指定加载项中所有自定义函数的命名空间以及 JavaScript、JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="18919-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="18919-165">下面的 XML 标记显示了 `<ExtensionPoint>` 和 `<Resources>` 元素的一个示例，必须在加载项清单中包含这些元素才能启用自定义函数。</span><span class="sxs-lookup"><span data-stu-id="18919-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. Can only contain alphanumeric characters and periods.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="18919-166">Excel 中的函数在前面追加 XML 清单文件中指定的命名空间作为前缀。</span><span class="sxs-lookup"><span data-stu-id="18919-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="18919-167">函数的命名空间在函数名称之前，并用句点分隔。</span><span class="sxs-lookup"><span data-stu-id="18919-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="18919-168">例如，若要在 Excel 工作表的单元格中调用函数 `ADD42`，需输入 `=CONTOSO.ADD42`，因为 `CONTOSO` 是命名空间，`ADD42` 是 JSON 文件中指定的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="18919-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="18919-169">命名空间旨在作为公司或加载项的标识符使用。</span><span class="sxs-lookup"><span data-stu-id="18919-169">The prefix is intended to be used as an identifier for your add-in.</span></span> <span data-ttu-id="18919-170">命名空间只能包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="18919-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="18919-171">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="18919-171">Functions that return data from external sources</span></span>

<span data-ttu-id="18919-172">如果自定义函数从外部源（如 Web）检索数据，则必须：</span><span class="sxs-lookup"><span data-stu-id="18919-172">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="18919-173">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="18919-173">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="18919-174">使用回调函数解析带有最终值的 Promise。</span><span class="sxs-lookup"><span data-stu-id="18919-174">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="18919-175">在 Excel 等待最终结果时，自定义函数会在单元格中显示一个 `#GETTING_DATA` 临时结果。</span><span class="sxs-lookup"><span data-stu-id="18919-175">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="18919-176">在等待结果时，用户可以与工作表的其余部分正常交互。</span><span class="sxs-lookup"><span data-stu-id="18919-176">Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="18919-177">在下面的代码示例中，`getTemperature()` 自定义函数检索温度计的当前温度。</span><span class="sxs-lookup"><span data-stu-id="18919-177">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="18919-178">注意，`sendWebRequest` 是一个假设函数（此处未指定），它使用 [XHR](custom-functions-runtime.md#xhr-example) 调用温度 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="18919-178">Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="18919-179">流式处理函数</span><span class="sxs-lookup"><span data-stu-id="18919-179">Streaming functions</span></span>

<span data-ttu-id="18919-180">流式处理自定义函数使用户能够在不需要用户显式请求数据刷新的情况下，随着时间的推移向单元格重复输出数据。</span><span class="sxs-lookup"><span data-stu-id="18919-180">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="18919-181">下面的代码示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="18919-181">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="18919-182">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="18919-182">Note the following about this code:</span></span>

- <span data-ttu-id="18919-183">Excel 使用 `setResult` 回调自动显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="18919-183">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="18919-184">当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数 `handler`。</span><span class="sxs-lookup"><span data-stu-id="18919-184">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="18919-185">`onCanceled` 回调定义取消函数时执行的函数。</span><span class="sxs-lookup"><span data-stu-id="18919-185">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="18919-186">对于任何流式处理函数，都必须实现此类取消处理程序。</span><span class="sxs-lookup"><span data-stu-id="18919-186">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="18919-187">有关详细信息，请参阅[取消函数](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="18919-187">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

<span data-ttu-id="18919-188">在 JSON 元数据文件中为流式处理函数指定元数据时，必须在 `options` 对象中设置属性 `"cancelable": true` 和 `"stream": true`，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="18919-188">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="18919-189">取消函数</span><span class="sxs-lookup"><span data-stu-id="18919-189">Canceling a function</span></span>

<span data-ttu-id="18919-190">在某些情况下，可能需要取消执行流式处理自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。</span><span class="sxs-lookup"><span data-stu-id="18919-190">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="18919-191">Excel 会在以下情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="18919-191">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="18919-192">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="18919-192">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="18919-193">函数的参数（输入）之一发生变化。</span><span class="sxs-lookup"><span data-stu-id="18919-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="18919-194">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="18919-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="18919-195">用户手动触发重新计算。</span><span class="sxs-lookup"><span data-stu-id="18919-195">When the user triggers recalculation manually.</span></span> <span data-ttu-id="18919-196">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="18919-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="18919-197">为了能够取消函数，必须在 JavaScript 函数中实现一个取消处理程序，并在说明函数的 JSON 元数据中指定 `options` 对象中的属性 `"cancelable": true`。</span><span class="sxs-lookup"><span data-stu-id="18919-197">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="18919-198">本文前一部分中的代码示例提供了这些方法的示例。</span><span class="sxs-lookup"><span data-stu-id="18919-198">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="18919-199">保存和共享状态</span><span class="sxs-lookup"><span data-stu-id="18919-199">Saving and sharing state</span></span>

<span data-ttu-id="18919-200">自定义函数可以将数据保存在全局 JavaScript 变量中，可用于后续调用。</span><span class="sxs-lookup"><span data-stu-id="18919-200">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="18919-201">当用户从多个单元格调用同一个自定义函数时，保存状态非常有用，因为函数的所有实例都可以访问该状态。</span><span class="sxs-lookup"><span data-stu-id="18919-201">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="18919-202">例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。</span><span class="sxs-lookup"><span data-stu-id="18919-202">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="18919-203">下面的代码示例演示温度流式处理函数的实现过程，该函数在全局范围内保存状态。</span><span class="sxs-lookup"><span data-stu-id="18919-203">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="18919-204">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="18919-204">Note the following about this code:</span></span>

- <span data-ttu-id="18919-205">`streamTemperature` 函数每秒更新单元格中显示的温度值，并使用 `savedTemperatures` 变量作为其数据源。</span><span class="sxs-lookup"><span data-stu-id="18919-205">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="18919-206">因为 `streamTemperature` 是一个流式处理函数，它将实现一个取消处理程序，当函数被取消时该处理程序将运行。</span><span class="sxs-lookup"><span data-stu-id="18919-206">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="18919-207">如果用户从 Excel 中的多个单元格调用 `streamTemperature` 函数，则 `streamTemperature` 函数在每次运行时都会从相同的 `savedTemperatures` 变量读取数据。</span><span class="sxs-lookup"><span data-stu-id="18919-207">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="18919-208">`refreshTemperature` 函数每秒读取特定温度计的温度，并将结果存储在 `savedTemperatures` 变量中。</span><span class="sxs-lookup"><span data-stu-id="18919-208">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="18919-209">因为 `refreshTemperature` 函数不在 Excel 中向最终用户显示，所以不需要在 JSON 文件中注册。</span><span class="sxs-lookup"><span data-stu-id="18919-209">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="18919-210">使用数据区域</span><span class="sxs-lookup"><span data-stu-id="18919-210">Working with ranges of data</span></span>

<span data-ttu-id="18919-211">自定义函数可以接受数据区域作为输入参数，也可以返回数据区域。</span><span class="sxs-lookup"><span data-stu-id="18919-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="18919-212">在 JavaScript，数据区域表示为一个二维数组。</span><span class="sxs-lookup"><span data-stu-id="18919-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="18919-213">例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。</span><span class="sxs-lookup"><span data-stu-id="18919-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="18919-214">下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。</span><span class="sxs-lookup"><span data-stu-id="18919-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="18919-215">请注意，在此函数的 JSON 元数据中，将参数的 `type` 属性设置为 `matrix`。</span><span class="sxs-lookup"><span data-stu-id="18919-215">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="18919-216">处理错误</span><span class="sxs-lookup"><span data-stu-id="18919-216">Handling errors</span></span>

<span data-ttu-id="18919-217">在生成定义自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="18919-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="18919-218">自定义函数的错误处理与 [Excel JavaScript API 的错误处理](excel-add-ins-error-handling.md)大致相同。</span><span class="sxs-lookup"><span data-stu-id="18919-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="18919-219">在以下代码示例中，`.catch` 将处理之前发生在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="18919-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;

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

## <a name="known-issues"></a><span data-ttu-id="18919-220">已知问题</span><span class="sxs-lookup"><span data-stu-id="18919-220">Known issues</span></span>

- <span data-ttu-id="18919-221">Excel 暂未使用帮助 URL 和参数说明。</span><span class="sxs-lookup"><span data-stu-id="18919-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="18919-222">移动客户端目前还不能在 Excel 上使用自定义函数。</span><span class="sxs-lookup"><span data-stu-id="18919-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="18919-223">尚不支持 Volatile 函数（每当电子表格中不相关的数据发生变化时都会自动重新计算的函数）。</span><span class="sxs-lookup"><span data-stu-id="18919-223">Volatile functions (those that recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="18919-224">尚未启用通过 Office 365 管理门户和 AppSource 进行部署。</span><span class="sxs-lookup"><span data-stu-id="18919-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="18919-225">在一段时间处于非活动状态后，Excel Online 中的自定义函数可能会在会话期间停止工作。</span><span class="sxs-lookup"><span data-stu-id="18919-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="18919-226">刷新浏览器页面 (F5) 并重新输入自定义函数以恢复该功能。</span><span class="sxs-lookup"><span data-stu-id="18919-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="18919-227">如果在 Excel for Windows 上运行多个加载项，可能会在工作表单元格中看到 **#GETTING_DATA** 临时结果。</span><span class="sxs-lookup"><span data-stu-id="18919-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="18919-228">关闭所有 Excel 窗口并重启 Excel。</span><span class="sxs-lookup"><span data-stu-id="18919-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="18919-229">将来可能会推出专门针对自定义函数的调试工具。</span><span class="sxs-lookup"><span data-stu-id="18919-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="18919-230">在此期间，可以使用 F12 开发人员工具在 Excel Online 上进行调试。</span><span class="sxs-lookup"><span data-stu-id="18919-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="18919-231">有关详细信息，请参阅[自定义函数最佳实践](custom-functions-best-practices.md)。</span><span class="sxs-lookup"><span data-stu-id="18919-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="18919-232">更改日志</span><span class="sxs-lookup"><span data-stu-id="18919-232">Changelog</span></span>

- <span data-ttu-id="18919-233">**2017 年 11 月 7 日**：发布了\*自定义函数（预览）和示例</span><span class="sxs-lookup"><span data-stu-id="18919-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="18919-234">**2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题</span><span class="sxs-lookup"><span data-stu-id="18919-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="18919-235">**2017 年 11 月 28 日**：发布了\*对取消异步函数的支持（需要对流式处理函数进行相应更改）</span><span class="sxs-lookup"><span data-stu-id="18919-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="18919-236">**2018 年 5 月 7 日**：发布了\*对 Mac、Excel Online 和在进程中运行的异步函数的支持</span><span class="sxs-lookup"><span data-stu-id="18919-236">**May 7, 2018**: Shipped\* support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="18919-237">**2018 年 9 月 20 日**：发布了对自定义函数 JavaScript 运行时的支持。</span><span class="sxs-lookup"><span data-stu-id="18919-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="18919-238">有关详细信息，请参阅 [Excel 自定义函数的运行时](custom-functions-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="18919-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>
- <span data-ttu-id="18919-239">**2018 年 10 月 20 日**：随着 [10 月预览体验内部版本](https://support.office.com/zh-CN/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24)的推出，自定义函数现在需要适用于 Windows Desktop 和 Online 的[自定义函数元数据](custom-functions-json.md)中的“id”参数。</span><span class="sxs-lookup"><span data-stu-id="18919-239">**October 20, 2018**: With the [October Insiders build](https://support.office.com/zh-CN/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), Custom Functions now requires the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online.</span></span> <span data-ttu-id="18919-240">在 Mac 上，应忽略此参数。</span><span class="sxs-lookup"><span data-stu-id="18919-240">On Mac, this parameter should be ignored.</span></span>


<span data-ttu-id="18919-241">\* 转到 [Office 预览体验成员](https://products.office.com/office-insider)频道（以前称为“预览体验成员 - 快”）</span><span class="sxs-lookup"><span data-stu-id="18919-241">\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")</span></span>

## <a name="see-also"></a><span data-ttu-id="18919-242">另请参阅</span><span class="sxs-lookup"><span data-stu-id="18919-242">See also</span></span>

* [<span data-ttu-id="18919-243">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="18919-243">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="18919-244">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="18919-244">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="18919-245">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="18919-245">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="18919-246">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="18919-246">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
