---
ms.date: 09/27/2018
description: 在 Excel 中使用 JavaScript 创建自定义的函数。
title: 在 Excel 中创建自定义函数（预览）
ms.openlocfilehash: 98e418f843f6f5574088cea9c7393afc4a42060b
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348799"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="0a73d-103">在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="0a73d-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="0a73d-p101">自定义函数使开发人员可以通过在 JavaScript 中定义这些函数作为加载项的一部分，将新函数添加到 Excel。然后，用户可以像使用 Excel 中的其他本机函数（例如 `SUM()`）一样访问自定义函数。本文介绍了如何在 Excel 中创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="0a73d-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`). This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0a73d-107">下图显示了最终用户将插入 Excel 工作表的单元格的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="0a73d-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="0a73d-108">`CONTOSO.ADD42` 自定义函数用于将 42 添加到用户指定为函数的输入参数的一对数字中。</span><span class="sxs-lookup"><span data-stu-id="0a73d-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="0a73d-109">下面的代码定义 `ADD42` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="0a73d-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="0a73d-110">本文后面的 [已知问题](#known-issues) 一节指定了自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="0a73d-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="0a73d-111">自定义函数加载项项目的组件</span><span class="sxs-lookup"><span data-stu-id="0a73d-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="0a73d-112">果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office) 创建 Excel 自定义函数加载项项目，将在项目中看到生成器创建的以下文件：</span><span class="sxs-lookup"><span data-stu-id="0a73d-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="0a73d-113">文件</span><span class="sxs-lookup"><span data-stu-id="0a73d-113">File</span></span> | <span data-ttu-id="0a73d-114">文件格式</span><span class="sxs-lookup"><span data-stu-id="0a73d-114">File format</span></span> | <span data-ttu-id="0a73d-115">说明</span><span class="sxs-lookup"><span data-stu-id="0a73d-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="0a73d-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="0a73d-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="0a73d-117">或</span><span class="sxs-lookup"><span data-stu-id="0a73d-117">or</span></span><br/><span data-ttu-id="0a73d-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="0a73d-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="0a73d-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="0a73d-119">JavaScript</span></span><br/><span data-ttu-id="0a73d-120">或</span><span class="sxs-lookup"><span data-stu-id="0a73d-120">or</span></span><br/><span data-ttu-id="0a73d-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="0a73d-121">TypeScript</span></span> | <span data-ttu-id="0a73d-122">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="0a73d-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="0a73d-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="0a73d-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="0a73d-124">JSON</span><span class="sxs-lookup"><span data-stu-id="0a73d-124">JSON</span></span> | <span data-ttu-id="0a73d-125">包含描述自定义函数的元数据，并使 Excel 能够注册自定义函数且使其可供最终用户使用。</span><span class="sxs-lookup"><span data-stu-id="0a73d-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="0a73d-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="0a73d-126">**./index.html**</span></span> | <span data-ttu-id="0a73d-127">HTML</span><span class="sxs-lookup"><span data-stu-id="0a73d-127">HTML</span></span> | <span data-ttu-id="0a73d-128">提供 &lt;脚本&gt; 定义自定义函数的 JavaScript 文件的引用。</span><span class="sxs-lookup"><span data-stu-id="0a73d-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="0a73d-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="0a73d-129">**Manifest.xml**</span></span> | <span data-ttu-id="0a73d-130">XML</span><span class="sxs-lookup"><span data-stu-id="0a73d-130">XML</span></span> | <span data-ttu-id="0a73d-131">此表中指定加载项中所有自定义函数的命名空间，以及前面列出的 JavaScript、JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="0a73d-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="0a73d-132">下面的章节将提供有关这些文件的更多信息。</span><span class="sxs-lookup"><span data-stu-id="0a73d-132">The following sections provide more information about those changes.</span></span>

### <a name="script-file"></a><span data-ttu-id="0a73d-133">脚本文件</span><span class="sxs-lookup"><span data-stu-id="0a73d-133">Script file</span></span> 

<span data-ttu-id="0a73d-134">脚本文件 （Yo Office 生成器创建项目中的 **./src/customfunctions.js** 或 **./src/customfunctions.ts**）包含定义自定义函数的代码，该代码还将自定义函数的名称映射到 [JSON 元数据文件](#json-metadata-file)中的对象。</span><span class="sxs-lookup"><span data-stu-id="0a73d-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="0a73d-135">例如，下面的代码示例定义自定义函数 `add` 和 `increment`，然后指定这两个函数的映射信息。</span><span class="sxs-lookup"><span data-stu-id="0a73d-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="0a73d-136">`add` 函数被映射到 JSON 元数据文件中的对象，其中 `id` 属性的值是 **ADD**，而 `increment` 函数被映射到元数据文件中的对象，其中 `id` 属性的值是 **INCREMENT**。</span><span class="sxs-lookup"><span data-stu-id="0a73d-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="0a73d-137">有关将脚本文件中的函数名称映射到 JSON 元数据文件中对象的详细信息，请参阅[自定义函数的最佳做法](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="0a73d-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="0a73d-138">JSON 元数据文件</span><span class="sxs-lookup"><span data-stu-id="0a73d-138">JSON metadata file</span></span> 

<span data-ttu-id="0a73d-139">自定义函数元数据文件（Yo Office 生成器所创建项目中的 **./config/customfunctions.json**）提供 Excel 注册自定义函数需要的信息，并将其提供给最终用户。</span><span class="sxs-lookup"><span data-stu-id="0a73d-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="0a73d-140">自定义函数是在用户第一次运行加载项时注册的。</span><span class="sxs-lookup"><span data-stu-id="0a73d-140">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="0a73d-141">之后，所有工作簿中的同一用户都可以使用它们 （即不仅在加载项最初运行的工作簿中。）</span><span class="sxs-lookup"><span data-stu-id="0a73d-141">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="0a73d-142">JSON 文件的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) 才能使自定义函数在 Excel Online 中正常工作。</span><span class="sxs-lookup"><span data-stu-id="0a73d-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="0a73d-143">下面在 **customfunctions.json** 中的代码指定前面所述 `add` 函数和 `increment` 函数的元数据。</span><span class="sxs-lookup"><span data-stu-id="0a73d-143">The following code in **customfunctions.json** specifies the metadata for the `add` function that was described previously in this article.</span></span> <span data-ttu-id="0a73d-144">此代码后示例的表格提供了有关此 JSON 对象中单独属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="0a73d-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="0a73d-145">有关指定 JSON 元数据文件中 `id` 和 `name` 属性值的详细信息，请参阅[自定义函数的最佳做法](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="0a73d-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="0a73d-146">下表列出了通常存在于 JSON 元数据文件的属性。</span><span class="sxs-lookup"><span data-stu-id="0a73d-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="0a73d-147">有关 JSON 元数据文件的详细信息，请参阅[自定义函数元数据](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="0a73d-147">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="0a73d-148">属性</span><span class="sxs-lookup"><span data-stu-id="0a73d-148">Property</span></span>  | <span data-ttu-id="0a73d-149">说明</span><span class="sxs-lookup"><span data-stu-id="0a73d-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="0a73d-150">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="0a73d-150">A unique ID for the group.</span></span> <span data-ttu-id="0a73d-151">设置之后，不应更改此 ID。</span><span class="sxs-lookup"><span data-stu-id="0a73d-151">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="0a73d-152">最终用户在 Excel 中看到的函数名称。</span><span class="sxs-lookup"><span data-stu-id="0a73d-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="0a73d-153">在 Excel 中，此函数名称将以[ XML ](#manifest-file)清单文件中指定的自定义函数命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="0a73d-153">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="0a73d-154">用户请求帮助时显示的页面的 Url。</span><span class="sxs-lookup"><span data-stu-id="0a73d-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="0a73d-155">介绍函数的用途。</span><span class="sxs-lookup"><span data-stu-id="0a73d-155">Describes what the function does.</span></span> <span data-ttu-id="0a73d-156">当函数是 Excel 中自动完成菜单中的选定项时，此值将显示为工具提示。</span><span class="sxs-lookup"><span data-stu-id="0a73d-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="0a73d-157">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="0a73d-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0a73d-158">`type` 子属性的值可以是**字符串**、**数字**或**布尔值**。</span><span class="sxs-lookup"><span data-stu-id="0a73d-158">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="0a73d-159">`dimensionality` 子属性的值可以是**scalar** 或 **matrix**（指定 `type` 值的二维数组）。</span><span class="sxs-lookup"><span data-stu-id="0a73d-159">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="0a73d-160">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="0a73d-160">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="0a73d-161">在 Excel intelliSense 中出现的 `name` 和 `description` 子属性。</span><span class="sxs-lookup"><span data-stu-id="0a73d-161">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="0a73d-162">`type` 子属性的值可以是**字符串**、**数字**或**布尔值**。</span><span class="sxs-lookup"><span data-stu-id="0a73d-162">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="0a73d-163">`dimensionality` 子属性的值可以是**scalar** 或 **matrix**（指定 `type` 值的二维数组）。</span><span class="sxs-lookup"><span data-stu-id="0a73d-163">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `options` | <span data-ttu-id="0a73d-164">使你可以自定义 Excel 执行函数的方式和时间等的某些方面。</span><span class="sxs-lookup"><span data-stu-id="0a73d-164">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0a73d-165">有关如何使用此属性的详细信息，请参阅本文后面的[流式函数](#streamed-functions)和[取消函数](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="0a73d-165">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="0a73d-166">清单文件</span><span class="sxs-lookup"><span data-stu-id="0a73d-166">Manifest file</span></span>

<span data-ttu-id="0a73d-167">定义自定义函数（Yo Office 生成器所创建项目中的 **./manifest.xml**）的加载项的 XML 清单文件指定加载项和 JavaScript、JSON 和 HTML 文件的位置中的所有自定义函数的命名空间。</span><span class="sxs-lookup"><span data-stu-id="0a73d-167">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="0a73d-168">以下 XML 标记显示 了`<ExtensionPoint>` 和 `<Resources>` 元素的示例，必须在加载项的清单中包含该实例，以启用自定义函数。</span><span class="sxs-lookup"><span data-stu-id="0a73d-168">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="0a73d-169">Excel 中的函数以 XML 清单文件中指定的命名空间为前缀。</span><span class="sxs-lookup"><span data-stu-id="0a73d-169">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="0a73d-170">函数的命名空间出现在函数名之前并由句点分隔。</span><span class="sxs-lookup"><span data-stu-id="0a73d-170">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="0a73d-171">例如，若要`ADD42`在 Excel 工作表的单元格中调用函数，则需要键入 `=CONTOSO.ADD42`，因为 CONTOSO 是命名空间并且`ADD42`是 JSON 文件中所指定函数的名称。</span><span class="sxs-lookup"><span data-stu-id="0a73d-171">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="0a73d-172">命名空间旨在用作公司或加载项的标识符。</span><span class="sxs-lookup"><span data-stu-id="0a73d-172">The prefix is intended to be used as an identifier for your add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="0a73d-173">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="0a73d-173">Functions that return data from external sources</span></span>

<span data-ttu-id="0a73d-174">如果自定义函数从 web 等外部源检索数据，它必须：</span><span class="sxs-lookup"><span data-stu-id="0a73d-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="0a73d-175">将 JavaScript 承诺返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="0a73d-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="0a73d-176">使用回调函数，用最终值解析承诺。</span><span class="sxs-lookup"><span data-stu-id="0a73d-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="0a73d-177">自定义函数在单元格中显示 `#GETTING_DATA` 临时结果，而 Excel 等待最终结果。</span><span class="sxs-lookup"><span data-stu-id="0a73d-177">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="0a73d-178">用户可以在等待结果时与电子表格的其余部分进行正常交互。</span><span class="sxs-lookup"><span data-stu-id="0a73d-178">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="0a73d-179">在下面的代码示例中，`getTemperature()` 自定义函数检索温度计当前温度。</span><span class="sxs-lookup"><span data-stu-id="0a73d-179">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="0a73d-180">注意，`sendWebRequest` 是一个假设的函数（这里没有指定），它使用 [XHR](custom-functions-runtime.md#xhr) 来调用温度 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="0a73d-180">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="0a73d-181">流式函数</span><span class="sxs-lookup"><span data-stu-id="0a73d-181">Streamed functions</span></span>

<span data-ttu-id="0a73d-182">流式自定义函数使你能够在一段时间内重复地将数据输出到单元格，而无需用户明确请求数据刷新。</span><span class="sxs-lookup"><span data-stu-id="0a73d-182">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="0a73d-183">以下示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="0a73d-183">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="0a73d-184">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="0a73d-184">Note the following about this code:</span></span>

- <span data-ttu-id="0a73d-185">Excel会自动使用 `setResult` 回调来显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="0a73d-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="0a73d-186">第二个输入参数 `handler` 在最终用户从自动完成菜单中选择函数时不在 Excel 中向他们显示。</span><span class="sxs-lookup"><span data-stu-id="0a73d-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="0a73d-187">`onCanceled` 回调定义函数被取消时执行的该函数。</span><span class="sxs-lookup"><span data-stu-id="0a73d-187">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="0a73d-188">必须为每个流式函数实施一个取消处理程序。</span><span class="sxs-lookup"><span data-stu-id="0a73d-188">You must implement a cancellation handler like this for any streamed function.</span></span> <span data-ttu-id="0a73d-189">有关详细信息，请参阅[取消函数](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="0a73d-189">For more information, see [Canceling a function](#canceling-a-function).</span></span> 

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

<span data-ttu-id="0a73d-190">指定 JSON 元数据文件中的流式函数元数据时，必须设置 `options` 对象内部的属性 `"cancelable": true` 和 `"stream": true`，如下面的示例中所示。</span><span class="sxs-lookup"><span data-stu-id="0a73d-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="0a73d-191">取消函数</span><span class="sxs-lookup"><span data-stu-id="0a73d-191">Canceling a function</span></span>

<span data-ttu-id="0a73d-192">在某些情况下，可能需要取消流式自定义函数的执行，以减少其带宽消耗、工作内存和 CPU 负载。</span><span class="sxs-lookup"><span data-stu-id="0a73d-192">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="0a73d-193">Excel 在下列情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="0a73d-193">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="0a73d-194">用户编辑或删除引用函数的单元格时。</span><span class="sxs-lookup"><span data-stu-id="0a73d-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="0a73d-195">函数的参数（输入）之一发生变化时。</span><span class="sxs-lookup"><span data-stu-id="0a73d-195">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="0a73d-196">在这种情况下，取消后触发新函数调用。</span><span class="sxs-lookup"><span data-stu-id="0a73d-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="0a73d-197">用户手动触发重新计算时。</span><span class="sxs-lookup"><span data-stu-id="0a73d-197">When the user triggers recalculation manually.</span></span> <span data-ttu-id="0a73d-198">在这种情况下，取消后触发新函数调用。</span><span class="sxs-lookup"><span data-stu-id="0a73d-198">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="0a73d-199">若要启用取消函数的功能，必须实施 JavaScript 函数中的取消处理程序，并在描述函数的 JSON 元数据中 `options` 对象内部指定属性 `"cancelable": true`。</span><span class="sxs-lookup"><span data-stu-id="0a73d-199">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="0a73d-200">本文的上一节中的代码示例提供了这些技术的示例。</span><span class="sxs-lookup"><span data-stu-id="0a73d-200">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="0a73d-201">保存和共享状态</span><span class="sxs-lookup"><span data-stu-id="0a73d-201">Saving and sharing state</span></span>

<span data-ttu-id="0a73d-202">自定义函数可以将数据保存在全局 JavaScript 变量中。</span><span class="sxs-lookup"><span data-stu-id="0a73d-202">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="0a73d-203">在后续调用中，自定义函数可以使用保存在这些变量中的值。</span><span class="sxs-lookup"><span data-stu-id="0a73d-203">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="0a73d-204">当用户将相同的自定义函数添加到多个单元格时，保存状态很有用，因为函数的所有实例都可以共享该状态。</span><span class="sxs-lookup"><span data-stu-id="0a73d-204">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="0a73d-205">例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。</span><span class="sxs-lookup"><span data-stu-id="0a73d-205">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="0a73d-206">下面的代码示例演示了全局保存状态的温度流式函数的实现。</span><span class="sxs-lookup"><span data-stu-id="0a73d-206">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="0a73d-207">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="0a73d-207">Note the following about this code:</span></span>

- <span data-ttu-id="0a73d-208">`refreshTemperature` 是一个流式函数，它会在每一秒内读取特定温度计的温度。</span><span class="sxs-lookup"><span data-stu-id="0a73d-208">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="0a73d-209">新的温度保存在 `savedTemperatures` 变量，但不直接更新单元格值。</span><span class="sxs-lookup"><span data-stu-id="0a73d-209">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="0a73d-210">它不应该直接从工作表单元格中调用，*因此未在 JSON 文件中注册*。</span><span class="sxs-lookup"><span data-stu-id="0a73d-210">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="0a73d-211">`streamTemperature` 每秒钟更新单元格中显示的温度值并使用 `savedTemperatures` 变量作为其数据源。</span><span class="sxs-lookup"><span data-stu-id="0a73d-211">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="0a73d-212">它必须在 JSON 文件中注册，并使用全大写字母命名：`STREAMTEMPERATURE`。</span><span class="sxs-lookup"><span data-stu-id="0a73d-212">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="0a73d-213">用户可以从 Excel UI 的多个单元格中调用 `streamTemperature`。</span><span class="sxs-lookup"><span data-stu-id="0a73d-213">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="0a73d-214">每次调用都从相同的 `savedTemperatures` 变量读取数据。</span><span class="sxs-lookup"><span data-stu-id="0a73d-214">Each call reads data from the same `savedTemperatures` variable.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="0a73d-215">使用数据区域</span><span class="sxs-lookup"><span data-stu-id="0a73d-215">Working with ranges of data</span></span>

<span data-ttu-id="0a73d-216">自定义的函数可接受一系列数据作为输入参数，或者可返回一系列数据。</span><span class="sxs-lookup"><span data-stu-id="0a73d-216">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="0a73d-217">在 JavaScript 中，数据区域表示为一个二维数组。</span><span class="sxs-lookup"><span data-stu-id="0a73d-217">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="0a73d-218">例如，假设函数从 Excel 中存储的一系列数字中返回第二个最大值。</span><span class="sxs-lookup"><span data-stu-id="0a73d-218">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="0a73d-219">下面的函数接受参数 `values`，其类型为 `Excel.CustomFunctionDimensionality.matrix`。</span><span class="sxs-lookup"><span data-stu-id="0a73d-219">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="0a73d-220">请注意，在此函数的 JSON 元数据中，可以将参数的 `type` 属性设置为 `matrix`。</span><span class="sxs-lookup"><span data-stu-id="0a73d-220">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="0a73d-221">处理错误</span><span class="sxs-lookup"><span data-stu-id="0a73d-221">Handling errors</span></span>

<span data-ttu-id="0a73d-222">在生成定义自定义函数的加载项时，请务必包含错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="0a73d-222">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="0a73d-223">自定义函数的错误处理与 [Excel JavaScript API 的错误处理大体相同](excel-add-ins-error-handling.md)。</span><span class="sxs-lookup"><span data-stu-id="0a73d-223">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="0a73d-224">在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="0a73d-224">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="0a73d-225">已知问题</span><span class="sxs-lookup"><span data-stu-id="0a73d-225">Known issues</span></span>

- <span data-ttu-id="0a73d-226">Excel 暂未使用帮助 URL 和参数说明。</span><span class="sxs-lookup"><span data-stu-id="0a73d-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="0a73d-227">自定义功能目前不适用于移动客户端的 Excel。</span><span class="sxs-lookup"><span data-stu-id="0a73d-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="0a73d-228">不支持可变函数（每当电子表格中不相关的数据更改时自动重新计算）。</span><span class="sxs-lookup"><span data-stu-id="0a73d-228">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="0a73d-229">尚未启用通过 Office 365 管理门户和 AppSource 进行的部署。</span><span class="sxs-lookup"><span data-stu-id="0a73d-229">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="0a73d-230">Excel Online 中的自定义功能，可能会在一段时间无活动后，在进程期间停止工作。</span><span class="sxs-lookup"><span data-stu-id="0a73d-230">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="0a73d-231">刷新浏览器页面 (F5) 并重新输入自定义函数以恢复该功能。</span><span class="sxs-lookup"><span data-stu-id="0a73d-231">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="0a73d-232">如果有多个加载项在 Excel for Windows 上运行，可能会在工作表单元格内看到 **#GETTING_DATA** 临时结果。</span><span class="sxs-lookup"><span data-stu-id="0a73d-232">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="0a73d-233">关闭 Excel 的所有窗口，并重新启动 Excel。</span><span class="sxs-lookup"><span data-stu-id="0a73d-233">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="0a73d-234">将来可能会提供专门用于自定义函数的调试工具。</span><span class="sxs-lookup"><span data-stu-id="0a73d-234">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="0a73d-235">同时，可以在 Excel Online 使用 F12 开发人员工具调试。</span><span class="sxs-lookup"><span data-stu-id="0a73d-235">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="0a73d-236">请参阅[自定义函数最佳做法](custom-functions-best-practices.md)中的详细信息。</span><span class="sxs-lookup"><span data-stu-id="0a73d-236">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="0a73d-237">更改日志</span><span class="sxs-lookup"><span data-stu-id="0a73d-237">Changelog</span></span>

- <span data-ttu-id="0a73d-238">**2017 年 11 月 7 日**：发布了\*自定义函数预览和示例</span><span class="sxs-lookup"><span data-stu-id="0a73d-238">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="0a73d-239">**2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题</span><span class="sxs-lookup"><span data-stu-id="0a73d-239">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="0a73d-240">**2017 年 11 月 28 日**：发布了\*对取消异步函数的支持（需要对流式函数进行相应更改）</span><span class="sxs-lookup"><span data-stu-id="0a73d-240">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="0a73d-241">**2018 年 5 月 7 日**：发布了\*对 Mac、Excel Online 和在进程中运行的同步函数的支持</span><span class="sxs-lookup"><span data-stu-id="0a73d-241">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="0a73d-242">**2018 年 9 月 20 日，** 发布了对自定义函数 JavaScript 运行时的支持。</span><span class="sxs-lookup"><span data-stu-id="0a73d-242">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="0a73d-243">有关详细信息，请参阅 [Excel 自定义函数的运行时](custom-functions-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="0a73d-243">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="0a73d-244">\* 至 Office 预览体验计划渠道</span><span class="sxs-lookup"><span data-stu-id="0a73d-244">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="0a73d-245">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0a73d-245">See also</span></span>

* [<span data-ttu-id="0a73d-246">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="0a73d-246">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0a73d-247">Excel 自定义函数运行时</span><span class="sxs-lookup"><span data-stu-id="0a73d-247">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0a73d-248">自定义函数最佳做法</span><span class="sxs-lookup"><span data-stu-id="0a73d-248">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0a73d-249">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="0a73d-249">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)