---
ms.date: 01/30/2019
description: 在 Excel 中使用 JavaScript 创建自定义函数。
title: 在 Excel 中创建自定义函数（预览）
localization_priority: Priority
ms.openlocfilehash: 3359962f3419f35692829444ab835d3f5cdc915a
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982025"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="845ab-103">在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="845ab-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="845ab-104">开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="845ab-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="845ab-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="845ab-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="845ab-106">本文介绍了如何在 Excel 中创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="845ab-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="845ab-107">下图演示最终用户将自定义函数插入到 Excel 工作表单元格的过程。</span><span class="sxs-lookup"><span data-stu-id="845ab-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="845ab-108">`CONTOSO.ADD42` 自定义函数旨在向用户指定作为函数输入参数的数字对添加 42。</span><span class="sxs-lookup"><span data-stu-id="845ab-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="845ab-109">以下代码定义 `ADD42` 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="845ab-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="845ab-110">本文后面的[已知问题](#known-issues)部分指定自定义函数的当前限制。</span><span class="sxs-lookup"><span data-stu-id="845ab-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="845ab-111">自定义函数加载项项目的组件</span><span class="sxs-lookup"><span data-stu-id="845ab-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="845ab-112">如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office)创建 Excel 自定义函数加载项项目，将在生成器创建的项目中看到以下文件：</span><span class="sxs-lookup"><span data-stu-id="845ab-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="845ab-113">文件</span><span class="sxs-lookup"><span data-stu-id="845ab-113">File</span></span> | <span data-ttu-id="845ab-114">文件格式</span><span class="sxs-lookup"><span data-stu-id="845ab-114">File format</span></span> | <span data-ttu-id="845ab-115">说明</span><span class="sxs-lookup"><span data-stu-id="845ab-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="845ab-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="845ab-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="845ab-117">或</span><span class="sxs-lookup"><span data-stu-id="845ab-117">or</span></span><br/><span data-ttu-id="845ab-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="845ab-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="845ab-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="845ab-119">JavaScript</span></span><br/><span data-ttu-id="845ab-120">或</span><span class="sxs-lookup"><span data-stu-id="845ab-120">or</span></span><br/><span data-ttu-id="845ab-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="845ab-121">TypeScript</span></span> | <span data-ttu-id="845ab-122">包含定义自定义函数的代码。</span><span class="sxs-lookup"><span data-stu-id="845ab-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="845ab-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="845ab-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="845ab-124">JSON</span><span class="sxs-lookup"><span data-stu-id="845ab-124">JSON</span></span> | <span data-ttu-id="845ab-125">包含描述自定义函数的元数据，使 Excel 能够注册自定义函数，并使其可供最终用户使用。</span><span class="sxs-lookup"><span data-stu-id="845ab-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="845ab-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="845ab-126">**./index.html**</span></span> | <span data-ttu-id="845ab-127">HTML</span><span class="sxs-lookup"><span data-stu-id="845ab-127">HTML</span></span> | <span data-ttu-id="845ab-128">提供对定义自定义函数的 JavaScript 文件的&lt;脚本&gt;引用。</span><span class="sxs-lookup"><span data-stu-id="845ab-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="845ab-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="845ab-129">**./manifest.xml**</span></span> | <span data-ttu-id="845ab-130">XML</span><span class="sxs-lookup"><span data-stu-id="845ab-130">XML</span></span> | <span data-ttu-id="845ab-131">指定加载项中所有自定义函数的命名空间以及此表中前面列出的 JavaScript、JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="845ab-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="845ab-132">下列部分将提供有关这些文件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="845ab-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="845ab-133">脚本文件</span><span class="sxs-lookup"><span data-stu-id="845ab-133">Script file</span></span>

<span data-ttu-id="845ab-134">脚本文件（Yo Office 生成器创建的项目中的 **./src/customfunctions.js** 或 **./src/customfunctions.ts**）包含定义自定义函数并将自定义函数名称映射到 [JSON 元数据文件](#json-metadata-file)中的对象的代码。</span><span class="sxs-lookup"><span data-stu-id="845ab-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="845ab-135">例如，以下代码定义自定义函数 `add` 和 `increment`，然后指定这两个函数的关联信息。</span><span class="sxs-lookup"><span data-stu-id="845ab-135">For example, the following code defines the custom functions `add` and `increment` and then specifies association information for both functions.</span></span> <span data-ttu-id="845ab-136">将 `add` 函数关联到 JSON 元数据文件中的对象，其中 `id` 属性的值为 **ADD**，将 `increment` 函数关联到元数据文件中的对象，其中 `id` 属性的值为 **INCREMENT**。</span><span class="sxs-lookup"><span data-stu-id="845ab-136">The `add` function is associated with the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is associated with the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="845ab-137">有关将脚本文件中的函数名称关联到 JSON 元数据文件中的对象的更多信息，请参阅[自定义函数最佳实践](custom-functions-best-practices.md#associating-function-names-with-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="845ab-137">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about associating function names in the script file to objects in the JSON metadata file.</span></span>

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a><span data-ttu-id="845ab-138">JSON 元数据文件</span><span class="sxs-lookup"><span data-stu-id="845ab-138">JSON metadata file</span></span>

<span data-ttu-id="845ab-139">自定义函数元数据文件（Yo Office 生成器创建的项目中的 **./config/customfunctions.json**）提供 Excel 注册自定义函数并使其可供最终用户使用所需的信息。</span><span class="sxs-lookup"><span data-stu-id="845ab-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="845ab-140">自定义函数在用户首次运行加载项时注册。</span><span class="sxs-lookup"><span data-stu-id="845ab-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="845ab-141">之后，它们可在所有工作簿（即，不仅仅是在加载项初始运行的工作簿）中供同一用户使用。</span><span class="sxs-lookup"><span data-stu-id="845ab-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="845ab-142">托管 JSON 文件的服务器上的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS)，以便自定义函数在 Excel Online 中正常工作。</span><span class="sxs-lookup"><span data-stu-id="845ab-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="845ab-143">**customfunctions.json** 中的以下代码指定上述 `add` 函数和 `increment` 函数的元数据。</span><span class="sxs-lookup"><span data-stu-id="845ab-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="845ab-144">此代码示例后面的表提供了有关此 JSON 对象中各个属性的详细信息。</span><span class="sxs-lookup"><span data-stu-id="845ab-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="845ab-145">有关在 JSON 元数据文件中指定 `id` 和 `name` 属性值的详细信息，请参阅[自定义函数最佳实践](custom-functions-best-practices.md#associating-function-names-with-json-metadata)。</span><span class="sxs-lookup"><span data-stu-id="845ab-145">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="845ab-146">下表列出了 JSON 元数据文件中的常见属性。</span><span class="sxs-lookup"><span data-stu-id="845ab-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="845ab-147">有关 JSON 元数据文件的更多详细信息，请参阅[自定义函数元数据](custom-functions-json.md)。</span><span class="sxs-lookup"><span data-stu-id="845ab-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="845ab-148">属性</span><span class="sxs-lookup"><span data-stu-id="845ab-148">Property</span></span>  | <span data-ttu-id="845ab-149">说明</span><span class="sxs-lookup"><span data-stu-id="845ab-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="845ab-150">函数的唯一 ID。</span><span class="sxs-lookup"><span data-stu-id="845ab-150">A unique ID for the function.</span></span> <span data-ttu-id="845ab-151">此 ID 只能包含字母数字字符和句点，设置后不应更改。</span><span class="sxs-lookup"><span data-stu-id="845ab-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="845ab-152">最终用户在 Excel 中看到的函数名称。</span><span class="sxs-lookup"><span data-stu-id="845ab-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="845ab-153">在 Excel 中，此函数名称将以 [XML 清单文件](#manifest-file)中指定的自定义函数命名空间作为前缀。</span><span class="sxs-lookup"><span data-stu-id="845ab-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="845ab-154">当用户请求帮助时显示的页面的 URL。</span><span class="sxs-lookup"><span data-stu-id="845ab-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="845ab-155">说明函数的功能。</span><span class="sxs-lookup"><span data-stu-id="845ab-155">Describes what the function does.</span></span> <span data-ttu-id="845ab-156">当函数是 Excel 自动完成菜单中的选中项时，此值将作为工具提示显示。</span><span class="sxs-lookup"><span data-stu-id="845ab-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="845ab-157">定义函数返回的信息类型的对象。</span><span class="sxs-lookup"><span data-stu-id="845ab-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="845ab-158">有关此对象的详细信息，请参阅[结果](custom-functions-json.md#result)。</span><span class="sxs-lookup"><span data-stu-id="845ab-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="845ab-159">定义函数的输入参数的数组。</span><span class="sxs-lookup"><span data-stu-id="845ab-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="845ab-160">有关此对象的详细信息，请参阅[参数](custom-functions-json.md#parameters)。</span><span class="sxs-lookup"><span data-stu-id="845ab-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="845ab-161">使用户能够自定义 Excel 执行函数的方式和时间。</span><span class="sxs-lookup"><span data-stu-id="845ab-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="845ab-162">有关如何使用此属性的详细信息，请参阅[流式处理函数](#streaming-functions)和[取消函数](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="845ab-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function).</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="845ab-163">清单文件</span><span class="sxs-lookup"><span data-stu-id="845ab-163">Manifest file</span></span>

<span data-ttu-id="845ab-164">定义自定义函数的加载项的 XML 清单文件（Yo Office 生成器创建的项目中的 **./manifest.xml**）指定加载项中所有自定义函数的命名空间以及 JavaScript、JSON 和 HTML 文件的位置。</span><span class="sxs-lookup"><span data-stu-id="845ab-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="845ab-165">下面的 XML 标记显示了 `<ExtensionPoint>` 和 `<Resources>` 元素的一个示例，必须在加载项清单中包含这些元素才能启用自定义函数。</span><span class="sxs-lookup"><span data-stu-id="845ab-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://localhost:8081/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://localhost:8081/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://localhost:8081/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="845ab-166">Excel 中的函数在前面追加 XML 清单文件中指定的命名空间作为前缀。</span><span class="sxs-lookup"><span data-stu-id="845ab-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="845ab-167">函数的命名空间在函数名称之前，并用句点分隔。</span><span class="sxs-lookup"><span data-stu-id="845ab-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="845ab-168">例如，若要在 Excel 工作表的单元格中调用函数 `ADD42`，需输入 `=CONTOSO.ADD42`，因为 `CONTOSO` 是命名空间，`ADD42` 是 JSON 文件中指定的函数的名称。</span><span class="sxs-lookup"><span data-stu-id="845ab-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="845ab-169">命名空间旨在作为公司或加载项的标识符使用。</span><span class="sxs-lookup"><span data-stu-id="845ab-169">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="845ab-170">命名空间只能包含字母数字字符和句点。</span><span class="sxs-lookup"><span data-stu-id="845ab-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="845ab-171">从外部源返回数据的函数</span><span class="sxs-lookup"><span data-stu-id="845ab-171">Functions that return data from external sources</span></span>

<span data-ttu-id="845ab-172">如果自定义函数从外部源（如 Web）检索数据，则必须：</span><span class="sxs-lookup"><span data-stu-id="845ab-172">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="845ab-173">将 JavaScript Promise 返回到 Excel。</span><span class="sxs-lookup"><span data-stu-id="845ab-173">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="845ab-174">使用回调函数解析带有最终值的 Promise。</span><span class="sxs-lookup"><span data-stu-id="845ab-174">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="845ab-175">在 Excel 等待最终结果时，自定义函数会在单元格中显示一个 `#GETTING_DATA` 临时结果。</span><span class="sxs-lookup"><span data-stu-id="845ab-175">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="845ab-176">在等待结果时，用户可以与工作表的其余部分正常交互。</span><span class="sxs-lookup"><span data-stu-id="845ab-176">Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="845ab-177">在下面的代码示例中，`getTemperature()` 自定义函数检索温度计的当前温度。</span><span class="sxs-lookup"><span data-stu-id="845ab-177">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="845ab-178">注意，`sendWebRequest` 是一个假设函数（此处未指定），它使用 [XHR](custom-functions-runtime.md#xhr-example) 调用温度 Web 服务。</span><span class="sxs-lookup"><span data-stu-id="845ab-178">Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="845ab-179">流式处理函数</span><span class="sxs-lookup"><span data-stu-id="845ab-179">Streaming functions</span></span>

<span data-ttu-id="845ab-180">流式处理自定义函数使用户能够在不需要用户显式请求数据刷新的情况下，随着时间的推移向单元格重复输出数据。</span><span class="sxs-lookup"><span data-stu-id="845ab-180">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="845ab-181">下面的代码示例是一个自定义函数，它每秒向结果添加一个数字。</span><span class="sxs-lookup"><span data-stu-id="845ab-181">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="845ab-182">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="845ab-182">Note the following about this code:</span></span>

- <span data-ttu-id="845ab-183">Excel 使用 `setResult` 回调自动显示每个新值。</span><span class="sxs-lookup"><span data-stu-id="845ab-183">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="845ab-184">当最终用户从自动完成菜单中选择函数时，不会在 Excel 中向其显示第二个输入参数 `handler`。</span><span class="sxs-lookup"><span data-stu-id="845ab-184">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="845ab-185">`onCanceled` 回调定义取消函数时执行的函数。</span><span class="sxs-lookup"><span data-stu-id="845ab-185">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="845ab-186">对于任何流式处理函数，都必须实现此类取消处理程序。</span><span class="sxs-lookup"><span data-stu-id="845ab-186">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="845ab-187">有关详细信息，请参阅[取消函数](#canceling-a-function)。</span><span class="sxs-lookup"><span data-stu-id="845ab-187">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="845ab-188">在 JSON 元数据文件中为流式处理函数指定元数据时，必须在 `options` 对象中设置属性 `"cancelable": true` 和 `"stream": true`，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="845ab-188">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="845ab-189">取消函数</span><span class="sxs-lookup"><span data-stu-id="845ab-189">Canceling a function</span></span>

<span data-ttu-id="845ab-190">在某些情况下，可能需要取消执行流式处理自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。</span><span class="sxs-lookup"><span data-stu-id="845ab-190">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="845ab-191">Excel 会在以下情况下取消函数的执行：</span><span class="sxs-lookup"><span data-stu-id="845ab-191">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="845ab-192">用户编辑或删除引用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="845ab-192">When the user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="845ab-193">函数的参数（输入）之一发生变化。</span><span class="sxs-lookup"><span data-stu-id="845ab-193">When one of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="845ab-194">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="845ab-194">In this case, a new function call is triggered following the cancellation.</span></span>

- <span data-ttu-id="845ab-195">用户手动触发重新计算。</span><span class="sxs-lookup"><span data-stu-id="845ab-195">When the user triggers recalculation manually.</span></span> <span data-ttu-id="845ab-196">在这种情况下，取消之后还会触发新的函数调用。</span><span class="sxs-lookup"><span data-stu-id="845ab-196">In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="845ab-197">为了能够取消函数，必须在 JavaScript 函数中实现一个取消处理程序，并在说明函数的 JSON 元数据中指定 `options` 对象中的属性 `"cancelable": true`。</span><span class="sxs-lookup"><span data-stu-id="845ab-197">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="845ab-198">本文前一部分中的代码示例提供了这些方法的示例。</span><span class="sxs-lookup"><span data-stu-id="845ab-198">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="845ab-199">保存和共享状态</span><span class="sxs-lookup"><span data-stu-id="845ab-199">Saving and sharing state</span></span>

<span data-ttu-id="845ab-200">自定义函数可以将数据保存在全局 JavaScript 变量中，可用于后续调用。</span><span class="sxs-lookup"><span data-stu-id="845ab-200">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="845ab-201">当用户从多个单元格调用同一个自定义函数时，保存状态非常有用，因为函数的所有实例都可以访问该状态。</span><span class="sxs-lookup"><span data-stu-id="845ab-201">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="845ab-202">例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。</span><span class="sxs-lookup"><span data-stu-id="845ab-202">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="845ab-203">下面的代码示例演示温度流式处理函数的实现过程，该函数在全局范围内保存状态。</span><span class="sxs-lookup"><span data-stu-id="845ab-203">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="845ab-204">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="845ab-204">Note the following about this code:</span></span>

- <span data-ttu-id="845ab-205">`streamTemperature` 函数每秒更新单元格中显示的温度值，并使用 `savedTemperatures` 变量作为其数据源。</span><span class="sxs-lookup"><span data-stu-id="845ab-205">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="845ab-206">因为 `streamTemperature` 是一个流式处理函数，它将实现一个取消处理程序，当函数被取消时该处理程序将运行。</span><span class="sxs-lookup"><span data-stu-id="845ab-206">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="845ab-207">如果用户从 Excel 中的多个单元格调用 `streamTemperature` 函数，则 `streamTemperature` 函数在每次运行时都会从相同的 `savedTemperatures` 变量读取数据。</span><span class="sxs-lookup"><span data-stu-id="845ab-207">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="845ab-208">`refreshTemperature` 函数每秒读取特定温度计的温度，并将结果存储在 `savedTemperatures` 变量中。</span><span class="sxs-lookup"><span data-stu-id="845ab-208">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="845ab-209">因为 `refreshTemperature` 函数不在 Excel 中向最终用户显示，所以不需要在 JSON 文件中注册。</span><span class="sxs-lookup"><span data-stu-id="845ab-209">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="845ab-210">使用数据区域</span><span class="sxs-lookup"><span data-stu-id="845ab-210">Working with ranges of data</span></span>

<span data-ttu-id="845ab-211">自定义函数可以接受数据区域作为输入参数，也可以返回数据区域。</span><span class="sxs-lookup"><span data-stu-id="845ab-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="845ab-212">在 JavaScript，数据区域表示为一个二维数组。</span><span class="sxs-lookup"><span data-stu-id="845ab-212">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="845ab-213">例如，假设函数从 Excel 中存储的数字区域返回第二个最高值。</span><span class="sxs-lookup"><span data-stu-id="845ab-213">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="845ab-214">下面的函数接受参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 类型。</span><span class="sxs-lookup"><span data-stu-id="845ab-214">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="845ab-215">请注意，在此函数的 JSON 元数据中，将参数的 `type` 属性设置为 `matrix`。</span><span class="sxs-lookup"><span data-stu-id="845ab-215">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="845ab-216">确定调用自定义函数的单元格</span><span class="sxs-lookup"><span data-stu-id="845ab-216">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="845ab-217">在某些情况下，需要获取调用自定义函数的单元格地址。</span><span class="sxs-lookup"><span data-stu-id="845ab-217">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="845ab-218">这在以下类型的应用场景中非常有用：</span><span class="sxs-lookup"><span data-stu-id="845ab-218">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="845ab-219">设置区域格式：将单元格地址用作键，以便将信息存储到 [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data) 中。</span><span class="sxs-lookup"><span data-stu-id="845ab-219">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](https://docs.microsoft.com/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="845ab-220">然后，使用 Excel 中的 [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) 从 `AsyncStorage` 加载该键。</span><span class="sxs-lookup"><span data-stu-id="845ab-220">Then, use [onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="845ab-221">显示缓存值：如果脱机使用函数，将显示 `AsyncStorage` 中使用 `onCalculated` 存储的缓存值。</span><span class="sxs-lookup"><span data-stu-id="845ab-221">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="845ab-222">协调：使用单元格地址发现原始单元格，以帮助你在处理时进行协调。</span><span class="sxs-lookup"><span data-stu-id="845ab-222">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="845ab-223">仅当函数 JSON 元数据文件中的 `requiresAddress` 被标记为 `true` 时，才会公开与单元格地址相关的信息。</span><span class="sxs-lookup"><span data-stu-id="845ab-223">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="845ab-224">以下示例诠释了此情况：</span><span class="sxs-lookup"><span data-stu-id="845ab-224">The following sample gives an example of this:</span></span>

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

<span data-ttu-id="845ab-225">此外，需要在脚本文件 (**./src/customfunctions.js** 或 **./src/customfunctions.ts**) 中添加 `getAddress` 函数，以查找单元格地址。</span><span class="sxs-lookup"><span data-stu-id="845ab-225">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="845ab-226">此函数可能会使用参数，如以下示例 `parameter1` 所示。</span><span class="sxs-lookup"><span data-stu-id="845ab-226">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="845ab-227">最后一个参数始终为 `invocationContext`，该对象包含 JSON 元数据文件中的 `requiresAddress` 被标记为 `true` 时 Excel 传递的单元格位置。</span><span class="sxs-lookup"><span data-stu-id="845ab-227">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="845ab-228">默认情况下，从 `getAddress` 函数返回的值遵循以下格式：`SheetName!CellNumber`。</span><span class="sxs-lookup"><span data-stu-id="845ab-228">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="845ab-229">例如，如果名为“Expense”的工作表中的 B2 单元格调用了函数，则返回的值为 `Expenses!B2`。</span><span class="sxs-lookup"><span data-stu-id="845ab-229">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="handling-errors"></a><span data-ttu-id="845ab-230">处理错误</span><span class="sxs-lookup"><span data-stu-id="845ab-230">Handling errors</span></span>

<span data-ttu-id="845ab-231">在生成定义自定义函数的加载项时，请务必加入错误处理逻辑，以便解决运行时错误。</span><span class="sxs-lookup"><span data-stu-id="845ab-231">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="845ab-232">自定义函数的错误处理与 [Excel JavaScript API 的错误处理](excel-add-ins-error-handling.md)大致相同。</span><span class="sxs-lookup"><span data-stu-id="845ab-232">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="845ab-233">在以下代码示例中，`.catch` 将处理之前发生在代码中的任何错误。</span><span class="sxs-lookup"><span data-stu-id="845ab-233">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="845ab-234">已知问题</span><span class="sxs-lookup"><span data-stu-id="845ab-234">Known issues</span></span>

<span data-ttu-id="845ab-235">在 [Excel 自定义功能 GitHub 存储库](https://github.com/OfficeDev/Excel-Custom-Functions/issues)上查看已知问题。</span><span class="sxs-lookup"><span data-stu-id="845ab-235">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span> 

## <a name="see-also"></a><span data-ttu-id="845ab-236">另请参阅</span><span class="sxs-lookup"><span data-stu-id="845ab-236">See also</span></span>

* [<span data-ttu-id="845ab-237">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="845ab-237">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="845ab-238">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="845ab-238">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="845ab-239">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="845ab-239">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="845ab-240">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="845ab-240">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="845ab-241">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="845ab-241">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)

