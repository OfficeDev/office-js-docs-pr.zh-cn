---
ms.date: 09/27/2018
description: 在 Excel 中使用 JavaScript 创建自定义的函数。
title: 在 Excel 中创建自定义函数（预览）
ms.openlocfilehash: c8a2d8755a68530ecf8743c4a8ab65a4bed5b849
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348154"
---
# <a name="create-custom-functions-in-excel-preview"></a>在 Excel 中创建自定义函数（预览）

自定义函数使开发人员可以通过在 JavaScript 中定义这些函数作为外接程序的一部分，将新函数添加到 Excel。然后，用户可以像使用 Excel 中的其他本机函数（例如 `SUM()`）一样访问自定义函数。本文介绍了如何在 Excel 中创建自定义函数。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

下图显示了最终用户将插入 Excel 工作表的单元格的自定义的函数。 `CONTOSO.ADD42` 自定义函数用于将 42 添加到用户指定为函数的输入参数的一对数字中。

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

下面的代码定义 `ADD42` 自定义函数。

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> 本文后面的 [已知问题](#known-issues) 一节指定了自定义函数的当前限制。

## <a name="components-of-a-custom-functions-add-in-project"></a>自定义函数外接程序项目的组件

如果使用 [Yo Office 生成器](https://github.com/OfficeDev/generator-office) 创建 Excel 自定义函数外接程序项目，将在项目中看到生成器创建的以下文件：

| 文件 | 文件格式 | 说明 |
|------|-------------|-------------|
| **./src/customfunctions.js**<br/>或<br/>**./src/customfunctions.ts** | JavaScript<br/>或<br/>TypeScript | 包含定义自定义函数的代码。 |
| **./config/customfunctions.json** | JSON | 包含描述自定义函数的元数据，并使 Excel 能够注册自定义函数且使其可供最终用户使用。 |
| **./index.html** | HTML | 提供 &lt;脚本&gt; 定义自定义函数的 JavaScript 文件的引用。 |
| **./manifest.xml** | XML | 此表中指定外接程序中所有自定义函数的命名空间，以及前面列出的JavaScript、JSON 和 HTML 文件的位置。 |

下面的章节将提供有关这些文件的更多信息。

### <a name="script-file"></a>脚本文件 

脚本文件 （Yo Office 生成器创建项目中的 **./src/customfunctions.js** 或 **./src/customfunctions.ts**）包含定义自定义函数的代码，该代码还将自定义函数的名称映射到 [JSON 元数据文件](#json-metadata-file)中的对象。 

例如，下面的代码示例定义自定义函数 `add` 和 `increment`，然后指定这两个函数的映射信息。  `add` 函数被映射到 JSON 元数据文件中的对象，其中 `id` 属性的值是 **ADD**，而 `increment` 函数被映射到元数据文件中的对象，其中 `id` 属性的值是 **INCREMENT**。 有关将脚本文件中的函数名称映射到 JSON 元数据文件中对象的详细信息，请参阅[自定义函数的最佳做法](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)。

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

### <a name="json-metadata-file"></a>JSON 元数据文件 

自定义函数元数据文件（Yo Office 生成器所创建项目中的 **./config/customfunctions.json**）提供 Excel 注册自定义函数需要的信息，并将其提供给最终用户。 自定义函数是在用户第一次运行加载项时注册的。 之后，所有工作簿中的同一用户都可以使用它们 （即不仅在加载项最初运行的工作簿中。）

> [!TIP]
> JSON 文件的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) 才能使自定义函数在 Excel Online 中正常工作。

下面在 **customfunctions.json** 中的代码指定前面所述 `add` 函数和 `increment` 函数的元数据。 此代码后示例的表格提供了有关此 JSON 对象中单独属性的详细信息。 有关指定 JSON 元数据文件中 `id`  和 `name` 属性值的详细信息，请参阅[自定义函数的最佳做法](custom-functions-best-practices.md#mapping-function-names-to-json-metadata)。

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

下表列出了通常存在于 JSON 元数据文件的属性。 有关 JSON 元数据文件的详细信息，请参阅[自定义函数元数据](custom-functions-json.md)。

| 属性  | 说明 |
|---------|---------|
| `id` | 函数的唯一 ID。 设置之后，不应更改此 ID。 |
| `name` | 最终用户在 Excel 中看到的函数名称。 在 Excel 中，此函数名称将以[  XML ](#manifest-file)清单文件中指定的自定义函数命名空间为前缀。 |
| `helpUrl` | 用户请求帮助时显示的页面的 Url。 |
| `description` | 介绍函数的用途。 当函数是 Excel 中自动完成菜单中的选定项时，此值将显示为工具提示。 |
| `result`  | 定义函数返回的信息类型的对象。 `type` 子属性的值可以是**字符串**、**数字**或**布尔值**。 `dimensionality` 子属性的值可以是**scalar** 或 **matrix**（指定 `type` 值的二维数组）。 |
| `parameters` | 定义函数的输入参数的数组。 在 Excel intelliSense 中出现的 `name` 和 `description` 子属性。 `type` 子属性的值可以是**字符串**、**数字**或**布尔值**。 `dimensionality` 子属性的值可以是**scalar** 或 **matrix**（指定 `type` 值的二维数组）。 |
| `options` | 使你可以自定义 Excel 执行函数的方式和时间等的某些方面。 有关如何使用此属性的详细信息，请参阅本文后面的[流式函数](#streamed-functions)和[取消函数](#canceling-a-function)。 |

### <a name="manifest-file"></a>清单文件

定义自定义函数（Yo Office 生成器所创建项目中的 **./manifest.xml**）的加载项的 XML 清单文件指定加载项和 JavaScript、JSON 和 HTML 文件的位置中的所有自定义函数的命名空间。 以下 XML 标记显示 了`<ExtensionPoint>` 和 `<Resources>` 元素的示例，必须在加载项的清单中包含该实例，以启用自定义函数。  

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
> Excel 中的函数以 XML 清单文件中指定的命名空间为前缀。 函数的命名空间出现在函数名之前并由句点分隔。 例如，若要`ADD42`在 Excel 工作表的单元格中调用函数，则需要键入 `=CONTOSO.ADD42`，因为 CONTOSO 是命名空间并且`ADD42`是 JSON 文件中所指定函数的名称。 命名空间旨在用作公司或加载项的标识符。 

## <a name="functions-that-return-data-from-external-sources"></a>从外部源返回数据的函数

从外部源返回数据的函数

1. 将 JavaScript 承诺返回到 Excel。

2. 使用回调函数，用最终值解析承诺。

自定义函数在单元格中显示 `#GETTING_DATA` 临时结果，而 Excel 等待最终结果。 用户可以在等待结果时与电子表格的其余部分进行正常交互。

在下面的代码示例中，`getTemperature()` 自定义函数检索温度计当前温度。 注意，`sendWebRequest` 是一个假设的函数（这里没有指定），它使用 [XHR](custom-functions-runtime.md#xhr)  来调用温度 Web 服务。

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>流式函数

流式自定义函数使你能够在一段时间内重复地将数据输出到单元格，而无需用户明确请求数据刷新。 以下示例是一个自定义函数，它每秒向结果添加一个数字。 关于此代码，请注意以下几点：

- Excel会自动使用 `setResult` 回调来显示每个新值。

- 第二个输入参数 `handler` 在最终用户从自动完成菜单中选择函数时不在 Excel 中向他们显示。

-  `onCanceled` 回调定义函数被取消时执行的该函数。 必须为每个流式函数实施一个取消处理程序。 有关详细信息，请参阅[取消函数](#canceling-a-function)。 

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

指定 JSON 元数据文件中的流式函数元数据时，必须设置 `options` 对象内部的属性 `"cancelable": true` 和 `"stream": true`，如下面的示例中所示。

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

## <a name="canceling-a-function"></a>取消函数

在某些情况下，可能需要取消流式自定义函数的执行，以减少其带宽消耗、工作内存和 CPU 负载。 Excel 在下列情况下取消函数的执行：

- 用户编辑或删除引用函数的单元格时。

- 函数的参数（输入）之一发生变化时。 在这种情况下，取消后触发新函数调用。

- 用户手动触发重新计算时。 在这种情况下，取消后触发新函数调用。

若要启用取消函数的功能，必须实施 JavaScript 函数中的取消处理程序，并在描述函数的 JSON 元数据中 `options` 对象内部指定属性 `"cancelable": true`。 本文的上一节中的代码示例提供了这些技术的示例。

## <a name="saving-and-sharing-state"></a>保存和共享状态

自定义函数可以将数据保存在全局 JavaScript 变量中。 在后续调用中，自定义函数可以使用保存在这些变量中的值。 当用户将相同的自定义函数添加到多个单元格时，保存状态很有用，因为函数的所有实例都可以共享该状态。 例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。

下面的代码示例演示了全局保存状态的温度流式函数的实现。 关于此代码，请注意以下几点：

- `refreshTemperature` 是一个流式函数，它会在每一秒内读取特定温度计的温度。 新的温度保存在 `savedTemperatures` 变量，但不直接更新单元格值。 它不应该直接从工作表单元格中调用，*所以它没有在 JSON 文件中注册*。

- `streamTemperature` 每秒钟更新单元格中显示的温度值并使用 `savedTemperatures` 变量作为其数据源。 它必须在 JSON 文件中注册，并使用全大写字母命名：`STREAMTEMPERATURE`。

- 用户可以从 Excel UI 的多个单元格中调用 `streamTemperature`。 每次调用都从相同的 `savedTemperatures` 变量读取数据。

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

## <a name="working-with-ranges-of-data"></a>使用数据区域

自定义的函数可接受一系列数据作为输入参数，或者可返回一系列数据。 JavaScript 中，数据区域表示为一个二维数组。

例如，假设函数从 Excel 中存储的一系列数字中返回第二个最大值。 下面的函数接受参数 `values`，其类型为 `Excel.CustomFunctionDimensionality.matrix`。 请注意，在此函数的 JSON 元数据中，可以将参数的 `type` 属性设置为 `matrix`。

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

## <a name="handling-errors"></a>处理错误

在生成定义自定义函数的加载项时，请务必包含错误处理逻辑，以便解决运行时错误。 自定义函数的错误处理与 [Excel JavaScript API 的错误处理大体相同](excel-add-ins-error-handling.md)。 在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。

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

## <a name="known-issues"></a>已知问题

- Excel 暂未使用帮助 URL 和参数说明。
- 自定义功能目前不适用于移动客户端的 Excel。
- 不支持可变函数（每当电子表格中不相关的数据更改时自动重新计算）。
- 尚未启用通过 Office 365 管理门户和 AppSource 进行的部署。
- Excel Online中的自定义功能，可能会在一段时间无活动后，在进程期间停止工作。 刷新浏览器页面 (F5) 并重新输入自定义函数以恢复该功能。
- 如果有多个加载项在 Excel for Windows 上运行，可能会在工作表单元格内看到 **#GETTING_DATA** 临时结果。 关闭 Excel 的所有窗口，并重新启动 Excel。
- 将来可能会提供专门用于自定义函数的调试工具。 同时，可以在 Excel Online 使用 F12 开发人员工具调试。 请参阅[自定义函数最佳做法](custom-functions-best-practices.md)中的详细信息。

## <a name="changelog"></a>更改日志

- **2017 年 11 月 7 日**：发布了*自定义函数预览和示例
- **2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题
- **2017 年 11 月 28 日**：发布了*对取消异步函数的支持（需要对流式函数进行相应更改）
- **2018 年 5 月 7 日**：发布了*对 Mac、Excel Online 和在进程中运行的同步函数的支持
- **2018 年 9 月 20 日，** 发布了对自定义函数 JavaScript 运行时的支持。 有关详细信息，请参阅 [Excel 自定义函数的运行时](custom-functions-runtime.md)。

\* 至 Office 预览体验计划渠道

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数运行时](custom-functions-runtime.md)
* [自定义函数最佳做法](custom-functions-best-practices.md)
