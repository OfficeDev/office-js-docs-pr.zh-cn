---
ms.date: 09/20/2018
description: 在 Excel 中使用 JavaScript 创建自定义的函数。
title: 在 Excel 中创建自定义函数（预览）
ms.openlocfilehash: b214329fe50955d0f39d50f674152f475ca24b4d
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005041"
---
# <a name="create-custom-functions-in-excel-preview"></a>在 Excel 中创建自定义函数（预览）

自定义函数使开发人员可以通过在 JavaScript 中定义这些函数作为外接程序的一部分，将新函数添加到 Excel。 然后，用户可以像使用 Excel 中的其他本机函数（例如 `SUM()`）一样访问自定义函数。 本文介绍了如何在 Excel 中创建自定义函数。

下图显示了最终用户将插入 Excel 工作表的单元格的自定义的函数。 自定义函数用于将 42 添加到用户指定为函数的输入参数的一对数字中。`CONTOSO.ADD42`

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

下面的代码定义 `ADD42` 自定义函数。

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

自定义函数现可在 Windows、Mac 和 Excel Online 的开发人员预览版中使用。 若要试用它们，请完成以下步骤：

1. 安装 Office（Windows 上的内部版本 10827 或 Mac 上的内部版本 13.329）并加入 [Office 预览体验计划](https://products.office.com/office-insider) 程序。 您必须加入 Office 预览体验计划才能访问自定义的函数；目前，除非您是 office 预览体验计划程序的成员，否则在所有 office 生成中都会禁用自定义函数。

2. 使用 [Yo Office](https://github.com/OfficeDev/generator-office) 创建 Excel 自定义函数外接程序项目，然后按照 [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) 中使用项目的说明。

3. 在 Excel 工作表的任意单元格键入 `=CONTOSO.ADD42(1,2)`，并按 **Enter** 运行自定义的函数。

> [!NOTE]
> 本文后面的 [已知问题](#known-issues) 一节指定了自定义函数的当前限制。

## <a name="learn-the-basics"></a>学习基础知识

在您使用 [Yo Office](https://github.com/OfficeDev/generator-office)创建的自定义函数项目中，您将看到以下文件：

| 文件 | 文件格式 | 说明 |
|------|-------------|-------------|
| **./src/customfunctions.js** | JavaScript | 包含定义自定义函数的代码。 |
| **./config/customfunctions.json** | JSON | 包含描述自定义函数的元数据，并使 Excel 能够注册自定义函数以使其可供最终用户使用。 |
| **./index.html** | HTML | 提供 &lt;脚本&gt; 定义自定义函数的 JavaScript 文件的引用。 |
| **./manifest.xml** | XML | 此表中指定外接程序中所有自定义函数的命名空间，以及前面列出的JavaScript、 JSON 和 HTML 文件的位置。 |

### <a name="manifest-file-manifestxml"></a>清单文件（./manifest.xml）

定义自定义函数的外接程序的 XML 清单文件指定外接程序和 JavaScript、 JSON 和 HTML 文件的位置中的所有自定义函数的命名空间。 以下 XML 标记显示 了`<ExtensionPoint>` 和 `<Resources>` 元素的示例，您必须在外接程序的清单中包含该实例，才能使 Excel 能够运行自定义函数。  

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
> Excel 中的函数由 XML 清单文件中指定的命名空间预置。 函数的命名空间出现在函数名之前并由句点分隔。 例如，若要`ADD42()`在 Excel 工作表的单元格中调用函数，则需要键入 `=CONTOSO.ADD42`，因为 CONTOSO 是命名空间并且`ADD42`是 JSON 文件中指定的函数的名称。 该命名空间旨在用作公司或加载项的标识符。 

### <a name="json-file-configcustomfunctionsjson"></a>JSON 文件 (./config/customfunctions.json)

自定义函数元数据文件提供 Excel 要求注册自定义函数并使其可供最终用户使用的信息。 自定义函数是在用户第一次运行加载项时注册的。 之后，所有工作簿中的同一用户都可以使用它们 （即，不仅在加载项最初运行的工作簿中。）

> [!TIP]
> 您的 JSON 文件的服务器设置必须启用 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) 才能使自定义函数在 Excel Online 中正常工作。

下面的代码 **customfunctions.json** 指定的元数据 `ADD42` 是以前本文中所述的函数。 此元数据定义的函数名称、说明、返回值、输入的参数等。 下表提供了此代码示例有关的 JSON 对象中的各个属性的详细信息。

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

下表列出了通常存在于 JSON 元数据文件的属性。 有关 JSON 元数据文件的详细信息，包括上一示例中未使用的选项，请参阅 [自定义函数元数据](custom-functions-json.md)。

| 属性  | 说明 |
|---------|---------|
| `id` | 函数的唯一 ID。 设置之后，不应更改此 ID。 |
| `name` | 当用户在单元格中键入公式时，自动完成菜单中显示函数的名称。 在自动完成菜单中，此值将由自定义函数的命名空间中的 XML 清单文件指定作为前缀。 |
| `helpUrl` | 当用户请求帮助显示的页面的 Url。 |
| `description` | 介绍函数的用途。 当函数是 Excel 中自动完成菜单中的选定项时，此值将显示为工具提示。 |
| `result`  | 定义函数返回的信息类型的对象。 子属性的值可以是 **字符串**、**数字**或 **布尔值**。`type` `dimensionality` 子属性的值可以是**scalar** 或 **matrix**（指定 `type` 值的二维数组）。 |
| `parameters` | 定义函数的输入参数的数组。 在 Excel intelliSense 中出现的 `name` 和 `description` 子属性。 和 `dimensionality`子属性与此表中`result`前面描述的对象的子属性相同。`type` |
| `options` | 使你可以自定义 Excel 执行函数的方式和时间等的某些方面。 有关如何使用此属性的详细信息，请参阅本文后面的 [Streamed 函数](#streamed-functions)和 [取消](#canceling-a-function)。 |

## <a name="functions-that-return-data-from-external-sources"></a>从外部源返回数据的函数

如果自定义的函数从 web 等外部源检索数据，它必须：

1. 将 JavaScript 承诺返回到 Excel。

2. 使用回调函数，用最终值解析承诺。

自定义函数在单元格中显示 `#GETTING_DATA` 临时结果，而 Excel 等待最终结果。 用户可以在等待结果时与电子表格的其余部分进行正常交互。

在下面的代码示例中，`getTemperature()` 自定义函数检索温度计当前温度。 注意，`sendWebRequest` 是一个假设的函数（这里没有指定），它使用 XHR 来调用温度 Web 服务。

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a>流式处理函数

流式自定义函数使您能够在一段时间内重复地将数据输出到单元格，而无需用户明确请求重新计算。 以下示例是一个自定义函数，它每秒向结果添加一个数字。 关于此代码，请注意以下几点：

- Excel会自动使用 `setResult` 回调来显示每个新值。

- 最后一个参数 `handler` 永远不会在注册代码中指定，当 Excel 用户输入该函数时，它不会显示在自动完成菜单中。 它是包含`setResult` 回调函数的对象，用于将数据从函数传递到 Excel，以更新单元格值。

- 为了让 Excel 在 `handler` 对象中传递 `setResult`函数，您必须在函数注册期间声明支持流，方法是在 JSON 元数据文件中为自定义函数的 `options` 属性设置选项 `"stream": true`。

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a>取消函数

在某些情况下，您可能需要取消执行流的自定义函数，以减少其带宽消耗、工作内存和 CPU 负载。 Excel 在下列情况下取消函数的执行：

- 用户编辑或删除引用函数的单元格。

- 当函数的一个参数（输入）发生变化时。 在这种情况下，取消后触发新函数调用。

- 用户手动触发重新计算。 在这种情况下，取消后触发新函数调用。

> [!NOTE]
> 您必须为每个流式传输功能实施取消处理程序。

若要使函数可取消，请在 JSON 元数据文件中的自定义函数的 `options` 属性中设置选项 `"cancelable": true`。

下面的代码显示以前描述的相同`incrementValue` 函数，但这一次实现了一个取消处理程序。 本示例中， `clearInterval()` 将运行时 `incrementValue` 取消函数。

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

## <a name="saving-and-sharing-state"></a>保存和共享状态

自定义函数可以将数据保存在全局 JavaScript 变量中。 在后续调用中，自定义函数可以使用保存在这些变量中的值。 当用户将相同的自定义函数添加到多个单元格时，保存状态很有用，因为该函数的所有实例都可以共享该状态。 例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。

下面的代码示例演示了以前全局保存状态的温度流函数的实现。 关于此代码，请注意以下几点：

- `refreshTemperature` 是一个流式处理函数，它会在每一秒内读取特定温度计的温度。 新的温度保存在 `savedTemperatures` 变量，但不直接更新单元格值。 它不应该直接从工作表单元格中调用， *所以它没有在JSON文件中注册*。

- `streamTemperature` 每秒钟更新单元格中显示的温度值并使用 `savedTemperatures` 变量作为其数据源。 它必须在JSON文件中注册，并用所有大写字母命名， `STREAMTEMPERATURE`。

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

您自定义的函数可能接受范围的数据作为输入参数，或它可能返回的数据范围。 JavaScript 中，数据范围表示为一个二维数组。

例如，假设函数从 Excel 中存储的一系列数字中返回第二个最大值。 下面的函数接受参数 `values`，其类型为 `Excel.CustomFunctionDimensionality.matrix`。 请注意，在该函数的 JSON 元数据中，您可以将该参数的 `type` 属性 设置为 `matrix`。

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

在生成定义自定义函数的加载项时，请务必包含错误处理逻辑，以便解决运行时错误。 自定义的函数的错误处理与 [Excel JavaScript API 的错误处理整体类同](excel-add-ins-error-handling.md)。 在下面的代码示例中，`.catch` 将处理之前出现在代码中的任何错误。

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi/comments/" + x;

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
- 自定义功能目前不适用于移动客户的Excel。
- 不支持可变函数（每当电子表格中不相关的数据更改时自动重新计算）。
- 尚未启用通过 Office 365 管理门户和 AppSource 进行的部署。
- Excel Online中的自定义功能，可能会在一段时间无活动后，在进程期间停止工作。 刷新浏览器页面 (F5) 并重新输入自定义函数以恢复该功能。
- 如果您有多个加载项在 Excel for Windows 上运行，您可能会看到 **#GETTING_DATA**临时结果单元格内的工作表。 关闭 Excel 的所有窗口，并重新启动 Excel。
- 将来可能会提供专门用于自定义函数的调试工具。 同时，您可以在 Excel Online 使用 F12 开发人员工具调试。 请参阅 [自定义函数最佳实践](custom-functions-best-practices.md)中的详细信息。

## <a name="changelog"></a>更改日志

- **2017 年 11 月 7 日**：发布了* 自定义函数预览和示例
- **2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题
- **2017 年 11 月 28 日**：发布了* 对取消异步函数的支持（需要对流式函数进行相应更改）
- **2018 年 5 月 7 日**：发布了* 对 Mac、Excel Online 和在进程中运行的同步函数的支持
- **2018 年 9 月 20 日，** 发布了支持自定义函数 JavaScript 的运行时。 有关详细信息，请参阅 [Excel 自定义函数的运行时](custom-functions-runtime.md)。

\* 至 Office 预览体验计划渠道

## <a name="see-also"></a>另请参阅

* [自定义函数元数据](custom-functions-json.md)
* [Excel 自定义函数运行运行时](custom-functions-runtime.md)
* [自定义函数的最佳做法](custom-functions-best-practices.md)
