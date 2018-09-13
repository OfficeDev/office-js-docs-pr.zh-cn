# <a name="create-custom-functions-in-excel-preview"></a>在 Excel 中创建自定义函数（预览）

借助自定义函数（类似于用户定义的函数 [UDF]），开发人员可以使用加载项向 Excel 添加任何 JavaScript 函数。 然后，用户可以像使用 Excel 中的其他本机函数（例如 `=SUM()`）一样访问自定义函数。 本文介绍了如何在 Excel 中创建自定义函数。

下图显示了最终用户如何将自定义函数插入到单元格中。 将 42 添加到一对数字的函数。

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

以下是相同自定义函数的代码。

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

自定义函数现可在 Windows、Mac 和 Excel Online 的开发人员预览版中使用。 若要试用，请按照以下步骤操作：

1. 安装 Office（Windows 的内部版本 9325 或 Mac 上的内部版本 13.329）并加入 [Office 预览体验成员](https://products.office.com/office-insider)计划。 （请注意，仅仅获取最新版本是不够的；在加入预览体验成员计划之前，任何版本的功能都将禁用）
2. 使用 [Yo Office](https://github.com/OfficeDev/generator-office) 创建 Excel 自定义函数的加载项项目，并按照 [project README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) 中的说明在 Excel 中启动加载项，更改代码并进行调试。
3. 在任意单元格中键入“`=CONTOSO.ADD42(1,2)`”，再按 **Enter** 运行自定义函数。

请参阅本文末尾的**已知问题**部分，其中包括自定义函数的当前限制，该部分将随时间进行更新。

## <a name="learn-the-basics"></a>学习基础知识

在克隆的示例存储库中，将看到以下文件：

- **./src/customfunctions.js**，其中包含自定义函数代码（请参阅上面 `ADD42` 函数的简单代码示例）。
- **./config/customfunctions.json**，其中包含将自定义函数告诉 Excel 的注册 JSON。 注册会使自定义函数显示在用户键入单元格时显示的可用函数列表中。
- **./index.html**，它提供 JS 文件的&lt;脚本&gt;引用。 该文件不在 Excel 中显示 UI。
- **./manifest.xml**，它将 HTML、JavaScript 和 JSON 文件的位置告诉 Excel；还为与该加载项一起安装的所有自定义函数指定一个命名空间。

### <a name="json-file-configcustomfunctionsjson"></a>JSON 文件 (./config/customfunctions.json)

customfunctions.json中的以下代码相同的 `ADD42` 功能指定元数据。

> [!NOTE]
> JSON文件的详细参考信息（包括本示例中未使用的选项）位于 [自定义函数注册JSON](custom-functions-json.md)。

请注意，对于这个例子：

- 只有一个自定义函数，所以只有 `functions` 阵列的一个成员。
- 该 `name` 属性定义了函数名称。 正如您在前面的动画gif中看到的，名称空间（`CONTOSO`）预先添加到Excel自动完成菜单中的函数名称。 此前缀在加载项清单中定义，如下所述。 前缀和函数名使用句点分隔，按照惯例，前缀和函数名都是大写。 要使用自定义函数，用户键入名称空间，后跟该函数的名称（`ADD42`）进入一个单元格，在这种情况下 `=CONTOSO.ADD42`。 前缀将用作公司或加载项的标识符。 
- `description` 将在 Excel 的自动完成菜单中显示。
- 当用户针对某个函数请求帮助时，Excel 将打开任务窗格并显示位于 `helpUrl` 所指定 URL 的网页。
- 该 `result` 属性指定函数返回给 Excel 之信息的类型。 该 `type` 子属性可以 `"string"`， `"number"`， 或 `"boolean"`。 该 `dimensionality` 属性可以 `scalar` 或 `matrix` （指定 `type`值的二维数组。）
- 该 `parameters` 数组 *按顺序*指定了传递给函数的每个参数中的数据类型。 该 `name` 和 `description` 在Excel智能感知中使用子属性。 该 `type` 和 `dimensionality` 子属性与上述 `result` 属性之子属性相同。
- 该 `options` 属性使您可以自定义Excel执行功能之方式和时间的某些方面。 本文后面有关于这些选项的更多信息。

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
> 自定义函数是在用户第一次运行加载项时注册的。 之后，对于同一用户，在所有工作簿中都可以使用它们（不仅是最初加载项运行的那个。）

您的JSON文件的服务器设置必须具有 [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) 启用以使自定义函数在Excel Online中正常工作。


### <a name="manifest-file-manifestxml"></a>清单文件 (./manifest.xml)


以下是一个例子 `<ExtensionPoint>` 和 `<Resources>` 您在加载项的清单中包含的标记使Excel能够运行您的函数。 请注意有关此标记的以下事实：

- 该 `<Script>` 元素及其相应的资源ID指定JavaScript文件在您的函数中的位置。
- 该 `<Page>` 元素及其相应的资源ID指定加载项之HTML页面的位置。 HTML页面包含一个 `<Script>` 加载JavaScript文件的标签（customfunctions.js）。 HTML 页面是一个隐藏页面，始终不会在 UI 中显示。
- 该 `<Metadata>` 元素及其相应的资源ID指定JSON文件的位置。
- 一个 `<Namespace>` 元素及其相应的资源ID指定加载项中所有自定义函数的前缀。


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

## <a name="initializing-custom-functions"></a>初始化自定义函数

您的代码在使用之前必须初始化自定义函数功能。 你可以在一个 &lt;脚本&gt; 在HTML文件（customfunctions.html）中的标记或JavaScript文件（customfunctions.js）的顶部。 在预览自定义函数期间，您可以选择两种初始化语法。 回购库中的HTML文件使用以下语法：

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

您还可以使用以下语法：

```js
Office.Preview.StartCustomFunctions();
```

## <a name="handling-errors"></a>处理错误
自定义的函数的错误处理与 [Excel JavaScript API 的错误处理整体类同](./excel-add-ins-error-handling.md)。 一般情况下，将使用 `.catch` 来处理错误。 下面的代码是 `.catch` 的例子。 

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

## <a name="synchronous-and-asynchronous-functions"></a>同步和异步功能

上面的 `ADD42` 功能是关于Excel同步的（通过设置在JSON文件中的 `"sync": true` 选项来指定）。 同步函数提供了快速的性能，因为它们与Excel运行的过程相同，并且在多线程计算过程中它们并行运行。   

另一方面，如果您的自定义函数从Web中检索数据，则它必须相对于Excel异步。 异步函数必须：

1. 将 JavaScript 承诺返回到 Excel。
3. 使用回调函数，用最终值解析Promise。

下面的代码显示用于检索温度计温度的异步自定义函数示例。 注意 `sendWebRequest` 是一个假设的功能，这里没有指定，它使用XHR来调用温度网络服务。

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

异步函数显示在Excel等待最终结果时，单元格中出现的一个 `GETTING_DATA` 暂时错误。 用户可以在等待结果时与电子表格的其余部分进行正常交互。

> [!NOTE]
> 自定义函数默认是异步的。 要将功能指定为同步，设置 `"sync": true` 选项在注册JSON文件中自定义函数的 `options` 属性中。

## <a name="streamed-functions"></a>流式处理函数

异步功能可以流式处理。 借助流式处理自定义函数，可以随时间推移将数据重复输出到单元格，而无需等待 Excel 或用户请求重新计算。 以下示例是一个自定义函数，它每秒向结果添加一个数字。 关于此代码，请注意以下几点：

- Excel会自动使用 `setResult` 回调来显示每个新值。
- 始终不会在注册代码中指定最后的 `caller` 参数，且当 Excel 用户输入此函数时，该参数不会在其自动完成菜单中显示。 它是包含`setResult` 回调函数的对象，用于将数据从函数传递到 Excel，以更新单元格值。
- 为了让Excel通过 `setResult` 功能在 `caller` 对象，您必须通过设置 `"stream": true` 选项在注册JSON文件中自定义函数的 `options` 属性里，来声明在函数注册期间支持流式处理。

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a>取消

可以取消流式处理函数和异步函数。 对于减少带宽消耗、工作内存和 CPU 负载，取消函数调用非常重要。 Excel 在以下情况下取消函数调用：

- 用户编辑或删除引用函数的单元格。
- 函数的参数（输入）之一发生变化。 在这种情况下，除了取消之外，还会触发新的函数调用。
- 用户手动触发重新计算。与上述情况一样，除了取消之外，还会触发新的函数调用。

您 *必须* 为每个流式传输功能实施取消处理程序。 异步、非流式功能可取消，也可不取消；由你决定。 同步功能无法取消。

要使功能可取消，请设置选项 `"cancelable": true` 在注册JSON文件中自定义函数的 `options` 属性里面。

下面的代码展示了已实现取消的上一个示例。 在此代码中，`caller` 对象包含一个为每个可取消的自定义函数定义的 `onCanceled` 函数。

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

## <a name="saving-and-sharing-state"></a>保存和共享状态

异步自定义函数可以将数据保存在全局 JavaScript 变量中。 在后续调用中，自定义函数可以使用保存在这些变量中的值。 当用户将相同的自定义函数添加到多个单元格时，保存状态很有用，因为该函数的所有实例都可以共享该状态。 例如，可以保存调用某个 Web 资源时返回的数据，以避免再次调用同一个 Web 资源。

下面的代码演示之前的温度流式处理函数的实施过程，该函数将状态保存在全局作用域。 关于此代码，请注意以下几点：

- `refreshTemperature` 是一个流式处理函数，它会在每一秒内读取特定温度计的温度。 新的温度保存在 `savedTemperatures` 变量，但不直接更新单元格值。 它不应该直接从工作表单元格中调用， *所以它没有在JSON文件中注册*。
- `streamTemperature` 每秒钟更新单元格中显示的温度值并使用 `savedTemperatures` 变量作为其数据源。 它必须在JSON文件中注册，并用所有大写字母命名， `STREAMTEMPERATURE`。
- 用户可以从 Excel UI 的多个单元格中调用 `streamTemperature`。 每次调用都从相同的 `savedTemperatures` 变量读取数据。

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
> 同步功能（通过在JSON文件中设置 `"sync": true` 选项指定）不能共享状态，因为Excel在多线程计算过程中将它们并行化。 只有异步函数可以共享状态，因为加载项的同步函数在每个进程中共享相同的JavaScript上下文。

## <a name="working-with-ranges-of-data"></a>使用数据区域

自定义函数可以将数据区域用作参数，或者可以从自定义函数返回数据区域。

例如，假设函数从 Excel 中存储的一系列数字中返回第二个最大值。 下面的函数需要使用参数 `values`，即 `Excel.CustomFunctionDimensionality.matrix` 参数类型。 请注意，在此函数的注册JSON中，您可以设置参数的 `type` 属性给 `matrix`。

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

如您所见，范围在JavaScript中以行矩阵的矩阵（如2维矩阵）处理。

## <a name="known-issues"></a>已知问题

- Excel 暂未使用帮助 URL 和参数说明。
- 自定义功能目前不适用于移动客户的Excel。
- 目前，加载项依赖隐藏的浏览器进程来运行异步自定义函数。 将来，JavaScript 将直接在某些平台上运行，以确保自定义函数运行速度更快并占用更少的内存。 此外，大多数平台将不再需要清单的 `<Page>` 元素所引用的 HTML 页面，因为 Excel 将直接运行 JavaScript。 若要为这一更改做准备，请确保自定义函数未使用网页 DOM。 使用GET或POST，用于访问网络所支持的主机APIs将会是 [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) 和 [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) 。
- 易失性函数（当电子表格中不相关数据发生变化时自动重新计算的函数）尚不受支持。
- 调试仅适用于Excel for Windows上的异步功能。
- 尚未启用通过Office 365管理门户和AppSource进行的部署。
- Excel Online中的自定义功能，可能会在一段时间无活动后，在进程期间停止工作。 刷新浏览器页面（F5）并重新输入自定义函数以恢复该功能。

## <a name="changelog"></a>更改日志

- **2017 年 11 月 7 日**：发布了自定义函数（预览）和示例
- **2017 年 11 月 20 日**：修复了使用内部版本 8801 及更高版本的函数的兼容性问题
- **2017 年 11 月 28 日**：发布了对取消异步函数的支持（需要对流式处理函数进行相应更改）
- **2018年5月7日**：提供对Mac、Excel Online和进程运行中的同步功能的支持
