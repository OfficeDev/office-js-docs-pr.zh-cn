---
title: Office 加载项中的异步编程
description: 了解 javaScript Office如何在加载项Office异步编程。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 0e16cf86c57d6020ecf63d077cd4e55b11059d2587ab2ef5c5e3f51fb5da3885
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57081397"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Office 加载项中的异步编程

[!include[information about the common API](../includes/alert-common-api-info.md)]

为什么 Office 外接程序 API 使用异步编程？ 因为 JavaScript 是单线程语言，如果脚本调用长时间运行的同步进程，则会阻止所有后续脚本执行，直至该进程完成。 由于针对 Office Web 客户端的某些操作 (但富客户端以及) 如果同步运行，它们可能会阻止执行，因此大多数 Office JavaScript API 设计为异步执行。 这确保Office外接程序具有响应性和快速性。 使用这些异步方法时，也通常会要求您编写回调函数。

API 中所有异步方法的名称以"Async"结尾，如 `Document.getSelectedDataAsync` 、 或 `Binding.getDataAsync` `Item.loadCustomPropertiesAsync` 方法。 调用某个“Async”方法时，该方法会立即执行，并且任何后续脚本执行都可以继续。 传递给“Async”方法的可选回调函数在数据或请求操作准备就绪后便会立即执行。 虽然是立即执行，但在它返回之前可能会略有延迟。

下图显示了对"Async"方法的调用的执行流，该方法可读取用户在基于服务器的 Word 或 Excel 中打开的文档中选择的数据。 在执行"Async"调用时，JavaScript 执行线程可以自由地执行任何其他客户端处理 (尽管图中未显示任何) 。 ）当“Async”方法返回时，回调在线程上恢复执行，外接程序可以访问数据、处理数据并显示结果。 在使用富客户端应用程序（如 Word 2013 或 Office 2013）时，相同的异步执行模式Excel保留。

*图 1. 异步编程执行流*

![显示一段时间与用户、外接程序页面和托管外接程序的 Web 应用程序服务器之间的命令执行交互的图表。](../images/office-addins-asynchronous-programming-flow.png)

在富客户端和 Web 客户端中支持此异步设计是 Office 加载项开发模型"写入一次，跨平台运行"设计目标的一部分。例如，可以使用将在 Excel 2013 和 Excel 网页版中运行的单一基本代码创建一个内容应用程序或任务窗格加载项。

## <a name="write-the-callback-function-for-an-async-method"></a>为"Async"方法编写回调函数

作为 callback 参数传递给"Async"方法的回调函数必须声明一个参数，当执行回调函数时，外接程序运行时将使用该参数提供对[AsyncResult](/javascript/api/office/office.asyncresult)对象的访问。  可以编写：

- 必须按照对作为"Async"方法的 _callback_ 参数的"Async"方法的调用直接写入和传递的匿名函数。

- 命名函数，作为"Async"方法的 _callback_ 参数传递该函数的名称。

如果您打算只使用一次代码，则可以使用匿名函数，这是因为该函数没有名称，您不能在代码的其他部分引用此代码。如果您打算重复将回调函数用于多个"Async"方法，则可以使用命名函数。

### <a name="write-an-anonymous-callback-function"></a>编写匿名回调函数

以下匿名回调函数声明一个名为 的参数，该参数在回调返回时从 `result` [AsyncResult.value](/javascript/api/office/office.asyncresult#value) 属性检索数据。

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

以下示例演示如何在方法的完整"Async"方法调用上下文中将此匿名回调函数内行 `Document.getSelectedDataAsync` 传递。

- 第一 _个 coercionType_ 参数 `Office.CoercionType.Text` 指定以文本字符串形式返回所选数据。

- 第二 _个 callback_ 参数是内行传递给该方法的匿名函数。 函数执行时，将使用 _result_ 参数访问对象的属性，以显示用户在 `value` `AsyncResult` 文档中选择的数据。

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    function (result) {
        write('Selected data: ' + result.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

您还可以使用回调函数的 参数访问对象的其他 `AsyncResult` 属性。 可以使用 [AsyncResult.status](/javascript/api/office/office.asyncresult#status) 属性，以确定调用是成功还是失败。 如果调用失败，你可以使用 [AsyncResult.error](/javascript/api/office/office.asyncresult#error) 属性访问 [Error](/javascript/api/office/office.error) 对象，以获取错误信息。

有关使用 方法的信息，请参阅在文档或电子表格中的活动选定内容中读取 `getSelectedDataAsync` [和写入数据](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。 

### <a name="write-a-named-callback-function"></a>编写命名回调函数

或者，您可以编写命名函数，并在其名称传递给"Async"方法的 _callback_ 参数。 例如，可以重写前一个示例，将名为 `writeDataCallback` 的函数作为 _callback_ 参数进行传递，如下所示。

```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    writeDataCallback);

// Callback to write the selected data to the add-in UI.
function writeDataCallback(result) {
    write('Selected data: ' + result.value);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a>返回 AsyncResult.value 属性的内容的差异

对象的 、 和 属性将相同类型的信息返回到传递给所有 `asyncContext` `status` `error` `AsyncResult` "Async"方法的回调函数。 但是，返回给属性 `AsyncResult.value` 的功能因"Async"方法的功能而异。

例如 `addHandlerAsync` [，Binding](/javascript/api/office/office.binding) (、CustomXmlPart、Document、RoamingSettings[](/javascript/api/office/office.document)和[](/javascript/api/outlook/office.roamingsettings)设置 对象[) ](/javascript/api/office/office.settings)的方法用于向这些对象表示的项添加事件处理程序函数。 [](/javascript/api/office/office.customxmlpart) 可以从传递给任何方法的回调函数访问属性，但由于添加事件处理程序时未访问任何数据或对象，因此，如果尝试访问该属性，则该属性始终返回 `AsyncResult.value` `addHandlerAsync` `value` **undefined。**

另一方面，如果调用 方法，它会将用户在文档中选择的数据返回到 `Document.getSelectedDataAsync` `AsyncResult.value` 回调中的 属性。 或者，如果调用 [Bindings.getAllAsync](/javascript/api/office/office.bindings#getAllAsync_options__callback_) 方法，它将返回文档中所有 `Binding` 对象的数组。 而且，如果你调用 [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) 方法，它将返回单个 `Binding` 对象。

有关返回到方法属性的内容的说明，请参阅该方法的参考主题的"回调值 `AsyncResult.value` `Async` "部分。 有关提供方法的所有对象的摘要，请参阅 `Async` [AsyncResult](/javascript/api/office/office.asyncresult) 对象主题底部的表。

## <a name="asynchronous-programming-patterns"></a>异步编程模式

JavaScript API Office两种类型的异步编程模式。

- 使用嵌套回调
- 使用承诺模式

使用回调函数的异步编程通常需要您将回调返回的结果嵌套在两个或更多回调中。如果您需要这么做，则可以使用来自 API 的所有"Async"方法的嵌套回调。

使用嵌套回调是大多数 JavaScript 开发人员都熟知的编程模式，但使用了深层嵌套回调的代码难以阅读和理解。 作为嵌套回调的替代方法，Office JavaScript API 还支持实现承诺模式。

> [!NOTE]
> 在当前版本的 Office JavaScript API 中，对承诺模式的内置支持仅适用于 Excel 电子表格和 Word[文档中的绑定代码](bind-to-regions-in-a-document-or-spreadsheet.md)。 但是，可以将具有回调的其他函数包装在你自己的自定义 Promise 返回函数中。 有关详细信息，请参阅在[Promise 返回函数中包装常见 API。](#wrap-common-apis-in-promise-returning-functions)

### <a name="asynchronous-programming-using-nested-callback-functions"></a>使用嵌套回调函数的异步编程

通常，完成一项任务需要执行两个或更多个异步操作。为实现此目的，可在一个调用中嵌套另一个"Async"调用。

以下代码示例内嵌两个异步调用。

- 首先，调用 [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getByIdAsync_id__options__callback_) 方法，以访问名为“MyBinding”的文档中的绑定。 返回到 `AsyncResult` 该回调 `result` 的参数的对象提供对属性中指定绑定对象的 `AsyncResult.value` 访问。
- 然后，使用第一个参数访问的绑定对象 `result` 调用 [Binding.getDataAsync](/javascript/api/office/office.binding#getDataAsync_options__callback_) 方法。
- 最后， `result2` 传递给该方法的回调的参数用于在 `Binding.getDataAsync` 绑定中显示数据。

```js
function readData() {
    Office.context.document.bindings.getByIdAsync("MyBinding", function (result) {
        result.value.getDataAsync({ coercionType: 'text' }, function (result2) {
            write(result2.value);
        });
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

此基本嵌套回调模式可用于 JavaScript API 中所有Office方法。

以下各节显示如何使用匿名函数或命名函数用于异步方法中的嵌套回调。

#### <a name="use-anonymous-functions-for-nested-callbacks"></a>将匿名函数用于嵌套回调

在下面的示例中，两个匿名函数内嵌声明，并作为嵌套回调传入 `getByIdAsync` 和 `getDataAsync` 方法。 由于这两个函数简单且为内嵌，因此实现的意图很清晰。

```js
Office.context.document.bindings.getByIdAsync('myBinding', function (bindingResult) {
    bindingResult.value.getDataAsync(function (getResult) {
        if (getResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Data has been read successfully.');
        }
    });
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

#### <a name="use-named-functions-for-nested-callbacks"></a>将命名函数用于嵌套回调

在复杂实现中，使用命名函数对于提高代码的可读性、可维护性和可重用性可能会有帮助。 在下面的示例中，上一节中的示例中的两个匿名函数已被重写为名为 和 `deleteAllData` 的函数 `showResult` 。 然后，这些命名函数将按名称 `getByIdAsync` `deleteAllDataValuesAsync` 作为回调传递到 和 方法中。

```js
Office.context.document.bindings.getByIdAsync('myBinding', deleteAllData);

function deleteAllData(asyncResult) {
    asyncResult.value.deleteAllDataValuesAsync(showResult);
}

function showResult(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Data has been deleted successfully.');
    }
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a>使用承诺模式访问绑定中的数据的异步编程

在继续执行之前，承诺编程模式会立即返回表示其预期结果的承诺对象，而不是传递回调函数并等待函数返回。然而，与真正同步编程不同的是，在 Office 外接程序运行时环境完成请求之前，承诺结果的实现在后台实际上是延迟的。提供 _onError_ 处理程序来覆盖请求无法满足的情况。

JavaScript API Office JavaScript API 提供了[Office.select](/javascript/api/office#Office_select_expression__callback_)方法，以支持用于处理现有绑定对象的承诺模式。 返回到 方法的承诺对象仅支持可以直接从 Binding 对象访问的四种方法 `Office.select` [](/javascript/api/office/office.binding)：getDataAsync、setDataAsync、addHandlerAsync[](/javascript/api/office/office.binding#setDataAsync_data__options__callback_)和[removeHandlerAsync](/javascript/api/office/office.binding#removeHandlerAsync_eventType__options__callback_)。 [](/javascript/api/office/office.binding#getDataAsync_options__callback_) [](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_)

使用绑定的承诺模式采用此形式。

**Office.select (** _selectorExpression_， _onError_ **) .**_BindingObjectAsyncMethod_

_selectorExpression_ 参数采用的形式为 ，其中 `"bindings#bindingId"` _bindingId_ 是之前使用集合的"addFrom"方法之一在文档或电子表格中创建的绑定的名称 ( () ：、 或 `id` `Bindings` `addFromNamedItemAsync` `addFromPromptAsync` `addFromSelectionAsync`) 。 例如，选择 `bindings#cities` 器表达式指定你要访问 **ID** 为"cities"的绑定。

_onError_ 参数是一个错误处理函数，该函数采用可用于访问对象的单个类型参数（如果该方法无法访问 `AsyncResult` `Error` `select` 指定的绑定）。 以下示例显示了一个可传递给 _onError_ 参数的基本错误处理程序函数。

```js
function onError(result){
    var err = result.error;
    write(err.name + ": " + err.message);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

将 _BindingObjectAsyncMethod_ 占位符替换为对承诺对象支持的四个对象方法之一的调用：、、 `Binding` `getDataAsync` 或 `setDataAsync` `addHandlerAsync` `removeHandlerAsync` 。 对这些方法的调用不支持其他的承诺。 你必须使用[嵌套回调函数模式](#asynchronous-programming-using-nested-callback-functions)来调用它们。

对象承诺实现后，可以在链式方法调用中重复使用它，就像它是绑定一样 (加载项运行时不会异步重试实现承诺 `Binding`) 。 如果对象承诺无法实现，加载项运行时将在下次调用其异步方法之一时再次尝试访问 `Binding` 绑定对象。

下面的代码示例使用 方法从集合中检索具有 " " 的绑定，然后调用 `select` `id` `cities` `Bindings` [addHandlerAsync](/javascript/api/office/office.binding#addHandlerAsync_eventType__handler__options__callback_) 方法为绑定的 [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) 事件添加事件处理程序。

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> `Binding`方法返回的对象承诺 `Office.select` 仅提供对对象的四个方法 `Binding` 的访问。 如果需要访问对象的其他任何成员，则必须使用 属性和方法检索 `Binding` `Document.bindings` `Bindings.getByIdAsync` `Bindings.getAllAsync` `Binding` 对象。 例如，如果需要访问对象的任何属性 (、 、 或 属性) ，或者需要访问 `Binding` `document` `id` `type` [MatrixBinding](/javascript/api/office/office.matrixbinding) 或 [TableBinding](/javascript/api/office/office.tablebinding) `getByIdAsync` `getAllAsync` 对象的属性 `Binding` ，则必须使用 或 方法检索对象。

## <a name="pass-optional-parameters-to-asynchronous-methods"></a>将可选参数传递给异步方法

所有"Async"方法的常见语法都遵循此模式。

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

所有异步方法都支持可选参数，这些可选参数作为包含一个或多个可选参数的 JavaScript 对象表示法 (JSON) 对象传入。包含可选参数的 JSON 对象是键-值对的无序集合，其中用":"字符来分隔键和值。对象中的每对用逗号分隔，整个对集合括在大括号中。键是参数名称，值是要为该参数传递的值。

可以创建包含可选参数内嵌的 JSON 对象，或者创建对象，并作为 options 参数进行 `options` 传递。 

### <a name="pass-optional-parameters-inline"></a>内联传递可选参数

例如，用可选参数内嵌调用 [Document.setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) 方法的语法类似如下：

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

在此形式的调用语法中，两个可选参数 _coercionType_ 和 _asyncContext_ 定义为括在括号中的 JSON 对象内嵌。

以下示例演示如何通过指定可选参数内嵌 `Document.setSelectedDataAsync` 来调用该方法。

```js
Office.context.document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    {coercionType: "html", asyncContext: 42},
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

> [!NOTE]
> 可以任何顺序在 JSON 对象中指定可选参数，只要指定正确的参数名称即可。

### <a name="pass-optional-parameters-in-an-options-object"></a>在 options 对象中传递可选参数

或者，也可以创建一个名为 的对象，该对象与方法调用分开指定可选参数，然后将该对象 `options` `options` 作为 _options_ 参数传递。

下面的示例演示创建对象的一种方法，其中 、 等是实际参数名称和值的 `options` `parameter1` `value1` 占位符。

```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

用于指定 [ValueFormat](/javascript/api/office/office.valueformat) 和 [FilterType](/javascript/api/office/office.filtertype) 参数时与以下示例类似。

```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

下面是创建对象的另一 `options` 种方法。

```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

当用于指定 和 参数时，其外观 `ValueFormat` `FilterType` 如以下示例所示：

```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> 使用任一方法创建对象时，只要正确指定了可选参数的名称，就可以按任意顺序 `options` 指定这些参数。

以下示例演示如何通过指定对象中的可选参数 `Document.setSelectedDataAsync` 来调用 `options` 方法。

```js
var options = {
   coercionType: "html",
   asyncContext: 42
};

document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    options,
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在这两个可选参数示例中 _，callback_ 参数指定为内嵌可选参数 (或 _options_ 参数对象之后的最后一) 。 还可以在内嵌 JSON 对象或  对象内指定 `options` 参数。 但是，只能在一个位置传递 _callback_ 参数：在 _option_ 对象内（内嵌或在外部创建），或作为最后一个参数，但不能同时在两个位置。

## <a name="wrap-common-apis-in-promise-returning-functions"></a>在 Promise 返回函数中包装常见 API

通用 API (和 Outlook API) 不会返回[Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。 因此，在 [异步操作](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) 完成之前，不能使用 await 暂停执行。 如果需要行为，可以将方法调用包装在显式创建的 `await` Promise 中。 

基本模式是创建一个异步方法，该方法立即返回 Promise 对象，在内部方法完成时解析 Promise 对象;如果该方法失败，则拒绝该对象。 下面展示了一个非常简单的示例。

```javascript
function getDocumentFilePath() {
    return new OfficeExtension.Promise(function (resolve, reject) {
        try {
            Office.context.document.getFilePropertiesAsync(function (asyncResult) {
                resolve(asyncResult.value.url);
            });
        }
        catch (error) {
            reject(WordMarkdownConversion.errorHandler(error));
        }
    })
}
```

当需要等待此方法时，可以使用关键字或作为传递给函数的函数 `await` 调用 `then` 此方法。

> [!NOTE]
> 当你需要在特定于应用程序的对象模型中的方法调用内调用其中一个通用 API 时，此 `run` 技术尤其有用。 有关以上函数的示例，请参阅示例 [ Word-Add-in-JavaScript-MDConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js)Home.js中的文件。

下面是使用 TypeScript 的示例。

```typescript
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](../reference/javascript-api-for-office.md)
