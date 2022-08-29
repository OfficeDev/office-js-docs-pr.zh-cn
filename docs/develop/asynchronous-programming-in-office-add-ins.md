---
title: Office 加载项中的异步编程
description: 了解 Office JavaScript 库如何在 Office 加载项中使用异步编程。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: ce317e2d0648d114fe3716fc47d8cc1315369fc4
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423067"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Office 加载项中的异步编程

[!include[information about the common API](../includes/alert-common-api-info.md)]

为什么 Office 外接程序 API 使用异步编程？ 因为 JavaScript 是单线程语言，如果脚本调用长时间运行的同步进程，则会阻止所有后续脚本执行，直至该进程完成。 由于针对 Office Web 客户端的某些操作 (但富客户端以及) 在同步运行时可能会阻止执行，因此大多数 Office JavaScript API 都设计为异步执行。 这可确保 Office 加载项响应迅速。 使用这些异步方法时，也通常会要求您编写回调函数。

API 中所有异步方法的名称以“异步”结尾，例如 `Document.getSelectedDataAsync`， `Binding.getDataAsync`或 `Item.loadCustomPropertiesAsync` 方法。 调用某个“Async”方法时，该方法会立即执行，并且任何后续脚本执行都可以继续。 传递给“Async”方法的可选回调函数在数据或请求操作准备就绪后便会立即执行。 虽然是立即执行，但在它返回之前可能会略有延迟。

下图显示了调用“Async”方法的执行流，该方法读取用户在基于服务器的 Word 或 Excel 中打开的文档中选择的数据。 在进行“异步”调用时，JavaScript 执行线程可以自由执行任何其他客户端处理 (尽管图表) 中未显示任何内容。 ）当“Async”方法返回时，回调在线程上恢复执行，外接程序可以访问数据、处理数据并显示结果。 使用 Office 富客户端应用程序（如 Word 2013 或 Excel 2013）时，将保留相同的异步执行模式。

*图 1. 异步编程执行流*

![显示与用户的命令执行交互、加载项页和托管加载项的 Web 应用服务器的示意图。](../images/office-addins-asynchronous-programming-flow.png)

在富客户端和 Web 客户端中支持此异步设计是 Office 加载项开发模型"写入一次，跨平台运行"设计目标的一部分。例如，可以使用将在 Excel 2013 和 Excel 网页版中运行的单一基本代码创建一个内容应用程序或任务窗格加载项。

## <a name="write-the-callback-function-for-an-async-method"></a>为“Async”方法编写回调函数

作为 *回调* 参数传递给“Async”方法的回调函数必须声明一个参数，加载项运行时将使用该参数在执行回调函数时提供对 [AsyncResult](/javascript/api/office/office.asyncresult) 对象的访问。 可以编写：

- 一个匿名函数，该函数必须直接按照调用“Async”方法作为“异步”方法的 *回调* 参数进行写入和传递。

- 命名函数，将该函数的名称作为“Async”方法的 *回调* 参数传递。

如果您打算只使用一次代码，则可以使用匿名函数，这是因为该函数没有名称，您不能在代码的其他部分引用此代码。如果您打算重复将回调函数用于多个"Async"方法，则可以使用命名函数。

### <a name="write-an-anonymous-callback-function"></a>编写匿名回调函数

下面的匿名回调函数声明了一个名为在 `result` 回调返回时从 [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member) 属性中检索数据的单个参数。

```js
function (result) {
        write('Selected data: ' + result.value);
}
```

以下示例演示如何在方法的完整“Async”方法调用的上下文中将此匿名回调 `Document.getSelectedDataAsync` 函数按行传递。

- 第一个 *coercionType 参数*`Office.CoercionType.Text`指定以文本字符串形式返回所选数据。

- 第二个 *回调* 参数是内联传递给方法的匿名函数。 当函数执行时，它将使用 *结果* 参数访问 `value` 对象的 `AsyncResult` 属性，以显示用户在文档中选择的数据。

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

还可以使用回调函数的参数来访问对象的其他属性 `AsyncResult` 。 可以使用 [AsyncResult.status](/javascript/api/office/office.asyncresult#office-office-asyncresult-status-member) 属性，以确定调用是成功还是失败。 如果调用失败，你可以使用 [AsyncResult.error](/javascript/api/office/office.asyncresult#office-office-asyncresult-error-member) 属性访问 [Error](/javascript/api/office/office.error) 对象，以获取错误信息。

有关使用该方法的 `getSelectedDataAsync` 详细信息，请参阅 [文档或电子表格中的活动选择读取和写入数据](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。

### <a name="write-a-named-callback-function"></a>编写命名回调函数

或者，可以编写命名函数，并将其名称传递给“Async”方法的 *回调* 参数。 例如，可以重写前一个示例，将名为 `writeDataCallback` 的函数作为 *callback* 参数进行传递，如下所示。

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

`status`对象`asyncContext`的`AsyncResult`属性`error`返回传递给所有“异步”方法的回调函数的相同类型的信息。 但是，返回到属性 `AsyncResult.value` 的内容因“异步”方法的功能而异。

例如，[绑定](/javascript/api/office/office.binding)、`addHandlerAsync`[CustomXmlPart](/javascript/api/office/office.customxmlpart)、[Document](/javascript/api/office/office.document)、[RoamingSettings](/javascript/api/outlook/office.roamingsettings) 和 [Settings](/javascript/api/office/office.settings) 对象)  (的方法用于向这些对象表示的项添加事件处理程序函数。 可以从传递给任何方法的`addHandlerAsync`回调函数访问`AsyncResult.value`该属性，但由于添加事件处理程序时未访问任何数据或对象，`value`因此，如果尝试访问该属性，该属性始终返回 **未定义**。

另一方面，如果调用该 `Document.getSelectedDataAsync` 方法，它会将用户在文档中选择的数据返回到 `AsyncResult.value` 回调中的属性。 或者，如果调用 [Bindings.getAllAsync](/javascript/api/office/office.bindings#office-office-bindings-getallasync-member(1)) 方法，它将返回文档中所有 `Binding` 对象的数组。 并且，如果调用 [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) 方法，它将返回一 `Binding` 个对象。

有关方法返回到 `AsyncResult.value` 属性 `Async` 的说明，请参阅该方法参考主题的“回调值”部分。 有关提供 `Async` 方法的所有对象的摘要，请参阅 [AsyncResult](/javascript/api/office/office.asyncresult) 对象主题底部的表。

## <a name="asynchronous-programming-patterns"></a>异步编程模式

Office JavaScript API 支持两种异步编程模式。

- 使用嵌套回调
- 使用承诺模式

使用回调函数的异步编程通常需要您将回调返回的结果嵌套在两个或更多回调中。如果您需要这么做，则可以使用来自 API 的所有"Async"方法的嵌套回调。

使用嵌套回调是大多数 JavaScript 开发人员都熟知的编程模式，但使用了深层嵌套回调的代码难以阅读和理解。 作为嵌套回调的替代方法，Office JavaScript API 还支持实施承诺模式。

> [!NOTE]
> 在当前版本的 Office JavaScript API 中，对承诺模式的 *内置* 支持仅适用于 [Excel 电子表格和 Word 文档中绑定的](bind-to-regions-in-a-document-or-spreadsheet.md)代码。 但是，可以包装在自定义 Promise-returning 函数中具有回调的其他函数。 有关详细信息，请参阅 [Promise-returning 函数中的 Wrap Common API](#wrap-common-apis-in-promise-returning-functions)。

### <a name="asynchronous-programming-using-nested-callback-functions"></a>使用嵌套回调函数的异步编程

通常，完成一项任务需要执行两个或更多个异步操作。为实现此目的，可在一个调用中嵌套另一个"Async"调用。

以下代码示例内嵌两个异步调用。

- 首先，调用 [Bindings.getByIdAsync](/javascript/api/office/office.bindings#office-office-bindings-getbyidasync-member(1)) 方法，以访问名为“MyBinding”的文档中的绑定。 `AsyncResult`返回到`result`该回调参数的对象提供对`AsyncResult.value`属性中指定绑定对象的访问权限。
- 然后，从第一个 `result` 参数访问绑定对象用于调用 [Binding.getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1)) 方法。
- 最后， `result2` 传递给 `Binding.getDataAsync` 方法的回调参数用于在绑定中显示数据。

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

此基本嵌套回调模式可用于 Office JavaScript API 中的所有异步方法。

以下各节显示如何使用匿名函数或命名函数用于异步方法中的嵌套回调。

#### <a name="use-anonymous-functions-for-nested-callbacks"></a>对嵌套回调使用匿名函数

在下面的示例中，两个匿名函数以内联方式声明并作为嵌套回调传递到 `getByIdAsync` 该函数和 `getDataAsync` 方法中。 由于这两个函数简单且为内嵌，因此实现的意图很清晰。

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

#### <a name="use-named-functions-for-nested-callbacks"></a>对嵌套回调使用命名函数

在复杂实现中，使用命名函数对于提高代码的可读性、可维护性和可重用性可能会有帮助。 在下面的示例中，上一部分中示例中的两个匿名函数已重写为命名和命名 `deleteAllData` 的函数 `showResult`。 然后，这些命名函数将作为名称作为回调传递到 `getByIdAsync` 该函数和 `deleteAllDataValuesAsync` 方法中。

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

在继续执行之前，承诺编程模式会立即返回表示其预期结果的承诺对象，而不是传递回调函数并等待函数返回。然而，与真正同步编程不同的是，在 Office 外接程序运行时环境完成请求之前，承诺结果的实现在后台实际上是延迟的。提供 *onError* 处理程序来覆盖请求无法满足的情况。

Office JavaScript API 提供 [Office.select](/javascript/api/office#Office_select_expression__callback_) 函数，以支持使用现有绑定对象的承诺模式。 返回到函数的 `Office.select` promise 对象仅支持可直接从 [Binding](/javascript/api/office/office.binding) 对象访问的四种方法： [getDataAsync](/javascript/api/office/office.binding#office-office-binding-getdataasync-member(1))、 [setDataAsync](/javascript/api/office/office.binding#office-office-binding-setdataasync-member(1))、 [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) 和 [removeHandlerAsync](/javascript/api/office/office.binding#office-office-binding-removehandlerasync-member(1))。

使用绑定的承诺模式采用此形式。

**Office.select (**_selectorExpression_， _onError_**) 。**_BindingObjectAsyncMethod_

*selectorExpression* 参数采用窗体`"bindings#bindingId"`，其中 *bindingId* 是之前在文档或电子表格中创建的绑定的名称 ( `id`) ， (使用集合的“addFrom”方法`Bindings`之一：`addFromNamedItemAsync``addFromPromptAsync`或`addFromSelectionAsync`) 。 例如，选择器表达式 `bindings#cities` 指定要访问 **ID** 为“cities”的绑定。

*onError* 参数是一个错误处理函数，如果`select`该函数无法访问指定的绑定，则该函数采用一个可用于访问对象的类型的`AsyncResult`单个`Error`参数。 以下示例显示了一个可传递给 *onError* 参数的基本错误处理程序函数。

```js
function onError(result){
    const err = result.error;
    write(err.name + ": " + err.message);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

将 *BindingObjectAsyncMethod* 占位符替换为对 promise 对象支持的四`Binding`个对象方法之一的调用：`getDataAsync`、或 `setDataAsync``addHandlerAsync``removeHandlerAsync`。 对这些方法的调用不支持其他的承诺。 你必须使用[嵌套回调函数模式](#asynchronous-programming-using-nested-callback-functions)来调用它们。

`Binding`实现对象承诺后，可以在链接的方法调用中重复使用它，就好像它是绑定 (加载项运行时不会异步重试履行承诺) 。 `Binding`如果无法实现对象承诺，加载项运行时将在下次调用其异步方法之一时再次尝试访问绑定对象。

下面的代码示例使用该`select`函数从`Bindings`集合中检索包含`id`“`cities`”的绑定，然后调用 [addHandlerAsync](/javascript/api/office/office.binding#office-office-binding-addhandlerasync-member(1)) 方法为绑定的 [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) 事件添加事件处理程序。

```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```

> [!IMPORTANT]
> 函 `Binding` 数返回的 `Office.select` 对象承诺仅提供对对象的四种方法的访问 `Binding` 权限。 如果需要访问对象的任何其他成员 `Binding` ，则必须使用 `Document.bindings` 该属性和 `Bindings.getByIdAsync` 方法 `Bindings.getAllAsync` 来检索 `Binding` 该对象。 例如，如果需要访问对象的任何`Binding`属性 (`type` `id``document`) 或需要访问 [MatrixBinding](/javascript/api/office/office.matrixbinding) 或 [TableBinding](/javascript/api/office/office.tablebinding) 对象的属性，则必须使用`getByIdAsync`或`getAllAsync`方法来检索`Binding`对象。

## <a name="pass-optional-parameters-to-asynchronous-methods"></a>将可选参数传递给异步方法

所有“异步”方法的通用语法都遵循此模式。

 *AsyncMethod* `(`*RequiredParameters*`, [`*OptionalParameters*`],`*CallbackFunction*`);`

所有异步方法都支持可选参数，这些参数作为包含一个或多个可选参数的 JavaScript 对象传入。 包含可选参数的对象是键值对的无序集合，其中包含分隔键和值的“：”字符。 对象中的每对用逗号分隔，整个对集合括在大括号中。 键是参数名称，值是要为该参数传递的值。

可以创建包含内联可选参数的对象，也可以创建一个 `options` 对象并将该对象作为 *选项* 参数传入。

### <a name="pass-optional-parameters-inline"></a>内联传递可选参数

例如，用可选参数内嵌调用 [Document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) 方法的语法类似如下：

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

在此形式的调用语法中，两个可选参数 *coercionType* 和 *asyncContext* 定义为内嵌在大括号中的匿名 JavaScript 对象。

以下示例演示如何通过内联指定可选参数来 `Document.setSelectedDataAsync` 调用方法。

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
> 只要正确指定参数的名称，就可以在参数对象中按任意顺序指定可选参数。

### <a name="pass-optional-parameters-in-an-options-object"></a>传递 options 对象中的可选参数

或者，可以创建一个命名`options`对象，该对象指定与方法调用分开的可选参数，然后将该对象作为 *选项* 参数传递`options`。

下面的示例演示创建`options`对象的一种方法，其中`parameter1``value1`，例如实际参数名称和值的占位符。

```js
const options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};
```

用于指定 [ValueFormat](/javascript/api/office/office.valueformat) 和 [FilterType](/javascript/api/office/office.filtertype) 参数时与以下示例类似。

```js
const options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

下面是创建 `options` 对象的另一种方法。

```js
const options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

当用于指定 `ValueFormat` 参数和 `FilterType` 参数时，如下例所示：

```js
const options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```

> [!NOTE]
> 使用任一方法创建 `options` 对象时，只要正确指定其名称，就可以按任意顺序指定可选参数。

以下示例演示如何通过在对象中`options`指定可选参数来`Document.setSelectedDataAsync`调用方法。

```js
const options = {
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

在这两个可选参数示例中， *回调* 参数都指定为后跟内联可选参数的最后一个参数 (，或遵循 *选项* 参数对象) 。 或者，可以在内联 JavaScript 对象或对象中`options`指定 *回调* 参数。 但是，只能在一个位置传递 *回调* 参数： `options` 在对象中 (内联或在外部创建) ，或者作为最后一个参数，但不能同时传递两者。

## <a name="wrap-common-apis-in-promise-returning-functions"></a>在 Promise-returning 函数中包装常见 API

通用 API (和 Outlook API) 方法不返回 [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。 因此，在异步操作完成之前，不能使用 [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) 来暂停执行。 如果需要 `await` 行为，可以将方法调用包装在显式创建的 Promise 中。

基本模式是创建一个异步方法，该方法立即返回 Promise 对象，并在内部方法完成时 *解析* 该 Promise 对象;如果方法失败，则 *拒绝* 该对象。 下面展示了一个非常简单的示例。

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

当需要等待此函数时，可以使用关键字调用 `await` 该函数，也可以传递给函数 `then` 。

> [!NOTE]
> 当需要在特定于应用程序的对象模型中的函数调用中调用公共 API 时， `run` 此技术特别有用。 有关以这种方式使用的函数的 `getDocumentFilePath` 示例，请参阅 [ 示例 Word-Add-in-JavaScript-MDConversion 中的文件Home.js](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion/blob/master/Word-Add-in-JavaScript-MDConversionWeb/Home.js)。

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
- [Office 加载项中的运行时](../testing/runtimes.md)
