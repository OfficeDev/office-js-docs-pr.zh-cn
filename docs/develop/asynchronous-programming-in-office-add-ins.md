---
title: Office 加载项中的异步编程
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: fc39bddbe050f8253769a0013be2d48b26dcb599
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324643"
---
# <a name="asynchronous-programming-in-office-add-ins"></a>Office 加载项中的异步编程

[!include[information about the common API](../includes/alert-common-api-info.md)]

为什么 Office 加载项 API 使用异步编程？由于 JavaScript 是单线程语言，如果脚本调用长时间运行的同步过程，则所有后续脚本执行都将被阻止，直到该过程完成。由于对 Office web 客户端执行某些操作（但也有丰富的客户端），因此大多数 Office JavaScript Api 都是以异步方式执行的。这可确保 Office 加载项快速响应和快速。在使用这些异步方法时，通常还需要编写回调函数。

API end 中所有异步方法的名称，其中包含 "Async"，如`Document.getSelectedDataAsync`、 `Binding.getDataAsync`或`Item.loadCustomPropertiesAsync`方法。当调用 "Async" 方法时，它会立即执行并可继续执行任何后续脚本。传递给 "Async" 方法的可选回调函数将在数据或请求的操作准备就绪后立即执行。这通常会立即发生，但在返回之前可能会稍有延迟。

下图显示了一个调用"Async"方法的执行流，该方法可读取用户在基于服务器的 Word 或 Excel 中打开的文档中选择的数据。“Async”调用开始时，JavaScript 执行线程空闲，可以执行任何额外的客户端处理（但图中没有显示）。当“Async”方法返回时，回调在线程上恢复执行，加载项可以访问数据、处理数据并显示结果。当使用 Office 富客户端主机应用程序（如，Word 2013 或 Excel 2013）时，可保持同样的异步执行模式。

*图 1. 异步编程执行流*

![异步编程线程执行流](../images/office-addins-asynchronous-programming-flow.png)

在富客户端和 Web 客户端中支持此异步设计是 Office 加载项开发模型"写入一次，跨平台运行"设计目标的一部分。例如，可以使用将在 Excel 2013 和 Excel 网页版中运行的单一基本代码创建一个内容应用程序或任务窗格加载项。

## <a name="writing-the-callback-function-for-an-async-method"></a>编写"Async"方法的回调函数


作为_callback_参数传递给 "Async" 方法的回调函数必须声明一个参数，外接程序运行时将使用该参数在回调函数执行时提供对[AsyncResult](/javascript/api/office/office.asyncresult)对象的访问权限。您可以编写：


- 必须编写并作为 "Async" 方法的_callback_参数与调用一起直接传递给 "async" 方法的匿名函数。

- 一个命名函数，用于将该函数的名称作为 "Async" 方法的_callback_参数进行传递。

如果您打算只使用一次代码，则可以使用匿名函数，这是因为该函数没有名称，您不能在代码的其他部分引用此代码。如果您打算重复将回调函数用于多个"Async"方法，则可以使用命名函数。


### <a name="writing-an-anonymous-callback-function"></a>编写匿名回调函数

以下匿名回调函数声明名为`result`的单个参数，该参数在回调返回时从[AsyncResult](/javascript/api/office/office.asyncresult#value)属性中检索数据。


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

下面的示例演示如何在对`Document.getSelectedDataAsync`方法的完整 "Async" 方法调用的上下文中以行为的方式传递此匿名回调函数。


- 第一个_coercionType_参数`Office.CoercionType.Text`指定将所选数据作为文本字符串返回。

- 第二个_回调_参数是以串联方式传递给方法的匿名函数。函数执行时，它使用_result_参数访问`value` `AsyncResult`对象的属性，以显示用户在文档中选择的数据。


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

您还可以使用回调函数的参数来访问该`AsyncResult`对象的其他属性。使用[AsyncResult](/javascript/api/office/office.asyncresult#status)属性可确定呼叫是成功还是失败。如果调用失败，可以使用[AsyncResult](/javascript/api/office/office.asyncresult#error)属性访问[error 对象，以获取错误消息](/javascript/api/office/office.error)。

有关使用`getSelectedDataAsync`方法的详细信息，请参阅[在文档或电子表格的活动选定内容中读取和写入数据](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。 


### <a name="writing-a-named-callback-function"></a>编写命名回调函数

或者，也可以编写一个命名的函数并将其名称传递给 "Async" 方法的_callback_参数。例如，可以重写上面的示例，以将名`writeDataCallback`为的_回调_参数的函数传递给此类。


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


对象的`asyncContext`、 `status`和`error`属性将向传递给所有 "Async" 方法的回调函数返回相同类型的信息。 `AsyncResult`但是，返回到`AsyncResult.value`属性的内容将根据 "Async" 方法的功能而有所不同。

例如，CustomXmlPart、 `addHandlerAsync` [Document](/javascript/api/office/office.document)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)和[Settings](/javascript/api/office/office.settings)对象[](/javascript/api/office/office.customxmlpart)的方法（用于将事件处理程序函数添加到这些对象所表示的[项目中）](/javascript/api/office/office.binding)。您可以从传递`AsyncResult.value`给任何`addHandlerAsync`方法的回调函数访问该属性，但由于在添加事件处理程序时没有要访问的数据或对象，因此，如果您`value`尝试访问该属性，该属性将始终返回**undefined** 。

另一方面，如果调用`Document.getSelectedDataAsync`方法，它会将用户在文档中选择的数据返回到回调中的`AsyncResult.value`属性。或者，如果调用[getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-)方法，它将返回一个数组，其中的所有`Binding`对象都在文档中。如果调用[getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)方法，则它将返回单个`Binding`对象。

有关`AsyncResult.value` `Async`方法的返回属性的说明，请参阅该方法的参考主题的 "回调值" 部分。有关提供`Async`方法的所有对象的摘要，请参阅[AsyncResult](/javascript/api/office/office.asyncresult)对象主题底部的表。


## <a name="asynchronous-programming-patterns"></a>异步编程模式


Office JavaScript API 支持两种类型的异步编程模式：


- 使用嵌套回调
    
- 使用承诺模式
    
使用回调函数的异步编程通常需要您将回调返回的结果嵌套在两个或更多回调中。如果您需要这么做，则可以使用来自 API 的所有"Async"方法的嵌套回调。

使用嵌套回调是一种熟悉大多数 JavaScript 开发人员的编程模式，但使用深度嵌套回调的代码很难阅读和理解。作为嵌套回调的替代方法，Office JavaScript API 还支持实施承诺模式。但是，在当前版本的 Office JavaScript API 中，承诺模式仅适用于[Excel 电子表格和 Word 文档中的绑定](bind-to-regions-in-a-document-or-spreadsheet.md)代码。

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a>使用嵌套回调函数的异步编程


通常，完成一项任务需要执行两个或更多个异步操作。为实现此目的，可在一个调用中嵌套另一个"Async"调用。

以下代码示例内嵌两个异步调用。


- 首先，调用[getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)方法来访问名为 "MyBinding" 的文档中的绑定。返回`AsyncResult`到该回调的`result`参数的对象提供对该`AsyncResult.value`属性中指定的 binding 对象的访问权限。

- 然后，使用从第一个`result`参数访问的 binding 对象调用[binding.getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)方法。

- 最后，传递`result2`给`Binding.getDataAsync`方法的回调参数用于显示绑定中的数据。


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


#### <a name="using-anonymous-functions-for-nested-callbacks"></a>将匿名函数用于嵌套回调

在下面的示例中，将内联声明两个匿名函数，并`getByIdAsync`将`getDataAsync`其作为嵌套回调传递给和方法。由于函数是简单且内嵌的，因此实现的意图将立即清除。


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


#### <a name="using-named-functions-for-nested-callbacks"></a>将命名函数用于嵌套回调

在复杂的实现中，使用命名的函数使代码更易于读取、维护和重用可能非常有用。在下面的示例中，前一节的示例中的两个匿名函数已重写为名`deleteAllData`为`showResult`和的函数。然后，通过名称将这些命名的`getByIdAsync`函数`deleteAllDataValuesAsync`作为回调传递给和方法。


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


Office JavaScript API 提供了[office. select](/javascript/api/office#office-select-expression--callback-)方法，以支持使用现有绑定对象的承诺模式。返回`Office.select`到方法的承诺对象仅支持您可以直接从[Binding](/javascript/api/office/office.binding)对象访问的四个方法： [binding.getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)、 [binding.setdataasync](/javascript/api/office/office.binding#setdataasync-data--options--callback-)、 [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)和[removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-)。


与绑定一起使用的承诺模式采用以下形式：

 **Office. select （**_selectorExpression_， _onError_**）。**_BindingObjectAsyncMethod_

_SelectorExpression_参数`"bindings#bindingId"`采用窗体，其中_bindingId_是您之前在文档`id`或电子表格中创建的绑定的名称（）（ `Bindings`使用集合的 "addFrom" 方法之一： `addFromNamedItemAsync`、 `addFromPromptAsync`或`addFromSelectionAsync`）。例如，选择器表达式`bindings#cities`指定要访问**id**为 "城市" 的绑定。

_OnError_参数是一个错误处理函数，它采用可用于访问`AsyncResult` `Error`对象的单个参数类型，前提是该`select`方法无法访问指定的绑定。下面的示例演示可传递给_onError_参数的基本错误处理程序函数。




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

将_BindingObjectAsyncMethod_占位符替换为对承诺对象支持的四个`Binding`对象方法中的任何一个`getDataAsync`： `setDataAsync`、、 `addHandlerAsync`或。 `removeHandlerAsync`对这些方法的调用不支持其他承诺。必须使用[嵌套回调函数模式](#AsyncProgramming_NestedCallbacks)调用它们。

在满足`Binding`对象承诺后，可以在连锁方法调用中重用它，就像它是一个绑定一样（加载项运行时不会异步重试满足承诺）。如果无法`Binding`满足对象承诺，加载项运行时将在下次调用其异步方法之一时再次尝试访问 binding 对象。

下面的`select`代码示例使用方法`id`从`cities` `Bindings`集合中检索带有 "" 的绑定，然后调用[addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)方法为绑定的[dataChanged](/javascript/api/office/office.bindingdatachangedeventargs)事件添加事件处理程序。




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> 方法返回的`Binding`对象承诺仅提供对该`Binding`对象的四个方法的访问。 `Office.select`如果需要访问`Binding` `Document.bindings`对象的任何其他成员，则必须使用属性和`Bindings.getByIdAsync`或`Bindings.getAllAsync`方法检索该`Binding`对象。例如， `Binding`如果需要访问对象的任何属性（ `document` `id`或`type`属性），或者需要访问[MatrixBinding](/javascript/api/office/office.matrixbinding)或[TableBinding](/javascript/api/office/office.tablebinding)对象的属性，则必须使用`getByIdAsync`或`getAllAsync`方法来检索`Binding`对象。


## <a name="passing-optional-parameters-to-asynchronous-methods"></a>向异步方法传递可选参数


所有"异步"方法的常用语法都遵循此模式：

 _AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`

所有异步方法都支持可选参数，这些可选参数作为包含一个或多个可选参数的 JavaScript 对象表示法 (JSON) 对象传入。包含可选参数的 JSON 对象是键-值对的无序集合，其中用":"字符来分隔键和值。对象中的每对用逗号分隔，整个对集合括在大括号中。键是参数名称，值是要为该参数传递的值。

您可以创建包含可选参数内嵌的 JSON 对象，或通过创建`options`对象并将其作为_options_参数传入。


### <a name="passing-optional-parameters-inline"></a>传递可选参数内嵌

例如，用可选参数内嵌调用 [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 方法的语法类似如下：

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

在这种形式的调用语法中，两个可选参数_coercionType_和_asyncContext_定义为括在大括号内的 JSON 对象内联。

下面的示例演示如何通过指定内嵌可选`Document.setSelectedDataAsync`参数来调用方法。


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


### <a name="passing-optional-parameters-in-an-options-object"></a>在 options 对象中传递可选参数

或者，也可以创建一个名为`options`的对象，该对象指定与方法调用分开的可选参数，然后`options`将对象作为_options_参数传递。

下面的示例演示创建`options`对象的一种方法，其中`parameter1`、 `value1`等是实际参数名称和值的占位符。




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

此处是创建`options`对象的另一种方法。




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

用于指定`ValueFormat`和`FilterType`参数时与以下示例类似：


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> 使用任一方法创建`options`对象时，只要可选参数的名称指定正确，就可以按任意顺序指定这些参数。

下面的示例演示如何通过在`Document.setSelectedDataAsync` `options`对象中指定可选参数来调用此方法。




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


在这两个可选参数示例中，将_回调_参数指定为最后一个参数（后面是内联可选参数，或在_options_参数对象之后）。或者，可以在内联 JSON __ 对象内或在`options`对象中指定 callback 参数。但是，只能在一个位置中传递_callback_参数：在_options_对象中（内联或在外部创建），或作为最后一个参数，但不能同时传递这两者。


## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md)
- [Office JavaScript API](/office/dev/add-ins/reference/javascript-api-for-office)
