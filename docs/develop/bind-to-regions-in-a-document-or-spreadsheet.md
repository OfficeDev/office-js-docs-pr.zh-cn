---
title: 绑定到文档或电子表格中的区域
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: c927f5ceb6be1ad038185e54706a55ab21b3f63a
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324629"
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a>绑定到文档或电子表格中的区域

基于绑定的数据访问使内容和任务窗格加载项能够通过与绑定相关联的标识符一致地访问文档或电子表格的特定区域。加载项首先需要通过调用将文档的某一部分与唯一标识符相关联的以下某个方法来建立绑定：[addFromPromptAsync]、[addFromSelectionAsync] 或 [addFromNamedItemAsync]。建立绑定后，加载项可以使用提供的标识符访问文档或电子表格的关联区域中包含的数据。创建绑定可为加载项提供以下值：


- 允许访问跨支持的 Office 应用程序的通用数据结构，例如：表、区域或文本（一系列连续字符）。

- 允许读/写操作，而不需要用户做出选择。

- 在加载项和文档中的数据之间建立关系。绑定会保留在文档中，以后可以进行访问。

建立绑定还允许您订阅仅限文档或电子表格的特定区域的数据和选择更改事件。这意味着，加载项只会收到绑定区域内发生的更改的通知，而不是收到整个文档或电子表格内的常规更改的通知。

[Bindings] 对象公开 [getAllAsync] 方法，通过该方法可以访问在文档或电子表格中建立的所有绑定的集合。可使用 Bindings.[getByIdAsync] 或 [Office.select] 方法通过 ID 访问单个绑定。可使用 [Bindings] 对象的以下方法之一建立新绑定和删除现有绑定：[addFromSelectionAsync]、[addFromPromptAsync]、[addFromNamedItemAsync] 或 [releaseByIdAsync]。


## <a name="binding-types"></a>绑定类型

使用[addFromSelectionAsync]、 [addFromPromptAsync]或[addFromNamedItemAsync]方法创建绑定时，可以使用_BindingType_参数指定[三种不同类型的绑定][BindingType] ：

1. **[文本绑定][TextBinding]** - 绑定到可以文本形式表示的文档区域。

    在 Word 中，大多数连续选区都是有效的，而在 Excel 中，只有单个单元格选区才能作为文本绑定的目标。在 Excel 中，只支持纯文本。在 Word 中，支持以下三种格式：纯文本、HTML 和 Open XML for Office。

2. **[矩阵绑定][MatrixBinding] ** -绑定到包含不带标题的表格数据的文档的固定区域。矩阵绑定中的数据作为二维**数组**写入或读取，在 JavaScript 中实现为数组的数组。例如，可以写入或读取两行两列的**字符串**值` [['a', 'b'], ['c', 'd']]`，并且可以写入或读取包含三行的单列`[['a'], ['b'], ['c']]`。

    在 Excel 中，任何连续的单元格选区都可用于建立矩阵绑定。在 Word 中，只有表格支持矩阵绑定。

3. **[表格绑定][TableBinding]** - 绑定到包含带标题的表格的文档区域。表格绑定中的数据作为 [TableData](/javascript/api/office/office.tabledata) 对象写入或读取。`TableData` 对象通过 `headers` 和 `rows` 属性公开数据。

    任何 Excel 或 Word 表格均可作为表格绑定的基础。建立表格绑定后，用户添加到表格中的每个新行或新列都自动包含在绑定中。

使用`Bindings`对象的三个 "addFrom" 方法之一创建绑定之后，可以使用相应对象的方法来处理绑定的数据和属性： [MatrixBinding]、 [TableBinding]或[textbinding 添加]。这三个对象都继承`Binding`对象的[binding.getdataasync]和[binding.setdataasync]方法，使您能够与绑定的数据进行交互。

> [!NOTE]
> **应该何时使用矩阵和表格绑定？** 当使用的表格数据包含一个总计行时，如果外接程序的脚本需要访问总计行中的值，或检测用户的选区是否在总计行中，则必须使用矩阵绑定。如果为包含总计行的表格数据建立了表格绑定，那么 [TableBinding.rowCount] 属性和事件处理程序中 [BindingSelectionChangedEventArgs] 对象的 `rowCount` 和 `startRow` 属性将不会在值中反映总计行。要解决此限制，必须建立矩阵绑定以使用总计行。

## <a name="add-a-binding-to-the-users-current-selection"></a>向用户当前所选内容中添加绑定

以下示例显示如何使用 [addFromSelectionAsync] 方法向文档中的当前所选内容中添加名为 `myBinding` 的文本绑定。


```js
Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在此示例中，指定的绑定类型为文本。这意味着将为所选内容创建 [TextBinding]。不同的绑定类型会公开不同的数据和操作。[Office.BindingType] 是可用的绑定类型值的枚举。

第二个可选参数是一个对象，它指定要创建的新绑定的 ID。如果不指定 ID，则会自动生成一个。

作为最后一个 _callback_ 参数传入函数的匿名函数会在绑定创建完成时执行。该函数用单个参数 `asyncResult` 来调用，通过该参数可访问提供调用状态的 [AsyncResult] 对象。`AsyncResult.value` 属性包含对 [Binding] 对象的引用，该对象属于为新创建的绑定指定的类型。可以使用此 [Binding] 对象来获取和设置数据。

## <a name="add-a-binding-from-a-prompt"></a>从提示中添加绑定

以下示例显示如何使用使用 [addFromPromptAsync] 方法添加名为 `myBinding` 的文本绑定。此方法允许用户使用应用程序内置的范围选择提示来指定绑定范围。


```js
function bindFromPrompt() {
    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
        }
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

在此示例中，指定的绑定类型为文本。这意味着，将为用户在提示中指定的所选内容创建 [TextBinding]。

第二个参数是一个对象，它包含创建的新绑定的 ID。如果不指定 ID，则会自动生成一个。

作为第三个  _callback_ 参数传入函数的匿名函数会在绑定创建完成时执行。执行回调函数时， [AsyncResult] 对象包含调用的状态和新创建的绑定。

图 1 显示 Excel 中内置的范围选择提示。


*图 1.Excel 选择数据 UI*

![Excel 选择数据 UI](../images/agave-api-overview-excel-selection-ui.png)


## <a name="add-a-binding-to-a-named-item"></a>向已命名项目添加绑定


下面的示例演示如何使用[addFromNamedItemAsync]方法将绑定添加到`myRange`现有的已命名项目，作为 "matrix" 绑定，并将该绑定分配`id`为 "myMatrix"。


```js
function bindNamedItem() {
    Office.context.document.bindings.addFromNamedItemAsync("myRange", "matrix", {id:'myMatrix'}, function (result) {
        if (result.status == 'succeeded'){
            write('Added new binding with type: ' + result.value.type + ' and id: ' + result.value.id);
            }
        else
            write('Error: ' + result.error.message);
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

**对于 Excel**，addFromNamedItemAsync `itemName`方法的参数[]可以引用现有的命名区域、使用`A1`引用样式`("A1:A3")`指定的区域或表。默认情况下，在 Excel 中添加表会为您添加的第一个表指定名称 "Table1"，为您添加的第二个表指定 "Table2"，依此类推。若要在 Excel UI 中为表分配有意义的名称，请`Table Name`使用 "表格工具" 上的属性 **|** 功能区的 "设计" 选项卡。


> [!NOTE]
> 在 Excel 中，当指定表作为已命名项目时，必须完全限定该名称以在表名称中包括工作表名称，格式如下：`"Sheet1!Table1"`

下面的示例在 Excel 中创建绑定到列 A （ `"A1:A3"`）中的前三个单元格，分配`"MyCities"`id，然后向该绑定写入三个城市名称。


```js
 function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", "matrix", {id: "MyCities" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                write('Error: ' + asyncResult.error.message);
            }
            else {
                // Write data to the new binding.
                Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            write('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

**对于 Word**，addFromNamedItemAsync `itemName`方法的参数[]引用`Title` `Rich Text`内容控件的属性。（不能绑定到内容控件之外的`Rich Text`内容控件。）

默认情况下，内容控件没有`Title*`分配值。若要在 Word UI 中分配有意义的名称，请在功能区的 "**开发人员**" 选项卡上的 "**控件**" 组中插入**格式文本**内容控件后，使用 "**控件**" 组中的 "**属性**" 命令显示 "**内容控件属性**" 对话框。然后，将`Title`内容控件的属性设置为要从代码中引用的名称。

下面的示例在 Word 中创建一个文本绑定，以指定一个名为`"FirstName"`的格式文本内容控件，分配该**id** `"firstName"`，然后显示该信息。


```js
function bindContentControl() {
    Office.context.document.bindings.addFromNamedItemAsync('FirstName', 
        Office.BindingType.Text, {id:'firstName'},
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                write('Control bound. Binding.id: '
                    + result.value.id + ' Binding.type: ' + result.value.type);
            } else {
                write('Error:', result.error.message);
            }
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

## <a name="get-all-bindings"></a>获取所有绑定


以下示例显示如何使用 Bindings.[getAllAsync] 方法获取文档中的所有绑定。


```js
Office.context.document.bindings.getAllAsync(function (asyncResult) {
    var bindingString = '';
    for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
    }
    write('Existing bindings: ' + bindingString);
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在操作完成时，将执行作为`callback`参数传入函数的匿名函数。使用单个参数`asyncResult`调用函数，其中包含文档中的绑定数组。对数组进行迭代，以生成包含绑定的 Id 的字符串。然后，在消息框中显示该字符串。


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a>使用 Bindings 对象的 getByIdAsync 方法按 ID 获取绑定


以下示例显示如何使用 [getByIdAsync] 方法通过指定绑定的 ID 获取文档中的绑定。此示例假定已使用本主题前面介绍的方法之一将名为 `'myBinding'` 的绑定添加到文档。


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } 
    else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在此示例中，第`id`一个参数是要检索的绑定的 ID。

在操作完成时，将执行作为第二个_回调_参数传入函数的匿名函数。使用单个参数_asyncResult_调用函数，该参数包含调用的状态和 ID 为 "myBinding" 的绑定。


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a>使用 Office 对象的 select 方法按 ID 获取绑定


以下示例演示如何使用 [Office.select] 方法通过在选择器字符串中指定 [Binding] 对象目标 ID 来获取文档中的该目标。然后，它会调用 Binding.[getDataAsync] 方法，从指定绑定中获取数据。此示例假定已使用本主题前面介绍的方法之一将名为 `'myBinding'` 的绑定添加到文档。


```js
Office.select("bindings#myBinding", function onError(){}).getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


> [!NOTE]
> 如果`select`方法承诺成功返回一个[Binding]对象，则该对象仅公开对象的以下四个方法： [binding.getdataasync]、 [binding.setdataasync]、 [addHandlerAsync]和[removeHandlerAsync]。如果承诺无法返回 Binding 对象，则`onError`可以使用回调访问[asyncResult]. error 对象，以获取详细信息。如果`select`需要调用 binding 对象的成员，而不是由方法返回的[绑定]对象承诺所公开的四个方法，请改为使用[getByIdAsync]方法，方法是使用[Document. bindings]属性和绑定。用于检索[Binding]对象的[getByIdAsync]方法。

## <a name="release-a-binding-by-id"></a>按 ID 释放绑定


以下示例显示如何使用 [releaseByIdAsync] 方法通过指定绑定的 ID 释放文档中的绑定。

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在此示例中，第一个 `id` 参数是要释放的绑定的 ID。

作为第二个参数传入函数的匿名函数是在操作完成时执行的回调。该函数用单个参数  [asyncResult] 来调用，其中包含调用的状态。


## <a name="read-data-from-a-binding"></a>从绑定中读取数据


以下示例显示如何使用 [getDataAsync] 方法从现有绑定中获取数据。


```js
myBinding.getDataAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write(asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 `myBinding` 是包含文档中的现有文本绑定的变量。也可以使用 [Office.select] 按照其 ID 访问绑定，并启动对 [getDataAsync] 方法的调用，如下所示： 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


传入函数的匿名函数是在操作完成时执行的回调。[AsyncResult].value 属性包含 `myBinding` 中的数据。值的类型取决于绑定类型。此示例中的绑定是文本绑定。因此，该值将包含字符串。有关使用矩阵和表格绑定的其他示例，请参阅 [getDataAsync] 方法主题。


## <a name="write-data-to-a-binding"></a>向绑定中写入数据

以下示例演示如何使用 [setDataAsync] 方法在现有绑定中设置数据。

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 `myBinding` 是包含文档中的现有文本绑定的变量。

在此示例中，第一个参数是要设置的`myBinding`值。由于这是文本绑定，因此值为`string`。不同的绑定类型接受不同类型的数据。

传入函数的匿名函数是在操作完成时执行的回调。使用单个参数`asyncResult`调用函数，其中包含结果的状态。

> [!NOTE]
> 自 Excel 2013 SP1 发行版及相应的 Excel 网页版内部版本起，现在可以[在绑定表中写入和更新数据时设置格式](../excel/excel-add-ins-tables.md)。


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a>检测绑定中数据或选择内的更改


下面的示例展示了如何向 ID 为“MyBinding”的绑定的 [DataChanged](/javascript/api/office/office.binding) 事件附加事件处理程序。


```js
function addHandler() {
Office.select("bindings#MyBinding").addHandlerAsync(
    Office.EventType.BindingDataChanged, dataChanged);
}
function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

`myBinding` 是包含文档中现有文本绑定的变量。

[AddHandlerAsync]方法_的第一个事件名称_参数指定要订阅的事件的名称。" [. 事件]类型" 是可用事件类型值的枚举。`Office.EventType.BindingDataChanged`计算结果为字符串 "bindingDataChanged"。

作为`dataChanged`第二个_handler_参数传递到函数中的函数是在绑定中的数据更改时执行的事件处理程序。使用单个参数_eventArgs_调用函数，其中包含对绑定的引用。此绑定可用于检索更新后的数据。

类似地，你可以通过向绑定的 [SelectionChanged] 事件附加事件处理程序来检测用户是否更改绑定中的选择。为此，请将 [addHandlerAsync] 方法的 `eventType` 参数指定为 `Office.EventType.BindingSelectionChanged` 或 `"bindingSelectionChanged"`。

可以为给定事件添加多个事件处理程序，方法是再次调用  [addHandlerAsync] 方法，并为 `handler` 参数传入一个其他事件处理程序函数。只要每个事件处理程序函数的名称保持唯一，此方法就有用。


### <a name="remove-an-event-handler"></a>删除事件处理程序


若要删除事件的事件处理程序，请调用 [removeHandlerAsync] 方法将事件类型作为第一个 _eventType_ 参数传入，将要删除的事件处理程序函数的名称作为第二个 _handler_ 参数传入。例如，以下函数将删除在上一节的示例中添加的 `dataChanged` 事件处理程序函数。


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


> [!IMPORTANT]
> 如果调用[removeHandlerAsync]方法时省略可选的`eventType` _handler_参数，则将删除指定的所有事件处理程序。


## <a name="see-also"></a>另请参阅

- [了解 Office JavaScript API](understanding-the-javascript-api-for-office.md) 
- [Office 外接程序中的异步编程](asynchronous-programming-in-office-add-ins.md)
- [读取数据并将其写入文档或电子表格中的活动选择区](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)

[Binding]:               /javascript/api/office/office.binding
[MatrixBinding]:         /javascript/api/office/office.matrixbinding
[TableBinding]:          /javascript/api/office/office.tablebinding
[TextBinding]:           /javascript/api/office/office.textbinding
[getDataAsync]:          /javascript/api/office/Office.Binding#getdataasync-options--callback-
[setDataAsync]:          /javascript/api/office/Office.Binding#setdataasync-data--options--callback-
[SelectionChanged]:      /javascript/api/office/office.bindingselectionchangedeventargs
[addHandlerAsync]:       /javascript/api/office/Office.Binding#addhandlerasync-eventtype--handler--options--callback-
[removeHandlerAsync]:    /javascript/api/office/Office.Binding#removehandlerasync-eventtype--options--callback-

[Bindings]:              /javascript/api/office/office.bindings
[getByIdAsync]:          /javascript/api/office/office.bindings#getbyidasync-id--options--callback- 
[getAllAsync]:           /javascript/api/office/office.bindings#getallasync-options--callback-
[addFromNamedItemAsync]: /javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-
[addFromSelectionAsync]: /javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-
[addFromPromptAsync]:    /javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-
[releaseByIdAsync]:      /javascript/api/office/office.bindings#releasebyidasync-id--options--callback-

[AsyncResult]:          /javascript/api/office/office.asyncresult
[Office.BindingType]:   /javascript/api/office/office.bindingtype
[Office.select]:        /javascript/api/office 
[Office.EventType]:     /javascript/api/office/office.eventtype 
[Document.bindings]:    /javascript/api/office/office.document


[TableBinding.rowCount]: /javascript/api/office/office.tablebinding
[BindingSelectionChangedEventArgs]: /javascript/api/office/office.bindingselectionchangedeventargs
