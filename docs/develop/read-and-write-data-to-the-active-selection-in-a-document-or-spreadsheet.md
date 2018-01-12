
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>在文档或电子表格的活动选择内容中读取和写入数据

通过 [Document](../../reference/shared/document.md) 对象公开的方法，你可以读取文档或电子表格中用户的当前选区或向其中写入内容。为此，**Document** 对象提供了 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。本主题还介绍了如何读取、写入和创建事件处理程序，以检测对用户选定内容所做的更改。

**getSelectedDataAsync** 方法仅使用用户当前选区。如果需要在文档中保留选区，以便使用相同的选区在运行加载项的各个会话中读取和写入，必须使用 [Bindings.addFromSelectionAsync](http://msdn.microsoft.com/en-us/library/edc99214-e63e-43f2-9392-97ead42fc155.aspx) 方法添加绑定（或创建一个与 [Bindings](http://msdn.microsoft.com/en-us/library/09979e31-3bfb-45be-adda-0f7cc2db1fe1.aspx) 对象其他“addFrom”方法的绑定）。有关创建对文档区域的绑定，然后读取和写入绑定的信息，请参阅[绑定到文档或电子表格中的区域](../../docs/develop/bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="read-selected-data"></a>读取选择的数据


以下示例演示如何使用 [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) 方法从文档的选定内容中获取数据。


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    }
    else {
        write('Selected data: ' + asyncResult.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

在此示例中，将第一个  _coercionType_ 参数指定为 **Office.CoercionType.Text**（还可以使用文本字符串 `"text"` 指定此参数）。这意味着在回调函数的 [asyncResult](../../reference/shared/asyncresult.status.md) 参数中提供的 [AsyncResult](../../reference/shared/asyncresult.md) 对象的 _value_ 属性将返回一个包含文档中选定文本的 **string**。指定不同的强制类型将产生不同的值。[Office.CoercionType](../../reference/shared/coerciontype-enumeration.md) 是可用的强制类型值的枚举。**Office.CoercionType.Text** 的计算结果为字符串“text”。


 >**提示：** **应该在何时使用矩阵和表格 coercionType 进行数据访问？**如果需要在添加行和列时使所选表格数据动态增大，且必须使用表格标题，则应该使用表格数据类型（通过将 **getSelectedDataAsync** 方法的 _coercionType_ 参数指定为 `"table"` 或 **Office.CoercionType.Table**）。表格数据和矩阵数据中都支持在数据结构内添加行和列，但仅支持对表格数据追加行和列。如果不计划添加行和列，且数据不需要标题功能，则应使用矩阵数据类型（通过将 **getSelecteDataAsync** 方法的 _coercionType_ 参数指定为 `"matrix"` 或 **Office.CoercionType.Matrix**），它提供了与数据交互更简单的模型。

作为第二个  _callback_ 参数传入函数的匿名函数会在 **getSelectedDataAsync** 操作完成时执行。调用该函数时使用单个参数 _asyncResult_，后者包含调用的结果和状态。如果调用失败，则  [AsyncResult](../../reference/shared/asyncresult.context.md) 对象的 **error** 属性会提供对 [Error](../../reference/shared/error.md) 对象的访问。您可以检查 [Error.name](../../reference/shared/error.name.md) 和 [Error.message](../../reference/shared/error.message.md) 属性的值，以确定设置操作失败的原因。否则，会显示文档中选定的文本。

[AsyncResult.status](../../reference/shared/asyncresult.error.md) 属性在 **if** 语句中用于测试调用是否成功。[Office.AsyncResultStatus](../../reference/shared/asyncresultstatus-enumeration.md) 是可用的 **AsyncResult.status** 属性值的枚举。**Office.AsyncResultStatus.Failed** 的计算结果为字符串“failed”（而且，还可以指定为该文本字符串）。


## <a name="write-data-to-the-selection"></a>向选定内容中写入数据


以下示例演示如何将选定内容设置为显示"Hello World!"。


```js
Office.context.document.setSelectedDataAsync("Hello World!", function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

为  _data_ 参数传入不同对象类型会得到不同结果。结果取决于当前在文档中选定的内容、承载加载项的是哪个应用程序以及传入的数据是否可以强制为当前选定内容。

作为  [callback](../../reference/shared/document.setselecteddataasync.md) 参数传入 _setSelectedDataAsync_ 方法的匿名函数在异步调用完成时执行。在您使用 **setSelectedDataAsync** 方法向选定内容中写入数据时，回调的 _asyncResult_ 参数只提供对调用状态以及 [Error](../../reference/shared/error.md) 对象（如果调用失败）的访问。

 **注意：**从 Excel 2013 SP1 的发行版及相应的 Excel Online 内部版本开始，你现在可以[在将表写入当前选定内容时设置格式](../../docs/excel/format-tables-in-add-ins-for-excel.md)。


## <a name="detect-changes-in-the-selection"></a>检测选定内容中的更改


以下示例演示如何通过使用 [Document.addHandlerAsync](../../reference/shared/document.addhandlerasync.md) 方法为文档中的 [SelectionChanged](../../reference/shared/document.selectionchanged.event.md) 事件添加事件处理程序来检测选定内容中的更改。


```
Office.context.document.addHandlerAsync("documentSelectionChanged", myHandler, function(result){} 
);

// Event handler function.
function myHandler(eventArgs){
write('Document Selection Changed');
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

第一个  _eventType_ 参数指定要订阅的事件的名称。传递此参数的 `"documentSelectionChanged"` 字符串等同于传递 **Office.EventType** 枚举的 [Office.EventType.DocumentSelectionChanged](../../reference/shared/eventtype-enumeration.md) 事件类型。

作为第二个 _handler_ 参数传入函数的 `myHander()` 函数是在文档中的选定内容更改时执行的事件处理程序。调用该函数时使用单个参数 _eventArgs_，后者在异步操作完成时将包含对 [DocumentSelectionChangedEventArgs](../../reference/shared/document.selectionchangedeventargs.md) 对象的引用。可以使用 [DocumentSelectionChangedEventArgs.document](../../reference/shared/document.selectionchangedeventargs.document.md) 属性访问引发事件的文档。


 >**注释**  可以为给定事件添加多个事件处理程序，方法是再次调用  **addHandlerAsync** 方法，并为 _handler_ 参数传入一个其他事件处理程序函数。只要每个事件处理程序函数的名称保持唯一，此方法就有用。


## <a name="stop-detecting-changes-in-the-selection"></a>停止检测选定内容中的更改


以下示例演示如何通过调用 [document.removeHandlerAsync](../../reference/shared/document.selectionchanged.event.md) 方法停止侦听 [Document.SelectionChanged](../../reference/shared/document.removehandlerasync.md) 事件。


```
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

作为第二个  _handler_ 参数传递的 `myHandler` 函数名称指定将从 **SelectionChanged** 事件中移除的事件处理程序。


 >**重要说明：**  如果调用  _removeHandlerAsync_ 方法时省略可选的 **handler** 参数，则会移除指定的 _eventType_ 的所有事件处理程序。

