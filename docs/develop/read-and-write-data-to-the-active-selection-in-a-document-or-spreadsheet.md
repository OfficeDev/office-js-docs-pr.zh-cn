---
title: 对文档或电子表格中的活动选择执行数据读取和写入操作
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 2fe847fcc04e3670db294a421388dbd2faad6f2f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871449"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>对文档或电子表格中的活动选择执行数据读取和写入操作

通过 [Document](/javascript/api/office/office.document) 对象公开的方法，你可以读取文档或电子表格中用户的当前选区或向其中写入内容。为此，**Document** 对象提供了 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。本主题还介绍了如何读取、写入和创建事件处理程序，以检测对用户选定内容所做的更改。

**getSelectedDataAsync** 方法仅使用用户当前选区。如果需要在文档中保留选区，以便使用相同的选区在运行加载项的各个会话中读取和写入，必须使用 [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) 方法添加绑定（或创建一个与 [Bindings](/javascript/api/office/office.bindings) 对象其他“addFrom”方法的绑定）。有关创建对文档区域的绑定，然后读取和写入绑定的信息，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="read-selected-data"></a>读取选择的数据


以下示例演示如何使用 [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 方法从文档的选定内容中获取数据。


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

在此示例中，将第一个  _coercionType_ 参数指定为 **Office.CoercionType.Text**（还可以使用文本字符串 `"text"` 指定此参数）。这意味着在回调函数的 [asyncResult](/javascript/api/office/office.asyncresult#status) 参数中提供的 [AsyncResult](/javascript/api/office/office.asyncresult) 对象的 _value_ 属性将返回一个包含文档中选定文本的 **string**。指定不同的强制类型将产生不同的值。[Office.CoercionType](/javascript/api/office/office.coerciontype) 是可用的强制类型值的枚举。**Office.CoercionType.Text** 的计算结果为字符串“text”。


> [!TIP]
> **何时应使用矩阵与表格 coercionType 数据访问？** 如果需要表格数据在添加行和列时动态增长，且必须处理表格标题，应使用表格数据类型（具体操作是将 **getSelectedDataAsync** 方法的 _coercionType_ 参数指定为 `"table"` 或 **Office.CoercionType.Table**）。虽然表格数据和矩阵数据都支持在数据结构内添加行和列，但只有表格数据支持追加行和列。如果不打算添加行和列，且数据不需要使用标题功能，应使用矩阵数据类型（具体操作是将 **getSelecteDataAsync** 方法的 _coercionType_ 参数指定为 `"matrix"` 或 **Office.CoercionType.Matrix**），它提供了更简单的数据交互模型。

作为第二个  _callback_ 参数传入函数的匿名函数会在 **getSelectedDataAsync** 操作完成时执行。调用该函数时使用单个参数 _asyncResult_，后者包含调用的结果和状态。如果调用失败，则  [AsyncResult](/javascript/api/office/office.asyncresult#asynccontext) 对象的 **error** 属性会提供对 [Error](/javascript/api/office/office.error) 对象的访问。您可以检查 [Error.name](/javascript/api/office/office.error#name) 和 [Error.message](/javascript/api/office/office.error#message) 属性的值，以确定设置操作失败的原因。否则，会显示文档中选定的文本。

[AsyncResult.status](/javascript/api/office/office.asyncresult#error) 属性在 **if** 语句中用于测试调用是否成功。[Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) 是可用的 **AsyncResult.status** 属性值的枚举。**Office.AsyncResultStatus.Failed** 的计算结果为字符串“failed”（而且，还可以指定为该文本字符串）。


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

作为  [callback](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 参数传入 _setSelectedDataAsync_ 方法的匿名函数在异步调用完成时执行。在您使用 **setSelectedDataAsync** 方法向选定内容中写入数据时，回调的 _asyncResult_ 参数只提供对调用状态以及 [Error](/javascript/api/office/office.error) 对象（如果调用失败）的访问。

> [!NOTE]
> 自 Excel 2013 SP1 发行版及相应的 Excel Online 版本起，现在可以[在将表格写入当前选择时设置格式](../excel/excel-add-ins-tables.md)。


## <a name="detect-changes-in-the-selection"></a>检测选择中的更改


以下示例演示如何通过使用 [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) 方法为文档中的 [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) 事件添加事件处理程序来检测选定内容中的更改。


```js
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

第一个  _eventType_ 参数指定要订阅的事件的名称。传递此参数的 `"documentSelectionChanged"` 字符串等同于传递 **Office.EventType** 枚举的 [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.eventtype) 事件类型。

作为第二个 _handler_ 参数传入函数的 `myHander()` 函数是在文档中的选定内容更改时执行的事件处理程序。调用该函数时使用单个参数 _eventArgs_，后者在异步操作完成时将包含对 [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) 对象的引用。可以使用 [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) 属性访问引发事件的文档。


> [!NOTE]
> 可以为给定事件添加多个事件处理程序，具体方法是再次调用 **addHandlerAsync** 方法，并为 _handler_ 参数传入其他事件处理程序函数。只要每个事件处理程序函数的名称是唯一的，这样做就可行。


## <a name="stop-detecting-changes-in-the-selection"></a>停止检测选择中的更改


以下示例演示如何通过调用 [document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) 方法停止侦听 [Document.SelectionChanged](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) 事件。


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

`myHandler` 函数名称作为第二个 _handler_ 参数传递，指定了将从 **SelectionChanged** 事件中删除的事件处理程序。


> [!IMPORTANT]
> 如果调用 **removeHandlerAsync** 方法时省略了可选的 _handler_ 参数，将会删除指定 _eventType_ 的所有事件处理程序。
