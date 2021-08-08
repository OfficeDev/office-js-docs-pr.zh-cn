---
title: 对文档或电子表格中的活动选择执行数据读取和写入操作
description: 了解如何在 Word 文档或电子表格的活动选定内容中读取和写入Excel数据。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 19cd1edad3835403355e07cdfcd4a43cb2aa9c9b7746823e2fd31aea319b9bf2
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080265"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>对文档或电子表格中的活动选择执行数据读取和写入操作

通过 [Document](/javascript/api/office/office.document) 对象公开的方法，你可以读取文档或电子表格中用户的当前选区或向其中写入内容。 为此，对象提供 `Document` `getSelectedDataAsync` 和 `setSelectedDataAsync` 方法。 本主题还介绍了如何读取、写入和创建事件处理程序，以检测对用户选定内容所做的更改。

`getSelectedDataAsync`该方法仅适用于用户的当前选择。 如果需要在文档中保留选区，以便使用相同的选区在运行加载项的各个会话中读取和写入，必须使用 [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_) 方法添加绑定（或创建一个与 [Bindings](/javascript/api/office/office.bindings) 对象其他“addFrom”方法的绑定）。 有关创建对文档区域的绑定，然后读取和写入绑定的信息，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。


## <a name="read-selected-data"></a>读取选择的数据


以下示例演示如何使用 [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) 方法从文档的选定内容中获取数据。


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

本示例中，第一个  _coercionType_ 参数指定为 (您也可以通过使用文本字符串 `Office.CoercionType.Text` `"text"`) 。 这意味着在回调函数的 [asyncResult](/javascript/api/office/office.asyncresult#status) 参数中提供的 [AsyncResult](/javascript/api/office/office.asyncresult) 对象的 _value_ 属性将返回一个包含文档中选定文本的 **string**。 指定不同强制类型将产生不同的值。 [Office.CoercionType](/javascript/api/office/office.coerciontype) 是可用的强制类型值的枚举。 `Office.CoercionType.Text` 计算结果为字符串"text"。


> [!TIP]
> **何时应使用矩阵与表格 coercionType 数据访问？** 如果需要在添加行和列时使选定的表格数据动态增长，并且必须使用表格标题，则应该将方法的 _coercionType_ 参数指定为 或 数据类型 (，从而使用 `getSelectedDataAsync` `"table"` `Office.CoercionType.Table`) 。 表格数据和矩阵数据中都支持在数据结构内添加行和列，但仅支持对表格数据追加行和列。 如果您不计划添加行和列，并且数据不需要标题功能，则应该通过将 method 的  _coercionType_ 参数指定为 或) 来使用矩阵 `getSelectedDataAsync` `"matrix"` `Office.CoercionType.Matrix` 数据类型 (，从而提供与数据交互的更简单的模型。

作为第二个  _callback_ 参数传入函数的匿名函数在操作完成时 `getSelectedDataAsync` 执行。 调用该函数时使用单个参数 _asyncResult_，后者包含调用的结果和状态。 如果调用失败，对象的 [error](/javascript/api/office/office.asyncresult#error) 属性 `AsyncResult` 将提供对 [Error 对象的](/javascript/api/office/office.error) 访问。 您可以检查 [Error.name](/javascript/api/office/office.error#name) 和 [Error.message](/javascript/api/office/office.error#message) 属性的值，以确定设置操作失败的原因。 否则，会显示文档中选定的文本。

[AsyncResult.status](/javascript/api/office/office.asyncresult#error) 属性在 **if** 语句中用于测试调用是否成功。 [Office。AsyncResultStatus](/javascript/api/office/office.asyncresult#status)是可用属性值 `AsyncResult.status` 的枚举。 `Office.AsyncResultStatus.Failed` 计算结果为字符串"failed" (，同样，还可以指定为该文本字符串) 。


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

为  _data_ 参数传入不同对象类型会得到不同结果。 结果取决于文档中当前选择的内容、Office哪个客户端应用程序托管您的外接程序，以及传入的数据是否可以强制转换为当前选择。

作为  [callback](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) 参数传入 _setSelectedDataAsync_ 方法的匿名函数在异步调用完成时执行。 使用 方法将数据写入选区时，回调的 `setSelectedDataAsync` _asyncResult_ 参数仅提供对调用状态和 [Error](/javascript/api/office/office.error) 对象的访问（如果调用失败）。

> [!NOTE]
> 自 Excel 2013 SP1 发行版及相应的 Excel 网页版起，现在可以[在将表格写入当前选择时设置格式](../excel/excel-add-ins-tables.md)。


## <a name="detect-changes-in-the-selection"></a>检测选择中的更改


以下示例演示如何通过使用 [Document.addHandlerAsync](/javascript/api/office/office.document#addHandlerAsync_eventType__handler__options__callback_) 方法为文档中的 [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) 事件添加事件处理程序来检测选定内容中的更改。


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

第一个  _eventType_ 参数指定要订阅的事件的名称。 为此参数 `"documentSelectionChanged"` 传递字符串等效于传递事件 `Office.EventType.DocumentSelectionChanged` 类型的[Office。EventType](/javascript/api/office/office.eventtype)枚举。

作为第二个 _handler_ 参数传入函数的 `myHander()` 函数是在文档中的选定内容更改时执行的事件处理程序。调用该函数时使用单个参数 _eventArgs_，后者在异步操作完成时将包含对 [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) 对象的引用。可以使用 [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) 属性访问引发事件的文档。


> [!NOTE]
> 你可以为给定事件添加多个事件处理程序，方法是再次调用 该方法，然后为 handler 参数传递一 `addHandlerAsync` 个额外的事件 _处理程序_ 函数。 只要每个事件处理程序函数的名称保持唯一，此方法就有用。


## <a name="stop-detecting-changes-in-the-selection"></a>停止检测选择中的更改


以下示例演示如何通过调用 [document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) 方法停止侦听 [Document.SelectionChanged](/javascript/api/office/office.document#removeHandlerAsync_eventType__options__callback_) 事件。


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

`myHandler`作为第二个 _handler_ 参数传递的函数名称指定将从事件中删除 `SelectionChanged` 的事件处理程序。


> [!IMPORTANT]
> 如果在调用  _方法_ 时省略可选的 handler 参数，将删除指定 `removeHandlerAsync` _eventType_ 的所有事件处理程序。
