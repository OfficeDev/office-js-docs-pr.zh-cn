---
title: 对文档或电子表格中的活动选择执行数据读取和写入操作
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 76f6f5f6a2d117b59e1a7794e35e181383022269
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457885"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="d0326-102">对文档或电子表格中的活动选择执行数据读取和写入操作</span><span class="sxs-lookup"><span data-stu-id="d0326-102">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="d0326-p101">通过 [Document](https://docs.microsoft.com/javascript/api/office/office.document) 对象公开的方法，你可以读取文档或电子表格中用户的当前选区或向其中写入内容。为此，**Document** 对象提供了 **getSelectedDataAsync** 和 **setSelectedDataAsync** 方法。本主题还介绍了如何读取、写入和创建事件处理程序，以检测对用户选定内容所做的更改。</span><span class="sxs-lookup"><span data-stu-id="d0326-p101">The [Document](https://docs.microsoft.com/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet. To do that, the **Document** object provides the **getSelectedDataAsync** and **setSelectedDataAsync** methods. This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="d0326-p102">**getSelectedDataAsync** 方法仅使用用户当前选区。如果需要在文档中保留选区，以便使用相同的选区在运行加载项的各个会话中读取和写入，必须使用 [Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) 方法添加绑定（或创建一个与 [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings) 对象其他“addFrom”方法的绑定）。有关创建对文档区域的绑定，然后读取和写入绑定的信息，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="d0326-p102">The  **getSelectedDataAsync** method only works against the user's current selection. If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings) object). For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="d0326-109">读取选择的数据</span><span class="sxs-lookup"><span data-stu-id="d0326-109">Read selected data</span></span>


<span data-ttu-id="d0326-110">以下示例演示如何使用 [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 方法从文档的选定内容中获取数据。</span><span class="sxs-lookup"><span data-stu-id="d0326-110">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


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

<span data-ttu-id="d0326-p103">在此示例中，将第一个  _coercionType_ 参数指定为 **Office.CoercionType.Text**（还可以使用文本字符串 `"text"` 指定此参数）。这意味着在回调函数的 [asyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult#status) 参数中提供的 [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult) 对象的 _value_ 属性将返回一个包含文档中选定文本的 **string**。指定不同的强制类型将产生不同的值。[Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype) 是可用的强制类型值的枚举。**Office.CoercionType.Text** 的计算结果为字符串“text”。</span><span class="sxs-lookup"><span data-stu-id="d0326-p103">In this example, the first  _coercionType_ parameter is specified as **Office.CoercionType.Text** (you can also specify this parameter by using the literal string `"text"`). This means that the [value](https://docs.microsoft.com/javascript/api/office/office.asyncresult#status) property of the [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document. Specifying different coercion types will result in different values. [Office.CoercionType](https://docs.microsoft.com/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values. **Office.CoercionType.Text** evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="d0326-p104">**何时应使用矩阵与表格 coercionType 数据访问？** 如果需要表格数据在添加行和列时动态增长，且必须处理表格标题，应使用表格数据类型（具体操作是将 **getSelectedDataAsync** 方法的 _coercionType_ 参数指定为 `"table"` 或 **Office.CoercionType.Table**）。虽然表格数据和矩阵数据都支持在数据结构内添加行和列，但只有表格数据支持追加行和列。如果不打算添加行和列，且数据不需要使用标题功能，应使用矩阵数据类型（具体操作是将 **getSelecteDataAsync** 方法的 _coercionType_ 参数指定为 `"matrix"` 或 **Office.CoercionType.Matrix**），它提供了更简单的数据交互模型。</span><span class="sxs-lookup"><span data-stu-id="d0326-p104">**When should you use the matrix versus table coercionType for data access?** If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the **getSelectedDataAsync** method as `"table"` or **Office.CoercionType.Table**). Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data. If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of **getSelecteDataAsync** method as `"matrix"` or **Office.CoercionType.Matrix**), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="d0326-p105">作为第二个  _callback_ 参数传入函数的匿名函数会在 **getSelectedDataAsync** 操作完成时执行。调用该函数时使用单个参数 _asyncResult_，后者包含调用的结果和状态。如果调用失败，则  [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult#asynccontext) 对象的 **error** 属性会提供对 [Error](https://docs.microsoft.com/javascript/api/office/office.error) 对象的访问。您可以检查 [Error.name](https://docs.microsoft.com/javascript/api/office/office.error#name) 和 [Error.message](https://docs.microsoft.com/javascript/api/office/office.error#message) 属性的值，以确定设置操作失败的原因。否则，会显示文档中选定的文本。</span><span class="sxs-lookup"><span data-stu-id="d0326-p105">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the **getSelectedDataAsync** operation is completed. The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call. If the call fails, the [error](https://docs.microsoft.com/javascript/api/office/office.asyncresult#asynccontext) property of the **AsyncResult** object provides access to the [Error](https://docs.microsoft.com/javascript/api/office/office.error) object. You can check the value of the [Error.name](https://docs.microsoft.com/javascript/api/office/office.error#name) and [Error.message](https://docs.microsoft.com/javascript/api/office/office.error#message) properties to determine why the set operation failed. Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="d0326-p106">[AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult#error) 属性在 **if** 语句中用于测试调用是否成功。[Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult#status) 是可用的 **AsyncResult.status** 属性值的枚举。**Office.AsyncResultStatus.Failed** 的计算结果为字符串“failed”（而且，还可以指定为该文本字符串）。</span><span class="sxs-lookup"><span data-stu-id="d0326-p106">The [AsyncResult.status](https://docs.microsoft.com/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded. [Office.AsyncResultStatus](https://docs.microsoft.com/javascript/api/office/office.asyncresult#status) is an enumeration of available **AsyncResult.status** property values. **Office.AsyncResultStatus.Failed** evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="d0326-128">向选定内容中写入数据</span><span class="sxs-lookup"><span data-stu-id="d0326-128">Write data to the selection</span></span>


<span data-ttu-id="d0326-129">以下示例演示如何将选定内容设置为显示"Hello World!"。</span><span class="sxs-lookup"><span data-stu-id="d0326-129">The following example shows how to set the selection to show "Hello World!".</span></span>


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

<span data-ttu-id="d0326-p107">为  _data_ 参数传入不同对象类型会得到不同结果。结果取决于当前在文档中选定的内容、承载加载项的是哪个应用程序以及传入的数据是否可以强制为当前选定内容。</span><span class="sxs-lookup"><span data-stu-id="d0326-p107">Passing in different object types for the  _data_ parameter will have different results. The result depends on what is currently selected in the document, which application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="d0326-p108">作为  [callback](https://docs.microsoft.com/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 参数传入 _setSelectedDataAsync_ 方法的匿名函数在异步调用完成时执行。在您使用 **setSelectedDataAsync** 方法向选定内容中写入数据时，回调的 _asyncResult_ 参数只提供对调用状态以及 [Error](https://docs.microsoft.com/javascript/api/office/office.error) 对象（如果调用失败）的访问。</span><span class="sxs-lookup"><span data-stu-id="d0326-p108">The anonymous function passed into the [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed. When you write data to the selection by using the **setSelectedDataAsync** method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](https://docs.microsoft.com/javascript/api/office/office.error) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="d0326-134">自 Excel 2013 SP1 发行版及相应的 Excel Online 版本起，现在可以[在将表格写入当前选择时设置格式](../excel/excel-add-ins-tables.md)。</span><span class="sxs-lookup"><span data-stu-id="d0326-134">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="d0326-135">检测选择中的更改</span><span class="sxs-lookup"><span data-stu-id="d0326-135">Detect changes in the selection</span></span>


<span data-ttu-id="d0326-136">以下示例演示如何通过使用 [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) 方法为文档中的 [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) 事件添加事件处理程序来检测选定内容中的更改。</span><span class="sxs-lookup"><span data-stu-id="d0326-136">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) event on the document.</span></span>


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

<span data-ttu-id="d0326-p109">第一个  _eventType_ 参数指定要订阅的事件的名称。传递此参数的 `"documentSelectionChanged"` 字符串等同于传递 **Office.EventType** 枚举的 [Office.EventType.DocumentSelectionChanged](https://docs.microsoft.com/javascript/api/office/office.eventtype) 事件类型。</span><span class="sxs-lookup"><span data-stu-id="d0326-p109">The first  _eventType_ parameter specifies the name of the event to subscribe to. Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the **Office.EventType.DocumentSelectionChanged** event type of the [Office.EventType](https://docs.microsoft.com/javascript/api/office/office.eventtype) enumeration.</span></span>

<span data-ttu-id="d0326-p110">作为第二个 _handler_ 参数传入函数的 `myHander()` 函数是在文档中的选定内容更改时执行的事件处理程序。调用该函数时使用单个参数 _eventArgs_，后者在异步操作完成时将包含对 [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) 对象的引用。可以使用 [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs#document) 属性访问引发事件的文档。</span><span class="sxs-lookup"><span data-stu-id="d0326-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="d0326-p111">可以为给定事件添加多个事件处理程序，具体方法是再次调用 **addHandlerAsync** 方法，并为 _handler_ 参数传入其他事件处理程序函数。只要每个事件处理程序函数的名称是唯一的，这样做就可行。</span><span class="sxs-lookup"><span data-stu-id="d0326-p111">You can add multiple event handlers for a given event by calling the  **addHandlerAsync** method again and passing in an additional event handler function for the _handler_ parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="d0326-144">停止检测选择中的更改</span><span class="sxs-lookup"><span data-stu-id="d0326-144">Stop detecting changes in the selection</span></span>


<span data-ttu-id="d0326-145">以下示例演示如何通过调用 [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) 方法停止侦听 [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) 事件。</span><span class="sxs-lookup"><span data-stu-id="d0326-145">The following example shows how to stop listening to the [Document.SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="d0326-146">`myHandler` 函数名称作为第二个 _handler_ 参数传递，指定了将从 **SelectionChanged** 事件中删除的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d0326-146">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the **SelectionChanged** event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="d0326-147">如果调用 **removeHandlerAsync** 方法时省略了可选的 _handler_ 参数，将会删除指定 _eventType_ 的所有事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d0326-147">If the optional  _handler_ parameter is omitted when the **removeHandlerAsync** method is called, all event handlers for the specified _eventType_ will be removed.</span></span>

