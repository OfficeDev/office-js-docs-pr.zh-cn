---
title: 对文档或电子表格中的活动选择执行数据读取和写入操作
description: 了解如何在 Word 文档或 Excel 电子表格的活动选定内容中读取和写入数据。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9eaf0aac406731a9c0033e69bd8946464a4d1a4f
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292741"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a><span data-ttu-id="96881-103">对文档或电子表格中的活动选择执行数据读取和写入操作</span><span class="sxs-lookup"><span data-stu-id="96881-103">Read and write data to the active selection in a document or spreadsheet</span></span>

<span data-ttu-id="96881-104">通过 [Document](/javascript/api/office/office.document) 对象公开的方法，你可以读取文档或电子表格中用户的当前选区或向其中写入内容。</span><span class="sxs-lookup"><span data-stu-id="96881-104">The [Document](/javascript/api/office/office.document) object exposes methods that let you read and write to the user's current selection in a document or spreadsheet.</span></span> <span data-ttu-id="96881-105">若要执行此操作， `Document` 对象将提供 `getSelectedDataAsync` 和 `setSelectedDataAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="96881-105">To do that, the `Document` object provides the `getSelectedDataAsync` and `setSelectedDataAsync` methods.</span></span> <span data-ttu-id="96881-106">本主题还介绍了如何读取、写入和创建事件处理程序，以检测对用户选定内容所做的更改。</span><span class="sxs-lookup"><span data-stu-id="96881-106">This topic also describes how to read, write, and create event handlers to detect changes to the user's selection.</span></span>

<span data-ttu-id="96881-107">该 `getSelectedDataAsync` 方法仅适用于用户的当前选择。</span><span class="sxs-lookup"><span data-stu-id="96881-107">The `getSelectedDataAsync` method only works against the user's current selection.</span></span> <span data-ttu-id="96881-108">如果需要在文档中保留选区，以便使用相同的选区在运行加载项的各个会话中读取和写入，必须使用 [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) 方法添加绑定（或创建一个与 [Bindings](/javascript/api/office/office.bindings) 对象其他“addFrom”方法的绑定）。</span><span class="sxs-lookup"><span data-stu-id="96881-108">If you need to persist the selection in the document, so that the same selection is available to read and write across sessions of running your add-in, you must add a binding using the [Bindings.addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) method (or create a binding with one of the other "addFrom" methods of the [Bindings](/javascript/api/office/office.bindings) object).</span></span> <span data-ttu-id="96881-109">有关创建对文档区域的绑定，然后读取和写入绑定的信息，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="96881-109">For information about creating a binding to a region of a document, and then reading and writing to a binding, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>


## <a name="read-selected-data"></a><span data-ttu-id="96881-110">读取选择的数据</span><span class="sxs-lookup"><span data-stu-id="96881-110">Read selected data</span></span>


<span data-ttu-id="96881-111">以下示例演示如何使用 [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) 方法从文档的选定内容中获取数据。</span><span class="sxs-lookup"><span data-stu-id="96881-111">The following example shows how to get data from a selection in a document by using the [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method.</span></span>


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

<span data-ttu-id="96881-112">在此示例中，将第一个  _coercionType_ 参数指定为 `Office.CoercionType.Text` (您还可以使用文字字符串) 指定此参数 `"text"` 。</span><span class="sxs-lookup"><span data-stu-id="96881-112">In this example, the first  _coercionType_ parameter is specified as `Office.CoercionType.Text` (you can also specify this parameter by using the literal string `"text"`).</span></span> <span data-ttu-id="96881-113">这意味着在回调函数的 [asyncResult](/javascript/api/office/office.asyncresult#status) 参数中提供的 [AsyncResult](/javascript/api/office/office.asyncresult) 对象的 _value_ 属性将返回一个包含文档中选定文本的 **string**。</span><span class="sxs-lookup"><span data-stu-id="96881-113">This means that the [value](/javascript/api/office/office.asyncresult#status) property of the [AsyncResult](/javascript/api/office/office.asyncresult) object that is available from the _asyncResult_ parameter in the callback function will return a **string** that contains the selected text in the document.</span></span> <span data-ttu-id="96881-114">指定不同强制类型将产生不同的值。</span><span class="sxs-lookup"><span data-stu-id="96881-114">Specifying different coercion types will result in different values.</span></span> <span data-ttu-id="96881-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) 是可用的强制类型值的枚举。</span><span class="sxs-lookup"><span data-stu-id="96881-115">[Office.CoercionType](/javascript/api/office/office.coerciontype) is an enumeration of available coercion type values.</span></span> <span data-ttu-id="96881-116">`Office.CoercionType.Text` 计算结果为字符串 "text"。</span><span class="sxs-lookup"><span data-stu-id="96881-116">`Office.CoercionType.Text` evaluates to the string "text".</span></span>


> [!TIP]
> <span data-ttu-id="96881-117">**何时应使用矩阵与表格 coercionType 数据访问？**</span><span class="sxs-lookup"><span data-stu-id="96881-117">**When should you use the matrix versus table coercionType for data access?**</span></span> <span data-ttu-id="96881-118">如果需要在添加行和列时动态增大选定的表格数据，并且必须使用表格标题，则应通过将方法的 _coercionType_ 参数指定 `getSelectedDataAsync` 为 `"table"` 或) 来使用 table 数据类型 (`Office.CoercionType.Table` 。</span><span class="sxs-lookup"><span data-stu-id="96881-118">If you need your selected tabular data to grow dynamically when rows and columns are added, and you must work with table headers, you should use the table data type (by specifying the _coercionType_ parameter of the `getSelectedDataAsync` method as `"table"` or `Office.CoercionType.Table`).</span></span> <span data-ttu-id="96881-119">表格数据和矩阵数据中都支持在数据结构内添加行和列，但仅支持对表格数据追加行和列。</span><span class="sxs-lookup"><span data-stu-id="96881-119">Adding rows and columns within the data structure is supported in both table and matrix data, but appending rows and columns is supported only for table data.</span></span> <span data-ttu-id="96881-120">如果您不打算添加行和列，并且数据不需要标头功能，则应通过指定方法的  _coercionType_ 参数 as 或) 来使用矩阵数据类型 (`getSelectedDataAsync` `"matrix"` `Office.CoercionType.Matrix` ，这将提供更简单的数据交互模型。</span><span class="sxs-lookup"><span data-stu-id="96881-120">If you are you aren't planning on adding rows and columns, and your data doesn't require header functionality, then you should use the matrix data type (by specifying the  _coercionType_ parameter of `getSelectedDataAsync` method as `"matrix"` or `Office.CoercionType.Matrix`), which provides a simpler model of interacting with the data.</span></span>

<span data-ttu-id="96881-121">在操作完成时，将执行作为第二个  _回调_ 参数传入函数的匿名函数 `getSelectedDataAsync` 。</span><span class="sxs-lookup"><span data-stu-id="96881-121">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the `getSelectedDataAsync` operation is completed.</span></span> <span data-ttu-id="96881-122">调用该函数时使用单个参数 _asyncResult_，后者包含调用的结果和状态。</span><span class="sxs-lookup"><span data-stu-id="96881-122">The function is called with a single parameter, _asyncResult_, which contains the result and the status of the call.</span></span> <span data-ttu-id="96881-123">如果调用失败，对象的 [error](/javascript/api/office/office.asyncresult#asynccontext) 属性将 `AsyncResult` 提供对 [error](/javascript/api/office/office.error) 对象的访问权限。</span><span class="sxs-lookup"><span data-stu-id="96881-123">If the call fails, the [error](/javascript/api/office/office.asyncresult#asynccontext) property of the `AsyncResult` object provides access to the [Error](/javascript/api/office/office.error) object.</span></span> <span data-ttu-id="96881-124">您可以检查 [Error.name](/javascript/api/office/office.error#name) 和 [Error.message](/javascript/api/office/office.error#message) 属性的值，以确定设置操作失败的原因。</span><span class="sxs-lookup"><span data-stu-id="96881-124">You can check the value of the [Error.name](/javascript/api/office/office.error#name) and [Error.message](/javascript/api/office/office.error#message) properties to determine why the set operation failed.</span></span> <span data-ttu-id="96881-125">否则，会显示文档中选定的文本。</span><span class="sxs-lookup"><span data-stu-id="96881-125">Otherwise, the selected text in the document is displayed.</span></span>

<span data-ttu-id="96881-126">[AsyncResult.status](/javascript/api/office/office.asyncresult#error) 属性在 **if** 语句中用于测试调用是否成功。</span><span class="sxs-lookup"><span data-stu-id="96881-126">The [AsyncResult.status](/javascript/api/office/office.asyncresult#error) property is used in the **if** statement to test whether the call succeeded.</span></span> <span data-ttu-id="96881-127">[AsyncResultStatus](/javascript/api/office/office.asyncresult#status) 是可用 `AsyncResult.status` 属性值的枚举。</span><span class="sxs-lookup"><span data-stu-id="96881-127">[Office.AsyncResultStatus](/javascript/api/office/office.asyncresult#status) is an enumeration of available `AsyncResult.status` property values.</span></span> <span data-ttu-id="96881-128">`Office.AsyncResultStatus.Failed` 计算结果为字符串 "failed" (，同样，也可以指定为该文本字符串) 。</span><span class="sxs-lookup"><span data-stu-id="96881-128">`Office.AsyncResultStatus.Failed` evaluates to the string "failed" (and, again, can also be specified as that literal string).</span></span>


## <a name="write-data-to-the-selection"></a><span data-ttu-id="96881-129">向选定内容中写入数据</span><span class="sxs-lookup"><span data-stu-id="96881-129">Write data to the selection</span></span>


<span data-ttu-id="96881-130">以下示例演示如何将选定内容设置为显示"Hello World!"。</span><span class="sxs-lookup"><span data-stu-id="96881-130">The following example shows how to set the selection to show "Hello World!".</span></span>


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

<span data-ttu-id="96881-131">为  _data_ 参数传入不同对象类型会得到不同结果。</span><span class="sxs-lookup"><span data-stu-id="96881-131">Passing in different object types for the  _data_ parameter will have different results.</span></span> <span data-ttu-id="96881-132">结果取决于文档中当前选定的内容、Office 客户端应用程序托管加载项的内容，以及传入的数据是否可以强制转换为当前所选内容。</span><span class="sxs-lookup"><span data-stu-id="96881-132">The result depends on what is currently selected in the document, which Office client application is hosting your add-in, and whether the data passed in can be coerced to the current selection.</span></span>

<span data-ttu-id="96881-133">作为  [callback](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 参数传入 _setSelectedDataAsync_ 方法的匿名函数在异步调用完成时执行。</span><span class="sxs-lookup"><span data-stu-id="96881-133">The anonymous function passed into the [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method as the _callback_ parameter is executed when the asynchronous call is completed.</span></span> <span data-ttu-id="96881-134">使用方法将数据写入选定内容 `setSelectedDataAsync` 中时，回调的 _asyncResult_ 参数仅提供对调用状态的访问权限，并在调用失败时向 [Error](/javascript/api/office/office.error) 对象提供访问权限。</span><span class="sxs-lookup"><span data-stu-id="96881-134">When you write data to the selection by using the `setSelectedDataAsync` method, the _asyncResult_ parameter of the callback provides access only to the status of the call, and to the [Error](/javascript/api/office/office.error) object if the call fails.</span></span>

> [!NOTE]
> <span data-ttu-id="96881-135">自 Excel 2013 SP1 发行版及相应的 Excel 网页版起，现在可以[在将表格写入当前选择时设置格式](../excel/excel-add-ins-tables.md)。</span><span class="sxs-lookup"><span data-stu-id="96881-135">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel on the web, you can now [set formatting when writing a table to the current selection](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-in-the-selection"></a><span data-ttu-id="96881-136">检测选择中的更改</span><span class="sxs-lookup"><span data-stu-id="96881-136">Detect changes in the selection</span></span>


<span data-ttu-id="96881-137">以下示例演示如何通过使用 [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) 方法为文档中的 [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) 事件添加事件处理程序来检测选定内容中的更改。</span><span class="sxs-lookup"><span data-stu-id="96881-137">The following example shows how to detect changes in the selection by using the [Document.addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event on the document.</span></span>


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

<span data-ttu-id="96881-138">第一个  _eventType_ 参数指定要订阅的事件的名称。</span><span class="sxs-lookup"><span data-stu-id="96881-138">The first  _eventType_ parameter specifies the name of the event to subscribe to.</span></span> <span data-ttu-id="96881-139">传递 `"documentSelectionChanged"` 此参数的字符串等效于传递 `Office.EventType.DocumentSelectionChanged` [Office.](/javascript/api/office/office.eventtype) 事件类型枚举的事件类型。</span><span class="sxs-lookup"><span data-stu-id="96881-139">Passing the string `"documentSelectionChanged"` for this parameter is equivalent to passing the `Office.EventType.DocumentSelectionChanged` event type of the [Office.EventType](/javascript/api/office/office.eventtype) enumeration.</span></span>

<span data-ttu-id="96881-p110">作为第二个 _handler_ 参数传入函数的 `myHander()` 函数是在文档中的选定内容更改时执行的事件处理程序。调用该函数时使用单个参数 _eventArgs_，后者在异步操作完成时将包含对 [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) 对象的引用。可以使用 [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) 属性访问引发事件的文档。</span><span class="sxs-lookup"><span data-stu-id="96881-p110">The  `myHander()` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the selection is changed on the document. The function is called with a single parameter, _eventArgs_, which will contain a reference to a [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) object when the asynchronous operation completes. You can use the [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) property to access the document that raised the event.</span></span>


> [!NOTE]
> <span data-ttu-id="96881-143">您可以通过 `addHandlerAsync` 再次调用方法并为 _处理程序_ 参数传递其他事件处理程序函数，为给定事件添加多个事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="96881-143">You can add multiple event handlers for a given event by calling the `addHandlerAsync` method again and passing in an additional event handler function for the _handler_ parameter.</span></span> <span data-ttu-id="96881-144">只要每个事件处理程序函数的名称保持唯一，此方法就有用。</span><span class="sxs-lookup"><span data-stu-id="96881-144">This will work correctly as long as the name of each event handler function is unique.</span></span>


## <a name="stop-detecting-changes-in-the-selection"></a><span data-ttu-id="96881-145">停止检测选择中的更改</span><span class="sxs-lookup"><span data-stu-id="96881-145">Stop detecting changes in the selection</span></span>


<span data-ttu-id="96881-146">以下示例演示如何通过调用 [document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) 方法停止侦听 [Document.SelectionChanged](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) 事件。</span><span class="sxs-lookup"><span data-stu-id="96881-146">The following example shows how to stop listening to the [Document.SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) event by calling the [document.removeHandlerAsync](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) method.</span></span>


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

<span data-ttu-id="96881-147">`myHandler`作为第二个_handler_参数传递的函数名称指定将从事件中移除的事件处理程序 `SelectionChanged` 。</span><span class="sxs-lookup"><span data-stu-id="96881-147">The  `myHandler` function name that is passed as the second _handler_ parameter specifies the event handler that will be removed from the `SelectionChanged` event.</span></span>


> [!IMPORTANT]
> <span data-ttu-id="96881-148">如果调用方法时省略可选的_handler_参数 `removeHandlerAsync` ，则将删除指定的事件名称的所有_eventType_事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="96881-148">If the optional  _handler_ parameter is omitted when the `removeHandlerAsync` method is called, all event handlers for the specified _eventType_ will be removed.</span></span>
