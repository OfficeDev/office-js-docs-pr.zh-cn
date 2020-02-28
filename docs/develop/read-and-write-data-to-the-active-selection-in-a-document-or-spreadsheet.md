---
title: 对文档或电子表格中的活动选择执行数据读取和写入操作
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 039631e935d2ff6fadb4eab9d99df73ac30dae4d
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325001"
---
# <a name="read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet"></a>对文档或电子表格中的活动选择执行数据读取和写入操作

[Document](/javascript/api/office/office.document)对象公开的方法使您可以读取和写入用户在文档或电子表格中的当前选定内容。若要执行此操作`Document` ，对象将`getSelectedDataAsync`提供`setSelectedDataAsync`和方法。本主题还介绍了如何读取、写入和创建事件处理程序，以检测对用户所做更改的更改。

该`getSelectedDataAsync`方法仅适用于用户的当前选择。如果需要将所选内容保留在文档中，以便在运行外接程序的会话中可以读取和写入相同的选择，则必须使用[addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-)方法添加绑定（或使用[binding 对象的](/javascript/api/office/office.bindings)其他 "addFrom" 方法之一创建绑定）。有关创建对文档区域的绑定，然后对绑定进行读取和写入的信息，请参阅[绑定到文档或电子表格中的区域](bind-to-regions-in-a-document-or-spreadsheet.md)。


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

在此示例中，将第一个_coercionType_参数指定`Office.CoercionType.Text`为（也可以使用文本字符串`"text"`指定此参数）。这意味着，从回调函数中的_asyncresult_参数提供的[asyncresult](/javascript/api/office/office.asyncresult)对象的[value](/javascript/api/office/office.asyncresult#status)属性将返回包含文档中选定文本的**字符串**。指定不同的强制类型将导致不同的值。[CoercionType](/javascript/api/office/office.coerciontype)是可用强制类型值的枚举。`Office.CoercionType.Text`计算结果为字符串 "text"。


> [!TIP]
> **何时应使用矩阵与表 coercionType 进行数据访问？** 如果需要在添加行和列时动态增大选定的表格数据，并且必须使用表格标题，则应使用 table 数据类型（通过将`getSelectedDataAsync`方法的`"table"` _coercionType_参数指定为或`Office.CoercionType.Table`）。在数据结构中添加行和列在表数据和矩阵数据中都受支持，但只支持追加行和列的表数据。如果您不打算添加行和列，并且您的数据不需要标头功能，则应使用矩阵数据类型（通过将`getSelectedDataAsync`方法的`"matrix"` _coercionType_参数指定为 or `Office.CoercionType.Matrix`），这提供了与数据交互的更简单的模型。

在`getSelectedDataAsync`操作完成时，将执行作为第二个_回调_参数传入函数的匿名函数。使用单个参数_asyncResult_调用函数，其中包含了调用的结果和状态。如果调用失败， `AsyncResult`对象的[error](/javascript/api/office/office.asyncresult#asynccontext)属性将提供对[error](/javascript/api/office/office.error)对象的访问权限。您可以检查[Error.name](/javascript/api/office/office.error#name)和 Error 的值[。消息](/javascript/api/office/office.error#message)属性以确定设置操作失败的原因。否则，将显示文档中的选定文本。

[AsyncResult](/javascript/api/office/office.asyncresult#error)属性在**if**语句中使用，以测试调用是否成功。[AsyncResultStatus](/javascript/api/office/office.asyncresult#status)是可用`AsyncResult.status`属性值的枚举。`Office.AsyncResultStatus.Failed`计算结果为字符串 "failed" （同样，也可以指定为该文本字符串）。


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

在异步调用完成时，将匿名函数作为_callback_参数传递给[document.setselecteddataasync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)方法。使用`setSelectedDataAsync`方法将数据写入选定内容中时，回调的_asyncResult_参数仅提供对调用状态的访问权限，并在调用失败时向[Error](/javascript/api/office/office.error)对象提供访问权限。

> [!NOTE]
> 自 Excel 2013 SP1 发行版及相应的 Excel 网页版起，现在可以[在将表格写入当前选择时设置格式](../excel/excel-add-ins-tables.md)。


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

_第一个事件名称_参数指定要订阅的事件的名称。传递此参数`"documentSelectionChanged"`的字符串等效于传递[Office.](/javascript/api/office/office.eventtype)事件`Office.EventType.DocumentSelectionChanged`类型枚举的事件类型。

作为第二个 _handler_ 参数传入函数的 `myHander()` 函数是在文档中的选定内容更改时执行的事件处理程序。调用该函数时使用单个参数 _eventArgs_，后者在异步操作完成时将包含对 [DocumentSelectionChangedEventArgs](/javascript/api/office/office.documentselectionchangedeventargs) 对象的引用。可以使用 [DocumentSelectionChangedEventArgs.document](/javascript/api/office/office.documentselectionchangedeventargs#document) 属性访问引发事件的文档。


> [!NOTE]
> 您可以通过再次调用`addHandlerAsync`方法并为_处理程序_参数传递其他事件处理程序函数，为给定事件添加多个事件处理程序。只要每个事件处理程序函数的名称是唯一的，这就会正常工作。


## <a name="stop-detecting-changes-in-the-selection"></a>停止检测选择中的更改


以下示例演示如何通过调用 [document.removeHandlerAsync](/javascript/api/office/office.documentselectionchangedeventargs) 方法停止侦听 [Document.SelectionChanged](/javascript/api/office/office.document#removehandlerasync-eventtype--options--callback-) 事件。


```js
Office.context.document.removeHandlerAsync("documentSelectionChanged", {handler:myHandler}, function(result){});
```

作为`myHandler`第二个_handler_参数传递的函数名称指定将从`SelectionChanged`事件中移除的事件处理程序。


> [!IMPORTANT]
> 如果调用`removeHandlerAsync`方法时省略可选的_handler_参数，则将_删除指定的事件名称_的所有事件处理程序。
