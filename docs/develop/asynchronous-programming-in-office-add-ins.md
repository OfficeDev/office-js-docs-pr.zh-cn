---
title: Office 加载项中的异步编程
description: ''
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 931ef17115885c8f96d41bf00143b3269a515d56
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596688"
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="24874-102">Office 加载项中的异步编程</span><span class="sxs-lookup"><span data-stu-id="24874-102">Asynchronous programming in Office Add-ins</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="24874-p101">为什么 Office 加载项 API 使用异步编程？由于 JavaScript 是单线程语言，如果脚本调用长时间运行的同步过程，则所有后续脚本执行都将被阻止，直到该过程完成。由于对 Office web 客户端执行某些操作（但也有丰富的客户端），因此大多数 Office JavaScript Api 都是以异步方式执行的。这可确保 Office 加载项快速响应和快速。在使用这些异步方法时，通常还需要编写回调函数。</span><span class="sxs-lookup"><span data-stu-id="24874-p101">Why does the Office Add-ins API use asynchronous programming? Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes. Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the Office JavaScript APIs are designed to execute asynchronously. This makes sure that Office Add-ins are responsive and fast. It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="24874-p102">API end 中所有异步方法的名称，其中包含 "Async"，如`Document.getSelectedDataAsync`、 `Binding.getDataAsync`或`Item.loadCustomPropertiesAsync`方法。当调用 "Async" 方法时，它会立即执行并可继续执行任何后续脚本。传递给 "Async" 方法的可选回调函数将在数据或请求的操作准备就绪后立即执行。这通常会立即发生，但在返回之前可能会稍有延迟。</span><span class="sxs-lookup"><span data-stu-id="24874-p102">The names of all asynchronous methods in the API end with "Async", such as the `Document.getSelectedDataAsync`, `Binding.getDataAsync`, or `Item.loadCustomPropertiesAsync` methods. When an "Async" method is called, it executes immediately and any subsequent script execution can continue. The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready. This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="24874-p103">下图显示了一个调用"Async"方法的执行流，该方法可读取用户在基于服务器的 Word 或 Excel 中打开的文档中选择的数据。“Async”调用开始时，JavaScript 执行线程空闲，可以执行任何额外的客户端处理（但图中没有显示）。当“Async”方法返回时，回调在线程上恢复执行，加载项可以访问数据、处理数据并显示结果。当使用 Office 富客户端主机应用程序（如，Word 2013 或 Excel 2013）时，可保持同样的异步执行模式。</span><span class="sxs-lookup"><span data-stu-id="24874-p103">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word or Excel. At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing (although none are shown in the diagram). When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result. The same asynchronous execution pattern holds when working with the Office rich client host applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="24874-116">*图 1. 异步编程执行流*</span><span class="sxs-lookup"><span data-stu-id="24874-116">*Figure 1. Asynchronous programming execution flow*</span></span>

![异步编程线程执行流](../images/office-addins-asynchronous-programming-flow.png)

<span data-ttu-id="24874-p104">在富客户端和 Web 客户端中支持此异步设计是 Office 加载项开发模型"写入一次，跨平台运行"设计目标的一部分。例如，可以使用将在 Excel 2013 和 Excel 网页版中运行的单一基本代码创建一个内容应用程序或任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="24874-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel on the web.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="24874-120">编写"Async"方法的回调函数</span><span class="sxs-lookup"><span data-stu-id="24874-120">Writing the callback function for an "Async" method</span></span>


<span data-ttu-id="24874-p105">作为_callback_参数传递给 "Async" 方法的回调函数必须声明一个参数，外接程序运行时将使用该参数在回调函数执行时提供对[AsyncResult](/javascript/api/office/office.asyncresult)对象的访问权限。您可以编写：</span><span class="sxs-lookup"><span data-stu-id="24874-p105">The callback function you pass as the _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](/javascript/api/office/office.asyncresult) object when the callback function executes. You can write:</span></span>


- <span data-ttu-id="24874-123">必须编写并作为 "Async" 方法的_callback_参数与调用一起直接传递给 "async" 方法的匿名函数。</span><span class="sxs-lookup"><span data-stu-id="24874-123">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the _callback_ parameter of the "Async" method.</span></span>

- <span data-ttu-id="24874-124">一个命名函数，用于将该函数的名称作为 "Async" 方法的_callback_参数进行传递。</span><span class="sxs-lookup"><span data-stu-id="24874-124">A named function, passing the name of that function as the _callback_ parameter of an "Async" method.</span></span>

<span data-ttu-id="24874-p106">如果您打算只使用一次代码，则可以使用匿名函数，这是因为该函数没有名称，您不能在代码的其他部分引用此代码。如果您打算重复将回调函数用于多个"Async"方法，则可以使用命名函数。</span><span class="sxs-lookup"><span data-stu-id="24874-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>


### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="24874-127">编写匿名回调函数</span><span class="sxs-lookup"><span data-stu-id="24874-127">Writing an anonymous callback function</span></span>

<span data-ttu-id="24874-128">以下匿名回调函数声明名为`result`的单个参数，该参数在回调返回时从[AsyncResult](/javascript/api/office/office.asyncresult#value)属性中检索数据。</span><span class="sxs-lookup"><span data-stu-id="24874-128">The following anonymous callback function declares a single parameter named `result` that retrieves data from the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property when the callback returns.</span></span>


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="24874-129">下面的示例演示如何在对`Document.getSelectedDataAsync`方法的完整 "Async" 方法调用的上下文中以行为的方式传递此匿名回调函数。</span><span class="sxs-lookup"><span data-stu-id="24874-129">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the `Document.getSelectedDataAsync` method.</span></span>


- <span data-ttu-id="24874-130">第一个_coercionType_参数`Office.CoercionType.Text`指定将所选数据作为文本字符串返回。</span><span class="sxs-lookup"><span data-stu-id="24874-130">The first _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>

- <span data-ttu-id="24874-p107">第二个_回调_参数是以串联方式传递给方法的匿名函数。函数执行时，它使用_result_参数访问`value` `AsyncResult`对象的属性，以显示用户在文档中选择的数据。</span><span class="sxs-lookup"><span data-stu-id="24874-p107">The second _callback_ argument is the anonymous function passed in-line to the method. When the function executes, it uses the _result_ parameter to access the `value` property of the `AsyncResult` object to display the data selected by the user in the document.</span></span>


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

<span data-ttu-id="24874-p108">您还可以使用回调函数的参数来访问该`AsyncResult`对象的其他属性。使用[AsyncResult](/javascript/api/office/office.asyncresult#status)属性可确定呼叫是成功还是失败。如果调用失败，可以使用[AsyncResult](/javascript/api/office/office.asyncresult#error)属性访问[error 对象，以获取错误消息](/javascript/api/office/office.error)。</span><span class="sxs-lookup"><span data-stu-id="24874-p108">You can also use the parameter of your callback function to access other properties of the `AsyncResult` object. Use the [AsyncResult.status](/javascript/api/office/office.asyncresult#status) property to determine if the call succeeded or failed. If your call fails you can use the [AsyncResult.error](/javascript/api/office/office.asyncresult#error) property to access an [Error](/javascript/api/office/office.error) object for error information.</span></span>

<span data-ttu-id="24874-136">有关使用`getSelectedDataAsync`方法的详细信息，请参阅[在文档或电子表格的活动选定内容中读取和写入数据](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)。</span><span class="sxs-lookup"><span data-stu-id="24874-136">For more information about using the `getSelectedDataAsync` method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 


### <a name="writing-a-named-callback-function"></a><span data-ttu-id="24874-137">编写命名回调函数</span><span class="sxs-lookup"><span data-stu-id="24874-137">Writing a named callback function</span></span>

<span data-ttu-id="24874-p109">或者，也可以编写一个命名的函数并将其名称传递给 "Async" 方法的_callback_参数。例如，可以重写上面的示例，以将名`writeDataCallback`为的_回调_参数的函数传递给此类。</span><span class="sxs-lookup"><span data-stu-id="24874-p109">Alternatively, you can write a named function and pass its name to the _callback_ parameter of an "Async" method. For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>


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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="24874-140">返回 AsyncResult.value 属性的内容的差异</span><span class="sxs-lookup"><span data-stu-id="24874-140">Differences in what's returned to the AsyncResult.value property</span></span>


<span data-ttu-id="24874-p110">对象的`asyncContext`、 `status`和`error`属性将向传递给所有 "Async" 方法的回调函数返回相同类型的信息。 `AsyncResult`但是，返回到`AsyncResult.value`属性的内容将根据 "Async" 方法的功能而有所不同。</span><span class="sxs-lookup"><span data-stu-id="24874-p110">The `asyncContext`, `status`, and `error` properties of the `AsyncResult` object return the same kinds of information to the callback function passed to all "Async" methods. However, what's returned to the `AsyncResult.value` property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="24874-p111">例如，CustomXmlPart、 `addHandlerAsync` [Document](/javascript/api/office/office.document)、 [RoamingSettings](/javascript/api/outlook/office.roamingsettings)和[Settings](/javascript/api/office/office.settings)对象[CustomXmlPart](/javascript/api/office/office.customxmlpart)的方法（用于将事件处理程序函数添加到这些对象所表示的[项目中）](/javascript/api/office/office.binding)。您可以从传递`AsyncResult.value`给任何`addHandlerAsync`方法的回调函数访问该属性，但由于在添加事件处理程序时没有要访问的数据或对象，因此，如果您`value`尝试访问该属性，该属性将始终返回**undefined** 。</span><span class="sxs-lookup"><span data-stu-id="24874-p111">For example, the `addHandlerAsync` methods (of the [Binding](/javascript/api/office/office.binding), [CustomXmlPart](/javascript/api/office/office.customxmlpart), [Document](/javascript/api/office/office.document), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [Settings](/javascript/api/office/office.settings) objects) are used to add event handler functions to the items represented by these objects. You can access the `AsyncResult.value` property from the callback function you pass to any of the `addHandlerAsync` methods, but since no data or object is being accessed when you add an event handler, the `value` property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="24874-p112">另一方面，如果调用`Document.getSelectedDataAsync`方法，它会将用户在文档中选择的数据返回到回调中的`AsyncResult.value`属性。或者，如果调用[getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-)方法，它将返回一个数组，其中的所有`Binding`对象都在文档中。如果调用[getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)方法，则它将返回单个`Binding`对象。</span><span class="sxs-lookup"><span data-stu-id="24874-p112">On the other hand, if you call the `Document.getSelectedDataAsync` method, it returns the data the user selected in the document to the `AsyncResult.value` property in the callback. Or, if you call the [Bindings.getAllAsync](/javascript/api/office/office.bindings#getallasync-options--callback-) method, it returns an array of all of the `Binding` objects in the document. And, if you call the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method, it returns a single `Binding` object.</span></span>

<span data-ttu-id="24874-p113">有关`AsyncResult.value` `Async`方法的返回属性的说明，请参阅该方法的参考主题的 "回调值" 部分。有关提供`Async`方法的所有对象的摘要，请参阅[AsyncResult](/javascript/api/office/office.asyncresult)对象主题底部的表。</span><span class="sxs-lookup"><span data-stu-id="24874-p113">For a description of what's returned to the `AsyncResult.value` property for an `Async` method, see the "Callback value" section of that method's reference topic. For a summary of all of the objects that provide `Async` methods, see the table at the bottom of the [AsyncResult](/javascript/api/office/office.asyncresult) object topic.</span></span>


## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="24874-150">异步编程模式</span><span class="sxs-lookup"><span data-stu-id="24874-150">Asynchronous programming patterns</span></span>


<span data-ttu-id="24874-151">Office JavaScript API 支持两种类型的异步编程模式：</span><span class="sxs-lookup"><span data-stu-id="24874-151">The Office JavaScript API supports two kinds of asynchronous programming patterns:</span></span>


- <span data-ttu-id="24874-152">使用嵌套回调</span><span class="sxs-lookup"><span data-stu-id="24874-152">Using nested callbacks</span></span>
    
- <span data-ttu-id="24874-153">使用承诺模式</span><span class="sxs-lookup"><span data-stu-id="24874-153">Using the promises pattern</span></span>
    
<span data-ttu-id="24874-p114">使用回调函数的异步编程通常需要您将回调返回的结果嵌套在两个或更多回调中。如果您需要这么做，则可以使用来自 API 的所有"Async"方法的嵌套回调。</span><span class="sxs-lookup"><span data-stu-id="24874-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="24874-p115">使用嵌套回调是一种熟悉大多数 JavaScript 开发人员的编程模式，但使用深度嵌套回调的代码很难阅读和理解。作为嵌套回调的替代方法，Office JavaScript API 还支持实施承诺模式。但是，在当前版本的 Office JavaScript API 中，承诺模式仅适用于[Excel 电子表格和 Word 文档中的绑定](bind-to-regions-in-a-document-or-spreadsheet.md)代码。</span><span class="sxs-lookup"><span data-stu-id="24874-p115">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand. As an alternative to nested callbacks, the Office JavaScript API also supports an implementation of the promises pattern. However, in the current version of the Office JavaScript API, the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="24874-159">使用嵌套回调函数的异步编程</span><span class="sxs-lookup"><span data-stu-id="24874-159">Asynchronous programming using nested callback functions</span></span>


<span data-ttu-id="24874-p116">通常，完成一项任务需要执行两个或更多个异步操作。为实现此目的，可在一个调用中嵌套另一个"Async"调用。</span><span class="sxs-lookup"><span data-stu-id="24874-p116">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span>

<span data-ttu-id="24874-162">以下代码示例内嵌两个异步调用。</span><span class="sxs-lookup"><span data-stu-id="24874-162">The following code example nests two asynchronous calls.</span></span>


- <span data-ttu-id="24874-p117">首先，调用[getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-)方法来访问名为 "MyBinding" 的文档中的绑定。返回`AsyncResult`到该回调的`result`参数的对象提供对该`AsyncResult.value`属性中指定的 binding 对象的访问权限。</span><span class="sxs-lookup"><span data-stu-id="24874-p117">First, the [Bindings.getByIdAsync](/javascript/api/office/office.bindings#getbyidasync-id--options--callback-) method is called to access a binding in the document named "MyBinding". The `AsyncResult` object returned to the `result` parameter of that callback provides access to the specified binding object from the `AsyncResult.value` property.</span></span>

- <span data-ttu-id="24874-165">然后，使用从第一个`result`参数访问的 binding 对象调用[binding.getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)方法。</span><span class="sxs-lookup"><span data-stu-id="24874-165">Then, the binding object accessed from the first `result` parameter is used to call the [Binding.getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-) method.</span></span>

- <span data-ttu-id="24874-166">最后，传递`result2`给`Binding.getDataAsync`方法的回调参数用于显示绑定中的数据。</span><span class="sxs-lookup"><span data-stu-id="24874-166">Finally, the `result2` parameter of the callback passed to the `Binding.getDataAsync` method is used to display the data in the binding.</span></span>


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

<span data-ttu-id="24874-167">此基本嵌套回调模式可用于 Office JavaScript API 中的所有异步方法。</span><span class="sxs-lookup"><span data-stu-id="24874-167">This basic nested callback pattern can be used for all asynchronous methods in the Office JavaScript API.</span></span>

<span data-ttu-id="24874-168">以下各节显示如何使用匿名函数或命名函数用于异步方法中的嵌套回调。</span><span class="sxs-lookup"><span data-stu-id="24874-168">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>


#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="24874-169">将匿名函数用于嵌套回调</span><span class="sxs-lookup"><span data-stu-id="24874-169">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="24874-p118">在下面的示例中，将内联声明两个匿名函数，并`getByIdAsync`将`getDataAsync`其作为嵌套回调传递给和方法。由于函数是简单且内嵌的，因此实现的意图将立即清除。</span><span class="sxs-lookup"><span data-stu-id="24874-p118">In the following example, two anonymous functions are declared inline and passed into the `getByIdAsync` and `getDataAsync` methods as nested callbacks. Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>


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


#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="24874-172">将命名函数用于嵌套回调</span><span class="sxs-lookup"><span data-stu-id="24874-172">Using named functions for nested callbacks</span></span>

<span data-ttu-id="24874-p119">在复杂的实现中，使用命名的函数使代码更易于读取、维护和重用可能非常有用。在下面的示例中，前一节的示例中的两个匿名函数已重写为名`deleteAllData`为`showResult`和的函数。然后，通过名称将这些命名的`getByIdAsync`函数`deleteAllDataValuesAsync`作为回调传递给和方法。</span><span class="sxs-lookup"><span data-stu-id="24874-p119">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse. In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named `deleteAllData` and `showResult`. These named functions are then passed into the `getByIdAsync` and `deleteAllDataValuesAsync` methods as callbacks by name.</span></span>


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


### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="24874-176">使用承诺模式访问绑定中的数据的异步编程</span><span class="sxs-lookup"><span data-stu-id="24874-176">Asynchronous programming using the promises pattern to access data in bindings</span></span>


<span data-ttu-id="24874-p120">在继续执行之前，承诺编程模式会立即返回表示其预期结果的承诺对象，而不是传递回调函数并等待函数返回。然而，与真正同步编程不同的是，在 Office 外接程序运行时环境完成请求之前，承诺结果的实现在后台实际上是延迟的。提供 _onError_ 处理程序来覆盖请求无法满足的情况。</span><span class="sxs-lookup"><span data-stu-id="24874-p120">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>


<span data-ttu-id="24874-p121">Office JavaScript API 提供了[office. select](/javascript/api/office#office-select-expression--callback-)方法，以支持使用现有绑定对象的承诺模式。返回`Office.select`到方法的承诺对象仅支持您可以直接从[Binding](/javascript/api/office/office.binding)对象访问的四个方法： [binding.getdataasync](/javascript/api/office/office.binding#getdataasync-options--callback-)、 [binding.setdataasync](/javascript/api/office/office.binding#setdataasync-data--options--callback-)、 [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)和[removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="24874-p121">The Office JavaScript API provides the [Office.select](/javascript/api/office#office-select-expression--callback-) method to support the promises pattern for working with existing binding objects. The promise object returned to the `Office.select` method supports only the four methods that you can access directly from the [Binding](/javascript/api/office/office.binding) object: [getDataAsync](/javascript/api/office/office.binding#getdataasync-options--callback-), [setDataAsync](/javascript/api/office/office.binding#setdataasync-data--options--callback-), [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-), and [removeHandlerAsync](/javascript/api/office/office.binding#removehandlerasync-eventtype--options--callback-).</span></span>


<span data-ttu-id="24874-182">与绑定一起使用的承诺模式采用以下形式：</span><span class="sxs-lookup"><span data-stu-id="24874-182">The promises pattern for working with bindings takes this form:</span></span>

 <span data-ttu-id="24874-183">**Office. select （**_selectorExpression_， _onError_**）。**_BindingObjectAsyncMethod_</span><span class="sxs-lookup"><span data-stu-id="24874-183">**Office.select(**_selectorExpression_, _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="24874-p122">_SelectorExpression_参数`"bindings#bindingId"`采用窗体，其中_bindingId_是您之前在文档`id`或电子表格中创建的绑定的名称（）（ `Bindings`使用集合的 "addFrom" 方法之一： `addFromNamedItemAsync`、 `addFromPromptAsync`或`addFromSelectionAsync`）。例如，选择器表达式`bindings#cities`指定要访问**id**为 "城市" 的绑定。</span><span class="sxs-lookup"><span data-stu-id="24874-p122">The _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where _bindingId_ is the name ( `id`) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the `Bindings` collection: `addFromNamedItemAsync`, `addFromPromptAsync`, or `addFromSelectionAsync`). For example, the selector expression `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="24874-p123">_OnError_参数是一个错误处理函数，它采用可用于访问`AsyncResult` `Error`对象的单个参数类型，前提是该`select`方法无法访问指定的绑定。下面的示例演示可传递给_onError_参数的基本错误处理程序函数。</span><span class="sxs-lookup"><span data-stu-id="24874-p123">The _onError_ parameter is an error handling function which takes a single parameter of type `AsyncResult` that can be used to access an `Error` object, if the `select` method fails to access the specified binding. The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>




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

<span data-ttu-id="24874-p124">将_BindingObjectAsyncMethod_占位符替换为对承诺对象支持的四个`Binding`对象方法中的任何一个`getDataAsync`： `setDataAsync`、、 `addHandlerAsync`或。 `removeHandlerAsync`对这些方法的调用不支持其他承诺。必须使用[嵌套回调函数模式](#AsyncProgramming_NestedCallbacks)调用它们。</span><span class="sxs-lookup"><span data-stu-id="24874-p124">Replace the _BindingObjectAsyncMethod_ placeholder with a call to any of the four `Binding` object methods supported by the promise object: `getDataAsync`, `setDataAsync`, `addHandlerAsync`, or `removeHandlerAsync`. Calls to these methods don't support additional promises. You must call them using the [nested callback function pattern](#AsyncProgramming_NestedCallbacks).</span></span>

<span data-ttu-id="24874-p125">在满足`Binding`对象承诺后，可以在连锁方法调用中重用它，就像它是一个绑定一样（加载项运行时不会异步重试满足承诺）。如果无法`Binding`满足对象承诺，加载项运行时将在下次调用其异步方法之一时再次尝试访问 binding 对象。</span><span class="sxs-lookup"><span data-stu-id="24874-p125">After a `Binding` object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise). If the `Binding` object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="24874-193">下面的`select`代码示例使用方法`id`从`cities` `Bindings`集合中检索带有 "" 的绑定，然后调用[addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-)方法为绑定的[dataChanged](/javascript/api/office/office.bindingdatachangedeventargs)事件添加事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="24874-193">The following code example uses the `select` method to retrieve a binding with the `id` "`cities`" from the `Bindings` collection, and then calls the [addHandlerAsync](/javascript/api/office/office.binding#addhandlerasync-eventtype--handler--options--callback-) method to add an event handler for the [dataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event of the binding.</span></span>




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> <span data-ttu-id="24874-p126">方法返回的`Binding`对象承诺仅提供对该`Binding`对象的四个方法的访问。 `Office.select`如果需要访问`Binding` `Document.bindings`对象的任何其他成员，则必须使用属性和`Bindings.getByIdAsync`或`Bindings.getAllAsync`方法检索该`Binding`对象。例如， `Binding`如果需要访问对象的任何属性（ `document` `id`或`type`属性），或者需要访问[MatrixBinding](/javascript/api/office/office.matrixbinding)或[TableBinding](/javascript/api/office/office.tablebinding)对象的属性，则必须使用`getByIdAsync`或`getAllAsync`方法来检索`Binding`对象。</span><span class="sxs-lookup"><span data-stu-id="24874-p126">The `Binding` object promise returned by the `Office.select` method provides access to only the four methods of the `Binding` object. If you need to access any of the other members of the `Binding` object, instead you must use the `Document.bindings` property and `Bindings.getByIdAsync` or `Bindings.getAllAsync` methods to retrieve the `Binding` object. For example, if you need to access any of the `Binding` object's properties (the `document`, `id`, or `type` properties), or need to access the properties of the [MatrixBinding](/javascript/api/office/office.matrixbinding) or [TableBinding](/javascript/api/office/office.tablebinding) objects, you must use the `getByIdAsync` or `getAllAsync` methods to retrieve a `Binding` object.</span></span>


## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="24874-197">向异步方法传递可选参数</span><span class="sxs-lookup"><span data-stu-id="24874-197">Passing optional parameters to asynchronous methods</span></span>


<span data-ttu-id="24874-198">所有"异步"方法的常用语法都遵循此模式：</span><span class="sxs-lookup"><span data-stu-id="24874-198">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="24874-199">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_`);`</span><span class="sxs-lookup"><span data-stu-id="24874-199">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="24874-p127">所有异步方法都支持可选参数，这些可选参数作为包含一个或多个可选参数的 JavaScript 对象表示法 (JSON) 对象传入。包含可选参数的 JSON 对象是键-值对的无序集合，其中用":"字符来分隔键和值。对象中的每对用逗号分隔，整个对集合括在大括号中。键是参数名称，值是要为该参数传递的值。</span><span class="sxs-lookup"><span data-stu-id="24874-p127">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="24874-204">您可以创建包含可选参数内嵌的 JSON 对象，或通过创建`options`对象并将其作为_options_参数传入。</span><span class="sxs-lookup"><span data-stu-id="24874-204">You can create the JSON object that contains optional parameters inline, or by creating an `options` object and passing that in as the _options_ parameter.</span></span>


### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="24874-205">传递可选参数内嵌</span><span class="sxs-lookup"><span data-stu-id="24874-205">Passing optional parameters inline</span></span>

<span data-ttu-id="24874-206">例如，用可选参数内嵌调用 [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) 方法的语法类似如下：</span><span class="sxs-lookup"><span data-stu-id="24874-206">For example, the syntax for calling the [Document.setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext: 'asyncContext'},callback);

```

<span data-ttu-id="24874-207">在这种形式的调用语法中，两个可选参数_coercionType_和_asyncContext_定义为括在大括号内的 JSON 对象内联。</span><span class="sxs-lookup"><span data-stu-id="24874-207">In this form of the calling syntax, the two optional parameters, _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="24874-208">下面的示例演示如何通过指定内嵌可选`Document.setSelectedDataAsync`参数来调用方法。</span><span class="sxs-lookup"><span data-stu-id="24874-208">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters inline.</span></span>


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
> <span data-ttu-id="24874-209">可以任何顺序在 JSON 对象中指定可选参数，只要指定正确的参数名称即可。</span><span class="sxs-lookup"><span data-stu-id="24874-209">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>


### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="24874-210">在 options 对象中传递可选参数</span><span class="sxs-lookup"><span data-stu-id="24874-210">Passing optional parameters in an options object</span></span>

<span data-ttu-id="24874-211">或者，也可以创建一个名为`options`的对象，该对象指定与方法调用分开的可选参数，然后`options`将对象作为_options_参数传递。</span><span class="sxs-lookup"><span data-stu-id="24874-211">Alternatively, you can create an object named `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="24874-212">下面的示例演示创建`options`对象的一种方法，其中`parameter1`、 `value1`等是实际参数名称和值的占位符。</span><span class="sxs-lookup"><span data-stu-id="24874-212">The following example shows one way of creating the `options` object, where `parameter1`, `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>




```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="24874-213">用于指定 [ValueFormat](/javascript/api/office/office.valueformat) 和 [FilterType](/javascript/api/office/office.filtertype) 参数时与以下示例类似。</span><span class="sxs-lookup"><span data-stu-id="24874-213">Which looks like the following example when used to specify the [ValueFormat](/javascript/api/office/office.valueformat) and [FilterType](/javascript/api/office/office.filtertype) parameters.</span></span>




```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="24874-214">此处是创建`options`对象的另一种方法。</span><span class="sxs-lookup"><span data-stu-id="24874-214">Here's another way of creating the `options` object.</span></span>




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="24874-215">用于指定`ValueFormat`和`FilterType`参数时与以下示例类似：</span><span class="sxs-lookup"><span data-stu-id="24874-215">Which looks like the following example when used to specify the `ValueFormat` and `FilterType` parameters:</span></span>


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> <span data-ttu-id="24874-216">使用任一方法创建`options`对象时，只要可选参数的名称指定正确，就可以按任意顺序指定这些参数。</span><span class="sxs-lookup"><span data-stu-id="24874-216">When using either method of creating the `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="24874-217">下面的示例演示如何通过在`Document.setSelectedDataAsync` `options`对象中指定可选参数来调用此方法。</span><span class="sxs-lookup"><span data-stu-id="24874-217">The following example shows how to call to the `Document.setSelectedDataAsync` method by specifying optional parameters in an `options` object.</span></span>




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


<span data-ttu-id="24874-p128">在这两个可选参数示例中，将_回调_参数指定为最后一个参数（后面是内联可选参数，或在_options_参数对象之后）。或者，可以在内联 JSON _callback_对象内或在`options`对象中指定 callback 参数。但是，只能在一个位置中传递_callback_参数：在_options_对象中（内联或在外部创建），或作为最后一个参数，但不能同时传递这两者。</span><span class="sxs-lookup"><span data-stu-id="24874-p128">In both optional parameter examples, the _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object). Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object. However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>


## <a name="see-also"></a><span data-ttu-id="24874-221">另请参阅</span><span class="sxs-lookup"><span data-stu-id="24874-221">See also</span></span>

- [<span data-ttu-id="24874-222">了解 Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="24874-222">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="24874-223">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="24874-223">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
