---
title: Office ?????????
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d251ebfd03227569b9a24bcd7f17baada6099938
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="d4f26-102">Office ?????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-102">Asynchronous programming in Office Add-ins</span></span>

<span data-ttu-id="d4f26-p101">??? Office ???? API ????????? JavaScript ??????????????????????????????????????????????? Office Web ??????????????????????????????????? Office ? JavaScript API ???????????????????? Office ?????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p101">Why does the Office Add-ins API use asynchronous programming? Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes. Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the methods in the JavaScript API for Office are designed to execute asynchronously. This makes sure that Office Add-ins are responsive and highly performing. It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="d4f26-p102">API ???????????????Async????? [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync)?[Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) ? [Item.loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) ????????Async??????????????????????????????????Async?????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p102">The names of all asynchronous methods in the API end with "Async", such as the  [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync), or [Item.loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) methods. When an "Async" method is called, it executes immediately and any subsequent script execution can continue. The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready. This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="d4f26-p103">?????????"Async"?????????????????????? Word Online ? Excel Online ?????????????"Async"??????JavaScript ?????????????????????????????????"Async"????????????????????????????????????????? Office ?????????????Word 2013 ? Excel 2013????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p103">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word Online or Excel Online. At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing. (Although none are shown in the diagram.) When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result. The same asynchronous execution pattern holds when working with the Office rich client host applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="d4f26-116">*? 1.???????*</span><span class="sxs-lookup"><span data-stu-id="d4f26-116">*Figure 1. Asynchronous programing execution flow*</span></span>

![?????????](../images/office15-app-async-prog-fig01.png)

<span data-ttu-id="d4f26-p104">?????? Web ???????????? Office ????????"??????????"??????????????????? Excel 2013 ? Excel Online ??????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel Online.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="d4f26-120">??"Async"???????</span><span class="sxs-lookup"><span data-stu-id="d4f26-120">Writing the callback function for an "Async" method</span></span>


<span data-ttu-id="d4f26-p105">??  _callback_ ?????"Async"???????????????????????????????????????? [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) ????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p105">The callback function you pass as the  _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object when the callback function executes. You can write:</span></span>


- <span data-ttu-id="d4f26-123">????????"Async"???  _callback_ ????????????"Async"????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-123">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the  _callback_ parameter of the "Async" method.</span></span>
    
- <span data-ttu-id="d4f26-124">??????"Async"???  _callback_ ????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-124">A named function, passing the name of that function as the  _callback_ parameter of an "Async" method.</span></span>
    
<span data-ttu-id="d4f26-p106">????????????????????????????????????????????????????????????????????"Async"?????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>


### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="d4f26-127">????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-127">Writing an anonymous callback function</span></span>

<span data-ttu-id="d4f26-128">???????????? `result` ?????????????????? [AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) ???????</span><span class="sxs-lookup"><span data-stu-id="d4f26-128">The following anonymous callback function declares a single parameter named  `result` that retrieves data from the [AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) property when the callback returns.</span></span>


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="d4f26-129">??????????????"Async"??????????????????  **Document.getSelectedDataAsync** ???</span><span class="sxs-lookup"><span data-stu-id="d4f26-129">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the  **Document.getSelectedDataAsync** method.</span></span>


- <span data-ttu-id="d4f26-130">???  _coercionType_ ?? `Office.CoercionType.Text` ?????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-130">The first  _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>
    
- <span data-ttu-id="d4f26-p107">???  _callback_ ?????????????????????????? _result_ ????? **AsyncResult** ??? **value** ??????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p107">The second  _callback_ argument is the anonymous function passed in-line to the method. When the function executes, it uses the _result_ parameter to access the **value** property of the **AsyncResult** object to display the data selected by the user in the document.</span></span>
    



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

<span data-ttu-id="d4f26-p108">??????????????? **AsyncResult** ???????????? [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) ???????????????????????????? [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) ???? [Error](https://dev.office.com/reference/add-ins/shared/error) ???????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p108">You can also use the parameter of your callback function to access other properties of the  **AsyncResult** object. Use the [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) property to determine if the call succeeded or failed. If your call fails you can use the [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) property to access an [Error](https://dev.office.com/reference/add-ins/shared/error) object for error information.</span></span>

<span data-ttu-id="d4f26-136">????  **getSelectedDataAsync** ??????????? [???????????????????????](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)?</span><span class="sxs-lookup"><span data-stu-id="d4f26-136">For more information about using the  **getSelectedDataAsync** method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 


### <a name="writing-a-named-callback-function"></a><span data-ttu-id="d4f26-137">????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-137">Writing a named callback function</span></span>

<span data-ttu-id="d4f26-p109">????????????????????????"Async"???  _callback_ ??????????????????? `writeDataCallback` ????? _callback_ ????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p109">Alternatively, you can write a named function and pass its name to the  _callback_ parameter of an "Async" method. For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>


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


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="d4f26-140">?? AsyncResult.value ????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-140">Differences in what's returned to the AsyncResult.value property</span></span>


<span data-ttu-id="d4f26-p110">**AsyncResult** ??? **asyncContext**?**status** ? **error** ????????????????????Async???????????????? **AsyncResult.value** ???????Async????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p110">The  **asyncContext**,  **status**, and  **error** properties of the **AsyncResult** object return the same kinds of information to the callback function passed to all "Async" methods. However, what's returned to the **AsyncResult.value** property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="d4f26-p111">????**Binding**?[CustomXmlPart](https://dev.office.com/reference/add-ins/shared/binding)?[Document](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart)?[RoamingSettings](https://dev.office.com/reference/add-ins/shared/document) ? [Settings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) ????[addHandlerAsync](https://dev.office.com/reference/add-ins/shared/settings) ?????????????????????????????????? **addHandlerAsync** ????????? **AsyncResult.value** ????????????????????????????????? **value** ????????? **undefined**?</span><span class="sxs-lookup"><span data-stu-id="d4f26-p111">For example, the  **addHandlerAsync** methods (of the [Binding](https://dev.office.com/reference/add-ins/shared/binding), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart), [Document](https://dev.office.com/reference/add-ins/shared/document), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings), and [Settings](https://dev.office.com/reference/add-ins/shared/settings) objects) are used to add event handler functions to the items represented by these objects. You can access the **AsyncResult.value** property from the callback function you pass to any of the **addHandlerAsync** methods, but since no data or object is being accessed when you add an event handler, the **value** property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="d4f26-p112">??????????  **Document.getSelectedDataAsync** ???????????????????????? **AsyncResult.value** ???????????? [Bindings.getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync) ???????????? **Binding** ?????????????? [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) ????????? **Binding** ???</span><span class="sxs-lookup"><span data-stu-id="d4f26-p112">On the other hand, if you call the  **Document.getSelectedDataAsync** method, it returns the data the user selected in the document to the **AsyncResult.value** property in the callback. Or, if you call the [Bindings.getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync) method, it returns an array of all of the **Binding** objects in the document. And, if you call the [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) method, it returns a single **Binding** object.</span></span>

<span data-ttu-id="d4f26-p113">??????Async??? **AsyncResult.value** ????????????????????????????????????Async????????????? [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) ??????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p113">For a description of what's returned to the  **AsyncResult.value** property for an "Async" method, see the "Callback value" section of that method's reference topic. For a summary of all of the objects that provide "Async" methods, see the table at the bottom of the [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object topic.</span></span>


## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="d4f26-150">??????</span><span class="sxs-lookup"><span data-stu-id="d4f26-150">Asynchronous programming patterns</span></span>


<span data-ttu-id="d4f26-151">??? Office ? JavaScript API ???????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-151">The JavaScript API for Office supports two kinds of asynchronous programming patterns:</span></span>


- <span data-ttu-id="d4f26-152">??????</span><span class="sxs-lookup"><span data-stu-id="d4f26-152">Using nested callbacks</span></span>
    
- <span data-ttu-id="d4f26-153">??????</span><span class="sxs-lookup"><span data-stu-id="d4f26-153">Using the promises pattern</span></span>
    
<span data-ttu-id="d4f26-p114">???????????????????????????????????????????????????? API ???"Async"????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="d4f26-p115">?????????? JavaScript ??????????????????????????????????????????????? Office ? JavaScript API ????????????????? Office ? JavaScript API ?????????????? [Excel ????? Word ??????](bind-to-regions-in-a-document-or-spreadsheet.md)????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p115">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand. As an alternative to nested callbacks, the JavaScript API for Office also supports an implementation of the promises pattern. However, in the current version of the JavaScript API for Office, the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="d4f26-159">?????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-159">Asynchronous programming using nested callback functions</span></span>


<span data-ttu-id="d4f26-p116">???????????????????????????????????????????"Async"???</span><span class="sxs-lookup"><span data-stu-id="d4f26-p116">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span> 

<span data-ttu-id="d4f26-162">???????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-162">The following code example nests two asynchronous calls.</span></span> 


- <span data-ttu-id="d4f26-p117">????? [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) ?????????MyBinding???????????????? `result` ??? **AsyncResult** ???????? **AsyncResult.value** ?????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p117">First, the [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) method is called to access a binding in the document named "MyBinding". The **AsyncResult** object returned to the `result` parameter of that callback provides access to the specified binding object from the **AsyncResult.value** property.</span></span>
    
- <span data-ttu-id="d4f26-165">??????? `result` ????????????? [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) ???</span><span class="sxs-lookup"><span data-stu-id="d4f26-165">Then, the binding object accessed from the first  `result` parameter is used to call the [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) method.</span></span>
    
- <span data-ttu-id="d4f26-166">??????  **Binding.getDataAsync** ?????? `result2` ?????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-166">Finally, the  `result2` parameter of the callback passed to the **Binding.getDataAsync** method is used to display the data in the binding.</span></span>
    



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

<span data-ttu-id="d4f26-167">???????????? ??? Office ? JavaScript API ?????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-167">This basic nested callback pattern can be used for all asynchronous methods in the JavaScript API for Office.</span></span>

<span data-ttu-id="d4f26-168">????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-168">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>


#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="d4f26-169">???????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-169">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="d4f26-p118">???????????????????????????????  **getByIdAsync** ? **getDataAsync** ????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p118">In the following example, two anonymous functions are declared inline and passed into the  **getByIdAsync** and **getDataAsync** methods as nested callbacks. Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>


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


#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="d4f26-172">???????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-172">Using named functions for nested callbacks</span></span>

<span data-ttu-id="d4f26-p119">????????????????????????????????????????????????????????????????????  `deleteAllData` ? `showResult` ???????????????????????? **getByIdAsync** ? **deleteAllDataValuesAsync** ???</span><span class="sxs-lookup"><span data-stu-id="d4f26-p119">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse. In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named  `deleteAllData` and `showResult`. These named functions are then passed into the  **getByIdAsync** and **deleteAllDataValuesAsync** methods as callbacks by name.</span></span>


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


### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="d4f26-176">???????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-176">Asynchronous programming using the promises pattern to access data in bindings</span></span>


<span data-ttu-id="d4f26-p120">????????????????????????????????????????????????????????????????? Office ???????????????????????????????????? _onError_ ?????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p120">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>

<span data-ttu-id="d4f26-p121">??? Office ? JavaScript API ????? [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) ???????????????????????? **Office.select** ????????????? [Binding](https://dev.office.com/reference/add-ins/shared/binding) ????????????[getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync)?[setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync)?[addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) ? [removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync)?</span><span class="sxs-lookup"><span data-stu-id="d4f26-p121">The JavaScript API for Office provides the [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) method to support the promises pattern for working with existing binding objects. The promise object returned to the **Office.select** method supports only the four methods that you can access directly from the [Binding](https://dev.office.com/reference/add-ins/shared/binding) object: [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync), [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync), [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value), and [removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync).</span></span>

<span data-ttu-id="d4f26-182">???????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-182">The promises pattern for working with bindings takes this form:</span></span>

 <span data-ttu-id="d4f26-183">**Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_</span><span class="sxs-lookup"><span data-stu-id="d4f26-183">**Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="d4f26-p122">_selectorExpression_ ???? `"bindings#bindingId"` ??????_bindingId_ ???????????????????? (**id**)??? **Bindings** ????addFrom??????**addFromNamedItemAsync**?**addFromPromptAsync** ? **addFromSelectionAsync**??????????? `bindings#cities` ?????? **id** ??cities?????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p122">The  _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where  _bindingId_ is the name ( **id**) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the  **Bindings** collection: **addFromNamedItemAsync**,  **addFromPromptAsync**, or  **addFromSelectionAsync**). For example, the selector expression  `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="d4f26-p123">_onError_ ??????????????????? **AsyncResult** ??????? **select** ????????????????? **Error** ???????????????? _onError_ ??????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p123">The  _onError_ parameter is an error handling function which takes a single parameter of type **AsyncResult** that can be used to access an **Error** object, if the **select** method fails to access the specified binding. The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>




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

<span data-ttu-id="d4f26-p124">???? _BindingObjectAsyncMethod_ ?????????????? **Binding** ???????????**getDataAsync**?**setDataAsync**?**addHandlerAsync** ? **emoveHandlerAsync**???????????????????????[????????](#AsyncProgramming_NestedCallbacks)??????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p124">Replace the  _BindingObjectAsyncMethod_ placeholder with a call to any of the four **Binding** object methods supported by the promise object: **getDataAsync**,  **setDataAsync**,  **addHandlerAsync**, or  **removeHandlerAsync**. Calls to these methods don't support additional promises. You must call them using the [nested callback function pattern](#AsyncProgramming_NestedCallbacks).</span></span>

<span data-ttu-id="d4f26-p125">**Binding** ???????????????????????????????????????????????? **Binding** ???????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p125">After a  **Binding** object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise). If the **Binding** object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="d4f26-193">???????? **select** ??? **Bindings** ???? **id** ??`cities`????????? [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) ???????? [dataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) ???????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-193">The following code example uses the  **select** method to retrieve a binding with the **id** " `cities`" from the  **Bindings** collection, and then calls the [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) method to add an event handler for the [dataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event of the binding.</span></span>




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> <span data-ttu-id="d4f26-p126">**Office.select** ????? **Binding** ???????? **Binding** ??????????????????? **Binding** ?????????????? **Document.bindings** ??? **Bindings.getByIdAsync** ? **Bindings.getAllAsync** ???? **Binding** ???????????? **Binding** ????????**document**?**id** ? **type** ????????? [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding) ? [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding) ?????????? **getByIdAsync** ? **getAllAsync** ???? **Binding** ???</span><span class="sxs-lookup"><span data-stu-id="d4f26-p126">The  **Binding** object promise returned by the **Office.select** method provides access to only the four methods of the **Binding** object. If you need to access any of the other members of the **Binding** object, instead you must use the **Document.bindings** property and **Bindings.getByIdAsync** or **Bindings.getAllAsync** methods to retrieve the **Binding** object. For example, if you need to access any of the **Binding** object's properties (the **document**,  **id**, or  **type** properties), or need to access the properties of the [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding) or [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding) objects, you must use the **getByIdAsync** or **getAllAsync** methods to retrieve a **Binding** object.</span></span>


## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="d4f26-197">???????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-197">Passing optional parameters to asynchronous methods</span></span>


<span data-ttu-id="d4f26-198">??"??"??????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-198">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="d4f26-199">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_ `);`</span><span class="sxs-lookup"><span data-stu-id="d4f26-199">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="d4f26-p127">?????????????????????????????????? JavaScript ????? (JSON) ???????????? JSON ????-???????????":"????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p127">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="d4f26-204">????????????? JSON ????????  `options` ??????? _options_ ?????</span><span class="sxs-lookup"><span data-stu-id="d4f26-204">You can create the JSON object that contains optional parameters inline, or by creating an  `options` object and passing that in as the _options_ parameter.</span></span>


### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="d4f26-205">????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-205">Passing optional parameters inline</span></span>

<span data-ttu-id="d4f26-206">???????????? [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) ??????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-206">For example, the syntax for calling the [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext:' asyncContext},callback);

```

<span data-ttu-id="d4f26-207">?????????????????  _coercionType_ ? _asyncContext_ ?????????? JSON ?????</span><span class="sxs-lookup"><span data-stu-id="d4f26-207">In this form of the calling syntax, the two optional parameters,  _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="d4f26-208">??????????????????????? **Document.setSelectedDataAsync** ???</span><span class="sxs-lookup"><span data-stu-id="d4f26-208">The following example shows how to call to the  **Document.setSelectedDataAsync** method by specifying optional parameters inline.</span></span>


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
> <span data-ttu-id="d4f26-209">??????? JSON ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-209">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>


### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="d4f26-210">? options ?????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-210">Passing optional parameters in an options object</span></span>

<span data-ttu-id="d4f26-211">????????????  `options` ??????????????????????? `options` ???? _options_ ??????</span><span class="sxs-lookup"><span data-stu-id="d4f26-211">Alternatively, you can create an object named  `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="d4f26-212">????????  `options` ?????????? `parameter1`? `value1` ???????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-212">The following example shows one way of creating the  `options` object, where `parameter1`,  `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>




```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="d4f26-213">???? [ValueFormat](https://dev.office.com/reference/add-ins/shared/valueformat-enumeration) ? [FilterType](https://dev.office.com/reference/add-ins/shared/filtertype-enumeration) ???????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-213">Which looks like the following example when used to specify the [ValueFormat](https://dev.office.com/reference/add-ins/shared/valueformat-enumeration) and [FilterType](https://dev.office.com/reference/add-ins/shared/filtertype-enumeration) parameters.</span></span>




```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="d4f26-214">?????  `options` ?????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-214">Here's another way of creating the  `options` object.</span></span>




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="d4f26-215">???? **ValueFormat** ? **FilterType** ????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-215">Which looks like the following example when used to specify the  **ValueFormat** and **FilterType** parameters.:</span></span>


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> <span data-ttu-id="d4f26-216">?????????? `options` ???????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-216">When using either method of creating the  `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="d4f26-217">????????????? `options` ???????????? **Document.setSelectedDataAsync** ???</span><span class="sxs-lookup"><span data-stu-id="d4f26-217">The following example shows how to call to the  **Document.setSelectedDataAsync** method by specifying optional parameters in an `options` object.</span></span>




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


<span data-ttu-id="d4f26-p128">???????????_callback_ ?????????????????????????? _options_ ?????????????? JSON ??? `options` ????? _callback_ ??????????????? _callback_ ???? _option_ ???????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d4f26-p128">In both optional parameter examples, the  _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object). Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object. However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>


## <a name="see-also"></a><span data-ttu-id="d4f26-221">????</span><span class="sxs-lookup"><span data-stu-id="d4f26-221">See also</span></span>

- [<span data-ttu-id="d4f26-222">????? Office ? JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d4f26-222">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md) 
- [<span data-ttu-id="d4f26-223">??? Office ? JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d4f26-223">JavaScript API for Office</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)
     
