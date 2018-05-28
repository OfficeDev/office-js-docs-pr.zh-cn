---
title: ??????????????
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: bd26aa12e5d6da145fb6a2a89daf937cf6e88f04
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="bind-to-regions-in-a-document-or-spreadsheet"></a><span data-ttu-id="a0ab9-102">??????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-102">Bind to regions in a document or spreadsheet</span></span>

<span data-ttu-id="a0ab9-p101">?????????????????????????????????????????????????????????????????????????????????????????????[addFromPromptAsync]?[addFromSelectionAsync] ? [addFromNamedItemAsync]????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p101">Binding-based data access enables content and task pane add-ins to consistently access a particular region of a document or spreadsheet through an identifier. The add-in first needs to establish the binding by calling one of the methods that associates a portion of the document with a unique identifier: [addFromPromptAsync], [addFromSelectionAsync], or [addFromNamedItemAsync]. After the binding is established, the add-in can use the provided identifier to access the data contained in the associated region of the document or spreadsheet. Creating bindings provides the following value to your add-in:</span></span>


- <span data-ttu-id="a0ab9-107">???????? Office ????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-107">Permits access to common data structures across supported Office applications, such as: tables, ranges, or text (a contiguous run of characters).</span></span>
    
- <span data-ttu-id="a0ab9-108">???/???????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-108">Enables read/write operations without requiring the user to make a selection.</span></span>
    
- <span data-ttu-id="a0ab9-p102">?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p102">Establishes a relationship between the add-in and the data in the document. Bindings are persisted in the document, and can be accessed at a later time.</span></span>
    
<span data-ttu-id="a0ab9-p103">?????????????????????????????????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p103">Establishing a binding also allows you to subscribe to data and selection change events that are scoped to that particular region of the document or spreadsheet. This means that the add-in is only notified of changes that happen within the bound region as opposed to general changes across the whole document or spreadsheet.</span></span>

<span data-ttu-id="a0ab9-p104">[Bindings] ???? [getAllAsync] ??????????????????????????????????? Bindings.[getByIdAsync] ? [Office.select] ???? ID ?????????? [Bindings] ??????????????????????[addFromSelectionAsync]?[addFromPromptAsync]?[addFromNamedItemAsync] ? [releaseByIdAsync]?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p104">The [Bindings] object exposes a [getAllAsync] method that gives access to the set of all bindings established on the document or spreadsheet. An individual binding can be accessed by its ID using either the Bindings.[getByIdAsync] or [Office.select] methods. You can establish new bindings as well as remove existing ones by using one of the following methods of the [Bindings] object: [addFromSelectionAsync], [addFromPromptAsync], [addFromNamedItemAsync], or [releaseByIdAsync].</span></span>


## <a name="binding-types"></a><span data-ttu-id="a0ab9-116">????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-116">Binding types</span></span>

<span data-ttu-id="a0ab9-117">??? [addFromSelectionAsync]?[addFromPromptAsync] ? [addFromNamedItemAsync] ??????????? _bindingType_ ????[?????????][Office.BindingType]?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-117">There are [three different types of bindings][Office.BindingType] that you specify with the  _bindingType_ parameter when you create a binding with the [addFromSelectionAsync], [addFromPromptAsync] or [addFromNamedItemAsync] methods:</span></span>

1. <span data-ttu-id="a0ab9-118">**[????][TextBinding]** - ?????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-118">**[Text Binding][TextBinding]** - Binds to a region of the document that can be represented as text.</span></span>

    <span data-ttu-id="a0ab9-p105">? Word ????????????????? Excel ???????????????????????? Excel ?????????? Word ???????????????HTML ? Open XML for Office?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p105">In Word, most contiguous selections are valid, while in Excel only single cell selections can be the target of a text binding. In Excel, only plain text is supported. In Word, three formats are supported: plain text, HTML, and Open XML for Office.</span></span>

2. <span data-ttu-id="a0ab9-p106">**[????][MatrixBinding]** - ????????????????????????????????????? **Array** ??????? JavaScript ????????????????????? **string** ????? ` [['a', 'b'], ['c', 'd']]` ?????????????????? `[['a'], ['b'], ['c']]` ??????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p106">**[Matrix Binding][MatrixBinding]** - Binds to a fixed region of a document that contains tabular data without headers.Data in a matrix binding is written or read as a two dimensional  **Array**, which in JavaScript is implemented as an array of arrays. For example, two rows of  **string** values in two columns can be written or read as ` [['a', 'b'], ['c', 'd']]`, and a single column of three rows can be written or read as  `[['a'], ['b'], ['c']]`.</span></span>

    <span data-ttu-id="a0ab9-p107">? Excel ???????????????????????? Word ?????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p107">In Excel, any contiguous selection of cells can be used to establish a matrix binding. In Word, only tables support matrix binding.</span></span>

3. <span data-ttu-id="a0ab9-p108">**[????][TableBinding]** - ??????????????????????????? [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) ????????`TableData` ???? `headers` ? `rows` ???????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p108">**[Table Binding][TableBinding]** - Binds to a region of a document that contains a table with headers.Data in a table binding is written or read as a [TableData](https://dev.office.com/reference/add-ins/shared/tabledata) object. The `TableData` object exposes the data through the `headers` and `rows` properties.</span></span>

    <span data-ttu-id="a0ab9-p109">?? Excel ? Word ????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p109">Any Excel or Word table can be the basis for a table binding. After you establish a table binding, each new row or column a user adds to the table is automatically included in the binding.</span></span>

<span data-ttu-id="a0ab9-p110">?? `Bindings` ??????addFrom?????????????????????????????????[MatrixBinding]?[TableBinding] ? [TextBinding]?????????? `Binding` ??? [getDataAsync] ? [setDataAsync] ????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p110">After a binding is created by using one of the three "addFrom" methods of the  `Bindings` object, you can work with the binding's data and properties by using the methods of the corresponding object: [MatrixBinding], [TableBinding], or [TextBinding]. All three of these objects inherit the [getDataAsync] and [setDataAsync] methods of the `Binding` object that enable you to interact with the bound data.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ab9-p111">**??????????????**?????????????????????????????????????????????????????????????????????????????????????? [TableBinding.rowCount] ?????????? [BindingSelectionChangedEventArgs] ??? `rowCount` ? `startRow` ????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p111">**When should you use matrix versus table bindings?** When the tabular data you are working with contains a total row, you must use a matrix binding if your add-in's script needs to access values in the total row or detect that the user's selection is in the total row. If you establish a table binding for tabular data that contains a total row, the [TableBinding.rowCount] property and the `rowCount` and `startRow` properties of the [BindingSelectionChangedEventArgs] object in event handlers won't reflect the total row in their values. To work around this limitation, you must use establish a matrix binding to work with the total row.</span></span>

## <a name="add-a-binding-to-the-users-current-selection"></a><span data-ttu-id="a0ab9-136">??????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-136">Add a binding to the user's current selection</span></span>

<span data-ttu-id="a0ab9-137">?????????? [addFromSelectionAsync] ?????????????????? `myBinding` ??????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-137">The following example shows how to add a text binding called  `myBinding` to the current selection in a document by using the [addFromSelectionAsync] method.</span></span>


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

<span data-ttu-id="a0ab9-p112">????????????????????????????? [TextBinding]????????????????????[Office.BindingType] ?????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p112">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection. Different binding types expose different data and operations. [Office.BindingType] is an enumeration of available binding type values.</span></span>

<span data-ttu-id="a0ab9-p113">???????????????????????? ID?????? ID??????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p113">The second optional parameter is an object that specifies the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="a0ab9-p114">?????? _callback_ ??????????????????????????????? `asyncResult` ??????????????????? [AsyncResult] ???`AsyncResult.value` ????? [Binding] ????????????????????????????? [Binding] ???????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p114">The anonymous function that is passed into the function as the final  _callback_ parameter is executed when the creation of the binding is complete. The function is called with a single parameter, `asyncResult`, which provides access to an [AsyncResult] object that provides the status of the call. The `AsyncResult.value` property contains a reference to a [Binding] object of the type that is specified for the newly created binding. You can use this [Binding] object to get and set data.</span></span>

## <a name="add-a-binding-from-a-prompt"></a><span data-ttu-id="a0ab9-148">????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-148">Add a binding from a prompt</span></span>

<span data-ttu-id="a0ab9-p115">???????????? [addFromPromptAsync] ?????? `myBinding` ????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p115">The following example shows how to add a text binding called  `myBinding` by using the [addFromPromptAsync] method. This method lets the user specify the range for the binding by using the application's built-in range selection prompt.</span></span>


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

<span data-ttu-id="a0ab9-p116">??????????????????????????????????????? [TextBinding]?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p116">In this example, the specified binding type is text. This means that a [TextBinding] will be created for the selection that the user specifies in the prompt.</span></span>

<span data-ttu-id="a0ab9-p117">????????????????????? ID?????? ID??????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p117">The second parameter is an object that contains the ID of the new binding being created. If an ID is not specified, one is generated automatically.</span></span>

<span data-ttu-id="a0ab9-p118">?????  _callback_ ??????????????????????????????? [AsyncResult] ?????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p118">The anonymous function passed into the function as the third  _callback_ parameter is executed when the creation of the binding is complete. When the callback function executes, the [AsyncResult] object contains the status of the call and the newly created binding.</span></span>

<span data-ttu-id="a0ab9-157">? 1 ?? Excel ???????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-157">Figure 1 shows the built-in range selection prompt in Excel.</span></span>


<span data-ttu-id="a0ab9-158">*? 1.Excel ???? UI*</span><span class="sxs-lookup"><span data-stu-id="a0ab9-158">*Figure 1. Excel Select Data UI*</span></span>

![Excel ???? UI](../images/agave-api-overview-excel-selection-ui.png)


## <a name="add-a-binding-to-a-named-item"></a><span data-ttu-id="a0ab9-160">??????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-160">Add a binding to a named item</span></span>


<span data-ttu-id="a0ab9-161">?????????? [addFromNamedItemAsync] ???????? `myRange` ?????????????????????? `id` ????myMatrix??</span><span class="sxs-lookup"><span data-stu-id="a0ab9-161">The following example shows how to add a binding to the existing  `myRange` named item as a "matrix" binding by using the [addFromNamedItemAsync] method, and assigns the binding's `id` as "myMatrix".</span></span>


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

<span data-ttu-id="a0ab9-p119">**?? Excel**?[addFromNamedItemAsync] ??? `itemName` ??????????????????? `A1` ???? `("A1:A3")` ???????????????? Excel ???????????????????Table1????????????????Table2?????????? Excel UI ????????????????????**???? | ??**???????**????**????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p119">**For Excel**, the  `itemName` parameter of the [addFromNamedItemAsync] method can refer to an existing named range, a range specified with the `A1` reference style `("A1:A3")`, or a table. By default, adding a table in Excel assigns the name "Table1" for the first table you add, "Table2" for the second table you add, and so on. To assign a meaningful name for a table in the Excel UI, use the **Table Name** property on the **Table Tools | Design** tab of the ribbon.</span></span>


> [!NOTE]
> <span data-ttu-id="a0ab9-165">? Excel ???????????????????????????????????????????  `"Sheet1!Table1"`</span><span class="sxs-lookup"><span data-stu-id="a0ab9-165">In Excel, when specifying a table as a named item, you must fully qualify the name to include the worksheet name in the name of the table in this format:  `"Sheet1!Table1"`</span></span>

<span data-ttu-id="a0ab9-166">?????? Excel ? A ???????? (`"A1:A3"`) ????????? id `"MyCities"`????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-166">The following example creates a binding in Excel to the first three cells in column A ( `"A1:A3"`), assigns the  id `"MyCities"`, and then writes three city names to that binding.</span></span>


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

<span data-ttu-id="a0ab9-p120">**?? Word**?[addFromNamedItemAsync] ??? `itemName` ???? `Rich Text` ????? `Title` ????????? `Rich Text` ???????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p120">**For Word**, the  `itemName` parameter of the [addFromNamedItemAsync] method refers to the `Title` property of a `Rich Text` content control. (You can't bind to content controls other than the `Rich Text` content control.)</span></span>

<span data-ttu-id="a0ab9-p121">?????????????? `Title*` ????? Word UI ?????????????????**????**???????**??**????????**????**??????????**??**?????**??**??????**??????**??????????????**??**??????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p121">By default, a content control has no  `Title*`value assigned. To assign a meaningful name in the Word UI, after inserting a **Rich Text** content control from the **Controls** group on the **Developer** tab of the ribbon, use the **Properties** command in the **Controls** group to display the **Content Control Properties** dialog box. Then set the **Title** property of the content control to the name you want to reference from your code.</span></span>

<span data-ttu-id="a0ab9-172">????? Word ?????????????  `"FirstName"` ?????????????????? **id**`"firstName"`??????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-172">The following example creates a text binding in Word to a rich text content control named  `"FirstName"`, assigns the  **id** `"firstName"`, and then displays that information.</span></span>


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

## <a name="get-all-bindings"></a><span data-ttu-id="a0ab9-173">??????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-173">Get all bindings</span></span>


<span data-ttu-id="a0ab9-174">?????????? Bindings.[getAllAsync] ?????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-174">The following example shows how to get all bindings in a document by using the Bindings.[getAllAsync] method.</span></span>


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

<span data-ttu-id="a0ab9-p122">?? `callback` ???????????????????????????? `asyncResult` ?????????????????????????????????? ID ?????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p122">The anonymous function that is passed into the function as the  `callback` parameter is executed when the operation is complete. The function is called with a single parameter, `asyncResult`, which contains an  array of the bindings in the document. The array is iterated to build a string that contains the IDs of the bindings. The string is then displayed in a message box.</span></span>


## <a name="get-a-binding-by-id-using-the-getbyidasync-method-of-the-bindings-object"></a><span data-ttu-id="a0ab9-179">?? Bindings ??? getByIdAsync ??? ID ????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-179">Get a binding by ID using the getByIdAsync method of the Bindings object</span></span>


<span data-ttu-id="a0ab9-p123">?????????? [getByIdAsync] ????????? ID ???????????????????????????????? `'myBinding'` ?????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p123">The following example shows how to use the [getByIdAsync] method to get a binding in a document by specifying its ID. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


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

<span data-ttu-id="a0ab9-182">????????? `id` ?????????? ID?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-182">In the example, the first  `id` parameter is the ID of the binding to retrieve.</span></span>

<span data-ttu-id="a0ab9-p124">?????  _callback_ ???????????????????????????? _asyncResult_ ????????????? ID ?"myBinding"????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p124">The anonymous function that is passed into the function as the second  _callback_ parameter is executed when the operation is completed. The function is called with a single parameter, _asyncResult_, which contains the status of the call and the binding with the ID "myBinding".</span></span>


## <a name="get-a-binding-by-id-using-the-select-method-of-the-office-object"></a><span data-ttu-id="a0ab9-185">?? Office ??? select ??? ID ????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-185">Get a binding by ID using the select method of the Office object</span></span>


<span data-ttu-id="a0ab9-p125">?????????? [Office.select] ?????????????? [Binding] ???? ID ?????????????????? Binding.[getDataAsync] ????????????????????????????????????? `'myBinding'` ?????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p125">The following example shows how to use the [Office.select] method to get a [Binding] object promise in a document by specifying its ID in a selector string. It then calls the Binding.[getDataAsync] method to get data from the specified binding. This example assumes that a binding named `'myBinding'` was added to the document using one of the methods described earlier in this topic.</span></span>


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
> <span data-ttu-id="a0ab9-p126">?? `select` ???????? [Binding] ?????????????????[getDataAsync]?[setDataAsync]?[addHandlerAsync] ? [removeHandlerAsync]????????? Binding ??????? `onError` ???? [asyncResult].error ????????????????? Binding ????????? `select` ????? Binding ??????????????? [getByIdAsync] ?????????? [Document.bindings] ??? Bindings.[getByIdAsync] ???? Binding** ???</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p126">If the  `select` method promise successfully returns a [Binding] object, that object exposes only the following four methods of the object: [getDataAsync], [setDataAsync], [addHandlerAsync], and [removeHandlerAsync]. If the promise cannot return a  Binding object, the `onError` callback can be used to access an [asyncResult].error object to get more information.If you need to call a member of the Binding object other than the four methods exposed by the Binding object promise returned by the `select` method, instead use the [getByIdAsync] method by using the [Document.bindings] property and Bindings.[getByIdAsync] method to retrieve the Binding** object.</span></span>

## <a name="release-a-binding-by-id"></a><span data-ttu-id="a0ab9-191">? ID ????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-191">Release a binding by ID</span></span>


<span data-ttu-id="a0ab9-192">?????????? [releaseByIdAsync] ????????? ID ?????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-192">The following example shows how use the [releaseByIdAsync] method to release a binding in a document by specifying its ID.</span></span>

```js
Office.context.document.bindings.releaseByIdAsync('myBinding', function (asyncResult) {
    write('Released myBinding!');
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="a0ab9-193">????????? `id` ?????????? ID?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-193">In the example, the first `id` parameter is the ID of the binding to release.</span></span>

<span data-ttu-id="a0ab9-p127">?????????????????????????????????????  [asyncResult] ??????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p127">The anonymous function that is passed into the function as the second parameter is a callback that is executed when the operation is complete. The function is called with a single parameter,  [asyncResult], which contains the status of the call.</span></span>


## <a name="read-data-from-a-binding"></a><span data-ttu-id="a0ab9-196">????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-196">Read data from a binding</span></span>


<span data-ttu-id="a0ab9-197">?????????? [getDataAsync] ?????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-197">The following example shows how to use the [getDataAsync] method to get data from an existing binding.</span></span>


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

 <span data-ttu-id="a0ab9-p128">`myBinding` ?????????????????????? [Office.select] ??? ID ????????? [getDataAsync] ???????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p128">`myBinding` is a variable that contains an existing text binding in the document. Alternatively, you could use the [Office.select] to access the binding by its ID, and start your call to the [getDataAsync] method, like this:</span></span> 

```js 
Office.select("bindings#myBindingID").getDataAsync
```


<span data-ttu-id="a0ab9-p129">??????????????????????[AsyncResult].value ???? `myBinding` ?????????????????????????????????????????????????????????????? [getDataAsync] ?????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p129">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The [AsyncResult].value property contains the data within `myBinding`. The type of the value depends on the binding type. The binding in this example is a text binding. Therefore, the value will contain a string. For additional examples of working with matrix and table bindings, see the [getDataAsync] method topic.</span></span>


## <a name="write-data-to-a-binding"></a><span data-ttu-id="a0ab9-206">????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-206">Write data to a binding</span></span>

<span data-ttu-id="a0ab9-207">?????????? [setDataAsync] ?????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-207">The following example shows how to use the [setDataAsync] method to set data in an existing binding.</span></span>

```js
myBinding.setDataAsync('Hello World!', function (asyncResult) { });
```

 <span data-ttu-id="a0ab9-208">`myBinding` ?????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-208">`myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="a0ab9-p130">?????????????? `myBinding` ??????????????????? `string`?????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p130">In the example, the first parameter is the value to set on  `myBinding`. Because this is a text binding, the value is a `string`. Different binding types accept different types of data.</span></span>

<span data-ttu-id="a0ab9-p131">?????????????????????????????? `asyncResult` ??????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p131">The anonymous function that is passed into the function is a callback that is executed when the operation is complete. The function is called with a single parameter,  `asyncResult`, which contains the status of the result.</span></span>

> [!NOTE]
> <span data-ttu-id="a0ab9-214">? Excel 2013 SP1 ??????? Excel Online ????????[????????????????????](../excel/excel-add-ins-tables.md)?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-214">Starting with the release of the Excel 2013 SP1 and the corresponding build of Excel Online, you can now [set formatting when writing and updating data in bound tables](../excel/excel-add-ins-tables.md).</span></span>


## <a name="detect-changes-to-data-or-the-selection-in-a-binding"></a><span data-ttu-id="a0ab9-215">??????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-215">Detect changes to data or the selection in a binding</span></span>


<span data-ttu-id="a0ab9-216">??????????? ID ??MyBinding????? [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) ???????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-216">The following example shows how to attach an event handler to the [DataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event of a binding with an id of "MyBinding".</span></span>


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

<span data-ttu-id="a0ab9-217">????????????????`myBinding`</span><span class="sxs-lookup"><span data-stu-id="a0ab9-217">The `myBinding` is a variable that contains an existing text binding in the document.</span></span>

<span data-ttu-id="a0ab9-p132">[addHandlerAsync] ?????? `eventType` ??????????????[Office.EventType] ????????????`Office.EventType.BindingDataChanged evaluates to the string `?bindingDataChanged??</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p132">The first  `eventType` parameter of the [addHandlerAsync] method specifies the name of the event to subscribe to. [Office.EventType] is an enumeration of available event type values. `Office.EventType.BindingDataChanged evaluates to the string `"bindingDataChanged"\`.</span></span>

<span data-ttu-id="a0ab9-p133">?????  _handler_ ??????? `dataChanged` ??????????????????????????????? _eventArgs_ ?????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p133">The  `dataChanged` function that is passed into the function as the second _handler_ parameter is an event handler that is executed when the data in the binding is changed. The function is called with a single parameter, _eventArgs_, which contains a reference to the binding. This binding can be used to retrieve the updated data.</span></span>

<span data-ttu-id="a0ab9-p134">????????????? [SelectionChanged] ??????????????????????????????? [addHandlerAsync] ??? `eventType` ????? `Office.EventType.BindingSelectionChanged` ? `"bindingSelectionChanged"`?</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p134">Similarly, you can detect when a user changes selection in a binding by attaching an event handler to the [SelectionChanged] event of a binding. To do that, specify the `eventType` parameter of the [addHandlerAsync] method as `Office.EventType.BindingSelectionChanged` or `"bindingSelectionChanged"`.</span></span>

<span data-ttu-id="a0ab9-p135">?????????????????????????  [addHandlerAsync] ????? `handler` ????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p135">You can add multiple event handlers for a given event by calling the [addHandlerAsync] method again and passing in an additional event handler function for the `handler` parameter. This will work correctly as long as the name of each event handler function is unique.</span></span>


### <a name="remove-an-event-handler"></a><span data-ttu-id="a0ab9-228">????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-228">Remove an event handler</span></span>


<span data-ttu-id="a0ab9-p136">????????????????? [removeHandlerAsync] ???????????? _eventType_ ?????????????????????????? _handler_ ?????????????????????????? `dataChanged` ?????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-p136">To remove an event handler for an event, call the [removeHandlerAsync] method passing in the event type as the first _eventType_ parameter, and the name of the event handler function to remove as the second _handler_ parameter. For example, the following function will remove the `dataChanged` event handler function added in the previous section's example.</span></span>


```js
function removeEventHandlerFromBinding() {
    Office.select("bindings#MyBinding").removeHandlerAsync(
        Office.EventType.BindingDataChanged, {handler:dataChanged});
}
```


> [!IMPORTANT]
> <span data-ttu-id="a0ab9-231">???? [removeHandlerAsync] ????????? _handler_ ????????? `eventType` ??????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-231">If the optional  _handler_ parameter is omitted when the [removeHandlerAsync] method is called, all event handlers for the specified `eventType` will be removed.</span></span>


## <a name="see-also"></a><span data-ttu-id="a0ab9-232">????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-232">See also</span></span>

- [<span data-ttu-id="a0ab9-233">????? Office ? JavaScript API</span><span class="sxs-lookup"><span data-stu-id="a0ab9-233">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md) 
- [<span data-ttu-id="a0ab9-234">Office ??????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-234">Asynchronous programming in Office Add-ins</span></span>](asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="a0ab9-235">???????????????????????</span><span class="sxs-lookup"><span data-stu-id="a0ab9-235">Read and write data to the active selection in a document or spreadsheet</span></span>](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
[Binding]:               https://dev.office.com/reference/add-ins/shared/binding
[MatrixBinding]:         https://dev.office.com/reference/add-ins/shared/binding.matrixbinding
[TableBinding]:          https://dev.office.com/reference/add-ins/shared/binding.tablebinding
[TextBinding]:           https://dev.office.com/reference/add-ins/shared/binding.textbinding
[getDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.getdataasync
[setDataAsync]:          https://dev.office.com/reference/add-ins/shared/binding.setdataasync
[SelectionChanged]:      https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedevent
[addHandlerAsync]:       https://dev.office.com/reference/add-ins/shared/binding.addhandlerasync
[removeHandlerAsync]:    https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync

[Bindings]:              https://dev.office.com/reference/add-ins/shared/bindings.bindings
[getByIdAsync]:          https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync 
[getAllAsync]:           https://dev.office.com/reference/add-ins/shared/bindings.getallasync
[addFromNamedItemAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromnameditemasync
[addFromSelectionAsync]: https://dev.office.com/reference/add-ins/shared/bindings.addfromselectionasync
[addFromPromptAsync]:    https://dev.office.com/reference/add-ins/shared/bindings.addfrompromptasync
[releaseByIdAsync]:      https://dev.office.com/reference/add-ins/shared/bindings.releasebyidasync

[AsyncResult]:          https://dev.office.com/reference/add-ins/shared/asyncresult
[Office.BindingType]:   https://dev.office.com/reference/add-ins/shared/bindingtype-enumeration
[Office.select]:        https://dev.office.com/reference/add-ins/shared/office.select 
[Office.EventType]:     https://dev.office.com/reference/add-ins/shared/eventtype-enumeration 
[Document.bindings]:    https://dev.office.com/reference/add-ins/shared/document.bindings


[TableBinding.rowCount]: https://dev.office.com/reference/add-ins/shared/binding.tablebinding.rowcount
[BindingSelectionChangedEventArgs]: https://dev.office.com/reference/add-ins/shared/binding.bindingselectionchangedeventargs
