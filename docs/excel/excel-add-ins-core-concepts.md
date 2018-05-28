---
title: Excel JavaScript API ????
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 1582268a3bdac2b7fe63c4b0a48cf1a19f85bd31
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="excel-javascript-api-core-concepts"></a><span data-ttu-id="d2d7d-102">Excel JavaScript API ????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-102">Excel JavaScript API core concepts</span></span>
 
<span data-ttu-id="d2d7d-103">???????? [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) ????? Excel 2016 ?????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-103">This article describes how to use the [Excel JavaScript API](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview) to build add-ins for Excel 2016.</span></span> <span data-ttu-id="d2d7d-104">?????????????????? API ??????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-104">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="d2d7d-105">Excel API ?????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-105">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="d2d7d-106">?? Web ? Excel ??????????????????????????? Office for Windows??? Office ???????? Office Online ?? HTML iFrame ????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-106">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online.</span></span> <span data-ttu-id="d2d7d-107">????????? Office.js API ?????????? Excel ??????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-107">Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations.</span></span> <span data-ttu-id="d2d7d-108">???Office.js ?? **sync()** API ???? [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)?? Excel ????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-108">Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions.</span></span> <span data-ttu-id="d2d7d-109">??????????????????????????????? **sync()** ?????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-109">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action.</span></span> <span data-ttu-id="d2d7d-110">??????????? **Excel.run()** ? **sync()** API ????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-110">The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="d2d7d-111">Excel.run</span><span class="sxs-lookup"><span data-stu-id="d2d7d-111">Excel.run</span></span>
 
<span data-ttu-id="d2d7d-112">**Excel.run** ???????????????? Excel ??????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-112">**Excel.run** executes a function where you specify the actions to perform against the Excel object model.</span></span> <span data-ttu-id="d2d7d-113">**Excel.run** ???????? Excel ?????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-113">**Excel.run** automatically creates a request context that you can use to interact with Excel objects.</span></span> <span data-ttu-id="d2d7d-114">?? **Excel.run** ?????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-114">When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="d2d7d-115">?????????? **Excel.run**?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-115">The following example shows how to use **Excel.run**.</span></span> <span data-ttu-id="d2d7d-116">catch ??????? **Excel.run** ???????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-116">The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
```js
Excel.run(function (context) {
  // You can use the Excel JavaScript API here in the batch function
  // to execute actions on the Excel object model.
  console.log('Your code goes here.');
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```

## <a name="request-context"></a><span data-ttu-id="d2d7d-117">?????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-117">Request context</span></span>
 
<span data-ttu-id="d2d7d-p105">Excel ????????????????????????????????? Excel ??????? **RequestContext** ?????????? Excel ???????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-p105">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="d2d7d-120">????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-120">Proxy objects</span></span>
 
<span data-ttu-id="d2d7d-121">??????????? Excel JavaScript ????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-121">The Excel JavaScript objects that you declare and use in an add-in are proxy objects.</span></span> <span data-ttu-id="d2d7d-122">?????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-122">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="d2d7d-123">??????????? `context.sync()`???? **sync()** ???????????????? Excel ????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-123">When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run.</span></span> <span data-ttu-id="d2d7d-124">???????Excel JavaScript API ??????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-124">The Excel JavaScript API is fundamentally batch-centric.</span></span> <span data-ttu-id="d2d7d-125">?????????????????????????? **sync()** ????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-125">You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="d2d7d-126">????????????? JavaScript ?? **selectedRange** ??? Excel ???????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-126">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object.</span></span> <span data-ttu-id="d2d7d-127">**SelectedRange** ???????????????????????????????????? Excel ??????????? **context.sync()**?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-127">The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="d2d7d-128">sync()</span><span class="sxs-lookup"><span data-stu-id="d2d7d-128">sync()</span></span>
 
<span data-ttu-id="d2d7d-129">????????? **sync()** ???? Excel ??????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-129">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document.</span></span> <span data-ttu-id="d2d7d-130">**Sync()** ??????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-130">The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="d2d7d-131">**sync()** ?????????????? [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise)?? **sync()** ?????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-131">The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="d2d7d-132">????????????????????? JavaScript ???? (**selectedRange**)?????????????? JavaScript Promises ???? **context.sync()** ??? Excel ????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-132">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
  selectedRange.load('address');
  return context.sync()
    .then(function () {
      console.log('The selected range is: ' + selectedRange.address);
  });
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
<span data-ttu-id="d2d7d-133">?????????? **selectedRange**????? **context.sync()** ???? **address** ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-133">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="d2d7d-134">?? **sync()** ????? promise ??????????? JavaScript ?????**??** promise?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-134">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="d2d7d-135">????????????????? **sync()** ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-135">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="d2d7d-136">???? **sync()** ??????????? [Excel JavaScript API ????](https://dev.office.com/reference/add-ins/excel/performance.md)?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-136">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://dev.office.com/reference/add-ins/excel/performance.md).</span></span>
 
### <a name="load"></a><span data-ttu-id="d2d7d-137">load()</span><span class="sxs-lookup"><span data-stu-id="d2d7d-137">load()</span></span>
 
<span data-ttu-id="d2d7d-138">?????????????????????????????? Excel ????????????????? **context.sync()**?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-138">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**.</span></span> <span data-ttu-id="d2d7d-139">??????????????????????????????? **address** ????????? **address** ????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-139">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it.</span></span> <span data-ttu-id="d2d7d-140">??????????????????????? **load()** ?????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-140">To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="d2d7d-141">??????????????????????? **load()** ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-141">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method.</span></span> <span data-ttu-id="d2d7d-142">?????????????????? **load()** ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-142">The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="d2d7d-p112">?????????????????????????????????????????????????????????? **sync()** ???????????????????????? **load()** ???????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-p112">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="d2d7d-145">????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-145">In the following example, only specific properties of the range are loaded.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
  myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);
 
  return context.sync()
    .then(function () {
      console.log (myRange.address);              // ok
      console.log (myRange.format.wrapText);      // ok
      console.log (myRange.format.fill.color);    // ok
      //console.log (myRange.format.font.color);  // not ok as it was not loaded
  });
}).then(function () {
  console.log('done');
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
<span data-ttu-id="d2d7d-146">?????????? `format/font` ??? **myRange.load()** ??????????? `format.font.color` ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-146">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="d2d7d-147">?????????????????? **load()** ?????????????????? [Excel JavaScript API????](performance.md)???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-147">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object.</span></span> <span data-ttu-id="d2d7d-148">?? **load()** ??????????? [Excel JavaScript API ????](excel-add-ins-advanced-concepts.md)?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-148">For more information about the **load()** method, see [Excel JavaScript API advanced concepts](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="d2d7d-149">null ?????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-149">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="d2d7d-150">?????? null ??</span><span class="sxs-lookup"><span data-stu-id="d2d7d-150">null input in 2-D Array</span></span>
 
<span data-ttu-id="d2d7d-151">? Excel ??????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-151">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns.</span></span> <span data-ttu-id="d2d7d-152">???????????????????????????????????????????????????????????????????? `null`?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-152">To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="d2d7d-153">????????????????????????????????????????????????????????????????????????? `null`?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-153">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells.</span></span> <span data-ttu-id="d2d7d-154">???????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-154">The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="d2d7d-155">??? null ??</span><span class="sxs-lookup"><span data-stu-id="d2d7d-155">null input for a property</span></span>
 
<span data-ttu-id="d2d7d-p116">`null` ?????????????????????????????? **values** ??????? `null`?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-p116">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="d2d7d-158">???????????????? `null` ?? **color** ???????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-158">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="d2d7d-159">???? null ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-159">null property values in the response</span></span>
 
<span data-ttu-id="d2d7d-160">???????????????? `size` ? `color` ????????????? `null` ??</span><span class="sxs-lookup"><span data-stu-id="d2d7d-160">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range.</span></span> <span data-ttu-id="d2d7d-161">???????????????? `format.font.color` ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-161">For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="d2d7d-162">??????????????????????? `range.format.font.color` ???????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-162">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="d2d7d-163">???????????????? `range.format.font.color` ? `null`?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-163">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="d2d7d-164">???????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-164">Blank input for a property</span></span>
 
<span data-ttu-id="d2d7d-p118">?????????????????????? `''`?????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-p118">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="d2d7d-167">?????? `values` ???????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-167">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="d2d7d-168">??? `numberFormat` ?????????????????? `General`?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-168">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="d2d7d-169">??? `formula` ??? `formulaLocale` ??????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-169">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="d2d7d-170">????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-170">Blank property values in the response</span></span>
 
<span data-ttu-id="d2d7d-171">??????????????????????????? `''`?????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-171">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value.</span></span> <span data-ttu-id="d2d7d-172">?????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-172">In the first example below, the first and last cell in the range contain no data.</span></span> <span data-ttu-id="d2d7d-173">????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-173">In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="d2d7d-174">?????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-174">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="d2d7d-175">??????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-175">Read an unbounded range</span></span>
 
<span data-ttu-id="d2d7d-p120">???????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-p120">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="d2d7d-178">??????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-178">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="d2d7d-179">???????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-179">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="d2d7d-180">API ????????????????`getRange('C:C')`????????????????? `values`?`text`?`numberFormat` ? `formula`?? `null` ??</span><span class="sxs-lookup"><span data-stu-id="d2d7d-180">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`.</span></span> <span data-ttu-id="d2d7d-181">???????? `address` ? `cellCount`?????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-181">Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="d2d7d-182">????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-182">Write to an unbounded range</span></span>
 
<span data-ttu-id="d2d7d-183">??????????????????????????????? `values`?`numberFormat` ? `formula`?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-183">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large.</span></span> <span data-ttu-id="d2d7d-184">???????????????????????? `values`?</span><span class="sxs-lookup"><span data-stu-id="d2d7d-184">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="d2d7d-185">????????????????????API ????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-185">The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="d2d7d-186">?????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-186">Read or write to a large range</span></span>
 
<span data-ttu-id="d2d7d-187">????????????????????/??????????????? API ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-187">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="d2d7d-188">API ???????????????????????????????????????????????????????????????? API ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-188">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="d2d7d-189">????????????????????????????????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-189">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="d2d7d-190">???????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-190">Update all cells in a range</span></span>
 
<span data-ttu-id="d2d7d-191">?????????????????????????????????????????????????????????????????? **range** ???????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-191">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="d2d7d-192">??????????? 20 ????????????????????? **3/11/2015** ????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-192">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
## <a name="error-messages"></a><span data-ttu-id="d2d7d-193">????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-193">Error messages</span></span>
 
<span data-ttu-id="d2d7d-194">?? API ????API ??????????? **error** ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-194">When an API error occurs, the API will return an **error** object that contains a code and a message.</span></span> <span data-ttu-id="d2d7d-195">????? API ??????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-195">The following table defines a list of errors that the API may return.</span></span>
 
|<span data-ttu-id="d2d7d-196">error.code</span><span class="sxs-lookup"><span data-stu-id="d2d7d-196">error.code</span></span> | <span data-ttu-id="d2d7d-197">error.message</span><span class="sxs-lookup"><span data-stu-id="d2d7d-197">error.message</span></span> |
|:----------|:--------------|
|<span data-ttu-id="d2d7d-198">InvalidArgument</span><span class="sxs-lookup"><span data-stu-id="d2d7d-198">InvalidArgument</span></span> |<span data-ttu-id="d2d7d-199">??????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-199">The argument is invalid or missing or has an incorrect format.</span></span>|
|<span data-ttu-id="d2d7d-200">InvalidRequest</span><span class="sxs-lookup"><span data-stu-id="d2d7d-200">InvalidRequest</span></span>  |<span data-ttu-id="d2d7d-201">????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-201">Cannot process the request.</span></span>|
|<span data-ttu-id="d2d7d-202">InvalidReference</span><span class="sxs-lookup"><span data-stu-id="d2d7d-202">InvalidReference</span></span>|<span data-ttu-id="d2d7d-203">????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-203">This reference is not valid for the current operation.</span></span>|
|<span data-ttu-id="d2d7d-204">InvalidBinding</span><span class="sxs-lookup"><span data-stu-id="d2d7d-204">InvalidBinding</span></span>  |<span data-ttu-id="d2d7d-205">??????????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-205">This object binding is no longer valid due to previous updates.</span></span>|
|<span data-ttu-id="d2d7d-206">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="d2d7d-206">InvalidSelection</span></span>|<span data-ttu-id="d2d7d-207">??????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-207">The current selection is invalid for this operation.</span></span>|
|<span data-ttu-id="d2d7d-208">Unauthenticated</span><span class="sxs-lookup"><span data-stu-id="d2d7d-208">Unauthenticated</span></span> |<span data-ttu-id="d2d7d-209">???????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-209">Required authentication information is either missing or invalid.</span></span>|
|<span data-ttu-id="d2d7d-210">?????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-210">AccessDenied</span></span> |<span data-ttu-id="d2d7d-211">???????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-211">You cannot perform the requested operation.</span></span>|
|<span data-ttu-id="d2d7d-212">ItemNotFound</span><span class="sxs-lookup"><span data-stu-id="d2d7d-212">ItemNotFound</span></span> |<span data-ttu-id="d2d7d-213">??????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-213">The requested resource doesn't exist.</span></span>|
|<span data-ttu-id="d2d7d-214">ActivityLimitReached</span><span class="sxs-lookup"><span data-stu-id="d2d7d-214">ActivityLimitReached</span></span>|<span data-ttu-id="d2d7d-215">????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-215">Activity limit has been reached.</span></span>|
|<span data-ttu-id="d2d7d-216">GeneralException</span><span class="sxs-lookup"><span data-stu-id="d2d7d-216">GeneralException</span></span>|<span data-ttu-id="d2d7d-217">????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-217">There was an internal error while processing the request.</span></span>|
|<span data-ttu-id="d2d7d-218">NotImplemented</span><span class="sxs-lookup"><span data-stu-id="d2d7d-218">NotImplemented</span></span>  |<span data-ttu-id="d2d7d-219">??????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-219">The requested feature isn't implemented.</span></span>|
|<span data-ttu-id="d2d7d-220">ServiceNotAvailable</span><span class="sxs-lookup"><span data-stu-id="d2d7d-220">ServiceNotAvailable</span></span>|<span data-ttu-id="d2d7d-221">??????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-221">The service is unavailable.</span></span>|
|<span data-ttu-id="d2d7d-222">??</span><span class="sxs-lookup"><span data-stu-id="d2d7d-222">Conflict</span></span>              |<span data-ttu-id="d2d7d-223">????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-223">Request could not be processed because of a conflict.</span></span>|
|<span data-ttu-id="d2d7d-224">ItemAlreadyExists</span><span class="sxs-lookup"><span data-stu-id="d2d7d-224">ItemAlreadyExists</span></span>|<span data-ttu-id="d2d7d-225">??????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-225">The resource being created already exists.</span></span>|
|<span data-ttu-id="d2d7d-226">UnsupportedOperation</span><span class="sxs-lookup"><span data-stu-id="d2d7d-226">UnsupportedOperation</span></span>|<span data-ttu-id="d2d7d-227">???????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-227">The operation being attempted is not supported.</span></span>|
|<span data-ttu-id="d2d7d-228">RequestAborted</span><span class="sxs-lookup"><span data-stu-id="d2d7d-228">RequestAborted</span></span>|<span data-ttu-id="d2d7d-229">??????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-229">The request was aborted during run time.</span></span>|
|<span data-ttu-id="d2d7d-230">ApiNotAvailable</span><span class="sxs-lookup"><span data-stu-id="d2d7d-230">ApiNotAvailable</span></span>|<span data-ttu-id="d2d7d-231">??? API ????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-231">The requested API is not available.</span></span>|
|<span data-ttu-id="d2d7d-232">InsertDeleteConflict</span><span class="sxs-lookup"><span data-stu-id="d2d7d-232">InsertDeleteConflict</span></span>|<span data-ttu-id="d2d7d-233">???????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-233">The insert or delete operation attempted resulted in a conflict.</span></span>|
|<span data-ttu-id="d2d7d-234">InvalidOperation</span><span class="sxs-lookup"><span data-stu-id="d2d7d-234">InvalidOperation</span></span>|<span data-ttu-id="d2d7d-235">????????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-235">The operation attempted is invalid on the object.</span></span>|
 
## <a name="see-also"></a><span data-ttu-id="d2d7d-236">????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-236">See also</span></span>
 
* [<span data-ttu-id="d2d7d-237">???? Excel ???</span><span class="sxs-lookup"><span data-stu-id="d2d7d-237">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="d2d7d-238">Excel ????????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-238">Excel add-ins code samples</span></span>](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [<span data-ttu-id="d2d7d-239">Excel JavaScript API????</span><span class="sxs-lookup"><span data-stu-id="d2d7d-239">Excel JavaScript API performance optimization</span></span>](https://dev.office.com/reference/add-ins/excel/performance.md)
* [<span data-ttu-id="d2d7d-240">Excel JavaScript API ??</span><span class="sxs-lookup"><span data-stu-id="d2d7d-240">Excel JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
