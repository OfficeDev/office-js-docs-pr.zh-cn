---
title: 使用 Excel JavaScript API 处理工作表
description: ''
ms.date: 04/18/2019
localization_priority: Priority
ms.openlocfilehash: 002c5763ebcfbbecbcfc5cb416d200b357c45bf2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448460"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="f78a4-102">使用 Excel JavaScript API 处理工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-102">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="f78a4-p101">本文提供了代码示例，介绍如何使用 Excel JavaScript API 对工作表执行常见任务。 有关 **Worksheet** 和 **WorksheetCollection** 对象支持的属性和方法的完整列表，请参阅 [Worksheet 对象 (Excel JavaScript API)](/javascript/api/excel/excel.worksheet) 和 [WorksheetCollection 对象 (Excel JavaScript API)](/javascript/api/excel/excel.worksheetcollection)。</span><span class="sxs-lookup"><span data-stu-id="f78a4-p101">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API. For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="f78a4-105">本文中的信息仅适用于常规工作表；不适用于“图表”或“宏”表。</span><span class="sxs-lookup"><span data-stu-id="f78a4-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="f78a4-106">获取工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-106">Get worksheets</span></span>

<span data-ttu-id="f78a4-107">下面的代码示例获取工作表集合，加载每个工作表的 **name** 属性，并向控制台写入消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length > 1) {
                console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
            } else {
                console.log(`There is one worksheet in the workbook:`);
            }
            for (var i in sheets.items) {
                console.log(sheets.items[i].name);
            }
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="f78a4-p102">工作表的 **id** 属性用于唯一标识指定工作簿中的工作表，即使工作表被重命名或移动，其值仍不变。 在 Excel for Mac 工作簿中删除工作表时，已删除工作表的 **id** 可能会重新分配到后续创建的新工作表。</span><span class="sxs-lookup"><span data-stu-id="f78a4-p102">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved. When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="f78a4-110">获取活动工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-110">Get the active worksheet</span></span>

<span data-ttu-id="f78a4-111">下面的代码示例获取活动工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="set-the-active-worksheet"></a><span data-ttu-id="f78a4-112">设置活动工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-112">Set the active worksheet</span></span>

<span data-ttu-id="f78a4-p103">下面的代码示例将活动工作表设置为名为 **Sample** 的工作表，加载其 **name** 属性，并向控制台写入一条消息。 如果没有使用该名称的工作表，**activate()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="f78a4-p103">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console. If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The active worksheet is "${sheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="f78a4-115">通过相对位置引用工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="f78a4-116">这些示例演示如何通过相对位置来引用工作表。</span><span class="sxs-lookup"><span data-stu-id="f78a4-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="f78a4-117">获取第一个工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-117">Get the first worksheet</span></span>

<span data-ttu-id="f78a4-118">下面的代码示例获取工作簿中的第一个工作表，加载其 **name** 属性，并向控制台中写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the first worksheet is "${firstSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-last-worksheet"></a><span data-ttu-id="f78a4-119">获取最后一个工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-119">Get the last worksheet</span></span>

<span data-ttu-id="f78a4-120">下面的代码示例获取工作簿中的最后一个工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the last worksheet is "${lastSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-next-worksheet"></a><span data-ttu-id="f78a4-121">获取下一个工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-121">Get the next worksheet</span></span>

<span data-ttu-id="f78a4-p104">下面的代码示例获取工作簿中活动工作表后面的工作表，加载其 **name** 属性，并向控制台写入一条消息。 如果活动工作表后没有工作表，**getNext()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="f78a4-p104">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

```js
 Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="f78a4-124">获取上一个工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-124">Get the previous worksheet</span></span>

<span data-ttu-id="f78a4-p105">下面的代码示例获取工作簿中活动工作表前面的工作表，加载其 **name** 属性，并向控制台写入一条消息。 如果活动工作表前没有工作表，**getPrevious()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="f78a4-p105">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console. If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    var previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="add-a-worksheet"></a><span data-ttu-id="f78a4-127">添加工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-127">Add a worksheet</span></span>

<span data-ttu-id="f78a4-p106">下面的代码示例向工作簿添加新工作表 **Sample**，加载它的 **name** 和 **position** 属性，并向控制台写入消息。新工作表添加在现有全部工作表的后面。</span><span class="sxs-lookup"><span data-stu-id="f78a4-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;

    var sheet = sheets.add("Sample");
    sheet.load("name, position");

    return context.sync()
        .then(function () {
            console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
        });
}).catch(errorHandlerFunction);
```

## <a name="delete-a-worksheet"></a><span data-ttu-id="f78a4-130">删除工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-130">Delete a worksheet</span></span>

<span data-ttu-id="f78a4-131">下面的代码示例删除工作簿中的最后一个工作表（前提是它不是工作簿中的唯一工作表），并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    return context.sync()
        .then(function () {
            if (sheets.items.length === 1) {
                console.log("Unable to delete the only worksheet in the workbook");
            } else {
                var lastSheet = sheets.items[sheets.items.length - 1];

                console.log(`Deleting worksheet named "${lastSheet.name}"`);
                lastSheet.delete();

                return context.sync();
            };
        });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="f78a4-132">不能使用 `delete` 方法删除可见性为 [VeryHidden](/javascript/api/excel/excel.sheetvisibility) 的工作表。</span><span class="sxs-lookup"><span data-stu-id="f78a4-132">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="f78a4-133">如果仍希望删除工作表，必须先更改可见性。</span><span class="sxs-lookup"><span data-stu-id="f78a4-133">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="f78a4-134">重命名工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-134">Rename a worksheet</span></span>

<span data-ttu-id="f78a4-135">下面的代码示例将活动工作表的名称更改为**新名称**。</span><span class="sxs-lookup"><span data-stu-id="f78a4-135">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="f78a4-136">移动工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-136">Move a worksheet</span></span>

<span data-ttu-id="f78a4-137">下面的代码示例将工作表从工作簿中的最后一个位置移动到工作簿中的第一个位置。</span><span class="sxs-lookup"><span data-stu-id="f78a4-137">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

```js
Excel.run(function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items");

    return context.sync()
        .then(function () {
            var lastSheet = sheets.items[sheets.items.length - 1];
            lastSheet.position = 0;

            return context.sync();
        });
}).catch(errorHandlerFunction);
```

## <a name="set-worksheet-visibility"></a><span data-ttu-id="f78a4-138">设置工作表可见性</span><span class="sxs-lookup"><span data-stu-id="f78a4-138">Set worksheet visibility</span></span>

<span data-ttu-id="f78a4-139">以下示例显示如何设置工作表的可见性。</span><span class="sxs-lookup"><span data-stu-id="f78a4-139">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="f78a4-140">隐藏工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-140">Hide a worksheet</span></span>

<span data-ttu-id="f78a4-141">下面的代码示例将名为 **Sample** 的工作表的可见性设置为隐藏，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-141">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is hidden`);
        });
}).catch(errorHandlerFunction);
```

### <a name="unhide-a-worksheet"></a><span data-ttu-id="f78a4-142">取消隐藏工作表</span><span class="sxs-lookup"><span data-stu-id="f78a4-142">Unhide a worksheet</span></span>

<span data-ttu-id="f78a4-143">下面的代码示例将名为 **Sample** 的工作表的可见性设置为可见，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-143">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    return context.sync()
        .then(function () {
            console.log(`Worksheet with name "${sheet.name}" is visible`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="f78a4-144">获取工作表中的单个单元格</span><span class="sxs-lookup"><span data-stu-id="f78a4-144">Get a single cell within a worksheet</span></span>

<span data-ttu-id="f78a4-145">下面的代码示例从名为 **Sample** 的工作表获取位于第 2 行第 5 列的单元格，加载其 **address** 和 **values** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-145">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="f78a4-146">传递给 `getCell(row: number, column:number)` 方法的值是要检索的单元格的零索引行号和列号。</span><span class="sxs-lookup"><span data-stu-id="f78a4-146">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var cell = sheet.getCell(1, 4);
    cell.load("address, values");

    return context.sync()
        .then(function() {
            console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
        })
}).catch(errorHandlerFunction);
```

## <a name="detect-data-changes"></a><span data-ttu-id="f78a4-147">检测数据更改</span><span class="sxs-lookup"><span data-stu-id="f78a4-147">Detect data changes</span></span>

<span data-ttu-id="f78a4-148">加载项可能需要回应对工作表中的数据进行更改的用户。</span><span class="sxs-lookup"><span data-stu-id="f78a4-148">Your add-in may need to react to users changing the data in a worksheet.</span></span> <span data-ttu-id="f78a4-149">若要检测这些更改，可以为工作表的 `onChanged` 事件[注册事件处理程序](excel-add-ins-events.md#register-an-event-handler)。</span><span class="sxs-lookup"><span data-stu-id="f78a4-149">To detect these changes, you can [register an event handler](excel-add-ins-events.md#register-an-event-handler) for the `onChanged` event of a worksheet.</span></span> <span data-ttu-id="f78a4-150">当事件触发时，`onChanged` 事件的事件处理程序将收到 [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) 对象。</span><span class="sxs-lookup"><span data-stu-id="f78a4-150">Event handlers for the `onChanged` event receive a [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) object when the event fires.</span></span>

<span data-ttu-id="f78a4-151">`WorksheetChangedEventArgs` 对象提供有关更改和来源的信息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-151">The `WorksheetChangedEventArgs` object provides information about the changes and the source.</span></span> <span data-ttu-id="f78a4-152">由于 `onChanged` 会在数据的格式或值发生变化时触发，因此让加载项检查值是否已实际更改可能很有用。</span><span class="sxs-lookup"><span data-stu-id="f78a4-152">Since `onChanged` fires when either the format or value of the data changes, it can be useful to have your add-in check if the values have actually changed.</span></span> <span data-ttu-id="f78a4-153">`details` 属性以 [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) 的形式封装此信息。</span><span class="sxs-lookup"><span data-stu-id="f78a4-153">The `details` property encapsulates this information as a [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail).</span></span> <span data-ttu-id="f78a4-154">以下代码示例演示如何显示已更改的单元格的之前和之后的值及类型。</span><span class="sxs-lookup"><span data-stu-id="f78a4-154">The following code sample shows how to display the before and after values and types of a cell that has been changed.</span></span>

> [!NOTE]
> <span data-ttu-id="f78a4-155">`WorksheetChangedEventArgs.details` 当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="f78a4-155">`WorksheetChangedEventArgs.details` is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

```js
// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
    Excel.run(function (context) {
        var details = eventArgs.details;
        var address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="find-all-cells-with-matching-text-preview"></a><span data-ttu-id="f78a4-156">查找具有匹配文本 （预览） 所有单元格</span><span class="sxs-lookup"><span data-stu-id="f78a4-156">Find all cells with matching text (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f78a4-p112">工作表对象的 `findAll` 函数当前仅适用于公共预览版。[!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]</span><span class="sxs-lookup"><span data-stu-id="f78a4-p112">The Worksheet object's `findAll` function is currently available only in public preview. [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]</span></span>

<span data-ttu-id="f78a4-158">`Worksheet` 对象具有 `find` 方法在工作表内搜索指定字符串。</span><span class="sxs-lookup"><span data-stu-id="f78a4-158">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="f78a4-159">返回 `RangeAreas` 对象，也就是可以进行一次性全部编辑的 `Range` 对象集。</span><span class="sxs-lookup"><span data-stu-id="f78a4-159">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="f78a4-160">以下代码示例查找值等于字符串 **完成** 的所有单元格，并标记为绿色。</span><span class="sxs-lookup"><span data-stu-id="f78a4-160">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="f78a4-161">请注意，若指定的字符串不存在于工作表中，`findAll` 将引发 `ItemNotFound` 错误。</span><span class="sxs-lookup"><span data-stu-id="f78a4-161">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="f78a4-162">若您预计到指定的字符串可能不存在工作表中，则可使用 [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) 方法，以便您的代码可正常处理该情况。</span><span class="sxs-lookup"><span data-stu-id="f78a4-162">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](excel-add-ins-advanced-concepts.md#ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    return context.sync()
        .then(function() {
            foundRanges.format.fill.color = "green"
    });
}).catch(errorHandlerFunction);
```

> [!NOTE]
> <span data-ttu-id="f78a4-163">本节介绍如何使用 `Worksheet` 对象函数查找单元格与区域。</span><span class="sxs-lookup"><span data-stu-id="f78a4-163">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="f78a4-164">更多区域检索信息可在特定对象文章中找到。</span><span class="sxs-lookup"><span data-stu-id="f78a4-164">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="f78a4-165">有关展示如何使用 `Range` 对象获取工作表中区域的示例，请参阅 [使用 Excel JavaScript API 处理区域](excel-add-ins-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="f78a4-165">For examples that show how to get a range within a worksheet using the `Range` object, see [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>
> - <span data-ttu-id="f78a4-166">有关展示如何从 `Table` 对象获取区域的示例，请参阅 [使用 Excel JavaScript API 处理表](excel-add-ins-tables.md)。</span><span class="sxs-lookup"><span data-stu-id="f78a4-166">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="f78a4-167">有关显示如何基于单元格特性进行多个子区域的较大区域搜索示例，请参阅 [使用 Excel 加载项同时处理多个区域](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="f78a4-167">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="filter-data"></a><span data-ttu-id="f78a4-168">筛选数据</span><span class="sxs-lookup"><span data-stu-id="f78a4-168">Filter data</span></span>

> [!NOTE]
> <span data-ttu-id="f78a4-169">`AutoFilter` 当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="f78a4-169">`AutoFilter` is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f78a4-170">[自动筛选](/javascript/api/excel/excel.autofilter)在工作表的一个范围内应用数据筛选器。</span><span class="sxs-lookup"><span data-stu-id="f78a4-170">An [AutoFilter](/javascript/api/excel/excel.autofilter) applies data filters across a range within the worksheet.</span></span> <span data-ttu-id="f78a4-171">这是通过 `Worksheet.autoFilter.apply` 创建的，它具有以下属性：</span><span class="sxs-lookup"><span data-stu-id="f78a4-171">This is created with `Worksheet.autoFilter.apply`, which has the following parameters:</span></span>

- <span data-ttu-id="f78a4-172">`range`：应用筛选器的范围，指定为 `Range` 对象或字符串。</span><span class="sxs-lookup"><span data-stu-id="f78a4-172">`range`: The range to which the filter is applied, specified as either a `Range` object or a string.</span></span>
- <span data-ttu-id="f78a4-173">`columnIndex`：从零开始的列索引，根据该索引评估筛选条件。</span><span class="sxs-lookup"><span data-stu-id="f78a4-173">`columnIndex`: The zero-based column index against which the filter criteria is evaluated.</span></span>
- <span data-ttu-id="f78a4-174">`criteria`：[FilterCriteria](/javascript/api/excel/excel.filtercriteria) 对象，该对象确定应基于列的单元格筛选哪些行。</span><span class="sxs-lookup"><span data-stu-id="f78a4-174">`criteria`: A [FilterCriteria](/javascript/api/excel/excel.filtercriteria) object determining which rows should be filtered based on the column's cell.</span></span>

<span data-ttu-id="f78a4-175">第一个代码示例显示如何将筛选器添加到工作表的已使用区域。</span><span class="sxs-lookup"><span data-stu-id="f78a4-175">The first code sample shows how to add a filter to the worksheet's used range.</span></span> <span data-ttu-id="f78a4-176">此筛选器将基于列 **3** 中的值，隐藏不在前 25% 内的条目。</span><span class="sxs-lookup"><span data-stu-id="f78a4-176">This filter will hide entries that are not in the top 25%, based on the values in column **3**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f78a4-177">下一个代码示例显示如何使用 `reapply` 方法刷新 auto-filter。</span><span class="sxs-lookup"><span data-stu-id="f78a4-177">The next code sample shows how to refresh the auto-filter using the `reapply` method.</span></span> <span data-ttu-id="f78a4-178">当范围中的数据更改时，应执行此操作。</span><span class="sxs-lookup"><span data-stu-id="f78a4-178">This should be done when the data in the range changes.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f78a4-179">最终的自动筛选代码示例显示如何使用 `remove` 方法将 auto-filter 从工作表移除。</span><span class="sxs-lookup"><span data-stu-id="f78a4-179">The final auto-filter code sample shows how to remove the auto-filter from the worksheet with the `remove` method.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f78a4-180">`AutoFilter` 也可应用到单个表。</span><span class="sxs-lookup"><span data-stu-id="f78a4-180">An `AutoFilter` can also be applied to individual tables.</span></span> <span data-ttu-id="f78a4-181">有关详细信息，请参阅[使用 Excel JavaScript API 处理表](excel-add-ins-tables.md#autofilter)。</span><span class="sxs-lookup"><span data-stu-id="f78a4-181">See [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md#autofilter) for more information.</span></span>

## <a name="data-protection"></a><span data-ttu-id="f78a4-182">数据保护</span><span class="sxs-lookup"><span data-stu-id="f78a4-182">Data protection</span></span>

<span data-ttu-id="f78a4-183">加载项可以控制用户能否编辑工作表中的数据。</span><span class="sxs-lookup"><span data-stu-id="f78a4-183">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="f78a4-184">工作表的 `protection` 属性是包含 `protect()` 方法的 [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) 对象。</span><span class="sxs-lookup"><span data-stu-id="f78a4-184">The worksheet's `protection` property is a [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="f78a4-185">下面的示例展示了关于切换活动工作表的完整保护的基本方案。</span><span class="sxs-lookup"><span data-stu-id="f78a4-185">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");

    return context.sync().then(function() {
        if (!activeSheet.protection.protected) {
            activeSheet.protection.protect();
        }
    })
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f78a4-186">`protect` 方法包含两个可选参数：</span><span class="sxs-lookup"><span data-stu-id="f78a4-186">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="f78a4-187">`options`：定义具体编辑限制的 [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) 对象。</span><span class="sxs-lookup"><span data-stu-id="f78a4-187">`options`: A [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="f78a4-188">`password`：表示用户规避保护并编辑工作表所需使用的密码的字符串。</span><span class="sxs-lookup"><span data-stu-id="f78a4-188">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="f78a4-189">[保护工作表](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6)一文详细介绍了工作表保护，以及如何通过 Excel UI 更改保护。</span><span class="sxs-lookup"><span data-stu-id="f78a4-189">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="page-layout-and-print-settings"></a><span data-ttu-id="f78a4-190">页面布局和打印设置</span><span class="sxs-lookup"><span data-stu-id="f78a4-190">Page layout and print settings</span></span>

> [!NOTE]
> <span data-ttu-id="f78a4-191">此部分中与页面布局关联的 API 目前仅在公共预览版中可用。</span><span class="sxs-lookup"><span data-stu-id="f78a4-191">The APIs in this section associated with page layout are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f78a4-192">加载项可以在工作表级别访问页面布局设置。</span><span class="sxs-lookup"><span data-stu-id="f78a4-192">Add-ins have access to page layout settings at a worksheet level.</span></span> <span data-ttu-id="f78a4-193">这些控制打印工作表的方式。</span><span class="sxs-lookup"><span data-stu-id="f78a4-193">These control how the sheet is printed.</span></span> <span data-ttu-id="f78a4-194">`Worksheet` 对象有三个与布局相关的属性：`horizontalPageBreaks`、`verticalPageBreaks`、`pageLayout`。</span><span class="sxs-lookup"><span data-stu-id="f78a4-194">A `Worksheet` object has three layout-related properties: `horizontalPageBreaks`, `verticalPageBreaks`, `pageLayout`.</span></span>

<span data-ttu-id="f78a4-195">`Worksheet.horizontalPageBreaks` 和 `Worksheet.verticalPageBreaks` 是 [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection)。</span><span class="sxs-lookup"><span data-stu-id="f78a4-195">`Worksheet.horizontalPageBreaks` and `Worksheet.verticalPageBreaks` are [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection).</span></span> <span data-ttu-id="f78a4-196">这些是 [PageBreaks](/javascript/api/excel/excel.pagebreak) 的集合，其中指定插入手动分页符的范围。</span><span class="sxs-lookup"><span data-stu-id="f78a4-196">These are collections of [PageBreaks](/javascript/api/excel/excel.pagebreak), which specify ranges where manual page breaks are inserted.</span></span> <span data-ttu-id="f78a4-197">以下代码示例在第 **21** 行上方添加了水平分页符。</span><span class="sxs-lookup"><span data-stu-id="f78a4-197">The following code sample adds a horizontal page break above row **21**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f78a4-198">`Worksheet.pageLayout` 是 [PageLayout](/javascript/api/excel/excel.pagelayout) 对象。</span><span class="sxs-lookup"><span data-stu-id="f78a4-198">`Worksheet.pageLayout` is a [PageLayout](/javascript/api/excel/excel.pagelayout) object.</span></span> <span data-ttu-id="f78a4-199">此对象包含不依赖于任何打印机特定实现的布局和打印设置。</span><span class="sxs-lookup"><span data-stu-id="f78a4-199">This object contains layout and print settings that are not dependant any printer-specific implementation.</span></span> <span data-ttu-id="f78a4-200">这些设置包括页边距、方向、页码编号、标题行，并打印区域。</span><span class="sxs-lookup"><span data-stu-id="f78a4-200">These settings include margins, orientation, page numbering, title rows, and print area.</span></span>

<span data-ttu-id="f78a4-201">以下代码示例使页面居中（垂直和水平），设置将在每页顶部打印的标题行，并将打印区域设置为工作表的子部分。</span><span class="sxs-lookup"><span data-stu-id="f78a4-201">The following code sample centers the page (both vertically and horizontally), sets a title row that will be printed at the top of every page, and sets the printed area to a subsection of the worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="f78a4-202">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f78a4-202">See also</span></span>

- [<span data-ttu-id="f78a4-203">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="f78a4-203">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
