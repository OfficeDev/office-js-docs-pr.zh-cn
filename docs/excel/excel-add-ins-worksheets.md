---
title: 使用 Excel JavaScript API 处理工作表
description: ''
ms.date: 02/15/2018
localization_priority: Priority
ms.openlocfilehash: 6d34807b1511573c507d43dad678811c5c1592ec
ms.sourcegitcommit: 03773fef3d2a380028ba0804739d2241d4b320e5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2019
ms.locfileid: "30091244"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="7b4be-102">使用 Excel JavaScript API 处理工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-102">Work with worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="7b4be-103">本文中的代码示例展示了如何使用 Excel JavaScript API 对工作表执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="7b4be-103">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span> <span data-ttu-id="7b4be-104">有关 **Worksheet** 和 **WorksheetCollection** 对象支持的属性和方法的完整列表，请参阅 [Worksheet 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) 和 [WorksheetCollection 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)。</span><span class="sxs-lookup"><span data-stu-id="7b4be-104">For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="7b4be-105">本文中的信息仅适用于常规工作表；不适用于“图表”或“宏”表。</span><span class="sxs-lookup"><span data-stu-id="7b4be-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="7b4be-106">获取工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-106">Get worksheets</span></span>

<span data-ttu-id="7b4be-107">下面的代码示例获取工作表集合，加载每个工作表的 **name** 属性，并向控制台写入消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

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
> <span data-ttu-id="7b4be-108">工作表的 **id** 属性用于唯一标识指定工作簿中的工作表，即使工作表被重命名或移动，其值仍不变。</span><span class="sxs-lookup"><span data-stu-id="7b4be-108">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved.</span></span> <span data-ttu-id="7b4be-109">在 Excel for Mac 工作簿中删除工作表时，已删除工作表的 **id** 可能会重新分配到后续创建的新工作表。</span><span class="sxs-lookup"><span data-stu-id="7b4be-109">When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="7b4be-110">获取活动工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-110">Get the active worksheet</span></span>

<span data-ttu-id="7b4be-111">下面的代码示例获取活动工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="7b4be-112">设置活动工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-112">Set the active worksheet</span></span>

<span data-ttu-id="7b4be-113">下面的代码示例将活动工作表设置为名为 **Sample** 的工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-113">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="7b4be-114">如果没有使用该名称的工作表，**activate()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="7b4be-114">If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="7b4be-115">通过相对位置引用工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="7b4be-116">这些示例演示如何通过相对位置来引用工作表。</span><span class="sxs-lookup"><span data-stu-id="7b4be-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="7b4be-117">获取第一个工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-117">Get the first worksheet</span></span>

<span data-ttu-id="7b4be-118">下面的代码示例获取工作簿中的第一个工作表，加载其 **name** 属性，并向控制台中写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="7b4be-119">获取最后一个工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-119">Get the last worksheet</span></span>

<span data-ttu-id="7b4be-120">下面的代码示例获取工作簿中的最后一个工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="7b4be-121">获取下一个工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-121">Get the next worksheet</span></span>

<span data-ttu-id="7b4be-122">下面的代码示例获取工作簿中活动工作表后面的工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-122">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="7b4be-123">如果活动工作表后没有工作表，**getNext()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="7b4be-123">If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="7b4be-124">获取上一个工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-124">Get the previous worksheet</span></span>

<span data-ttu-id="7b4be-125">下面的代码示例获取工作簿中活动工作表前面的工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-125">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="7b4be-126">如果活动工作表前没有工作表，**getPrevious()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="7b4be-126">If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="7b4be-127">添加工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-127">Add a worksheet</span></span>

<span data-ttu-id="7b4be-p106">下面的代码示例向工作簿添加新工作表 **Sample**，加载它的 **name** 和 **position** 属性，并向控制台写入消息。新工作表添加在现有全部工作表的后面。</span><span class="sxs-lookup"><span data-stu-id="7b4be-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

## <a name="delete-a-worksheet"></a><span data-ttu-id="7b4be-130">删除工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-130">Delete a worksheet</span></span>

<span data-ttu-id="7b4be-131">下面的代码示例删除工作簿中的最后一个工作表（前提是它不是工作簿中的唯一工作表），并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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
> <span data-ttu-id="7b4be-132">不能使用 `delete` 方法删除可见性为 [VeryHidden](/javascript/api/excel/excel.sheetvisibility) 的工作表。</span><span class="sxs-lookup"><span data-stu-id="7b4be-132">A worksheet with a visibility of "[Very Hidden](/javascript/api/excel/excel.sheetvisibility)" cannot be deleted with the `delete` method.</span></span> <span data-ttu-id="7b4be-133">如果仍希望删除工作表，必须先更改可见性。</span><span class="sxs-lookup"><span data-stu-id="7b4be-133">If you wish to delete the worksheet anyway, you must first change the visibility.</span></span>

## <a name="rename-a-worksheet"></a><span data-ttu-id="7b4be-134">重命名工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-134">Rename a worksheet</span></span>

<span data-ttu-id="7b4be-135">下面的代码示例将活动工作表的名称更改为**新名称**。</span><span class="sxs-lookup"><span data-stu-id="7b4be-135">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="7b4be-136">移动工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-136">Move a worksheet</span></span>

<span data-ttu-id="7b4be-137">下面的代码示例将工作表从工作簿中的最后一个位置移动到工作簿中的第一个位置。</span><span class="sxs-lookup"><span data-stu-id="7b4be-137">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="7b4be-138">设置工作表可见性</span><span class="sxs-lookup"><span data-stu-id="7b4be-138">Set worksheet visibility</span></span>

<span data-ttu-id="7b4be-139">以下示例显示如何设置工作表的可见性。</span><span class="sxs-lookup"><span data-stu-id="7b4be-139">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="7b4be-140">隐藏工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-140">Hide a worksheet</span></span>

<span data-ttu-id="7b4be-141">下面的代码示例将名为 **Sample** 的工作表的可见性设置为隐藏，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-141">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="7b4be-142">取消隐藏工作表</span><span class="sxs-lookup"><span data-stu-id="7b4be-142">Unhide a worksheet</span></span>

<span data-ttu-id="7b4be-143">下面的代码示例将名为 **Sample** 的工作表的可见性设置为可见，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-143">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="get-a-single-cell-within-a-worksheet"></a><span data-ttu-id="7b4be-144">获取工作表中的单个单元格</span><span class="sxs-lookup"><span data-stu-id="7b4be-144">Get a single cell within a worksheet</span></span>

<span data-ttu-id="7b4be-145">下面的代码示例从名为 **Sample** 的工作表获取位于第 2 行第 5 列的单元格，加载其 **address** 和 **values** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="7b4be-145">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="7b4be-146">传递给 `getCell(row: number, column:number)` 方法的值是要检索的单元格的零索引行号和列号。</span><span class="sxs-lookup"><span data-stu-id="7b4be-146">The values that are passed into the `getCell(row: number, column:number)` method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="find-all-cells-with-matching-text-preview"></a><span data-ttu-id="7b4be-147">查找具有匹配文本 （预览） 所有单元格</span><span class="sxs-lookup"><span data-stu-id="7b4be-147">Find all cells with matching text (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="7b4be-148">工作表对象的 `findAll` 函数当前仅适用于公共预览版（beta 版本）。</span><span class="sxs-lookup"><span data-stu-id="7b4be-148">The Worksheet object's `findAll` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="7b4be-149">若要使用此功能，必须使用 Office.js CDN 的 beta 版库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="7b4be-149">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="7b4be-150">如果使用的是 TypeScript 或代码编辑器将 TypeScript 类型定义文件用于 IntelliSense，则使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。</span><span class="sxs-lookup"><span data-stu-id="7b4be-150">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="7b4be-151">`Worksheet` 对象具有 `find` 方法在工作表内搜索指定字符串。</span><span class="sxs-lookup"><span data-stu-id="7b4be-151">The `Worksheet` object has a `find` method to search for a specified string within the worksheet.</span></span> <span data-ttu-id="7b4be-152">返回 `RangeAreas` 对象，也就是可以进行一次性全部编辑的 `Range` 对象集。</span><span class="sxs-lookup"><span data-stu-id="7b4be-152">It returns a `RangeAreas` object, which is a collection of `Range` objects that can be edited all at once.</span></span> <span data-ttu-id="7b4be-153">以下代码示例查找值等于字符串 **完成** 的所有单元格，并标记为绿色。</span><span class="sxs-lookup"><span data-stu-id="7b4be-153">The following code sample finds all cells with values equal to the string **Complete** and colors them green.</span></span> <span data-ttu-id="7b4be-154">请注意，若指定的字符串不存在于工作表中，`findAll` 将引发 `ItemNotFound` 错误。</span><span class="sxs-lookup"><span data-stu-id="7b4be-154">Note that `findAll` will throw an `ItemNotFound` error if the specified string doesn't exist in the worksheet.</span></span> <span data-ttu-id="7b4be-155">若您预计到指定的字符串可能不存在工作表中，则可使用 [findAllOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) 方法，以便您的代码可正常处理该情况。</span><span class="sxs-lookup"><span data-stu-id="7b4be-155">If you expect that the specified string may not exist in the worksheet, use the [findAllOrNullObject](excel-add-ins-advanced-concepts.md#42ornullobject-methods) method instead, so your code gracefully handles that scenario.</span></span>

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
> <span data-ttu-id="7b4be-156">本节介绍如何使用 `Worksheet` 对象函数查找单元格与区域。</span><span class="sxs-lookup"><span data-stu-id="7b4be-156">This section describes how to find cells and ranges using the `Worksheet` object's functions.</span></span> <span data-ttu-id="7b4be-157">更多区域检索信息可在特定对象文章中找到。</span><span class="sxs-lookup"><span data-stu-id="7b4be-157">More range retrieval information can be found in object-specific articles.</span></span>
> - <span data-ttu-id="7b4be-158">有关展示如何使用 `Range` 对象获取工作表中区域的示例，请参阅 [使用 Excel JavaScript API 处理区域](excel-add-ins-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="7b4be-158">For examples that show how to get a range within a worksheet using the `Range` object, see [Work with ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>
> - <span data-ttu-id="7b4be-159">有关展示如何从 `Table` 对象获取区域的示例，请参阅 [使用 Excel JavaScript API 处理表](excel-add-ins-tables.md)。</span><span class="sxs-lookup"><span data-stu-id="7b4be-159">For examples that show how to get ranges from a `Table` object, see [Work with tables using the Excel JavaScript API](excel-add-ins-tables.md).</span></span>
> - <span data-ttu-id="7b4be-160">有关显示如何基于单元格特性进行多个子区域的较大区域搜索示例，请参阅 [使用 Excel 加载项同时处理多个区域](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="7b4be-160">For examples that show how to search a large range for multiple sub-ranges based on cell characteristics, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="data-protection"></a><span data-ttu-id="7b4be-161">数据保护</span><span class="sxs-lookup"><span data-stu-id="7b4be-161">Data protection</span></span>

<span data-ttu-id="7b4be-162">加载项可以控制用户能否编辑工作表中的数据。</span><span class="sxs-lookup"><span data-stu-id="7b4be-162">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="7b4be-163">工作表的 `protection` 属性是包含 `protect()` 方法的 [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) 对象。</span><span class="sxs-lookup"><span data-stu-id="7b4be-163">The worksheet's `protection` property is a [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="7b4be-164">下面的示例展示了关于切换活动工作表的完整保护的基本方案。</span><span class="sxs-lookup"><span data-stu-id="7b4be-164">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

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

<span data-ttu-id="7b4be-165">`protect` 方法包含两个可选参数：</span><span class="sxs-lookup"><span data-stu-id="7b4be-165">The `protect` method has two optional parameters:</span></span>

- <span data-ttu-id="7b4be-166">`options`：定义具体编辑限制的 [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) 对象。</span><span class="sxs-lookup"><span data-stu-id="7b4be-166">`options`: A [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
- <span data-ttu-id="7b4be-167">`password`：表示用户规避保护并编辑工作表所需使用的密码的字符串。</span><span class="sxs-lookup"><span data-stu-id="7b4be-167">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="7b4be-168">[保护工作表](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6)一文详细介绍了工作表保护，以及如何通过 Excel UI 更改保护。</span><span class="sxs-lookup"><span data-stu-id="7b4be-168">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="7b4be-169">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7b4be-169">See also</span></span>

- [<span data-ttu-id="7b4be-170">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="7b4be-170">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
