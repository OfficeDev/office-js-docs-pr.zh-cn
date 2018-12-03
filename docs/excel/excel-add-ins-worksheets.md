---
title: 使用 Excel JavaScript API 处理工作表
description: ''
ms.date: 11/27/2018
ms.openlocfilehash: ef74dc622f3e857314874763a54df67bcff1d8ff
ms.sourcegitcommit: 026437bd3819f4e9cd4153ebe60c98ab04e18f4e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/30/2018
ms.locfileid: "26992224"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a><span data-ttu-id="74336-102">使用 Excel JavaScript API 处理工作表</span><span class="sxs-lookup"><span data-stu-id="74336-102">Work with Worksheets using the Excel JavaScript API</span></span>

<span data-ttu-id="74336-103">本文中的代码示例展示了如何使用 Excel JavaScript API 对工作表执行常见任务。</span><span class="sxs-lookup"><span data-stu-id="74336-103">This article provides code samples that show how to perform common tasks with worksheets using the Excel JavaScript API.</span></span> <span data-ttu-id="74336-104">有关 **Worksheet** 和 **WorksheetCollection** 对象支持的属性和方法的完整列表，请参阅 [Worksheet 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) 和 [WorksheetCollection 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection)。</span><span class="sxs-lookup"><span data-stu-id="74336-104">For the complete list of properties and methods that the **Worksheet** and **WorksheetCollection** objects support, see [Worksheet Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheet) and [WorksheetCollection Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.worksheetcollection).</span></span>

> [!NOTE]
> <span data-ttu-id="74336-105">本文中的信息仅适用于常规工作表；不适用于“图表”或“宏”表。</span><span class="sxs-lookup"><span data-stu-id="74336-105">The information in this article applies only to regular worksheets; it does not apply to "chart" sheets or "macro" sheets.</span></span>

## <a name="get-worksheets"></a><span data-ttu-id="74336-106">获取工作表</span><span class="sxs-lookup"><span data-stu-id="74336-106">Get worksheets</span></span>

<span data-ttu-id="74336-107">下面的代码示例获取工作表集合，加载每个工作表的 **name** 属性，并向控制台写入消息。</span><span class="sxs-lookup"><span data-stu-id="74336-107">The following code sample gets the collection of worksheets, loads the **name** property of each worksheet, and writes a message to the console.</span></span>

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
> <span data-ttu-id="74336-108">工作表的 **id** 属性用于唯一标识指定工作簿中的工作表，即使工作表被重命名或移动，其值仍不变。</span><span class="sxs-lookup"><span data-stu-id="74336-108">The **id** property of a worksheet uniquely identifies the worksheet in a given workbook and its value will remain the same even when the worksheet is renamed or moved.</span></span> <span data-ttu-id="74336-109">在 Excel for Mac 工作簿中删除工作表时，已删除工作表的 **id** 可能会重新分配到后续创建的新工作表。</span><span class="sxs-lookup"><span data-stu-id="74336-109">When a worksheet is deleted from a workbook in Excel for Mac, the **id** of the deleted worksheet may be reassigned to a new worksheet that is subsequently created.</span></span>

## <a name="get-the-active-worksheet"></a><span data-ttu-id="74336-110">获取活动工作表</span><span class="sxs-lookup"><span data-stu-id="74336-110">Get the active worksheet</span></span>

<span data-ttu-id="74336-111">下面的代码示例获取活动工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-111">The following code sample gets the active worksheet, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="set-the-active-worksheet"></a><span data-ttu-id="74336-112">设置活动工作表</span><span class="sxs-lookup"><span data-stu-id="74336-112">Set the active worksheet</span></span>

<span data-ttu-id="74336-113">下面的代码示例将活动工作表设置为名为 **Sample** 的工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-113">The following code sample sets the active worksheet to the worksheet named **Sample**, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="74336-114">如果没有使用该名称的工作表，**activate()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="74336-114">If there is no worksheet with that name, the **activate()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="reference-worksheets-by-relative-position"></a><span data-ttu-id="74336-115">通过相对位置引用工作表</span><span class="sxs-lookup"><span data-stu-id="74336-115">Reference worksheets by relative position</span></span>

<span data-ttu-id="74336-116">这些示例演示如何通过相对位置来引用工作表。</span><span class="sxs-lookup"><span data-stu-id="74336-116">These examples show how to reference a worksheet by its relative position.</span></span>

### <a name="get-the-first-worksheet"></a><span data-ttu-id="74336-117">获取第一个工作表</span><span class="sxs-lookup"><span data-stu-id="74336-117">Get the first worksheet</span></span>

<span data-ttu-id="74336-118">下面的代码示例获取工作簿中的第一个工作表，加载其 **name** 属性，并向控制台中写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-118">The following code sample gets the first worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-last-worksheet"></a><span data-ttu-id="74336-119">获取最后一个工作表</span><span class="sxs-lookup"><span data-stu-id="74336-119">Get the last worksheet</span></span>

<span data-ttu-id="74336-120">下面的代码示例获取工作簿中的最后一个工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-120">The following code sample gets the last worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="get-the-next-worksheet"></a><span data-ttu-id="74336-121">获取下一个工作表</span><span class="sxs-lookup"><span data-stu-id="74336-121">Get the next worksheet</span></span>

<span data-ttu-id="74336-122">下面的代码示例获取工作簿中活动工作表后面的工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-122">The following code sample gets the worksheet that follows the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="74336-123">如果活动工作表后没有工作表，**getNext()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="74336-123">If there is no worksheet after the active worksheet, the **getNext()** method throws an **ItemNotFound** error.</span></span>

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

### <a name="get-the-previous-worksheet"></a><span data-ttu-id="74336-124">获取上一个工作表</span><span class="sxs-lookup"><span data-stu-id="74336-124">Get the previous worksheet</span></span>

<span data-ttu-id="74336-125">下面的代码示例获取工作簿中活动工作表前面的工作表，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-125">The following code sample gets the worksheet that precedes the active worksheet in the workbook, loads its **name** property, and writes a message to the console.</span></span> <span data-ttu-id="74336-126">如果活动工作表前没有工作表，**getPrevious()** 方法将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="74336-126">If there is no worksheet before the active worksheet, the **getPrevious()** method throws an **ItemNotFound** error.</span></span>

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

## <a name="add-a-worksheet"></a><span data-ttu-id="74336-127">添加工作表</span><span class="sxs-lookup"><span data-stu-id="74336-127">Add a worksheet</span></span>

<span data-ttu-id="74336-p106">下面的代码示例向工作簿添加新工作表 **Sample**，加载它的 **name** 和 **position** 属性，并向控制台写入消息。新工作表添加在现有全部工作表的后面。</span><span class="sxs-lookup"><span data-stu-id="74336-p106">The following code sample adds a new worksheet named **Sample** to the workbook, loads its **name** and **position** properties, and writes a message to the console. The new worksheet is added after all existing worksheets.</span></span>

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

## <a name="delete-a-worksheet"></a><span data-ttu-id="74336-130">删除工作表</span><span class="sxs-lookup"><span data-stu-id="74336-130">Delete a worksheet</span></span>

<span data-ttu-id="74336-131">下面的代码示例删除工作簿中的最后一个工作表（前提是它不是工作簿中的唯一工作表），并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-131">The following code sample deletes the final worksheet in the workbook (as long as it's not the only sheet in the workbook) and writes a message to the console.</span></span>

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

## <a name="rename-a-worksheet"></a><span data-ttu-id="74336-132">重命名工作表</span><span class="sxs-lookup"><span data-stu-id="74336-132">Rename a worksheet</span></span>

<span data-ttu-id="74336-133">下面的代码示例将活动工作表的名称更改为**新名称**。</span><span class="sxs-lookup"><span data-stu-id="74336-133">The following code sample changes the name of the active worksheet to **New Name**.</span></span>

```js
Excel.run(function (context) {
    var currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    return context.sync();
}).catch(errorHandlerFunction);
```

## <a name="move-a-worksheet"></a><span data-ttu-id="74336-134">移动工作表</span><span class="sxs-lookup"><span data-stu-id="74336-134">Move a worksheet</span></span>

<span data-ttu-id="74336-135">下面的代码示例将工作表从工作簿中的最后一个位置移动到工作簿中的第一个位置。</span><span class="sxs-lookup"><span data-stu-id="74336-135">The following code sample moves a worksheet from the last position in the workbook to the first position in the workbook.</span></span>

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

## <a name="set-worksheet-visibility"></a><span data-ttu-id="74336-136">设置工作表可见性</span><span class="sxs-lookup"><span data-stu-id="74336-136">Set worksheet visibility</span></span>

<span data-ttu-id="74336-137">以下示例显示如何设置工作表的可见性。</span><span class="sxs-lookup"><span data-stu-id="74336-137">These examples show how to set the visibility of a worksheet.</span></span>

### <a name="hide-a-worksheet"></a><span data-ttu-id="74336-138">隐藏工作表</span><span class="sxs-lookup"><span data-stu-id="74336-138">Hide a worksheet</span></span>

<span data-ttu-id="74336-139">下面的代码示例将名为 **Sample** 的工作表的可见性设置为隐藏，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-139">The following code sample sets the visibility of worksheet named **Sample** to hidden, loads its **name** property, and writes a message to the console.</span></span>

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

### <a name="unhide-a-worksheet"></a><span data-ttu-id="74336-140">取消隐藏工作表</span><span class="sxs-lookup"><span data-stu-id="74336-140">Unhide a worksheet</span></span>

<span data-ttu-id="74336-141">下面的代码示例将名为 **Sample** 的工作表的可见性设置为可见，加载其 **name** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-141">The following code sample sets the visibility of worksheet named **Sample** to visible, loads its **name** property, and writes a message to the console.</span></span>

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

## <a name="get-a-cell-within-a-worksheet"></a><span data-ttu-id="74336-142">获取工作表中的单元格</span><span class="sxs-lookup"><span data-stu-id="74336-142">Get a cell within a worksheet</span></span>

<span data-ttu-id="74336-143">下面的代码示例从名为 **Sample** 的工作表获取位于第 2 行第 5 列的单元格，加载其 **address** 和 **values** 属性，并向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="74336-143">The following code sample gets the cell that is located in row 2, column 5 of the worksheet named **Sample**, loads its **address** and **values** properties, and writes a message to the console.</span></span> <span data-ttu-id="74336-144">传递给 **getCell(row: number, column:number)** 方法的值是要检索的单元格的零索引行号和列号。</span><span class="sxs-lookup"><span data-stu-id="74336-144">The values that are passed into the **getCell(row: number, column:number)** method are the zero-indexed row number and column number for the cell that is being retrieved.</span></span>

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

## <a name="get-a-range-within-a-worksheet"></a><span data-ttu-id="74336-145">获取工作表中的区域</span><span class="sxs-lookup"><span data-stu-id="74336-145">Get a range within a worksheet</span></span>

<span data-ttu-id="74336-146">有关展示了如何获取工作表中区域的示例，请参阅[使用 Excel JavaScript API 处理区域](excel-add-ins-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="74336-146">For examples that show how to get a range within a worksheet, see [Work with Ranges using the Excel JavaScript API](excel-add-ins-ranges.md).</span></span>

## <a name="data-protection"></a><span data-ttu-id="74336-147">数据保护</span><span class="sxs-lookup"><span data-stu-id="74336-147">Data protection</span></span>

<span data-ttu-id="74336-148">加载项可以控制用户能否编辑工作表中的数据。</span><span class="sxs-lookup"><span data-stu-id="74336-148">Your add-in can control a user's ability to edit data in a worksheet.</span></span> <span data-ttu-id="74336-149">工作表的 `protection` 属性是包含 `protect()` 方法的 [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) 对象。</span><span class="sxs-lookup"><span data-stu-id="74336-149">The worksheet's `protection` property is a [WorksheetProtection](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotection) object with a `protect()` method.</span></span> <span data-ttu-id="74336-150">下面的示例展示了关于切换活动工作表的完整保护的基本方案。</span><span class="sxs-lookup"><span data-stu-id="74336-150">The following example shows a basic scenario toggling the complete protection of the active worksheet.</span></span>

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

<span data-ttu-id="74336-151">`protect` 方法包含两个可选参数：</span><span class="sxs-lookup"><span data-stu-id="74336-151">The AddDependentLookup`protect` method has two parameters:</span></span>

 - <span data-ttu-id="74336-152">`options`：定义具体编辑限制的 [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) 对象。</span><span class="sxs-lookup"><span data-stu-id="74336-152">`options`: A [WorksheetProtectionOptions](https://docs.microsoft.com/javascript/api/excel/excel.worksheetprotectionoptions) object defining specific editing restrictions.</span></span>
 - <span data-ttu-id="74336-153">`password`：表示用户规避保护并编辑工作表所需使用的密码的字符串。</span><span class="sxs-lookup"><span data-stu-id="74336-153">`password`: A string representing the password needed for a user to bypass protection and edit the worksheet.</span></span>

<span data-ttu-id="74336-154">[保护工作表](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6)一文详细介绍了工作表保护，以及如何通过 Excel UI 更改保护。</span><span class="sxs-lookup"><span data-stu-id="74336-154">The article [Protect a worksheet](https://support.office.com/article/protect-a-worksheet-3179efdb-1285-4d49-a9c3-f4ca36276de6) has more information about worksheet protection and how to change it through the Excel UI.</span></span>

## <a name="see-also"></a><span data-ttu-id="74336-155">另请参阅</span><span class="sxs-lookup"><span data-stu-id="74336-155">See also</span></span>

- [<span data-ttu-id="74336-156">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="74336-156">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)

