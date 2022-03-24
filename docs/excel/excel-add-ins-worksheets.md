---
title: 使用 Excel JavaScript API 处理工作表
description: 显示如何使用 JavaScript API 对工作表执行常见Excel示例。
ms.date: 02/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 565a186220fb9b9a33d97ad73954fe405658cf97
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743391"
---
# <a name="work-with-worksheets-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理工作表

本文提供了代码示例，介绍如何使用 Excel JavaScript API 对工作表执行常见任务。 有关 `Worksheet` 和 `WorksheetCollection` 对象支持的属性和方法的完整列表，请参阅 [Worksheet 对象 (Excel JavaScript API)](/javascript/api/excel/excel.worksheet) 和 [WorksheetCollection 对象 (Excel JavaScript API)](/javascript/api/excel/excel.worksheetcollection)。

> [!NOTE]
> 本文中的信息仅适用于常规工作表；不适用于“图表”或“宏”表。

## <a name="get-worksheets"></a>获取工作表

下面的代码示例获取工作表集合，加载每个工作表的 `name` 属性，并向控制台写入一条消息。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();
    
    if (sheets.items.length > 1) {
        console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
    } else {
        console.log(`There is one worksheet in the workbook:`);
    }

    sheets.items.forEach(function (sheet) {
        console.log(sheet.name);
    });
});
```

> [!NOTE]
> 工作表的 `id` 属性用于唯一标识指定工作簿中的工作表，即使工作表被重命名或移动，此属性的值也仍保持不变。如果工作表从 Mac 版 Excel 的工作簿中删除，已删除工作表的 `id` 可能会重新分配给后续创建的新工作表。

## <a name="get-the-active-worksheet"></a>获取活动工作表

下面的代码示例获取活动工作表，加载其 `name` 属性，并向控制台写入一条消息。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    await context.sync();
    console.log(`The active worksheet is "${sheet.name}"`);
});
```

## <a name="set-the-active-worksheet"></a>设置活动工作表

下面的代码示例将活动工作表设置为名为 **Sample** 的工作表，加载其 `name` 属性，并向控制台写入一条消息。 如果没有使用该名称的工作表，`activate()` 方法将引发 `ItemNotFound` 错误。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    sheet.activate();
    sheet.load("name");

    await context.sync();
    console.log(`The active worksheet is "${sheet.name}"`);
});
```

## <a name="reference-worksheets-by-relative-position"></a>通过相对位置引用工作表

这些示例演示如何通过相对位置来引用工作表。

### <a name="get-the-first-worksheet"></a>获取第一个工作表

下面的代码示例获取工作簿中的第一个工作表，加载其 `name` 属性，并向控制台中写入一条消息。

```js
await Excel.run(async (context) => {
    let firstSheet = context.workbook.worksheets.getFirst();
    firstSheet.load("name");

    await context.sync();
    console.log(`The name of the first worksheet is "${firstSheet.name}"`);
});
```

### <a name="get-the-last-worksheet"></a>获取最后一个工作表

下面的代码示例获取工作簿中的最后一个工作表，加载其 `name` 属性，并向控制台写入一条消息。

```js
await Excel.run(async (context) => {
    let lastSheet = context.workbook.worksheets.getLast();
    lastSheet.load("name");

    await context.sync();
    console.log(`The name of the last worksheet is "${lastSheet.name}"`);
});
```

### <a name="get-the-next-worksheet"></a>获取下一个工作表

下面的代码示例获取工作簿中活动工作表后面的工作表，加载其 `name` 属性，并向控制台写入一条消息。 如果活动工作表后没有工作表，`getNext()` 方法将引发 `ItemNotFound` 错误。

```js
await Excel.run(async (context) => {
    let currentSheet = context.workbook.worksheets.getActiveWorksheet();
    let nextSheet = currentSheet.getNext();
    nextSheet.load("name");

    await context.sync();
    console.log(`The name of the sheet that follows the active worksheet is "${nextSheet.name}"`);
});
```

### <a name="get-the-previous-worksheet"></a>获取上一个工作表

下面的代码示例获取工作簿中活动工作表前面的工作表，加载其 `name` 属性，并向控制台写入一条消息。 如果活动工作表前没有工作表，`getPrevious()` 方法将引发 `ItemNotFound` 错误。

```js
await Excel.run(async (context) => {
    let currentSheet = context.workbook.worksheets.getActiveWorksheet();
    let previousSheet = currentSheet.getPrevious();
    previousSheet.load("name");

    await context.sync();
    console.log(`The name of the sheet that precedes the active worksheet is "${previousSheet.name}"`);
});
```

## <a name="add-a-worksheet"></a>添加工作表

下面的代码示例向工作簿添加新工作表 **Sample**，加载它的 `name` 和 `position` 属性，并向控制台写入消息。新工作表添加在现有全部工作表的后面。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;

    let sheet = sheets.add("Sample");
    sheet.load("name, position");

    await context.sync();
    console.log(`Added worksheet named "${sheet.name}" in position ${sheet.position}`);
});
```

### <a name="copy-an-existing-worksheet"></a>复制现有工作表

`Worksheet.copy` 通过复制现有工作表添加新工作表。 新工作表的名称将在末尾附加一个数字，格式与通过 Excel UI 复制工作表一致（例如 **MySheet (2)**）。 `Worksheet.copy` 可采用两个参数，且两者都是可选参数：

- `positionType` - 一个 [WorksheetPositionType](/javascript/api/excel/excel.worksheetpositiontype) 枚举，指定在工作簿中添加新工作表的位置。
- `relativeTo` - 如果 `positionType` 为 `Before` 或 `After`，则需要指定一个参考工作表，新工作表将相对于此工作表进行添加（此参数回答的问题是“在什么之前或之后？”）。

下面的代码示例复制当前工作表，并将新工作表直接插入到当前工作表之后。

```js
await Excel.run(async (context) => {
    let myWorkbook = context.workbook;
    let sampleSheet = myWorkbook.worksheets.getActiveWorksheet();
    let copiedSheet = sampleSheet.copy(Excel.WorksheetPositionType.after, sampleSheet);
    await context.sync();
});
```

## <a name="delete-a-worksheet"></a>删除工作表

下面的代码示例删除工作簿中的最后一个工作表（前提是它不是工作簿中的唯一工作表），并向控制台写入一条消息。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();
    if (sheets.items.length === 1) {
        console.log("Unable to delete the only worksheet in the workbook");
    } else {
        let lastSheet = sheets.items[sheets.items.length - 1];

        console.log(`Deleting worksheet named "${lastSheet.name}"`);
        lastSheet.delete();

        await context.sync();
    }
});
```

> [!NOTE]
> 不能使用 `delete` 方法删除可见性为 [VeryHidden](/javascript/api/excel/excel.sheetvisibility) 的工作表。 如果仍希望删除工作表，必须先更改可见性。

## <a name="rename-a-worksheet"></a>重命名工作表

下面的代码示例将活动工作表的名称更改为 **新名称**。

```js
await Excel.run(async (context) => {
    let currentSheet = context.workbook.worksheets.getActiveWorksheet();
    currentSheet.name = "New Name";

    await context.sync();
});
```

## <a name="move-a-worksheet"></a>移动工作表

下面的代码示例将工作表从工作簿中的最后一个位置移动到工作簿中的第一个位置。

```js
await Excel.run(async (context) => {
    let sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();

    let lastSheet = sheets.items[sheets.items.length - 1];
    lastSheet.position = 0;
    await context.sync();
});
```

## <a name="set-worksheet-visibility"></a>设置工作表可见性

以下示例显示如何设置工作表的可见性。

### <a name="hide-a-worksheet"></a>隐藏工作表

下面的代码示例将名为 **Sample** 的工作表的可见性设置为隐藏，加载其 `name` 属性，并向控制台写入一条消息。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.hidden;
    sheet.load("name");

    await context.sync();
    console.log(`Worksheet with name "${sheet.name}" is hidden`);
});
```

### <a name="unhide-a-worksheet"></a>取消隐藏工作表

下面的代码示例将名为 **Sample** 的工作表的可见性设置为可见，加载其 `name` 属性，并向控制台写入一条消息。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    sheet.visibility = Excel.SheetVisibility.visible;
    sheet.load("name");

    await context.sync();
    console.log(`Worksheet with name "${sheet.name}" is visible`);
});
```

## <a name="get-a-single-cell-within-a-worksheet"></a>获取工作表中的单个单元格

下面的代码示例从名为 **Sample** 的工作表获取位于第 2 行第 5 列的单元格，加载其 `address` 和 `values` 属性，并向控制台写入一条消息。 传递给 `getCell(row: number, column:number)` 方法的值是要检索的单元格的零索引行号和列号。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let cell = sheet.getCell(1, 4);
    cell.load("address, values");

    await context.sync();
    console.log(`The value of the cell in row 2, column 5 is "${cell.values[0][0]}" and the address of that cell is "${cell.address}"`);
});
```

## <a name="detect-data-changes"></a>检测数据更改

加载项可能需要回应对工作表中的数据进行更改的用户。 若要检测这些更改，可以为工作表的 `onChanged` 事件[注册事件处理程序](excel-add-ins-events.md#register-an-event-handler)。 当事件触发时，`onChanged` 事件的事件处理程序将收到 [WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs) 对象。

`WorksheetChangedEventArgs` 对象提供有关更改和来源的信息。 由于 `onChanged` 会在数据的格式或值发生变化时触发，因此让加载项检查值是否已实际更改可能很有用。 `details` 属性以 [ChangedEventDetail](/javascript/api/excel/excel.changedeventdetail) 的形式封装此信息。 以下代码示例演示如何显示已更改的单元格的之前和之后的值及类型。

```js
// This function would be used as an event handler for the Worksheet.onChanged event.
function onWorksheetChanged(eventArgs) {
    Excel.run(function (context) {
        let details = eventArgs.details;
        let address = eventArgs.address;

        // Print the before and after types and values to the console.
        console.log(`Change at ${address}: was ${details.valueBefore}(${details.valueTypeBefore}),`
            + ` now is ${details.valueAfter}(${details.valueTypeAfter})`);
        return context.sync();
    });
}
```

## <a name="detect-formula-changes"></a>检测公式更改

加载项可以跟踪对工作表中的公式所做的更改。 当工作表连接到外部数据库时，这很有用。 当工作表中的公式发生更改时，此方案中的事件将触发外部数据库中的相应更新。

若要检测对公式的更改， [请](excel-add-ins-events.md#register-an-event-handler) 为工作表的 [onFormulaChanged 事件注册](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member) 事件处理程序。 事件的事件处理程序在 `onFormulaChanged` 事件触发时接收 [WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs) 对象。

> [!IMPORTANT]
> 该事件 `onFormulaChanged` 检测公式本身何时更改，而不是由公式计算产生的数据值。

下面的代码示例演示如何注册事件处理程序，使用 对象检索已更改公式的 [formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-formuladetails-member) 数组，然后使用 [FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail) 属性打印有关已更改公式的详细信息。`onFormulaChanged` `WorksheetFormulaChangedEventArgs`

> [!NOTE]
> 此代码示例仅在更改单个公式时有效。

```js
async function run() {
    await Excel.run(async (context) => {
        // Retrieve the worksheet named "Sample".
        let sheet = context.workbook.worksheets.getItem("Sample");
    
        // Register the formula changed event handler for this worksheet.
        sheet.onFormulaChanged.add(formulaChangeHandler);
    
        await context.sync();
    });
}

async function formulaChangeHandler(event) {
    await Excel.run(async (context) => {
        // Retrieve details about the formula change event.
        // Note: This method assumes only a single formula is changed at a time. 
        let cellAddress = event.formulaDetails[0].cellAddress;
        let previousFormula = event.formulaDetails[0].previousFormula;
        let source = event.source;
    
        // Print out the change event details.
        console.log(
          `The formula in cell ${cellAddress} changed. 
          The previous formula was: ${previousFormula}. 
          The source of the change was: ${source}.`
        );         
    });
}
```

## <a name="handle-sorting-events"></a>处理排序事件

`onColumnSorted` 和 `onRowSorted` 事件表示工作表数据已排序。 这些事件连接到各 `Worksheet` 对象和工作簿的 `WorkbookCollection` 无论是通过编程排序还是通过 Excel 用户界面手动执行排序，它们都会触发。

> [!NOTE]
> 通过从左到右排序操作对列排序时，触发 `onColumnSorted` 通过从上到下排序操作对行排序时，触发 `onRowSorted` 使用列标题上的下拉菜单对表格进行排序时，将触发 `onRowSorted` 事件。 该事件对应于正在移动的内容，而不是排序条件。

`onColumnSorted` 和 `onRowSorted` 事件为它们的回叫分别提供 [WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs) 或 [WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs) 它们提供有关事件的更多详细信息。 特别的一点是，两个 `EventArgs` 都有 `address` 属性，表示排序操作移动的行或列。 已添加包含排序内容的所有单元格，即使单元格的值未包含在排序条件中，也是如此。

下图显示了排序事件的 `address` 属性返回的范围。 首先是排序前的示例数据：

![排序前Excel表中的数据。](../images/excel-sort-event-before.png)

如果对"**Q1**"执行从上到下排序 (**"B**") 中，则返回以下突出显示的行 `WorksheetRowSortedEventArgs.address`。

![从上到下排序后 Excel 中的表格数据。 已移动的行会突出显示。](../images/excel-sort-event-after-row.png)

如果对"**Quinces**"对象执行从左到右排序 (原始数据上的"**4**") 中的值，则返回以下突出显示的列 `WorksheetColumnsSortedEventArgs.address`。

![从左到右排序后 Excel 中的表格数据。 已移动的列会突出显示。](../images/excel-sort-event-after-column.png)

下面的代码示例演示如何为 `Worksheet.onRowSorted` 事件注册事件处理程序。 处理程序的回叫会清除该范围的填充颜色，然后填充已移动行的单元格。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // This will fire whenever a row has been moved as the result of a sort action.
    sheet.onRowSorted.add(async (event) => {
        await Excel.run(async (context) => {
            console.log("Row sorted: " + event.address);
            let sheet = context.workbook.worksheets.getActiveWorksheet();

            // Clear formatting for section, then highlight the sorted area.
            sheet.getRange("A1:E5").format.fill.clear();
            if (event.address !== "") {
                sheet.getRanges(event.address).format.fill.color = "yellow";
            }

            await context.sync();
        });
    });

    await context.sync();
});
```

## <a name="find-all-cells-with-matching-text"></a>查找所有包含匹配文本的单元格

`Worksheet` 对象具有 `find` 方法在工作表内搜索指定字符串。 返回 `RangeAreas` 对象，也就是可以进行一次性全部编辑的 `Range` 对象集。 以下代码示例查找值等于字符串 **完成** 的所有单元格，并标记为绿色。 请注意，若指定的字符串不存在于工作表中，`findAll` 将引发 `ItemNotFound` 错误。 若您预计到指定的字符串可能不存在工作表中，则可使用 [findAllOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) 方法，以便您的代码可正常处理该情况。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let foundRanges = sheet.findAll("Complete", {
        completeMatch: true, // findAll will match the whole cell value
        matchCase: false // findAll will not match case
    });

    await context.sync();
    foundRanges.format.fill.color = "green"
});
```

> [!NOTE]
> 本节介绍如何使用 `Worksheet` 对象函数查找单元格与区域。 更多区域检索信息可在特定对象文章中找到。
>
> - 有关显示如何使用 对象获取`Range`工作表中的区域的示例，请参阅使用 [JavaScript API](excel-add-ins-ranges-get.md) 获取Excel区域。
> - 有关展示如何从 `Table` 对象获取区域的示例，请参阅 [使用 Excel JavaScript API 处理表](excel-add-ins-tables.md)。
> - 有关显示如何基于单元格特性进行多个子区域的较大区域搜索示例，请参阅 [使用 Excel 加载项同时处理多个区域](excel-add-ins-multiple-ranges.md)。

## <a name="filter-data"></a>筛选数据

[自动筛选](/javascript/api/excel/excel.autofilter)在工作表的一个范围内应用数据筛选器。 这是通过 创建的 `Worksheet.autoFilter.apply`，它具有以下参数。

- `range`：应用筛选器的范围，指定为 `Range` 对象或字符串。
- `columnIndex`：从零开始的列索引，根据该索引评估筛选条件。
- `criteria`：[FilterCriteria](/javascript/api/excel/excel.filtercriteria) 对象，该对象确定应基于列的单元格筛选哪些行。

第一个代码示例显示如何将筛选器添加到工作表的已使用区域。 此筛选器将基于列 **3** 中的值，隐藏不在前 25% 内的条目。

```js
// This method adds a custom AutoFilter to the active worksheet
// and applies the filter to a column of the used range.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let farmData = sheet.getUsedRange();

    // This filter will only show the rows with the top 25% of values in column 3.
    sheet.autoFilter.apply(farmData, 3, { criterion1: "25", filterOn: Excel.FilterOn.topPercent });
    await context.sync();
});
```

下一个代码示例显示如何使用 `reapply` 方法刷新 auto-filter。 当范围中的数据更改时，应执行此操作。

```js
// This method refreshes the AutoFilter to ensure that changes are captured.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.reapply();
    await context.sync();
});
```

下面的代码示例演示如何 `clearColumnCriteria` 使用 方法从一个列清除自动筛选，同时使筛选器在其他列上保持活动状态。

```js
// This method clears the AutoFilter setting from one column.
await Excel.run(async (context) => {
    // Retrieve the active worksheet.
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Clear the filter from only column 3.
    sheet.autoFilter.clearColumnCriteria(3);
    await context.sync();
});
```

最终的自动筛选代码示例显示如何使用 `remove` 方法将 auto-filter 从工作表移除。

```js
// This method removes all AutoFilters from the active worksheet.
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.autoFilter.remove();
    await context.sync();
});
```

`AutoFilter` 也可应用到单个表。 有关详细信息，请参阅[使用 Excel JavaScript API 处理表](excel-add-ins-tables.md#autofilter)。

## <a name="data-protection"></a>数据保护

加载项可以控制用户能否编辑工作表中的数据。 工作表的 `protection` 属性是包含 `protect()` 方法的 [WorksheetProtection](/javascript/api/excel/excel.worksheetprotection) 对象。 下面的示例展示了关于切换活动工作表的完整保护的基本方案。

```js
await Excel.run(async (context) => {
    let activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("protection/protected");
    await context.sync();

    if (!activeSheet.protection.protected) {
        activeSheet.protection.protect();
    }
});
```

`protect` 方法包含两个可选参数：

- `options`：定义具体编辑限制的 [WorksheetProtectionOptions](/javascript/api/excel/excel.worksheetprotectionoptions) 对象。
- `password`：表示用户规避保护并编辑工作表所需使用的密码的字符串。

[保护工作表](https://support.microsoft.com/office/3179efdb-1285-4d49-a9c3-f4ca36276de6)一文详细介绍了工作表保护，以及如何通过 Excel UI 更改保护。

### <a name="detect-changes-to-the-worksheet-protection-state"></a>检测对工作表保护状态所做的更改

加载项可以通过加载项或加载项 UI 更改工作表Excel状态。 若要检测对保护状态所做的更改，请 [为](excel-add-ins-events.md#register-an-event-handler) 工作表的事件注册 [`onProtectionChanged`](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member) 事件处理程序。 事件的事件处理程序在 `onProtectionChanged` 事件触发时 [`WorksheetProtectionChangedEventArgs`](/javascript/api/excel/excel.worksheetprotectionchangedeventargs) 接收对象。

下面的代码示例演示如何注册`onProtectionChanged`事件处理程序并使用 `WorksheetProtectionChangedEventArgs` 对象`isProtected`检索事件的 `worksheetId``source` 、 和 属性。

```js
// This method registers an event handler for the onProtectionChanged event of a worksheet.
async function run() {
    await Excel.run(async (context) => {
        // Retrieve the worksheet named "Sample".
        let sheet = context.workbook.worksheets.getItem("Sample");
    
        // Register the onProtectionChanged event handler.
        sheet.onProtectionChanged.add(checkProtection);
        await context.sync();
    });
}

// This method is an event handler that returns the protection state of a worksheet 
// and information about the changed worksheet.
async function checkProtection(event) {
    await Excel.run(async (context) => {
        // Retrieve the protection, worksheet ID, and source properties of the event.
        let protectionStatus = event.isProtected;
        let worksheetId = event.worksheetId;
        let source = event.source;

        // Print the event properties to the console.
        console.log("Protection status changed. Protection status is now: " + protectionStatus);
        console.log("    ID of changed worksheet: " + worksheetId);
        console.log("    Source of change event: " + source);    
    });
}
```

## <a name="page-layout-and-print-settings"></a>页面布局和打印设置

加载项可以在工作表级别访问页面布局设置。 这些控制打印工作表的方式。 `Worksheet` 对象有三个与布局相关的属性：`horizontalPageBreaks`、`verticalPageBreaks`、`pageLayout`。

`Worksheet.horizontalPageBreaks` 和 `Worksheet.verticalPageBreaks` 是 [PageBreakCollections](/javascript/api/excel/excel.pagebreakcollection)。 这些是 [PageBreaks](/javascript/api/excel/excel.pagebreak) 的集合，其中指定插入手动分页符的范围。 以下代码示例在第 **21** 行上方添加了水平分页符。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.horizontalPageBreaks.add("A21:E21"); // The page break is added above this range.
    await context.sync();
});
```

`Worksheet.pageLayout` 是 [PageLayout](/javascript/api/excel/excel.pagelayout) 对象。 此对象包含不依赖于任何打印机特定实现的布局和打印设置。 这些设置包括页边距、方向、页码编号、标题行，并打印区域。

以下代码示例使页面居中（垂直和水平），设置将在每页顶部打印的标题行，并将打印区域设置为工作表的子部分。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();

    // Center the page in both directions.
    sheet.pageLayout.centerHorizontally = true;
    sheet.pageLayout.centerVertically = true;

    // Set the first row as the title row for every page.
    sheet.pageLayout.setPrintTitleRows("$1:$1");

    // Limit the area to be printed to the range "A1:D100".
    sheet.pageLayout.setPrintArea("A1:D100");

    await context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
