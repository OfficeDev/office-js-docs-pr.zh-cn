---
title: 同时在 Excel 加载项中处理多个区域
description: ''
ms.date: 12/26/2018
ms.openlocfilehash: ab7cd9757adaedf2b6cc43fdcc604b98a60b6ecd
ms.sourcegitcommit: 8d248cd890dae1e9e8ef1bd47e09db4c1cf69593
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/27/2018
ms.locfileid: "27447230"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>同时在 Excel 加载项中处理多个区域（预览）

Excel JavaScript 库允许你使用加载项同时在多个区域上执行操作和设置属性。 这些区域不必是连续区域。 除了简化代码以外，这种设置属性的方法还比为每个区域单独设置相同的属性要快得多。

> [!NOTE]
> 本文中所述的 API 需要 **Office 2016 即点即用版本 1809 内部版本 10820.20000** 或更高版本。 （您可能需要加入 [Office 预览体验成员计划](https://products.office.com/office-insider)才能获取相应的内部版本。）此外，您还必须从 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 加载 Office JavaScript 库的 Beta 版本。 最后，我们还没有提供这些 API 的参考页面。 但是，以下定义类型文件提供了相关说明：[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。

## <a name="rangeareas"></a>RangeAreas

一组区域（可能是非连续区域）由 `Excel.RangeAreas` 对象表示。 它具有与 `Range` 类型类似的属性和方法（许多具有相同或相似的名称），但已对以下对象进行了调整：

- 属性和 Setter 及 Getter 行为的数据类型。
- 方法参数和方法行为的数据类型。
- 方法返回值的数据类型。

例如：

- `RangeAreas` 具有 `address` 属性，它将返回一串以逗号分隔的区域地址，而不是像 `Range.address` 属性一样只返回一个地址。
- `RangeAreas` 具有 `dataValidation` 属性，它将返回一个 `DataValidation` 对象，用来表示 `RangeAreas` 中的所有区域的数据验证（如果保持一致）。 如果相同的 `DataValidation` 对象未应用到 `RangeAreas` 中的所有区域，则该属性将为 `null`。 对于 `RangeAreas` 对象，这是一般原则，而非通用原则：*如果某个属性在 `RangeAreas` 的所有区域上没有一致的值，则该属性将为 `null`。* 请参阅[读取 RangeAreas 的属性](#read-properties-of-rangeareas)，以了解详细信息和某些例外情况。
- `RangeAreas.cellCount` 将获取 `RangeAreas` 中的所有区域的单元格总数。
- `RangeAreas.calculate` 将重新计算 `RangeAreas` 中的所有区域的单元格数。
- `RangeAreas.getEntireColumn` 和 `RangeAreas.getEntireRow` 将返回另一个 `RangeAreas` 对象，用来表示 `RangeAreas` 中的所有区域的列数（或行数）。 例如，如果 `RangeAreas` 表示“A1:C4”和“F14:L15”，则 `RangeAreas.getEntireColumn` 将返回一个表示“A:C”和“F:L”的 `RangeAreas` 对象。
- `RangeAreas.copyFrom` 可以采用 `Range` 或 `RangeAreas` 参数，用来表示复制操作的源区域。

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>RangeAreas 还提供了区域成员的完整列表

##### <a name="properties"></a>属性

在编写用于读取任何所列属性的代码之前，请先熟悉[读取 RangeAreas 的属性](#read-properties-of-rangeareas)。 返回的内容存在细微差别。

- address
- addressLocal
- cellCount
- conditionalFormats
- context
- dataValidation
- format
- isEntireColumn
- isEntireRow
- style
- worksheet

##### <a name="methods"></a>Methods

将标记预览中的区域方法。

- calculate()
- clear()
- convertDataTypeToText()（预览）
- convertToLinkedDataType()（预览）
- copyFrom()（预览）
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange()（在 RangeAreas 对象上名为 getOffsetRangeAreas）
- getSpecialCells()（预览）
- getSpecialCellsOrNullObject()（预览）
- getTables()（预览）
- getUsedRange()（在 RangeAreas 对象上名为 getUsedRangeAreas）
- getUsedRangeOrNullObject()（在 RangeAreas 对象上名为 getUsedRangeAreasOrNullObject）
- load()
- set()
- setDirty()（预览）
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>特定于 RangeArea 的属性和方法

`RangeAreas` 类型具有一些未包含在 `Range` 对象中的属性和方法。 以下是其中的一部分：

- `areas`：一种 `RangeCollection` 对象，它包含由 `RangeAreas` 对象表示的所有区域。 `RangeCollection` 也是新对象，与其他 Excel 集合对象类似。 它具有 `items` 属性，它是一组表示区域的 `Range` 对象。
- `areaCount`：`RangeAreas` 中的区域总数。
- `getOffsetRangeAreas`：与 [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 的作用类似，不同之处在于，前者将返回 `RangeAreas` 并且包含多个区域，每个区域都是原始 `RangeAreas` 中的区域的偏移。

## <a name="create-rangeareas"></a>创建 RangeAreas

可以通过两种基本方法创建 `RangeAreas` 对象：

- 调用 `Worksheet.getRanges()` 并向其传递具有以逗号分隔的区域地址的字符串。 如果要包含的任何区域已插入到 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) 中，则可以在字符串中包含名称而不是地址。
- 调用 `Workbook.getSelectedRanges()`。 此方法将返回 `RangeAreas`，它表示在当前活动工作表上选择的所有区域。

获得 `RangeAreas` 对象后，你可以在返回 `RangeAreas` 的对象上使用该方法创建其他对象，例如 `getOffsetRangeAreas` 和 `getIntersection`。

> [!NOTE]
> 你不能直接将其他区域添加到 `RangeAreas` 对象。 例如，`RangeAreas.areas` 中的集合不具有 `add` 方法。

> [!WARNING]
> 不要尝试直接添加或删除 `RangeAreas.areas.items` 数组的成员。 这将导致代码中出现不需要的行为。 例如，可能会将其他 `Range` 对象推送到数组上，但这样做会导致错误，因为 `RangeAreas` 属性和方法将表现为如同新项目并不存在一样。 例如，`areaCount` 属性不包含通过这种方法推送的区域，并且如果 `index` 大于 `areasCount-1`，则 `RangeAreas.getItemAt(index)` 将引发错误。 同样，删除 `RangeAreas.areas.items` 数组中的 `Range` 对象（通过获取对它的引用并调用其 `Range.delete` 方法）也会导致错误：尽管 `Range` 对象*已被*删除，但父 `RangeAreas` 对象的属性和方法将表现为或尝试表现为如同它仍然存在一样。 例如，如果你的代码调用 `RangeAreas.calculate`，Office 将尝试计算区域，但这会引发错误，因为区域对象并不存在。

## <a name="set-properties-on-multiple-ranges"></a>在多个区域设置属性

在 `RangeAreas` 对象上设置属性会在 `RangeAreas.areas` 集合中的所有区域上设置相应的属性。

以下是在多个区域上设置属性的示例。 函数将突出显示区域 **F3:F5** 和 **H3:H5**。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

此示例适用于可以对传递给 `getRanges` 的区域地址进行硬编码或在运行时轻松进行计算的应用场景。 一些适用的应用场景包括：

- 代码在已知模板的上下文中运行。
- 代码在导入数据的上下文中运行，其中数据架构是已知的。

## <a name="get-special-cells-from-multiple-ranges"></a>从多个区域获取特殊单元格

`RangeAreas` 对象上的 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法原理上类似于 `Range` 对象上同名方法。 这些方法从 `RangeAreas.areas` 集合中所有区域返回包含指定特征的单元格。 请参阅 [查找区域内特殊单元格](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range-preview) 部分了解特殊单元格更多详细信息。

调用 `RangeAreas` 对象上的 `getSpecialCells` 或 `getSpecialCellsOrNullObject` 方法时：

- 如果传递 `Excel.SpecialCellType.sameConditionalFormat` 作为第一个参数，该方法返回具有相同条件格式的所有单元格作为 `RangeAreas.areas` 集合第一个区域左上角单元格。
- 如果传递 `Excel.SpecialCellType.sameDataValidation` 作为第一个参数，该方法返回具有相同数据验证规则的所有单元格作为 `RangeAreas.areas` 集合第一个区域左上角单元格。

## <a name="read-properties-of-rangeareas"></a>读取 RangeAreas 的属性

读取 `RangeAreas` 的属性值时须小心操作，因为对于 `RangeAreas` 内的不同区域，给定的属性可能具有不同的值。 一般规则是，如果*可以*返回一致的值，则系统会返回该值。 例如，在以下代码中，RGB 粉色代码 (`#FFC0CB`) 和 `true` 将记录到控制台，因为 `RangeAreas` 对象中的两个区域都具有粉色填充，并且都是整列。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

如果无法实现一致性，则情况将变得更加复杂。 `RangeAreas` 属性的行为遵循以下三个原则：

- 除非所有成员区域的属性均为 true，否则 `RangeAreas` 对象的布尔属性将返回 `false`。
- 除非所有成员区域上的对应属性都具有相同的值，否则非布尔属性（`address` 属性除外）将返回 `null`。
- `address` 属性将返回一串以逗号分隔的成员区域地址。

例如，以下代码将创建 `RangeAreas`，其中只有一个区域是整列，并且只有一个区域具有粉色填充。 控制台将为填充颜色显示 `null`，为 `isEntireRow` 属性显示 `false`，并为 `address` 属性显示“Sheet1!F3:F5, Sheet1!H:H”（假设工作表名称为“Sheet1”）。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 基本编程概念](../reference/overview/excel-add-ins-reference-overview.md)
- [使用 Excel JavaScript API 对区域执行操作（基本）](excel-add-ins-ranges.md)
- [使用 Excel JavaScript API 对区域执行操作（高级）](excel-add-ins-ranges-advanced.md)