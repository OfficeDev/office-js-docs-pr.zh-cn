---
title: 同时在 Excel 加载项中处理多个区域
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: 37f9c8a9f3127d78e1cc794aea9e6d1502cdeaf9
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270976"
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

## <a name="create-rangeareas-and-set-properties"></a>创建 RangeAreas 并设置属性

可以通过两种基本方法创建 `RangeAreas` 对象：

- 调用 `Worksheet.getRanges()` 并向其传递具有以逗号分隔的区域地址的字符串。 如果要包含的任何区域已插入到 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) 中，则可以在字符串中包含名称而不是地址。
- 调用 `Workbook.getSelectedRanges()`。 此方法将返回 `RangeAreas`，它表示在当前活动工作表上选择的所有区域。

获得 `RangeAreas` 对象后，你可以在返回 `RangeAreas` 的对象上使用该方法创建其他对象，例如 `getOffsetRangeAreas` 和 `getIntersection`。

> [!NOTE]
> 你不能直接将其他区域添加到 `RangeAreas` 对象。 例如，`RangeAreas.areas` 中的集合不具有 `add` 方法。


> [!WARNING] 
> 不要尝试直接添加或删除 `RangeAreas.areas.items` 数组的成员。 这将导致代码中出现不需要的行为。 例如，可能会将其他 `Range` 对象推送到数组上，但这样做会导致错误，因为 `RangeAreas` 属性和方法将表现为如同新项目并不存在一样。 例如，`areaCount` 属性不包含通过这种方法推送的区域，并且如果 `index` 大于 `areasCount-1`，则 `RangeAreas.getItemAt(index)` 将引发错误。 同样，删除 `RangeAreas.areas.items` 数组中的 `Range` 对象（通过获取对它的引用并调用其 `Range.delete` 方法）也会导致错误：尽管 `Range` 对象*已被*删除，但父 `RangeAreas` 对象的属性和方法将表现为或尝试表现为如同它仍然存在一样。 例如，如果你的代码调用 `RangeAreas.calculate`，Office 将尝试计算区域，但这会引发错误，因为区域对象并不存在。

在 `RangeAreas` 上设置属性会在 `RangeAreas.areas` 集合中的所有区域上设置相应的属性。

以下是在多个区域上设置属性的示例。 函数将突出显示区域 **F3:F5** 和 **H3:H5**。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

此示例适用于可以对传递给 `getRanges` 的区域地址进行硬编码或在运行时轻松进行计算的应用场景。 一些适用的应用场景包括： 

- 代码在已知模板的上下文中运行。
- 代码在导入数据的上下文中运行，其中数据架构是已知的。

如果你在编码时不知道要对哪个区域执行操作，则必须在运行时发现它们。 下一节将讨论这些应用场景。

### <a name="discover-range-areas-programmatically"></a>以编程方式发现区域

`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法使你能够在运行时根据单元格特征和单元格的值类型查找要对其执行操作的区域。 以下是 TypeScript 数据类型文件中的方法签名：

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

以下是使用第一种方法的示例。 关于此代码，请注意以下几点：

- 它通过先调用 `Worksheet.getUsedRange` 并仅调用该区域的 `getSpecialCells` 来限制需要搜索的工作表部分。
- 它以参数形式传递给 `getSpecialCells`，即 `Excel.SpecialCellType` 枚举中的值的字符串版本。 某些可以传递的其他值包括：适用于空单元格的“空白”，适用于包含文本值而不是公式的单元格的“常量”，以及适用于与 `usedRange` 中的第一个单元格具有相同条件格式的单元格的“SameConditionalFormat”。 第一个单元格是指最左上角的单元格。 有关枚举中的值的完整列表，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。
- `getSpecialCells` 方法将返回 `RangeAreas` 对象，因此包含公式的单元格都会变成粉色，即使它们并非都是连续的单元格。 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

有时，区域未包含*任何*具有目标特征的单元格。 如果 `getSpecialCells` 未找到任何单元格，它将引发 **ItemNotFound** 错误。 这会将控制流转移到 `catch` 信息块/方法（如果存在）。 如果不存在，则该错误会终止函数。 可能存在这样的应用场景，即你正好希望当不存在具有目标特征的单元格时引发错误。 

但在没有匹配单元格（这是正常现象，但可能并不常见）的应用场景中，你的代码应该会检查这种可能的情况并按正常方式处理它，而不会引发错误。 对于这些应用场景，请使用 `getSpecialCellsOrNullObject` 方法并测试 `RangeAreas.isNullObject` 属性。 示例如下。 关于此代码，请注意以下几点：

- `getSpecialCellsOrNullObject` 方法将始终返回代理对象，因此在一般的 JavaScript 认知中，它从不为 `null`。 但是，如果没有找到匹配的单元格，则对象的 `isNullObject` 属性将设置为 `true`。
- 在测试 `isNullObject` 属性*之前*，它将调用 `context.sync`。 这是所有 `*OrNullObject` 方法和属性的要求，因为你必须始终加载和同步属性才能读取它。 但是，不必*明确*加载 `isNullObject` 属性。 即使未在对象上调用 `load`，`context.sync` 也会自动加载该属性。 有关详细信息，请参阅 [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)。
- 你可以测试此代码，方法是先选择没有公式单元格的区域并运行它。 然后选择至少包含一个带公式的单元格的区域，并再次运行它。

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

为简单起见，本文中的所有其他示例都使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject`。

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>通过单元格值类型缩小目标单元格的范围

有一个枚举类型为 `Excel.SpecialCellValueType` 的第二个可选参数，它可以进一步缩小要定位的单元格范围。 仅当将“公式”或“常量”传递给 `getSpecialCells` 或 `getSpecialCellsOrNullObject` 时，才能使用它。 该参数指定你只需要具有特定值类型的单元格。 有四种基本类型：“错误”、“逻辑”（它表示布尔）、“数字”和“文本”。 （除了这四种类型以外，枚举还具有其他值，将在下文对此展开讨论。）以下是一个示例。 关于此代码，请注意以下几点：

- 它只会突出显示具有文本数值的单元格。 它既不会突出显示具有公式的单元格（即使结果是数字），也不会突出显示布尔、文本或错误状态单元格。
- 要测试代码，请确保工作表中的某些单元格包含文本数值，某些包含其他类型的文本值，而某些则包含公式。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

有时，你需要对多种单元格值类型执行操作，例如所有文本值和所有布尔值（“逻辑”）单元格。 `Excel.SpecialCellValueType` 枚举提供的值允许你组合不同的类型。 例如，“LogicalText”将定向所有布尔值和所有文本值单元格。 你可以组合四种基本类型中的任意两种或任意三种。 这些用于组合基本类型的枚举值的名称始终按字母顺序排列。 因此，要组合错误值、文本值和布尔值单元格，请使用“ErrorLogicalText”，而不是“LogicalErrorText”或“TextErrorLogical”。 默认参数“全部”将组合所有四种类型。 以下示例突出显示包含用于生成数字或布尔值的公式的所有单元格：

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> 仅当 `Excel.SpecialCellType` 参数为“公式”或“常量”时，才能使用 `Excel.SpecialCellValueType` 参数。

### <a name="get-rangeareas-within-rangeareas"></a>获取 RangeAreas 内的 RangeAreas

`RangeAreas` 类型本身也具有 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法，它们采用两个相同的参数。 这些方法将返回 `RangeAreas.areas` 集合中所有区域的所有目标单元格。 在 `RangeAreas` 对象而不是 `Range` 对象上调用时，这些方法的行为存在一个小差异：如果将“SameConditionalFormat”作为第一个参数进行传递，则该方法将返回与 `RangeAreas.areas` 集合中第一个区域*最左上角的单元格具有相同条件格式的所有单元格*。 这一点也适用于“SameDataValidation”：如果传递给 `Range.getSpecialCells`，它将返回与区域最左上角的单元格*具有相同数据验证规则的所有单元格*。 但是，如果传递给 `RangeAreas.getSpecialCells`，它将返回与 `RangeAreas.areas` 集合中第一个区域*最左上角的单元格具有相同数据验证规则的所有单元格*。

## <a name="read-properties-of-rangeareas"></a>读取 RangeAreas 的属性

读取 `RangeAreas` 的属性值时须小心操作，因为对于 `RangeAreas` 内的不同区域，给定的属性可能具有不同的值。 一般规则是，如果*可以*返回一致的值，则系统会返回该值。 例如，在以下代码中，RGB 粉色代码 (`#FFC0CB`) 和 `true` 将记录到控制台，因为 `RangeAreas` 对象中的两个区域都具有粉色填充，并且都是整列。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
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
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
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

- [Excel JavaScript API 基本编程概念](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Range 对象（适用于 Excel 的 JavaScript API）](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [Range 对象（适用于 Excel 的 JavaScript API）](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas)（当 API 处于预览状态时，此链接可能无效。 作为替代方法，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。)