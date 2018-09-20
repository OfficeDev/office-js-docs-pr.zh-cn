---
title: 在 Excel 加载项中同时处理多个范围
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016456"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>在 Excel 加载项同时处理多个范围（预览）

Excel JavaScript 库使加载项能够执行操作，并同时在多个范围设置属性。 这些范围不需要连续。 这样的设置属性方式除了能简化您的代码，其运行速度也快于单独地为每个范围设置相同属性。

> [!NOTE]
> 本文中介绍的 API 要求 **Office 2016 即点即用版本 1809 内部版本 10820.20000** 或更高版本。 （您可能需要加入 [Office 预览体验计划](https://products.office.com/office-insider) 以获取适当的版本。）此外，您必须从 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)加载 beta 版本的 Office JavaScript 库。 最后，我们尚未 为这些 API 提供参考页。 但下面的定义类型文件提供 了相关说明：[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。

## <a name="rangeareas"></a>RangeAreas

一组（可能不连续的）范围以 `Excel.RangeAreas` 对象表示。 其属性和方法类似于 `Range` 类型 （许多拥有相同或类似的名称），但也存在以下调整：

- 属性的数据类型以及 getter 和 setter的行为。
- 方法参数的数据类型和方法的行为。
- 返回值的数据类型。

例如：

- `RangeAreas` 具有 `address` 属性，返回使用逗号分隔的范围地址字符串，而不是像 `Range.address` 属性那样返回一个地址。
- `RangeAreas` 具有 `dataValidation` 属性，返回的 `DataValidation` 对象代表 `RangeAreas` 中所有范围的数据验证（如果一致）。 如果相同的 `DataValidation` 对象不应用于 `RangeAreas` 中的所有范围，则属性为 `null`。 对象有一项常规、但不是通用的原则：*如果属性在 `RangeAreas` 的所有区域上具有一致的值，则它是 `null`。*`RangeAreas` 要了解有关详细信息和一些例外，请参阅[读取 RangeAreas 的属性](#reading-properties-of-rangeareas)。
- `RangeAreas.cellCount` 获取 `RangeAreas` 中所有范围的单元格总数目。
- `RangeAreas.calculate` 重新计算 `RangeAreas` 中所有范围的单元格。
- `RangeAreas.getEntireColumn` 和 `RangeAreas.getEntireRow` 返回另一个 `RangeAreas` 对象，表示 `RangeAreas` 中所有范围的所有列 （或行）。 例如，如果 `RangeAreas` 表示"A:C" 和 "F:L"，则 `RangeAreas.getEntireColumn` 返回表示 "A1:C4" 和 "F14:L15" 的 `RangeAreas` 对象。
- `RangeAreas.copyFrom` 可以是 `Range` 或 `RangeAreas` 参数，表示复制操作的源范围。

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>在 RangeAreas 上也可用的“范围”成员的完整列表

##### <a name="properties"></a>属性

请先熟悉[读取 RangeAreas 的属性](#reading-properties-of-rangeareas)，然后再编写可读取列出的所有属性的代码。 返回什么有不少微妙之处。

- address
- addressLocal
- cellCount
- conditionalFormats
- context
- dataValidation
- 格式
- isEntireColumn
- isEntireRow
- 样式
- 工作表

##### <a name="methods"></a>方法

预览版中的范围方法是做了标记的。

- calculate()
- clear()
- convertDataTypeToText()（预览）
- convertToLinkedDataType()（预览）
- copyFrom()（预览）
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange()（名为 RangeAreas 对象上的 getOffsetRangeAreas）
- getSpecialCells()（预览）
- getSpecialCellsOrNullObject()（预览）
- getTables()（预览）
- getUsedRange()（名为 RangeAreas 对象上的 getUsedRangeAreas）
- getUsedRangeOrNullObject() （名为 RangeAreas 对象上的 getUsedRangeAreasOrNullObject）
- load()
- set()
- setDirty()（预览）
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>特定于 RangeArea 的属性和方法

`RangeAreas` 类型具有某些不在 `Range` 对象上的属性和方法。 以下是其中的选定内容：

- `areas`：`RangeCollection` 对象，其中包含 `RangeAreas` 对象所表示的所有范围。 对象也是新增，并类似于其他 Excel 集合对象。`RangeCollection` 它具有 `items` 属性，是表示范围的 `Range` 对象数组。
- `areaCount`：`RangeAreas` 中范围的总数目。
- `getOffsetRangeAreas`：工作原理和 [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 一样，只不过返回一个 `RangeAreas`，它包含的范围与原 `RangeAreas` 中的某一范围产生偏移。

## <a name="create-rangeareas-and-set-properties"></a>创建 RangeAreas 并设置属性

您可以通过两种基本方式创建 `RangeAreas` 对象：

- 调用 `Worksheet.getRanges()` 并向它传递以逗号分隔的范围地址的字符串。 如果您想要包括的任何区域已变成 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem)，可以在字符串中包括该名称，而不是该地址。
- 调用 `Workbook.getSelectedRanges()`。 此方法返回一个 `RangeAreas`，它表示当前活动工作表的所有选定范围。

拥有 `RangeAreas` 对象后，您可以使用返回如 `getOffsetRangeAreas` 和 `getIntersection` 等 `RangeAreas` 的对象创建其他对象。

> [!NOTE]
> 不能直接添加其他范围到 `RangeAreas` 对象。 例如，`RangeAreas.areas` 中的集合中没有 `add` 方法。


> [!WARNING] 
> 切勿尝试直接添加或删除 `RangeAreas.areas.items` 数组的成员。 这将导致代码中的不正常行为。 例如，可以将额外的 `Range` 对象推送到数组，但这样做将导致错误，因为 `RangeAreas` 属性和方法的行为不会因新项目的添加而有所不同。 例如，`areaCount` 属性不包括以这种方式推送的范围，并且如果 `index` 大于 `areasCount-1`，`RangeAreas.getItemAt(index)` 将抛出错误。 同样，通过获取引用并调用其 方法删除  数组中的  对象会导致错误：虽然  对象已删除，父  对象仍将视其为存在。`Range``RangeAreas.areas.items``Range.delete``Range`**`RangeAreas` 例如，如果您的代码调用 `RangeAreas.calculate`，Office 会尝试计算该范围，但将发生错误，因为 range 对象不存在。

在 `RangeAreas` 上设置属性会在 `RangeAreas.areas` 集合的所有范围上设置相应的属性。

以下是在多个范围上设置属性的示例。 该函数高亮显示 **F3:F5** 至 **H3:H5** 之间的范围。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

本示例适用于可以硬编码或运行时轻松计算传递给 `getRanges` 的范围地址的场景。 一些符合条件的方案包括： 

- 代码在已知模板的上下文中运行。
- 代码在导入数据的上下文中运行，其中数据架构为已知。

当你不能在编码时确定需要执行操作的范围时，必须在运行时确定。 下一节讨论了这些方案。

### <a name="discover-range-areas-programmatically"></a>以编程方式发现范围

和 `Range.getSpecialCellsOrNullObject()` 方法使你能够在运行时根据单元格和单元格值类型的特征而确定对其执行操作的范围。`Range.getSpecialCells()` 下面是 TypeScript 数据类型文件中的方法签名：

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

下面是使用第一个方法的示例。 关于此代码，请注意以下几点：

- 它通过首先调用 `Worksheet.getUsedRange` 并针对仅该范围调用 `getSpecialCells`，将搜索限制于工作表的所需部分。
- 它将 `Excel.SpecialCellType` 枚举中值的字符串版本作为参数传递到 `getSpecialCells`。 某些无法传递的其他值 包括： 空白单元格 为 "Blanks" ，具有文字值而不是公式的单元格为 "Constants"，与 `usedRange` 中第一个单元格具有相同条件格式的单元格 为 "SameConditionalFormat"。 第一个单元格为左上角的单元格。 要了解枚举中值的完整列表，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。
- `getSpecialCells` 方法返回 `RangeAreas` 对象，所以所有带公式的单元格都将以粉红色标示，即使它们并非全都相邻。 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

有时，范围不具有*任何*带目标特征的单元格。 如果 `getSpecialCells` 找不到任何单元格，将引发 **ItemNotFound** 错误。 如果存在 `catch` 块/方法，控制流将重定向到该块/方法。 如果不存在，错误将暂停函数的执行。 在某些时候，如果不存在具有目标特征的单元格，你可能希望抛出错误。 

但在其他时候，不存在匹配单元格的情况属于正常但可能不常见；你的代码应检查这种可能性并恰当处理而不抛出错误。 对于这些方案，请使用 `getSpecialCellsOrNullObject` 方法并测试 `RangeAreas.isNullObject` 属性。 示例如下。 关于此代码，请注意以下几点：

- 方法总是返回代理对象，因此永远不会成为一般 JavaScript 意义上的  `null`。`getSpecialCellsOrNullObject` 但是，如果找到匹配的单元格，对象的 `isNullObject` 属性将设置为 `true`。
- 它将在测试 `isNullObject` 属性*之前*调用 `context.sync`。 这是所有 `*OrNullObject` 方法和属性的要求，因为你始终需要加载和同步属性才能读取它。 但是，不需要*显式*加载 `isNullObject` 属性。 它会由 `context.sync` 自动加载，即使对象没有调用 `load`。 有关详细信息，请参阅 [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)。
- 你可以通过首先选择没有公式的单元格范围然后运行它来测试此代码。 然后选择其中至少具有一个带公式的单元格的范围并再次运行。

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

为了简单起见，在此文章中所有其他示例均使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject`。

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>通过单元格的值类型缩小目标单元格范围

可选的第二个参数为枚举类型的 `Excel.SpecialCellValueType`，它进一步缩小目标单元格的范围。 仅当你向 `getSpecialCells` 或 `getSpecialCellsOrNullObject`传递 "Formulas" 或 "Constants" 时可以使用。 该参数指定单元格必须拥有特定类型的值。 有四个基本类型："Error"、"Logical"（ 即布尔值）、"Numbers" 和 "Text"。 （除了这四个值，该枚举还有其他值，下文将会讨论。）下面是一个示例。 关于此代码，请注意以下几点：

- 它仅将突出显示包含文本数字值的单元格。 它不会突出显示包含公式 （即使结果是数字）的单元格或布尔值、文本或错误状态的单元格。
- 若要测试代码，请确保工作表的一些单元格拥有文本数字值、一些单元格带有其他类型的文本值，另有一些含有公式。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

有时，你需要对多个单元格值类型执行操作—— 例如所有文本值和所有布尔值 ("Logical")。 枚举具有让你组合多个类型的值。`Excel.SpecialCellValueType` 例如，"LogicalText" 将匹配所有布尔值和所有文本值单元格。 可以将这四种基本类型中的任何两种或任何三种类型组合使用。 这些枚举值的基本类型组合名称通常以字母顺序列出。 因此，若要合并错误值、文本值和布尔值的单元格，请使用 "ErrorLogicalText"，而不是 "LogicalErrorText" 或 "TextErrorLogical"。 默认参数 "All" 合并所有四种类型。 下面的示例突出显示带有生成数字或布尔值的公式的所有单元格：

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
> `Excel.SpecialCellValueType` 参数仅在 `Excel.SpecialCellType` 参数为 "Formulas" 或 "Constants" 时可以使用。

### <a name="get-rangeareas-within-rangeareas"></a>获取 RangeAreas 中的 RangeAreas

类型本身也有 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法，它们采用相同的两个参数。`RangeAreas` 这些方法返回 `RangeAreas.areas` 集合所有范围的所有目标单元格。 方法调用于  对象而不是  对象的行为中存在一个细小区别：将 "SameConditionalFormat" 作为第一个参数传递时，方法会返回  集合第一个范围中与左上角单元格具有同一条件格式的所有单元格。`RangeAreas``Range`*`RangeAreas.areas`* 这一点同样适用于 "SameDataValidation"：传递给 `Range.getSpecialCells` 时，它将返回与* 范围中*左上角单元格具有相同数据验证规则的的单元格。 但当传递给 `RangeAreas.getSpecialCells` 时，将返回与 *`RangeAreas.areas` 集合第一个范围中*的左上角单元格具有相同数据验证规则的所有单元格。

## <a name="read-properties-of-rangeareas"></a>读取 RangeAreas 的属性

读取 `RangeAreas` 的属性值时需要小心，因为给定的属性可能在 `RangeAreas` 的不同范围内具有不同的值。 一般规则是，如果*可以*返回一致的值，它就会返回。 例如，在下面的代码中，粉红色的 RGB 代码 (`#FFC0CB`) 和 `true` 将记录到控制台中，因为 `RangeAreas` 对象中的这两个范围中具有粉红色填充且两者均为整个列。

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

当不可能达成一致性时，事情会变得更为复杂。 属性的行为遵循以下三个原则：`RangeAreas`

- 对象的布尔值属性返回 `false`，除非该属性对于所有成员范围均为 true。`RangeAreas`
- 除 `address` 属性外的非布尔值属性将返回 `null`，除非所有成员范围的相应属性具有相同的值。
- 属性返回成员范围地址的以逗号分隔 字符串。`address`

例如，下面的代码生成一个 `RangeAreas`，其中只有一个范围是整列，也只有一个 范围填充粉红色。 控制台将显示填充颜色为 `null`，`isEntireRow` 属性为 `false` 且 `address` 属性为 "Sheet1!F3:F5, Sheet1!H:H"（假定工作表名称为 "Sheet1"）。 

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

- [Excel JavaScript API 核心概念](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Range 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [RangeAreas 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas)（该链接可能无法在预览版 API 中生效）。 作为替代方法，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。）