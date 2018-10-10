---
title: 在 Excel 加载项中同时处理多个范围
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: 2387be8dc17d85028b1d086cb192ac1accf167d5
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459194"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>在 Excel 加载项中同时处理多个范围（预览）

Excel JavaScript 库同时在多个范围启用加载项完成操作，并且设置属性 。范围不必是相邻的。除了使代码更加简单以外，这种设置属性的方式比单独为每个范围设置相同属性的方式运行速度要快很多。

> [!NOTE]
> 本文介绍的 API 要求 **Office 2016 Click-to-Run version 1809 Build 10820.20000** 或更高版本。（可能需要加入 [Office Insider 程序](https://products.office.com/office-insider) 来获取适当的内部版本。）此外，必须从 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 加载 beta 版本的 Office JavaScript 库。最后，我们没有这些 API 的参考页。但以下定义类型文件具有它们的说明： [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 。

## <a name="rangeareas"></a>RangeAreas

`Excel.RangeAreas` 对象表示一组范围（可能不相邻）。它具有类似于 `Range` 类型的属性和方法 （很多具有相同或类似的名称），但已经进行了调整：

- 属性的数据类型以及 getter 和 setter 的行为。
- 方法参数的数据类型和方法的行为。
- 返回值的数据类型。

例如：

- `RangeAreas` 具有返回范围地址的逗号分隔字符串的 `address` 属性，而不是像 `Range.address` 属性那样只是一个地址。
- `RangeAreas` 如果是一致的，具有返回 `DataValidation` 对象的 `dataValidation` 属性，该对象表示 `RangeAreas` 内所有范围的数据有效性。如果相同的 `DataValidation` 对象没有应用于 `RangeAreas` 内的所有范围则对象为 `null` 。这是 `RangeAreas` 对象的一个常规但非通用原则： *如果属性在 `RangeAreas` 内所有范围上没有一致的值，它便是 `null` 。* 有关详细信息和一些例外，请参阅 [读取 RangeAreas 的属性](#reading-properties-of-rangeareas) 。
- `RangeAreas.cellCount` 获取 `RangeAreas` 内所有范围的单元格总数。
- `RangeAreas.calculate` 重新计算 `RangeAreas` 内所有范围的单元格。
- `RangeAreas.getEntireColumn` `RangeAreas.getEntireRow` 返回另一个 `RangeAreas` 对象，该对象表示 `RangeAreas` 内所有范围的所有列 （或行）。例如，如果 `RangeAreas` 表示“ A1:C4 ”和“ F14:L15 ”，则 `RangeAreas.getEntireColumn` 返回表示“ A:C ”和“ F:L ”的 `RangeAreas` 对象。
- `RangeAreas.copyFrom` 可以采用 `Range` 或 `RangeAreas` 参数，表示复制操作的源范围。

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>在 RangeAreas 上也可用的“范围”成员的完整列表

##### <a name="properties"></a>属性

在编写读取任何所列属性的代码之前应熟知 [读取 RangeAreas 的属性](#reading-properties-of-rangeareas) 。返回内容存在细微差异。

- 地址
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

### <a name="rangearea-specific-properties-and-methods"></a>RangeArea 特定属性和方法

 `RangeAreas` 类型具有某些不在 `Range` 对象上的属性和方法。以下是它们的选择：

- `areas`: `RangeCollection` 对象，其中包含所有 `RangeAreas` 对象表示的范围。 `RangeCollection` 对象也是新对象并且类似于其他 Excel 集合对象。它具有 `items` 属性，是一个 `Range` 对象的数组，表示范围。
- `areaCount`： `RangeAreas` 内范围的总数。
- `getOffsetRangeAreas`：工作原理和 [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 一样，只不过返回一个 `RangeAreas`，它包含的范围与原 `RangeAreas` 中的某一范围产生偏移。

## <a name="create-rangeareas-and-set-properties"></a>创建 RangeAreas 并设置属性

可以通过两种基本方式创建 `RangeAreas` 对象：

- 调用 `Worksheet.getRanges()` ，将带逗号分隔范围地址的字符串传递给它。如果想要包含的任何范围已变成 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) ，可以在字符串中包含名称，而不是地址。
- 调用 `Workbook.getSelectedRanges()` 。此方法返回一个 `RangeAreas` ，表示在当前活动工作表上选择的所有范围。

拥有 `RangeAreas` 对象后，可以使用返回诸如 `getOffsetRangeAreas` 和 `getIntersection` 等的 `RangeAreas` 对象创建其他对象。

> [!NOTE]
> 不能向 `RangeAreas` 对象直接添加其他范围。例如， `RangeAreas.areas` 中的集合不具有 `add` 方法。


> [!WARNING] 
> 不要尝试直接添加或者删除 `RangeAreas.areas.items` 数组的成员。这样会引起代码出现异常。例如，可以将其他 `Range` 对象推送给数组，但是此举会引发错误，因为 `RangeAreas` 属性和方法并不表现出数组中存在新的项目。例如， `areaCount` 属性不包含以这种方式推送的范围，如果 `index` 大于 `areasCount-1` ， `RangeAreas.getItemAt(index)` 会引发一个错误。同样，通过获取引用或者调用其 `Range.delete`  方法来删除 `RangeAreas.areas.items` 数组中的 `Range` 对象会引发 bug ： 尽管 `Range` 对象 *被* 删除，但父级 `RangeAreas` 对象的属性和方法表现出，或者试图表现出，对象仍然存在。例如，如果代码调用 `RangeAreas.calculate` ， Office 便会对该范围进行计算，但会出错，因为该范围对象已经不存在。

在 `RangeAreas` 上设置属性会在 `RangeAreas.areas` 集合内的所有范围上设置相应的属性。

以下是在多个范围设置属性的示例。该功能突出显示 **F3:F5** 和 **H3:H5** 范围。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

本示例应用于可以硬编码传递给 `getRanges` 的范围地址或在运行时对其进行轻松计算的方案。一些可以真正将此包含在其中的方案： 

- 代码在已知模板的上下文中运行。
- 代码在导入数据的上下文中运行，其中数据架构为已知。

当不能在编码时确定需要执行操作的范围时，必须在运行时发现它们。下一节讨论这些方案。

### <a name="discover-range-areas-programmatically"></a>以编程方式发现范围

 `Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法能够根据单元格的特征和单元格值的类型，实现在运行时找到需要进行操作的范围。下面是 TypeScript 数据类型文件中的方法签名：

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

下面是使用第一个的示例。有关此代码，请注意：

- 它通过首先调用 `Worksheet.getUsedRange` 并针对仅该范围调用 `getSpecialCells` ，将搜索限制于工作表的所需部分。
- 它将作为参数传递给 `getSpecialCells` 来自 `Excel.SpecialCellType` 枚举值的字符串版本。某些可以传递的其他值可以是空单元格的“空白”、非公式文本值单元格的"常量"以及拥有与 `usedRange` 中第一单元格相同条件格式的单元格的“ SameConditionalFormat ”。第一个单元格是最左上角单元格。有关枚举中值的完整列表，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 。
-  `getSpecialCells` 方法返回 `RangeAreas` 对象，因此所有带公式的单元格都将以粉红色标示，即使它们并非全部连续。 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

有时候范围没有 *任何* 具有目标特征的单元格。如果 `getSpecialCells` 找不到任何这样的单元格，它将引发 **ItemNotFound** 错误。如果有，这会将控制流转移至 `catch` block/ 方法。如果没有，错误暂停函数。可以有在不存在具有目标特征的单元格时如愿引发错误的方案。 

但在一些它是正常的方案中，但可能不常见，因为有不匹配的单元格；代码应检查这种可能性并从容地对其进行处理，而不引发错误。对于这些方案，使用 `getSpecialCellsOrNullObject` 方法并测试 `RangeAreas.isNullObject` 属性。下面是一个示例。请注意此代码：

-  `getSpecialCellsOrNullObject` 方法总是返回代理对象，因此从一般 JavaScript 意义上说，它不会是 `null` 。但是，如果找不到匹配的单元格， 对象的 `isNullObject` 属性设置为 `true` 。
- 它将调用 `context.sync` ， *之前* 它测试 `isNullObject` 属性。这是所有 `*OrNullObject` 方法和属性的要求，因为始终要加载和同步属性以便读取它。但是，不需要 *明确地* 加载 `isNullObject` 属性。它是由 `context.sync` 自动加载的，即便在对象上不会对 `load` 进行调用。有关详细信息，请参阅 [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods) 。
- 可以通过首先选择无公式单元格的范围来测试此代码并运行它。然后选择至少具有一个带公式单元格的范围并再次运行它。

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

为了简单起见，此文中所有其他示例均使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject` 。

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>通过单元格的值类型缩小目标单元格范围

存在一个枚举类型 `Excel.SpecialCellValueType` 的可选第二参数，这进一步缩小了目标单元格范围。仅在将“公式”或者“常量”传递给 `getSpecialCells` 或者 `getSpecialCellsOrNullObject` 时才能使用它。此参数仅指定所要的具有某些值类型的单元格。有四种基本类型：“错误”、“逻辑”（即布尔值）、“数字”和“文本”。（枚举拥有这四种类型以外其他值，以下将对其进行讨论）下面是个示例。有关此代码，请注意：

- 它仅突出显示具有文本数字值的单元格。它不会突出显示具有公式（即便结果是数字）或布尔值、 文本或错误状态的单元格。
- 若要测试代码，请确保工作表拥有一些带文本数字值的单元格、一些带其他种类文本值的单元格，以及一些带公式的单元格。

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

有时候需要对一个以上的单元格值类型进行操作，例如所有文本值和所有布尔值（“逻辑”）单元格。 `Excel.SpecialCellValueType` 枚举具有允许合并类型的值。例如，“逻辑文本”将针对所有布尔值和所有文本值。可以合并四种基本类型中的任意两种或任意三种类型。这些合并基本类型的枚举值的名称总是按照字母排序。因此要合并错误值、文本值以及布尔值单元格，应使用“ ErrorLogicalText ”，而不是“ LogicalErrorText ”或“ TextErrorLogical ”。"所有"默认参数合并全部四种类型。下面的示例突出显示所有生成数字或者布尔值的带公式单元格。

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
> `Excel.SpecialCellValueType` 参数仅在 `Excel.SpecialCellType` 参数为“公式”或“常量”时才可以使用。

### <a name="get-rangeareas-within-rangeareas"></a>获取 RangeAreas 中的 RangeAreas

 `RangeAreas` 类型本身也有 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法，它们采用相同的两个参数。这些方法从 `RangeAreas.areas` 集合内的所有范围返回全部目标单元格。当在 `RangeAreas` 对象而不是 `Range` 对象上调用时，这些方法的行为存在一个小差异： 当“ SameConditionalFormat ”作为第一参数传递时，此方法会返回具有与 *`RangeAreas.areas` 集合内第一个范围中* 最左上角单元格相同条件格式的所有单元格。同一个点适用于“ SameDataValidation ”：在传递给 `Range.getSpecialCells` 时，它将返回具有与 *此范围内* 最左上角单元格相同数据验证规则的全部单元格。但当传递给 `RangeAreas.getSpecialCells` 时，它将返回具与 * `RangeAreas.areas` 集合内第一个范围中* 最左上角单元格相同数据验证规则的全部单元格。

## <a name="read-properties-of-rangeareas"></a>读取 RangeAreas 的属性

读取 `RangeAreas` 属性值需要小心，因为对于 `RangeAreas` 内的不同范围，给定的属性可能具有不同的值。一般规则是，如果 *可以* 返回一个一致的值则会将其返回。例如，在下面的代码中，粉红色 (`#FFC0CB`) 和 `true` 的 RGB 代码将被记录到控制台，因为 `RangeAreas` 对象中的两个范围均被充填粉红色而且两者均是整列。

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

当一致性为不可能时，情况会变得更加复杂。 `RangeAreas` 属性的行为遵循以下三个原则：

-  `RangeAreas` 对象的布尔值属性返回 `false` ，除非该属性对于所有成员范围均为真。
- 除 `address` 属性外的非布尔值属性将返回 `null`，除非所有成员范围的相应属性具有相同的值。
-  `address` 属性返回成员范围地址的逗号分隔字符串。

例如，下面的代码创建 `RangeAreas` ，其中只有一个范围是整个列，只有一个填充粉红色。控制台将显示填充颜色的 `null` 、 `isEntireRow` 属性的 `false` 和 `address`  属性的“ Sheet1 ！F3:F5，Sheet1 ！H:H ”（假定工作表名称为“ Sheet1 ”）。 

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

- [使用 Excel JavaScript API 的基本编程概念](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [Range 对象（ JavaScript API for Excel ）](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [RangeAreas 对象（ JavaScript API for Excel ）](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) （预览 API 时该链接可能无法使用。作为替代方法，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 。）