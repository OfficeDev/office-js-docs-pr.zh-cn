---
title: 使用 JavaScript API 查找Excel单元格
description: 了解如何使用 JavaScript API Excel查找特殊单元格，例如包含公式、错误或数字的单元格。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 2705f71b5b253801493b07143b8320a1f69062e9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744796"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a>使用 JavaScript API 查找Excel单元格

本文提供的代码示例使用 JavaScript API 查找Excel单元格。 有关对象支持的属性和方法`Range`的完整列表，请参阅Excel[。Range 类](/javascript/api/excel/excel.range)。

## <a name="find-ranges-with-special-cells"></a>查找包含特殊单元格的范围

[Range.getSpecialCells](/javascript/api/excel/excel.range#excel-excel-range-getspecialcells-member(1)) 和 [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#excel-excel-range-getspecialcellsornullobject-member(1)) 方法根据单元格的特征及其单元格的值类型查找区域。 这两种方法都返回 `RangeAreas` 对象。 以下是 TypeScript 数据类型文件中方法的签名：

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

下面的代码示例使用 `getSpecialCells` 方法查找包含公式的所有单元格。 关于此代码，请注意以下几点：

- 它通过先调用 `Worksheet.getUsedRange` 并仅调用该区域的 `getSpecialCells` 来限制需要搜索的工作表部分。
- `getSpecialCells` 方法将返回 `RangeAreas` 对象，因此包含公式的单元格都会变成粉色，即使它们并非都是连续的单元格。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    await context.sync();
});
```

如果区域中不存在具有目标特征的单元格，`getSpecialCells` 会引发 **ItemNotFound** 错误。 这会将控制流转移到 `catch` 信息块（如果存在）。 如果没有块，错误 `catch` 将终止方法。

如果你希望具有目标特征的单元格始终存在，则你可能想要代码在没有这些单元格的时候引发错误。 若没有匹配单元格是一个有效应用场景，代码应该会检查这种可能的情况并按正常方式处理它，而不会引发错误。 可以用此 `getSpecialCellsOrNullObject` 方法及其返回的 `isNullObject` 属性实现此行为。 下面的代码示例使用此模式。 关于此代码，请注意以下几点：

- 方法 `getSpecialCellsOrNullObject` 始终返回代理对象，因此在 `null` 普通 JavaScript 意义上，它永远不会返回。 但是，如果没有找到匹配的单元格，则对象的 `isNullObject` 属性将设置为 `true`。
- 在测试 `isNullObject` 属性 *之前*，它将调用 `context.sync`。 这是所有 `*OrNullObject` 方法和属性的要求，因为你必须始终加载和同步属性才能读取它。 但是，不需要显式 *加载*`isNullObject`属性。 即使未在 对象上调用 `context.sync` ， `load` 它也会自动加载。 有关详细信息，请参阅 [\*OrNullObject 方法和属性](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)。
- 你可以测试此代码，方法是先选择没有公式单元格的区域并运行它。 然后选择至少包含一个带公式的单元格的区域，并再次运行它。

```js
await Excel.run(async (context) => {
    let range = context.workbook.getSelectedRange();
    let formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    await context.sync();
        
    if (formulaRanges.isNullObject) {
        console.log("No cells have formulas");
    }
    else {
        formulaRanges.format.fill.color = "pink";
    }
    
    await context.sync();
});
```

为简单起见，本文中的所有其他代码示例都使用 `getSpecialCells` 方法而不是  `getSpecialCellsOrNullObject`。

## <a name="narrow-the-target-cells-with-cell-value-types"></a>通过单元格值类型缩小目标单元格的范围

`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法接受一个可选第二参数，用于进一步缩小目标单元格。 此第二参数是你用于指定只希望包含特定数值类型单元格的一个 `Excel.SpecialCellValueType`。

> [!NOTE]
> 当且仅当 `Excel.SpecialCellType` 为 `Excel.SpecialCellType.formulas` 或 `Excel.SpecialCellType.constants` 时才使用 `Excel.SpecialCellValueType` 参数。

### <a name="test-for-a-single-cell-value-type"></a>测试单个单元格值类型

`Excel.SpecialCellValueType` 枚举有四种基本类型 （本节后续部分所述其他组合值除外）：

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical`（意味着布尔值）
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

下面的代码示例查找数值常量的特殊单元格，并设置这些单元格的粉色。 关于此代码，请注意以下几点：

- 它只突出显示具有文字数字值的单元格。 它不会突出显示具有公式单元格 (即使结果是数字或布尔) 、文本或错误状态单元格。
- 要测试代码，请确保工作表中的某些单元格包含文本数值，某些包含其他类型的文本值，而某些则包含公式。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

### <a name="test-for-multiple-cell-value-types"></a>测试多个单元格值类型

有时，你需要对多种单元格值类型执行操作，例如所有文本值和所有布尔值（`Excel.SpecialCellValueType.logical`）单元格。 `Excel.SpecialCellValueType` 枚举具有组合类型的值。 例如，`Excel.SpecialCellValueType.logicalText` 将定向所有布尔值和所有文本值单元格。 `Excel.SpecialCellValueType.all` 是默认值，并不限制返回的单元格值类型。 下面的代码示例使用产生数字或布尔值的公式来设置所有单元格的颜色。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    await context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API Excel字符串](excel-add-ins-ranges-string-match.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
