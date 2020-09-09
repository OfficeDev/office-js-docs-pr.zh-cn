---
title: 使用 Excel JavaScript API 对区域执行操作（高级）
description: 高级的 range 对象函数和方案，如特殊单元格、删除重复项以及使用日期。
ms.date: 08/26/2020
localization_priority: Normal
ms.openlocfilehash: b3854d15a85db20e1c544ebfa6e8a63712e958d9
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408444"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a>使用 Excel JavaScript API 对区域执行操作（高级）

本文基于[使用 Excel JavaScript API 对区域执行操作（基本）](excel-add-ins-ranges.md)中包含的信息，它提供了显示如何使用 Excel JavaScript API 对区域执行更多高级任务的代码示例。 有关该对象支持的属性和方法的完整列表 `Range` ，请参阅 [Range 对象 (适用于 Excel 的 JavaScript API) ](/javascript/api/excel/excel.range)。

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a>使用 Moment-MSDate 插件处理日期

[时刻 JavaScript 库](https://momentjs.com/)提供了使用日期和时间戳的便捷方式。 [Moment-MSDate 插件](https://www.npmjs.com/package/moment-msdate)可将时刻格式转换为 Excel 所需的格式。 这是 [NOW 函数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)返回的相同格式。

以下代码显示如何将 **B4** 处的范围设置为时刻的时间戳：

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

这是一项类似于在单元格之外获取日期并将其转换为时刻或其他格式的技术，如以下代码中所示：

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

你的加载项将必须对范围进行格式化才能以更可读的形式显示日期。 `"[$-409]m/d/yy h:mm AM/PM;@"` 的示例显示类似“12/3/18 3:57 PM”的时间。 有关日期和时间数字格式的详细信息，请参阅[查看自定义数字格式的准则](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)一文中的“日期和时间格式的准则”。

## <a name="work-with-multiple-ranges-simultaneously"></a>同时处理多个区域

[RangeAreas](/javascript/api/excel/excel.rangeareas)对象允许外接程序一次在多个区域上执行操作。 这些区域可能但不必是连续区域。 `RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。

## <a name="find-special-cells-within-a-range"></a>查找区域中的特殊单元格

[GetSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)和[getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)方法根据单元格的特征和单元格的值类型来查找区域。 这两种方法都返回 `RangeAreas` 对象。 以下是 TypeScript 数据类型文件中方法的签名：

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

下面示例使用 `getSpecialCells` 方法来查找有公式的所有单元格。 关于此代码，请注意以下几点：

- 它通过先调用 `Worksheet.getUsedRange` 并仅调用该区域的 `getSpecialCells` 来限制需要搜索的工作表部分。
- `getSpecialCells` 方法将返回 `RangeAreas` 对象，因此包含公式的单元格都会变成粉色，即使它们并非都是连续的单元格。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

如果区域中不存在具有目标特征的单元格，`getSpecialCells` 会引发 **ItemNotFound**错误。 这会将控制流转移到 `catch` 信息块（如果存在）。 如果没有 `catch` 块，则错误停止方法。

如果你希望具有目标特征的单元格始终存在，则你可能想要代码在没有这些单元格的时候引发错误。 若没有匹配单元格是一个有效应用场景，代码应该会检查这种可能的情况并按正常方式处理它，而不会引发错误。 可以用此 `getSpecialCellsOrNullObject` 方法及其返回的 `isNullObject` 属性实现此行为。 此示例使用此模式。 关于此代码，请注意以下几点：

- `getSpecialCellsOrNullObject` 方法将始终返回代理对象，因此在一般的 JavaScript 认知中，它从不为 `null`。 但是，如果没有找到匹配的单元格，则对象的 `isNullObject` 属性将设置为 `true`。
- 在测试 `isNullObject` 属性*之前*，它将调用 `context.sync`。 这是所有 `*OrNullObject` 方法和属性的要求，因为你必须始终加载和同步属性才能读取它。 但是，不必*明确*加载 `isNullObject` 属性。 即使未在对象上调用 `load`，`context.sync` 也会自动加载该属性。 有关详细信息，请参阅[ \* OrNullObject 方法和属性](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)。
- 你可以测试此代码，方法是先选择没有公式单元格的区域并运行它。 然后选择至少包含一个带公式的单元格的区域，并再次运行它。

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

### <a name="narrow-the-target-cells-with-cell-value-types"></a>通过单元格值类型缩小目标单元格的范围

`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法接受一个可选第二参数，用于进一步缩小目标单元格。 此第二参数是你用于指定只希望包含特定数值类型单元格的一个 `Excel.SpecialCellValueType`。

> [!NOTE]
> 当且仅当 `Excel.SpecialCellType` 为 `Excel.SpecialCellType.formulas` 或 `Excel.SpecialCellType.constants` 时才使用 `Excel.SpecialCellValueType` 参数。

#### <a name="test-for-a-single-cell-value-type"></a>测试单个单元格值类型

`Excel.SpecialCellValueType` 枚举有四种基本类型 （本节后续部分所述其他组合值除外）：

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical`（意味着布尔值）
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

以下示例查找数值常量的特殊单元格，并将这些单元格设置为粉色。 关于此代码，请注意以下几点：

- 它只会突出显示具有文本数值的单元格。 它既不会突出显示具有公式的单元格（即使结果是数字），也不会突出显示布尔、文本或错误状态单元格。
- 要测试代码，请确保工作表中的某些单元格包含文本数值，某些包含其他类型的文本值，而某些则包含公式。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### <a name="test-for-multiple-cell-value-types"></a>测试多个单元格值类型

有时，你需要对多种单元格值类型执行操作，例如所有文本值和所有布尔值（`Excel.SpecialCellValueType.logical`）单元格。 `Excel.SpecialCellValueType` 枚举具有组合类型的值。 例如，`Excel.SpecialCellValueType.logicalText` 将定向所有布尔值和所有文本值单元格。 `Excel.SpecialCellValueType.all` 是默认值，并不限制返回的单元格值类型。 以下示例设置了包含用于生成数字或布尔值的公式的所有单元格颜色。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="cut-copy-and-paste"></a>剪切、复制和粘贴

### <a name="copy-and-paste"></a>Copy and paste

[CopyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)方法复制 Excel UI 的**复制**和**粘贴**操作。 调用 `copyFrom` 的区域对象是目标。 将要复制的源作为一个范围或一个表示范围的字符串地址进行传递。

以下代码示例将数据从“A1:E1”**** 复制到“G1”**** 开始的范围（粘贴到“G1:K1”**** 结束）。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

`Range.copyFrom` 具有三个可选参数。

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` 指定将哪些数据从源复制到目标。

- `Excel.RangeCopyType.formulas` 传输源单元格中的公式，并保留这些公式区域的相对定位。 将原样复制任何非公式条目。
- `Excel.RangeCopyType.values` 复制数据值，如果是公式，则复制公式的结果。
- `Excel.RangeCopyType.formats` 复制范围的格式设置（包括字体、颜色和其他格式），但不会复制任何值。
- `Excel.RangeCopyType.all` (默认选项) 复制数据和格式，并保留单元格的公式（如果找到）。

`skipBlanks` 设置是否将空白单元格复制到目标。 如果为 true，`copyFrom` 将跳过源范围中的空白单元格。
跳过的单元格不会覆盖目标范围中其对应单元格的现有数据。 默认值为 false。

`transpose` 确定是否将数据转置（即切换其行和列）到源位置。
转置范围沿主对角线翻转，因此，行“1”****、“2”**** 和“3”**** 将成为列“A”****、“B”**** 和“C”****。

以下代码示例和图像在一个简单的方案中演示此行为。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

*在上一个函数已运行之前。*

![Excel 中的数据，在区域的复制方法已运行之前](../images/excel-range-copyfrom-skipblanks-before.png)

*在上一个函数已运行之后。*

![Excel 中区域的复制方法已运行后的数据](../images/excel-range-copyfrom-skipblanks-after.png)

### <a name="cut-and-paste-move-cells"></a>剪切并粘贴 (移动) 单元格

该 [范围的 moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) 方法将单元格移动到工作簿中的新位置。 此单元格移动行为与单元格移动时的工作方式相同， [拖动区域边框](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) 或执行 **剪切** 和 **粘贴** 操作时。 将区域的格式和值移到指定作为参数的位置 `destinationRange` 。

下面的代码示例显示了使用方法移动的范围 `Range.moveTo` 。 请注意，如果目标区域小于源，它将被扩展以包含源内容。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="remove-duplicates"></a>删除重复项

[RemoveDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)方法删除指定列中具有重复条目的行。 该方法将从最小值索引到范围 (从上到下) 的最高值索引的范围中的每一行进行遍历。 如果指定列中的值之前显示在区域中，则会删除该行。 在区域内位于已删除行下方的行将上移。 `removeDuplicates` 不影响该区域外的单元格位置。

`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。 此数组从零开始并且与区域而不是与工作表相关。 此方法还采用一个布尔参数，用于指定第一行是否为标头。 如果为 **true**，则在考虑重复项时将忽略顶行。 `removeDuplicates`方法返回一个 `RemoveDuplicatesResult` 对象，该对象指定删除的行数和剩余的唯一行数。

使用区域的方法时 `removeDuplicates` ，请记住以下几点：

- `removeDuplicates` 会考虑单元格值，而不是函数结果。 如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。
- `removeDuplicates` 不会忽略空单元格。 空单元格的值与任何其他值具有相同的处理方式。 这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。

以下示例显示删除第一列中具有重复值的条目。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

*在上一个函数已运行之前。*

![Excel 中的数据之前，已运行区域的 "删除重复方法" 方法](../images/excel-ranges-remove-duplicates-before.png)

*在上一个函数已运行之后。*

![Excel 中已运行区域的 "删除重复项" 方法后 Excel 中的数据](../images/excel-ranges-remove-duplicates-after.png)

## <a name="group-data-for-an-outline"></a>分级显示的组数据

可以将区域中的行或列组合在一起，以创建 [分级显示](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)。 可以对这些组进行折叠和扩展以隐藏和显示相应的单元格。 这样可以更轻松地快速分析顶线数据。 使用 [Range](/javascript/api/excel/excel.range#group-groupoption-) 可以创建这些分级显示组。

大纲可以有层次结构，其中较小的组嵌套在更大的组下。 这样，可以在不同的级别查看大纲。 更改可见大纲级别可以通过 [showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) 方法以编程方式完成。 请注意，Excel 仅支持八种级别的分级显示组。

下面的代码示例演示如何创建一个大纲，其中包含两个级别的行和列的组。 随后的图像显示该轮廓的分组。 请注意，在代码示例中，要分组的区域不包括大纲控件的行或列 (本示例的 "总计") 。 组定义将折叠的内容，而不是控件的行或列。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![具有两个级别的两维轮廓的范围](../images/excel-outline.png)

若要取消行或列组的分组，请使用 [Range](/javascript/api/excel/excel.range#ungroup-groupoption-) 方法。 这将从大纲中删除最外面的级别。 如果同一行或列类型的多个组在指定区域中的同一级别，则所有这些组都将被取消组合。

## <a name="handle-dynamic-arrays-and-spilling-preview"></a>处理动态数组和 spilling (预览) 

> [!NOTE]
> 动态数组和 range spilling Api 当前处于预览阶段。 [!INCLUDE [Information about using preview Excel APIs](../includes/using-excel-preview-apis.md)]

有些 Excel 公式返回 [动态数组](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)。 这些值填充公式的原始单元格外部的多个单元格的值。 此值溢出称为 "溢出"。 您的外接程序可以使用 getSpillingToRange 方法查找用于溢出的范围[。](/javascript/api/excel/excel.range#getspillingtorange--) 此外，还有一个 [* OrNullObject 版本](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties) `Range.getSpillingToRangeOrNullObject` 。

下面的示例演示将区域的内容复制到单元格中的基本公式，这些单元格会扩散到相邻的单元格中。 然后，外接程序会记录包含溢出的范围。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

您还可以使用 getSpillParent 方法，在给定单元格 [内](/javascript/api/excel/excel.range#getspillparent--) 查找负责 spilling 的单元格。 请注意， `getSpillParent` 仅当 range 对象为单个单元格时才起作用。 `getSpillParent`对包含多个单元格的区域进行调用将导致 (引发错误，或返回) 的 null 范围 `Range.getSpillParentOrNullObject` 。

## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 对区域执行操作](excel-add-ins-ranges.md)
- [Office 外接程序中的 Excel JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
