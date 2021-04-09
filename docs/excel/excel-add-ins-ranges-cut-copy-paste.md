---
title: 使用 Excel JavaScript API 剪切、复制和粘贴区域
description: 了解如何使用 Excel JavaScript API 剪切、复制和粘贴区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8cf92ef148c24613674930140cec762c9cd8c4a4
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652812"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 剪切、复制和粘贴区域

本文提供使用 Excel JavaScript API 剪切、复制和粘贴区域的代码示例。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a>Copy and paste

[Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)方法复制 Excel UI的 **复制** 和粘贴操作。 目标为 `Range` 所 `copyFrom` 调用的对象。 将要复制的源作为一个范围或一个表示范围的字符串地址进行传递。

以下代码示例将数据从“A1:E1”复制到“G1”开始的范围（粘贴到“G1:K1”结束）。

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

- `Excel.RangeCopyType.formulas` 传输源单元格中的公式，并保留这些公式区域的相对位置。 将原样复制任何非公式条目。
- `Excel.RangeCopyType.values` 复制数据值，如果是公式，则复制公式的结果。
- `Excel.RangeCopyType.formats` 复制范围的格式设置（包括字体、颜色和其他格式），但不会复制任何值。
- `Excel.RangeCopyType.all` (默认选项) 复制数据和格式，并保留单元格的公式（如果找到）。

`skipBlanks` 设置是否将空白单元格复制到目标。 如果为 true，`copyFrom` 将跳过源范围中的空白单元格。
跳过的单元格不会覆盖目标范围中其对应单元格的现有数据。 默认值为 false。

`transpose` 确定是否将数据转置（即切换其行和列）到源位置。
转置范围沿主对角线翻转，因此，行“1”、“2”和“3”将成为列“A”、“B”和“C”。

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

### <a name="data-before-range-is-copied-and-pasted"></a>复制和粘贴区域之前的数据

![Excel 中运行区域复制方法之前的数据](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a>复制和粘贴区域之后的数据

![Excel 中运行区域复制方法后的数据](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a>剪切并粘贴 (单元格) 移动

[Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-)方法将单元格移动到工作簿中的新位置。 此单元格移动行为的工作方式与通过拖动区域边框或执行"[](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e)剪切"和"粘贴"操作移动单元格 **时相同**。 区域的格式和值都移至指定为 参数 `destinationRange` 的位置。

下面的代码示例使用 方法移动 `Range.moveTo` 区域。 请注意，如果目标区域小于源范围，它将扩展以包含源内容。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [使用 Excel JavaScript API 删除重复项](excel-add-ins-ranges-remove-duplicates.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
