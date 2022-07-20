---
title: 使用 Excel JavaScript API 删除重复项
description: 了解如何使用 Excel JavaScript API 删除重复项。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9ece7c9f35b341dbb8d0d90e8ca4bda5215580ed
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889140"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 删除重复项

本文提供了一个代码示例，该示例使用 Excel JavaScript API 删除区域中的重复条目。 有关对象支持的属性和方法的 `Range` 完整列表，请参阅 [Excel.Range 类](/javascript/api/excel/excel.range)。

## <a name="remove-rows-with-duplicate-entries"></a>删除具有重复条目的行

[Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) 方法删除指定列中具有重复条目的行。 该方法将遍经范围中的每一行，从最低值索引到从上到下)  (范围内的最高值索引。 如果指定列中的值之前显示在区域中，则会删除该行。 在区域内位于已删除行下方的行将上移。 `removeDuplicates` 不影响该区域外的单元格位置。

`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。 此数组从零开始并且与区域而不是与工作表相关。 该方法还采用一个布尔参数，该参数指定第一行是否为标头。 当 `true`考虑重复项时，将忽略顶部行。 该 `removeDuplicates` 方法返回一个 `RemoveDuplicatesResult` 对象，该对象指定删除的行数和剩余的唯一行数。

使用区域 `removeDuplicates` 的方法时，请记住以下事项。

- `removeDuplicates` 会考虑单元格值，而不是函数结果。 如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。
- `removeDuplicates` 不会忽略空单元格。 空单元格的值与任何其他值具有相同的处理方式。 这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。

下面的代码示例显示删除第一列中具有重复值的条目。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B2:D11");

    let deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    await context.sync();

    console.log(deleteResult.removed + " entries with duplicate names removed.");
    console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
});
```

### <a name="data-before-duplicate-entries-are-removed"></a>删除重复条目之前的数据

![在运行范围的删除重复方法之前，Excel 中的数据。](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>删除重复条目后的数据

![在运行范围的删除重复方法后，Excel 中的数据。](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [使用 Excel JavaScript API 剪切、复制和粘贴范围](excel-add-ins-ranges-cut-copy-paste.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
