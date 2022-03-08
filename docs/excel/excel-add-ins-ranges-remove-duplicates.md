---
title: 使用 JavaScript API Excel重复项
description: 了解如何使用 JavaScript EXCEL删除重复项。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 80e1227e06f177d0e37cc2750a7830c727a59436
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340573"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>使用 JavaScript API Excel重复项

本文提供了一个代码示例，该示例使用 JavaScript API 删除Excel条目。 有关对象支持的属性和方法`Range`的完整列表，请参阅Excel[。Range 类](/javascript/api/excel/excel.range)。

## <a name="remove-rows-with-duplicate-entries"></a>删除条目重复的行

[Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) 方法删除指定列中具有重复条目的行。 方法将浏览区域的每一行，从最低值索引到区域最高值索引， (从上到下) 。 如果指定列中的值之前显示在区域中，则会删除该行。 在区域内位于已删除行下方的行将上移。 `removeDuplicates` 不影响该区域外的单元格位置。

`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。 此数组从零开始并且与区域而不是与工作表相关。 该方法还采用一个布尔参数，该参数指定第一行是否是标题。 如果为 **true**，则在考虑重复项时将忽略顶行。 该方法 `removeDuplicates` 返回一 `RemoveDuplicatesResult` 个对象，该对象指定删除的行数和剩余的唯一行数。

使用区域的方法时 `removeDuplicates` ，请牢记以下事项。

- `removeDuplicates` 会考虑单元格值，而不是函数结果。 如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。
- `removeDuplicates` 不会忽略空单元格。 空单元格的值与任何其他值具有相同的处理方式。 这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。

下面的代码示例演示删除第一列中具有重复值的条目。

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

![区域Excel重复项方法运行之前的数据。](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>删除重复条目后的数据

![区域Excel重复项方法运行后的数据。](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API Excel、复制和粘贴区域](excel-add-ins-ranges-cut-copy-paste.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
