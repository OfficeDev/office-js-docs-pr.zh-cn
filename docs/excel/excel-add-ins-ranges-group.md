---
title: 使用 JavaScript API Excel组范围
description: 了解如何将一个范围的行或列组合在一起，以使用 JavaScript API Excel大纲。
ms.date: 04/05/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9ec3f9e23f5099c703fbbf53fdc6fbb800acba6d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149265"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a>使用 JavaScript API 的大纲Excel区域

本文提供了一个代码示例，演示如何使用 JavaScript API 对大纲Excel分组。 有关对象支持的属性和方法的完整列表， `Range` 请参阅[Excel。Range 类](/javascript/api/excel/excel.range)。

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a>对分级显示区域行或列进行分组

可以将范围的行或列组合在一起以 [创建大纲](https://support.microsoft.com/office/08ce98c4-0063-4d42-8ac7-8278c49e9aff)。 可以折叠和展开这些组，以隐藏和显示相应的单元格。 这使得快速分析首行数据变得更容易。 使用 [Range.group](/javascript/api/excel/excel.range#group_groupOption_) 可创建这些大纲组。

大纲可以具有层次结构，其中较小的组嵌套在较大的组下。 这允许在不同级别查看大纲。 可以通过 [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_) 方法以编程方式更改可见大纲级别。 请注意，Excel仅支持八个级别的大纲组。

下面的代码示例为行和列创建包含两个级别的组的大纲。 后续图像显示该轮廓的分组。 在代码示例中，分组的范围不包括大纲控件的行或列 (此示例的"总计") 。 组定义要折叠的项，而不是控件的行或列。

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

![具有两级、二维轮廓的范围。](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a>从区域行或列中删除分组

若要取消分组行或列组，请使用 [Range.ungroup](/javascript/api/excel/excel.range#ungroup_groupOption_) 方法。 这将从大纲中删除最外面的级别。 如果同一行或列类型的多个组位于指定范围内的同一级别，则所有这些组将取消分组。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
