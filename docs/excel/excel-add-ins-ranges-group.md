---
title: 使用 Excel JavaScript API 的组范围
description: 了解如何将区域行或列组合在一起，以使用 Excel JavaScript API 创建大纲。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 32f65cf88c23bd6368b37318d3ba20fde95b8436
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652799"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 的大纲组范围

本文提供了一个代码示例，演示如何使用 Excel JavaScript API 对大纲区域进行分组。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a>对分级显示区域行或列进行分组

可以将范围的行或列组合在一起以 [创建大纲](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)。 可以折叠和展开这些组，以隐藏和显示相应的单元格。 这使得快速分析首行数据变得更容易。 使用 [Range.group](/javascript/api/excel/excel.range#group-groupoption-) 可创建这些大纲组。

大纲可以具有层次结构，其中较小的组嵌套在较大的组下。 这允许在不同级别查看大纲。 可以通过 [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) 方法以编程方式更改可见大纲级别。 请注意，Excel 仅支持八个级别的大纲组。

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

![具有两级、二维轮廓的范围](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a>从区域行或列中删除分组

若要取消行或列组的分组，请使用 [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) 方法。 这将从大纲中删除最外面的级别。 如果同一行或列类型的多个组位于指定范围内的同一级别，则所有这些组将取消分组。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)