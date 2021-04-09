---
title: 使用 Excel JavaScript API 处理公式引用单元格
description: 了解如何使用 Excel JavaScript API 检索公式引用单元格。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652796"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 获取公式引用单元格

本文提供使用 Excel JavaScript API 检索公式引用单元格的代码示例。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。

## <a name="get-formula-precedents"></a>获取公式引用单元格

Excel 公式通常引用其他单元格。 当单元格向公式提供数据时，它称为公式"precedent"。 若要了解有关与单元格之间的关系相关的 Excel 功能，请参阅 [显示公式和单元格之间的关系](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)。 

使用 [Range.getDirectPrecedents，](/javascript/api/excel/excel.range#getdirectprecedents--)加载项可以定位公式的直接引用单元格。 `Range.getDirectPrecedents` 返回 `WorkbookRangeAreas` 一个对象。 此对象包含工作簿中所有引用单元格的地址。 对于每个包含 `RangeAreas` 至少一个公式引用单元格的工作表，它都有一个单独的对象。 有关使用对象的信息，请参阅在 Excel 加载项中同时处理 `RangeAreas` [多个区域](excel-add-ins-multiple-ranges.md)。

在 Excel UI 中 **，"追踪引用** 单元格"按钮绘制从引用单元格到所选公式的箭头。 与 Excel UI 按钮不同， `getDirectPrecedents` 该方法不绘制箭头。 

> [!IMPORTANT]
> `getDirectPrecedents`方法无法跨工作簿检索引用单元格。 

下面的代码示例获取活动区域的直接引用单元格，然后将这些引用单元格的背景色更改为黄色。 

> [!NOTE]
> 活动区域必须包含一个公式，该公式引用同一工作簿中的其他单元格，使突出显示正常工作。 

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
