---
title: 使用 JavaScript API 处理公式引用Excel依赖项
description: 了解如何使用 JavaScript API Excel引用单元格和依赖项。
ms.date: 11/30/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 60da910879fc48f1564d43cf3f87c2a5bf930fbe
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242059"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>使用 JavaScript API 获取公式引用Excel依赖项

Excel公式通常引用其他单元格。 这些跨单元格引用称为"引用单元格"和"从属单元格"。 引用单元格是向公式提供数据的单元格。 从属单元格是包含引用其他单元格的公式的单元格。 若要了解有关与Excel关系相关的功能，请参阅显示[公式和单元格之间的关系](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507)。

引用单元格可以有自己的引用单元格。 此引用单元格链中的每个引用单元格仍是原始单元格的引用单元格。 依赖项存在相同的关系。 受其他单元格影响的任何单元格都依赖于该单元格。 "直接引用单元格"是此序列中前面的第一组单元格，类似于父子关系中父级的概念。 "直接从属"是序列中第一个从属单元格组，类似于父子关系中的子级。

本文提供使用 JavaScript API 检索公式引用单元格和从属Excel示例。 有关对象支持的属性和方法的完整列表，请参阅 `Range` Range Object [ (JavaScript API for Excel) 。 ](/javascript/api/excel/excel.range)

## <a name="get-the-precedents-of-a-formula"></a>获取公式引用单元格

使用 [Range.getPrecedents](/javascript/api/excel/excel.range#getPrecedents__)查找公式的引用单元格。 `Range.getPrecedents` 返回 `WorkbookRangeAreas` 一个对象。 此对象包含工作簿中所有引用单元格的地址。 对于每个包含 `RangeAreas` 至少一个公式引用单元格的工作表，它都有一个单独的对象。 若要详细了解对象，请参阅在加载项中同时Excel `RangeAreas` [多个区域](excel-add-ins-multiple-ranges.md)。

若要仅查找公式的直接引用单元格，请使用 [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getDirectPrecedents__)。 `Range.getDirectPrecedents` 的工作方式 `Range.getPrecedents` 与 和 返回 `WorkbookRangeAreas` 一个包含直接引用单元格地址的对象。

以下屏幕截图显示了在用户界面中 **选择"追踪** 引用单元格"Excel的结果。 此按钮绘制从引用单元格到选定单元格的箭头。 选定的单元格 **E3** 包含公式"=C3 * D3"，因此 **C3** 和 **D3 都是** 引用单元格。 与 Excel UI 按钮不同， `getPrecedents` 和 `getDirectPrecedents` 方法不绘制箭头。

![箭头跟踪活动 UI 中的引用单元格Excel单元格。](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> `getPrecedents`和 `getDirectPrecedents` 方法不检索工作簿中的引用单元格。

下面的代码示例演示如何使用 `Range.getPrecedents` 和 `Range.getDirectPrecedents` 方法。 该示例获取活动区域引用单元格，然后更改这些引用单元格的背景色。 直接引用单元格的背景色设置为黄色，其他引用单元格的背景色设置为橙色。

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
Excel.run(function (context) {
  var range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  var precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  var directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  return context.sync()
    .then(function () {
      console.log(`All precedent cells of ${range.address}:`);
      
      // Use the precedents API to loop through all precedents of the active cell.
      for (var i = 0; i < precedents.areas.items.length; i++) {
        // Highlight and print out the address of all precedent cells.
        precedents.areas.items[i].format.fill.color = "Orange";
        console.log(`  ${precedents.areas.items[i].address}`);
      }

      console.log(`Direct precedent cells of ${range.address}:`);

      // Use the direct precedents API to loop through direct precedents of the active cell.
      for (var i = 0; i < directPrecedents.areas.items.length; i++) {
        // Highlight and print out the address of each direct precedent cell.
        directPrecedents.areas.items[i].format.fill.color = "Yellow";
        console.log(`  ${directPrecedents.areas.items[i].address}`);
      }
    });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula"></a>获取公式的直接依赖项

使用 [Range.getDirectDependents 查找公式的直接从属单元格](/javascript/api/excel/excel.range#getDirectDependents__)。 与 `Range.getDirectPrecedents` 类似 `Range.getDirectDependents` ，也返回 `WorkbookRangeAreas` 对象。 此对象包含工作簿中所有直接依赖项的地址。 对于每个包含至少一个依赖公式的工作表，它 `RangeAreas` 都有一个单独的对象。 有关使用对象的信息，请参阅在加载项中同时处理Excel `RangeAreas` [区域](excel-add-ins-multiple-ranges.md)。

以下屏幕截图显示了在导航 UI 中选择"**跟踪从属项**"Excel的结果。 此按钮绘制从从属单元格到选定单元格的箭头。 选定的单元格 **D3** 将单元格 **E3** 作为从属单元格。 **E3** 包含公式"=C3 * D3"。 与 Excel UI 按钮不同， `getDirectDependents` 该方法不绘制箭头。

![箭头跟踪 UI 中的Excel单元格。](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> `getDirectDependents`方法不检索工作簿中的从属单元格。

下面的代码示例获取活动区域的直接从属单元格，然后将这些从属单元格的背景色更改为黄色。

```js
// This code sample shows how to find and highlight the dependents of the currently selected cell.
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
