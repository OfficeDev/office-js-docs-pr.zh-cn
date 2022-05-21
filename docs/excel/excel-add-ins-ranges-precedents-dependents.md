---
title: 使用 Excel JavaScript API 处理公式先例和依赖项
description: 了解如何使用 Excel JavaScript API 检索公式先例和依赖项。
ms.date: 05/19/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ca432b7eb6825781960e995af2ed2193c7caa5e2
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628094"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 获取公式先例和依赖项

Excel公式通常引用其他单元格。 这些跨单元格引用称为“先例”和“依赖项”。 先例是向公式提供数据的单元格。 依赖项是包含引用其他单元格的公式的单元格。 若要详细了解Excel与单元格之间的关系相关的功能，请参阅[显示公式和单元格之间的关系](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507)。

一个先例单元可能有自己的先例单元格。 这一系列先例中的每一个先例单元仍然是原始单元格的先例。 依赖者存在相同的关系。 受另一个单元格影响的任何单元格都是该单元格的依赖项。 “直接先例”是此序列中前一组单元格，类似于父子关系中父母的概念。 “直接依赖”是序列中第一个依赖单元格组，类似于父子关系中的子级。

本文提供使用 Excel JavaScript API 检索公式的先例和依赖项的代码示例。 有关对象支持的属性和方法`Range`的完整列表，请[参阅 Range Object (JavaScript API for Excel) ](/javascript/api/excel/excel.range)。

## <a name="get-the-precedents-of-a-formula"></a>获取公式的先例

使用 [Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)) 查找公式的先例单元格。 `Range.getPrecedents` 返回一个 `WorkbookRangeAreas` 对象。 此对象包含工作簿中所有先例的地址。 它为每个工作表具有一个单独 `RangeAreas` 的对象，其中至少包含一个公式先例。 若要了解有关对象的`RangeAreas`详细信息，请参阅[Excel加载项中同时使用多个范围](excel-add-ins-multiple-ranges.md)。

若要仅查找公式的直接先例单元格，请使用 [Range.getDirectPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1))。 `Range.getDirectPrecedents` 工作方式类似于 `Range.getPrecedents` 并返回包含 `WorkbookRangeAreas` 直接先例地址的对象。

以下屏幕截图显示了在 Excel UI 中选择 **“跟踪先例**”按钮的结果。 此按钮将箭头从前置单元格绘制到所选单元格。 所选单元格 **E3** 包含公式“=C3 * D3”，因此 **C3** 和 **D3** 都是先例单元格。 与Excel UI 按钮不同，`getPrecedents`这些和`getDirectPrecedents`方法不会绘制箭头。

![箭头跟踪Excel UI 中的先例单元格。](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> 这些 `getPrecedents` 和 `getDirectPrecedents` 方法不会在工作簿中检索先例单元格。

下面的代码示例演示如何使用这些 `Range.getPrecedents` 方法和 `Range.getDirectPrecedents` 方法。 该示例获取活动区域的先例，然后更改这些先例单元格的背景色。 直接的先例单元格的背景色设置为黄色，其他前例单元格的背景色设置为橙色。

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
await Excel.run(async (context) => {
  let range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  let precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  let directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  await context.sync();

  console.log(`All precedent cells of ${range.address}:`);
  
  // Use the precedents API to loop through all precedents of the active cell.
  for (let i = 0; i < precedents.areas.items.length; i++) {
    // Highlight and print out the address of all precedent cells.
    precedents.areas.items[i].format.fill.color = "Orange";
    console.log(`  ${precedents.areas.items[i].address}`);
  }

  console.log(`Direct precedent cells of ${range.address}:`);

  // Use the direct precedents API to loop through direct precedents of the active cell.
  for (let i = 0; i < directPrecedents.areas.items.length; i++) {
    // Highlight and print out the address of each direct precedent cell.
    directPrecedents.areas.items[i].format.fill.color = "Yellow";
    console.log(`  ${directPrecedents.areas.items[i].address}`);
  }
});
```

## <a name="get-the-dependents-of-a-formula"></a>获取公式的依赖项

使用 [Range.getDependents](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1)) 查找公式的依赖单元格。 同样 `Range.getPrecedents`， `Range.getDependents` 还返回一个 `WorkbookRangeAreas` 对象。 此对象包含工作簿中所有依赖项的地址。 它为每个工作表具有一个单独 `RangeAreas` 的对象，其中至少包含一个依赖公式的工作表。 有关使用对象的`RangeAreas`详细信息，请参阅[Excel加载项中同时使用多个范围](excel-add-ins-multiple-ranges.md)。

若要仅查找公式的直接依赖单元格，请使用 [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1))。 `Range.getDirectDependents` 工作方式类似于 `Range.getDependents` 并返回包含 `WorkbookRangeAreas` 直接依赖项地址的对象。

以下屏幕截图显示了在 Excel UI 中选择 **“跟踪依赖项**”按钮的结果。 此按钮将箭头从所选单元格绘制到依赖单元格。 所选单元格 **D3** 将单元 **格 E3** 作为依赖项。 **E3** 包含公式“=C3 * D3”。 与Excel UI 按钮不同，`getDependents`这些和`getDirectDependents`方法不会绘制箭头。

![箭头跟踪Excel UI 中的依赖单元格。](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> 这些 `getDependents` 和 `getDirectDependents` 方法不会检索工作簿中的依赖单元格。

下面的代码示例获取活动范围的直接依赖项，然后将这些依赖单元格的背景色更改为黄色。

下面的代码示例演示如何使用这些 `Range.getDependents` 方法和 `Range.getDirectDependents` 方法。 该示例获取活动区域的依赖项，然后更改这些依赖单元格的背景色。 直接依赖单元格的背景色设置为黄色，其他从属单元格的背景色设置为橙色。

```js
// This code sample shows how to find and highlight the dependents 
// and direct dependents of the currently selected cell.
await Excel.run(async (context) => {
    let range = context.workbook.getActiveCell();
    // Dependents are all cells that contain formulas that refer to other cells.
    let dependents = range.getDependents();  
    // Direct dependents are the child cells, or the first succeeding group of cells in a sequence of cells that refer to other cells.
    let directDependents = range.getDirectDependents();

    range.load("address");
    dependents.areas.load("address");    
    directDependents.areas.load("address");
    
    await context.sync();

    console.log(`All dependent cells of ${range.address}:`);
    
    // Use the dependents API to loop through all dependents of the active cell.
    for (let i = 0; i < dependents.areas.items.length; i++) {
      // Highlight and print out the addresses of all dependent cells.
      dependents.areas.items[i].format.fill.color = "Orange";
      console.log(`  ${dependents.areas.items[i].address}`);
    }

    console.log(`Direct dependent cells of ${range.address}:`);

    // Use the direct dependents API to loop through direct dependents of the active cell.
    for (let i = 0; i < directDependents.areas.items.length; i++) {
      // Highlight and print the address of each dependent cell.
      directDependents.areas.items[i].format.fill.color = "Yellow";
      console.log(`  ${directDependents.areas.items[i].address}`);
    }
});
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
