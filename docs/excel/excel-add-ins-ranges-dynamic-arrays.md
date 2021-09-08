---
title: 使用 JavaScript API 处理动态数组Excel溢出
description: 了解如何使用 JavaScript API 处理动态数组和Excel溢出。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d95546b4cff3f0ba7410d9ceaa73e19b7e684985
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938736"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>使用 JavaScript API 处理动态Excel溢出

本文提供了一个代码示例，该示例使用 JavaScript API 处理动态数组Excel溢出。 有关对象支持的属性和方法的完整列表， `Range` 请参阅[Excel。Range 类](/javascript/api/excel/excel.range)。

## <a name="dynamic-arrays"></a>动态数组

某些Excel返回[动态数组](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)。 这些填充公式原始单元格之外的多个单元格的值。 此值溢出称为"溢出"。 外接程序可以使用 [Range.getSpillingToRange](/javascript/api/excel/excel.range#getSpillingToRange__) 方法查找用于溢出的范围。 还有 [*OrNullObject 版本](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) `Range.getSpillingToRangeOrNullObject` 。

以下示例显示一个基本公式，该公式将区域的内容复制到单元格中，该公式会溢出到相邻的单元格中。 然后，外接程序记录包含溢出的范围。

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

## <a name="range-spilling"></a>区域溢出

使用 [Range.getSpillParent](/javascript/api/excel/excel.range#getSpillParent__) 方法查找负责溢出到给定单元格的单元格。 请注意， `getSpillParent` 仅在 range 对象是单个单元格时有效。 对 `getSpillParent` 具有多个单元格的范围调用将导致在返回 (或返回空区域时 `Range.getSpillParentOrNullObject`) 。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
