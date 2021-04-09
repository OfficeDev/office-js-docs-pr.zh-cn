---
title: 使用 Excel JavaScript API 处理单元格。
description: 了解单元格的 Excel JavaScript API 定义，并了解如何使用单元格。
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652874"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理单元格

Excel JavaScript API 没有"Cell"对象或类。 相反，所有 Excel 单元格都是 `Range` 对象。 Excel UI 中的单个单元格转换为 Excel JavaScript API 中具有一个单元格 `Range` 的对象。

对象 `Range` 还可以包含多个连续单元格。 连续单元格形成一个不间断的矩形 (包括单个行或) 。 若要了解如何处理不连续的单元格，请参阅使用 [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object)对象处理不连续的单元格。

有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。

## <a name="excel-javascript-apis-that-mention-cells"></a>提及单元格的 Excel JavaScript API

即使 Excel JavaScript API 没有"Cell"对象或类，许多 API 名称也提及单元格。 这些 API 控制单元格属性，如颜色、文本格式和字体。

以下 Excel JavaScript API 列表引用单元格。

- [CellBorder](/javascript/api/excel/excel.cellborder)
- [CellBorderCollection](/javascript/api/excel/excel.cellbordercollection)
- [CellProperties](/javascript/api/excel/excel.cellproperties)
- [CellPropertiesFill](/javascript/api/excel/excel.cellpropertiesfill)
- [CellPropertiesFont](/javascript/api/excel/excel.cellpropertiesfont)
- [CellPropertiesFormat](/javascript/api/excel/excel.cellpropertiesformat)
- [CellPropertiesProtection](/javascript/api/excel/excel.cellpropertiesprotection)
- [CellValueConditionalFormat](/javascript/api/excel/excel.cellvalueconditionalformat)
- [ConditionalCellValueRule](/javascript/api/excel/excel.conditionalcellvaluerule)
- [SettableCellProperties](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>使用 RangeAreas 对象处理不连续单元格

[RangeAreas](/javascript/api/excel/excel.rangeareas)对象允许您的外接程序一次对多个区域执行操作。 这些区域可能是连续的，但不必是。 `RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 获取区域](excel-add-ins-ranges-get.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
