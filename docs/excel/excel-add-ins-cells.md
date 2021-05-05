---
title: 使用 Excel JavaScript API 处理单元格。
description: 了解单元格的 Excel JavaScript API 定义，并了解如何使用单元格。
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917098"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 处理单元格

The Excel JavaScript API 没有“Cell”对象或类。 相反，所有 Excel 单元格都是 `Range` 对象。 Excel UI 中的单个单元格转换为 Excel JavaScript API 中包含一个单元格的 `Range` 对象。

对象 `Range` 还可以包含多个连续单元格。 连续单元格形成一个不间断的矩形 (包括单个行或) 。 若要了解如何处理不连续的单元格，请参阅使用 [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object)对象处理不连续的单元格。

有关对象支持的属性和方法的完整列表，请参阅 `Range` Range Object [ (JavaScript API for Excel) 。 ](/javascript/api/excel/excel.range)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a>使用 RangeAreas 对象处理不连续单元格

[RangeAreas](/javascript/api/excel/excel.rangeareas)对象允许您的外接程序一次对多个区域执行操作。 这些区域可能是连续的，但不必是。 `RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 获取区域](excel-add-ins-ranges-get.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)