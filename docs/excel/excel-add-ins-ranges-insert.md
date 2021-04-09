---
title: 使用 Excel JavaScript API 插入区域
description: 了解如何使用 Excel JavaScript API 插入单元格区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 401a08dd10b3775012738ab9c80ec6ab367555ec
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652798"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 插入单元格区域

本文提供了一个代码示例，该示例使用 Excel JavaScript API 插入一系列单元格。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>插入多个单元格

下面的代码示例将多个单元格插入位置 **B4:E4**，并将其他单元格下移，以便为新的单元格提供空间。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a>插入区域之前的数据

![Excel 中插入区域之前的数据](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>插入区域之后的数据

![Excel 中插入区域之后的数据](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [使用 Excel JavaScript API 清除或删除区域](excel-add-ins-ranges-clear-delete.md)
