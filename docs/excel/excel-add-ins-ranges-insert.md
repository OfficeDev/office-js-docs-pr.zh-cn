---
title: 使用 JavaScript API Excel区域
description: 了解如何使用 JavaScript API 插入Excel单元格。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad75b8c107005777047418ff9a1824665552cb5cca06c1e858f3645172f12e7c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084643"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>使用 JavaScript API 插入Excel单元格

本文提供了一个代码示例，该示例使用 JavaScript API 插入Excel单元格。 有关对象支持的属性和方法的完整列表， `Range` 请参阅[Excel。Range 类](/javascript/api/excel/excel.range)。

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

![插入Excel之前数据。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>插入区域之后的数据

![插入Excel后数据。](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API 清除或删除Excel区域](excel-add-ins-ranges-clear-delete.md)
