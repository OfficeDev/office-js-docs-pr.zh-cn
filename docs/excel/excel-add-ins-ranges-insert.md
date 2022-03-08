---
title: 使用 JavaScript API Excel区域
description: 了解如何使用 JavaScript API 插入Excel单元格。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 0e1ed6d2302bcdb4a11688cd6d77448811f8a93b
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340545"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>使用 JavaScript API 插入Excel单元格

本文提供了一个代码示例，该示例使用 JavaScript API 插入Excel单元格。 有关对象支持的属性和方法`Range`的完整列表，请参阅Excel[。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>插入多个单元格

下面的代码示例将多个单元格插入位置 **B4:E4**，并将其他单元格下移，以便为新的单元格提供空间。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    await context.sync();
});
```

### <a name="data-before-range-is-inserted"></a>插入区域之前的数据

![插入Excel之前数据。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>插入区域之后的数据

![插入Excel后数据。](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API 清除或删除Excel区域](excel-add-ins-ranges-clear-delete.md)
