---
title: 使用 JavaScript API 清除或删除Excel区域
description: 了解如何使用 JavaScript API 清除或删除Excel区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a1bd99db3aa9af3903552d9cefc6ec6d21701136
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075829"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>使用 JavaScript API 清除或删除Excel区域

本文提供的代码示例使用 JavaScript API 清除和删除Excel区域。 有关对象支持的属性和方法的完整 `Range` 列表，请参阅[Excel。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a>清除多个单元格内容

下面的代码示例清除区域 **E2:E5** 中的所有内容和单元格格式设置。  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a>清除区域之前的数据

![清除Excel之前的数据。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>清除区域之后的数据

![清除区域Excel中数据。](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>删除多个单元格

下面的代码示例删除 **区域 B4：E4** 中的单元格，并上移其他单元格以填充已删除单元格空出的空间。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a>删除区域之前的数据

![删除Excel之前的数据。](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>删除区域之后的数据

![删除Excel区域之后的数据。](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a>另请参阅

- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API Excel和获取范围](excel-add-ins-ranges-set-get.md)
- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
