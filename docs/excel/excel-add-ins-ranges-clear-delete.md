---
title: 使用 Excel JavaScript API 清除或删除区域
description: 了解如何使用 Excel JavaScript API 清除或删除区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7e030c6b5ba7ba6e6c54e9be0524cd93c2516bcb
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652870"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 清除或删除区域

本文提供的代码示例使用 Excel JavaScript API 清除和删除区域。 有关对象支持的属性和方法的完整 `Range` 列表，请参阅 [Excel.Range 类](/javascript/api/excel/excel.range)。

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

![Excel 中清除区域之前的数据](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>清除区域之后的数据

![Excel 中清除区域之后的数据](../images/excel-ranges-after-clear.png)

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

![Excel 中删除区域之前的数据](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>删除区域之后的数据

![Excel 中删除区域之后的数据](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a>另请参阅

- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [使用 Excel JavaScript API 设置和获取区域](excel-add-ins-ranges-set-get.md)
- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
