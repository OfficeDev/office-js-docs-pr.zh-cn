---
title: 使用 Excel JavaScript API 设置区域的格式
description: 了解如何使用 Excel JavaScript API 设置区域的格式。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: fdd78ea69fc38cbefb9d240dbc61554891c73c21
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652794"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>使用 Excel JavaScript API 设置区域格式

本文提供的代码示例使用 Excel JavaScript API 为区域单元格设置字体颜色、填充颜色和数字格式。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>设置字体颜色和填充颜色

下面的代码示例为区域 **B2:E2** 中的单元格设置字体颜色和填充颜色。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>区域中设置字体颜色和填充颜色之前的数据

![Excel 中设置格式之前的数据](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>区域中设置字体颜色和填充颜色之后的数据

![Excel 中设置格式之后的数据](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>设置数字格式

下面的代码示例为区域 **D3:E5** 中的单元格设置数字格式。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-number-format-is-set"></a>区域中设置数字格式之前的数据

![设置数字格式之前 Excel 中的数据](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>区域中设置数字格式之后的数据

![设置数字格式后 Excel 中的数据](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 Excel JavaScript API 处理单元格](excel-add-ins-cells.md)
- [使用 Excel JavaScript API 设置和获取区域](excel-add-ins-ranges-set-get.md)
- [使用 Excel JavaScript API 设置和获取区域值、文本或公式](excel-add-ins-ranges-set-get-values.md)
