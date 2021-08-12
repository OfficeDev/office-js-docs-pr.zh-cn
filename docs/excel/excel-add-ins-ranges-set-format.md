---
title: 使用 JavaScript API 设置Excel格式
description: 了解如何使用 Excel JavaScript API 设置区域的格式。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 3e9c8dd58216de695fc0bc634a6eaed7dbc724468bd9e02bc23bf34c394b5e16
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086491"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>使用 JavaScript API Excel区域格式

本文提供的代码示例使用 JavaScript API 为区域单元格设置字体颜色、填充颜色和数字Excel格式。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。

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

![设置Excel之前的数据。](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>区域中设置字体颜色和填充颜色之后的数据

![设置Excel格式后的数据。](../images/excel-ranges-format-font-and-fill.png)

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

![设置数字Excel之前的数据。](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>区域中设置数字格式之后的数据

![设置数字Excel之后的数据。](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API Excel和获取范围](excel-add-ins-ranges-set-get.md)
- [使用 JavaScript API 设置和获取区域Excel文本或公式](excel-add-ins-ranges-set-get-values.md)
