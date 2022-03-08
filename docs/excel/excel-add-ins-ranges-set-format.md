---
title: 使用 JavaScript API 设置Excel格式
description: 了解如何使用 Excel JavaScript API 设置区域的格式。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 41727f6fd71636be24bdc1bb8416cb3ba07c06e1
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340349"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>使用 JavaScript API Excel区域格式

本文提供的代码示例使用 JavaScript API 为区域单元格设置字体颜色、填充颜色和数字Excel格式。 有关对象支持的属性和方法`Range`的完整列表，请参阅Excel[。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>设置字体颜色和填充颜色

下面的代码示例为区域 **B2:E2** 中的单元格设置字体颜色和填充颜色。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    await context.sync();
});
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>区域中设置字体颜色和填充颜色之前的数据

![设置Excel之前的数据。](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>区域中设置字体颜色和填充颜色之后的数据

![设置Excel格式之后的数据。](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>设置数字格式

下面的代码示例为区域 **D3:E5** 中的单元格设置数字格式。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    let range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    await context.sync();
});
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
