---
title: 使用 JavaScript API 设置并Excel区域
description: 了解如何使用 javaScript API Excel JavaScript API 设置和获取Excel范围。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0bd4a4f4bcf40e7899ee429cdc631a43ba176077
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075773"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a>使用 JavaScript API Excel和获取范围

本文提供了使用 JavaScript API 设置和获取区域Excel示例。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>设置所选区域

下面的代码示例选择活动工作表中的区域 **B2:E6**。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a>选定的区域 B2:E6

![选定区域Excel。](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>获取所选区域

下面的代码示例获取所选区域、加载其 `address` 属性，然后向控制台写入一条消息。

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API 设置和获取区域Excel文本或公式](excel-add-ins-ranges-set-get-values.md)
- [使用 JavaScript API Excel区域格式](excel-add-ins-ranges-set-format.md)
