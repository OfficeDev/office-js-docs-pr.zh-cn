---
title: 使用 JavaScript API Excel区域
description: 了解如何使用 JavaScript API Excel区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d48d69a45e964db2d5797e2f0927f776795bcca0365f0ccef245fcd3682a3a72
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084715"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>使用 JavaScript API Excel区域

本文提供的示例显示了使用 JavaScript API 在工作表内获取区域Excel方法。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>按地址获取区域

下面的代码示例从名为 **Sample** 的工作表获取地址 **为 B2：C5** 的范围，加载其 属性，然后向控制台写入 `address` 一条消息。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a>按名称获取区域

下面的代码示例从名为 Sample 的工作表获取名为 的范围，加载其 属性，然后 `MyRange` 向控制台 `address` 写入一条消息。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a>获取使用的区域

下面的代码示例从名为 **Sample** 的工作表获取已用区域，加载其 属性，然后 `address` 向控制台写入一条消息。 使用的区域是包含工作表中分配了值或格式的任意单元格的最小区域。 如果整个工作表为空， `getUsedRange()` 该方法将返回仅由左上单元格组成的区域。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a>获取整个区域

下面的代码示例从名为 **Sample** 的工作表获取整个工作表区域，加载其 属性，然后向控制台写入 `address` 一条消息。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API Excel区域](excel-add-ins-ranges-insert.md)
