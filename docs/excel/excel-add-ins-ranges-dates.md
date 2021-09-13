---
title: 使用 JavaScript API Excel日期
description: 使用 Moment-MSDate JavaScript API Excel插件处理日期。
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 4e09badf3bf6f848f96fc9e5864f383f781ef63f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149266"
---
# <a name="work-with-dates-using-the-excel-javascript-api-and-the-moment-msdate-plug-in"></a>使用 JavaScript API Excel应用程序插件处理Moment-MSDate日期

本文提供的代码示例显示如何使用 javaScript API 和[Excel-MSDate 插件处理日期](https://www.npmjs.com/package/moment-msdate)。 有关对象支持的属性和方法的完整列表， `Range` 请参阅[Excel。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="use-the-moment-msdate-plug-in-to-work-with-dates"></a>使用Moment-MSDate插件处理日期

[时刻 JavaScript 库](https://momentjs.com/)提供了使用日期和时间戳的便捷方式。 [Moment-MSDate 插件](https://www.npmjs.com/package/moment-msdate)可将时刻格式转换为 Excel 所需的格式。 这是 [NOW 函数](https://support.microsoft.com/office/3337fd29-145a-4347-b2e6-20c904739c46)返回的相同格式。

以下代码演示如何将 **B4** 区域设置为时刻的时间戳。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

下面的代码示例演示了一种类似的技术，用于从单元格重新获取日期并将其转换为 `Moment` 或其他格式。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

外接程序必须设置范围的格式，以更可读的形式显示日期。 例如，显示 `"[$-409]m/d/yy h:mm AM/PM;@"` "12/3/18 3：57 PM"。 有关日期和时间数字格式详细信息，请参阅查看自定义数字格式指南文章中的"日期和时间 [格式指南](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) "。


## <a name="see-also"></a>另请参阅

- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
