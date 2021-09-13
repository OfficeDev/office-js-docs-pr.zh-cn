---
title: 使用 JavaScript API Excel字符串
description: 了解如何使用 JavaScript API 查找Excel字符串。
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: c143acdfb94928b3c59e4fa92eab41ca635f021a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152277"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>使用 JavaScript API 查找Excel字符串

本文提供了一个代码示例，该示例使用 JavaScript API 查找Excel字符串。 有关对象支持的属性和方法的完整 `Range` 列表，请参阅[Excel。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>匹配范围内的字符串

`Range` 对象具有 `find` 方法在区域内搜索指定字符串。 返回有匹配文本的第一个单元格区域。

以下代码示例查找值等于字符串 **食品** 的第一个单元格，并将其地址记录到控制台。 请注意，若指定的字符串不存在于区域中，`find` 将引发 `ItemNotFound` 错误。 若您预计到指定的字符串可能不存在区域中，则可使用 [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) 方法，以便您的代码可正常处理该情况。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

在表示一个单元格的区域调用 `find` 方法时，将在整个工作表进行搜索。 搜索开始于该单元格，并按照 `SearchCriteria.searchDirection` 指定的方向进行，如有需要在工作表结束的地方换行。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API 查找Excel单元格](excel-add-ins-ranges-special-cells.md)
