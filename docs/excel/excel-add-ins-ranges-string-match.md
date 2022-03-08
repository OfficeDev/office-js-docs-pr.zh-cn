---
title: 使用 JavaScript API Excel字符串
description: 了解如何使用 JavaScript API 查找Excel字符串。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 042465e01af55bbb3f4325ea44edc27174d558f2
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340986"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>使用 JavaScript API 查找Excel字符串

本文提供了一个代码示例，该示例使用 JavaScript API 查找Excel字符串。 有关对象支持的属性和方法`Range`的完整列表，请参阅Excel[。Range 类](/javascript/api/excel/excel.range)。

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>匹配范围内的字符串

`Range` 对象具有 `find` 方法在区域内搜索指定字符串。 返回有匹配文本的第一个单元格区域。

以下代码示例查找值等于字符串 **食品** 的第一个单元格，并将其地址记录到控制台。 请注意，若指定的字符串不存在于区域中，`find` 将引发 `ItemNotFound` 错误。 若您预计到指定的字符串可能不存在区域中，则可使用 [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties) 方法，以便您的代码可正常处理该情况。

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let table = sheet.tables.getItem("ExpensesTable");
    let searchRange = table.getRange();
    let foundRange = searchRange.find("Food", {
        completeMatch: true, // Match the whole cell value.
        matchCase: false, // Don't match case.
        searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address");
    await context.sync();

    console.log(foundRange.address);
});
```

在表示一个单元格的区域调用 `find` 方法时，将在整个工作表进行搜索。 搜索开始于该单元格，并按照 `SearchCriteria.searchDirection` 指定的方向进行，如有需要在工作表结束的地方换行。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API 查找Excel单元格](excel-add-ins-ranges-special-cells.md)
