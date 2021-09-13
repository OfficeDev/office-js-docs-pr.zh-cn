---
title: 使用 JavaScript API 读取或写入无限Excel区域
description: 了解如何使用 Excel JavaScript API 读取或写入无限区域。
ms.date: 04/05/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: a7b2a564377d0dab73d4f3ad6d3aacf2219ddeae
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152274"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>使用 JavaScript API 读取或写入无限Excel区域

本文介绍如何使用 JavaScript API 读取和写入无限Excel范围。 有关对象支持的属性和方法的完整 `Range` 列表，请参阅[Excel。Range 类](/javascript/api/excel/excel.range)。

无限区域地址是指定整列或整行的范围地址。 例如：

- 由整列组成的区域地址：<ul><li>`C:C`</li><li>`A:F`</li></ul>
- 由整行组成的区域地址：<ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a>读取无限区域

API 发出请求以检索无限区域时（例如，`getRange('C:C')`），该响应将包含单元格级别属性（如 `null`、`values`、`text` 和 `numberFormat`）的 `formula` 值。 其他区域属性（如 `address` 和 `cellCount`）将包含无限区域的有效值。

## <a name="write-to-an-unbounded-range"></a>写入一个无限区域

由于输入请求过大，无法在无限区域上设置单元格级属性（如 、 和 `values` `numberFormat` `formula` ）。 例如，下面的代码示例无效，因为它尝试指定 `values` 无限区域。 如果您尝试为无限区域设置单元格级属性，API 将返回错误。

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API 读取或写入Excel区域](excel-add-ins-ranges-large.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
