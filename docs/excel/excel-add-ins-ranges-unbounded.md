---
title: 使用 JavaScript API 读取或写入无限Excel区域
description: 了解如何使用 JavaScript API Excel读取或写入无限区域。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e9df563c9bac688e0806f5280139da8faa5b94870d4d32a852c919cd7fd03ad0
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087647"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>使用 JavaScript API 读取或写入无限Excel区域

本文介绍如何使用 JavaScript API 对无限区域进行Excel写入。 有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。

无限区域地址是指定整列或整行的范围地址。 例如：

- 由整列组成的区域地址：<ul><li>`C:C`</li><li>`A:F`</li></ul>
- 由整行组成的区域地址：<ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a>读取无限区域

API 发出请求以检索无限区域时（例如，`getRange('C:C')`），该响应将包含单元格级别属性（如 `null`、`values`、`text` 和 `numberFormat`）的 `formula` 值。 其他区域属性（如 `address` 和 `cellCount`）将包含无限区域的有效值。

## <a name="write-to-an-unbounded-range"></a>写入一个无限区域

由于输入请求过大，无法在无限区域上设置单元格级属性（如 、 和 `values` `numberFormat` `formula` ）。 例如，下面的代码示例无效，因为它尝试指定 `values` 无限区域。 如果您尝试为无限区域设置单元格级别属性，API 将返回错误。

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
