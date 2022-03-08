---
title: 使用 JavaScript API 读取或写入无限Excel区域
description: 了解如何使用 Excel JavaScript API 读取或写入无限区域。
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 6e9b0c56dfd04cd53e01c41fea23fbf826a6fa14
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340951"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>使用 JavaScript API 读取或写入无限Excel区域

本文介绍如何使用 JavaScript API 对无限区域进行Excel写入。 有关对象支持的属性和方法`Range`的完整列表，请参阅Excel[。Range 类](/javascript/api/excel/excel.range)。

无限区域地址是指定整列或整行的范围地址。 例如：

- 由整列组成的区域地址。
  - `C:C`
  - `A:F`
- 由整行组成的区域地址。
  - `2:2`
  - `1:4`

## <a name="read-an-unbounded-range"></a>读取无限区域

API 发出请求以检索无限区域时（例如，`getRange('C:C')`），该响应将包含单元格级别属性（如 `null`、`values`、`text` 和 `numberFormat`）的 `formula` 值。 其他区域属性（如 `address` 和 `cellCount`）将包含无限区域的有效值。

## <a name="write-to-an-unbounded-range"></a>写入一个无限区域

由于输入请求过大`values``numberFormat``formula`，无法在无限区域上设置单元格级属性（如 、 和 ）。 例如，下面的代码示例无效，因为它 `values` 尝试指定无限区域。 如果您尝试为无限区域设置单元格级属性，API 将返回错误。

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
let range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [使用 JavaScript API Excel单元格](excel-add-ins-cells.md)
- [使用 JavaScript API 读取或写入Excel区域](excel-add-ins-ranges-large.md)
- [ 同时在 Excel 加载项中处理多个区域 ](excel-add-ins-multiple-ranges.md)
