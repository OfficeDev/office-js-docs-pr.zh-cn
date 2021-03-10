---
title: 从自定义函数调用 Excel JavaScript API
description: 了解可以从自定义函数调用的 Excel JavaScript API。
ms.date: 03/05/2021
localization_priority: Normal
ms.openlocfilehash: 4be1b1ee8ea4ae8b2f5d1d27195be18f7aa841da
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613904"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>从自定义函数调用 Excel JavaScript API

从自定义函数调用 Excel JavaScript API 以获取区域数据，并获取更多计算上下文。 通过自定义函数调用 Excel JavaScript API 在：

- 自定义函数需要在计算之前从 Excel 获取信息。 此信息可能包括文档属性、范围格式、自定义 XML 部件、工作簿名称或其他特定于 Excel 的信息。
- 自定义函数将在计算后设置返回值的单元格编号格式。

> [!IMPORTANT]
> 若要从自定义函数调用 Excel JavaScript API，你需要使用共享的 JavaScript 运行时。 查看 [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md) 以了解更多信息。

## <a name="code-sample"></a>代码示例

若要从自定义函数调用 Excel JavaScript API，首先需要上下文。 使用 [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) 对象获取上下文。 然后使用上下文调用工作簿中所需的 API。

下面的代码示例演示如何用于从工作簿 `Excel.RequestContext` 中的单元格获取值。 在此示例中， `address` 参数将传递到 Excel JavaScript API [Worksheet.getRange](/javascript/api/excel/excel.worksheet#getRange_address_) 方法中，并且必须以字符串形式输入。 例如，输入到 Excel UI 中的自定义函数必须遵循模式，其中要检索值的单元格 `=CONTOSO.GETRANGEVALUE("A1")` `"A1"` 的地址。

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 var context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>通过自定义函数调用 Excel JavaScript API 的限制

不要从更改 Excel 环境的自定义函数调用 Excel JavaScript API。 这意味着自定义函数不应执行下列任何操作：

- 在电子表格中插入、删除或设置单元格的格式。
- 更改另一个单元格的值。
- 移动、重命名、删除或向工作簿添加工作表。
- 更改任何环境选项，如计算模式或屏幕视图。
- 向工作簿添加名称。
- 设置属性或执行大多数方法。

更改 Excel 可能会导致性能不佳、时间不足和无限循环。 自定义函数计算不应在 Excel 重新计算时运行，因为它将导致不可预知的结果。

相反，请从功能区按钮或任务窗格的上下文中对 Excel 进行更改。

## <a name="next-steps"></a>后续步骤

- [Excel JavaScript API 基本编程概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>另请参阅

- [在 Excel 自定义函数和任务窗格教程之间共享数据和事件](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
