---
title: 从自定义函数调用 Excel JavaScript API
description: 了解可以从自定义函数调用的 Excel JavaScript API。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: aa22cb007bb4803863c17e0f72876cc58c15b992
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423186"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>从自定义函数调用 Excel JavaScript API

从自定义函数调用 Excel JavaScript API 以获取范围数据并获取更多计算上下文。 在以下情况下，通过自定义函数调用 Excel JavaScript API 会很有帮助：

- 自定义函数需要在计算前从 Excel 获取信息。 此信息可能包括文档属性、范围格式、自定义 XML 部件、工作簿名称或其他特定于 Excel 的信息。
- 自定义函数将在计算后为返回值设置单元格的数字格式。

> [!IMPORTANT]
> 若要从自定义函数调用 Excel JavaScript API，需要使用 [共享运行时](../testing/runtimes.md#shared-runtime)。 请参阅 [“配置 Office 外接程序”以使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md) 了解详细信息。

## <a name="code-sample"></a>代码示例

若要从自定义函数调用 Excel JavaScript API，首先需要上下文。 使用 [Excel.RequestContext](/javascript/api/excel/excel.requestcontext) 对象获取上下文。 然后使用上下文调用工作簿中所需的 API。

下面的代码示例演示如何使用 `Excel.RequestContext` 它从工作簿中的单元格获取值。 在此示例中 `address` ，参数将传递到 Excel JavaScript API [Worksheet.getRange 方法中](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1)) ，并且必须以字符串的形式输入。 例如，在 Excel UI 中输入的自定义函数必须遵循模式 `=CONTOSO.GETRANGEVALUE("A1")`，从中检索值的单元格的地址在哪里 `"A1"` 。

```JavaScript
/**
 * @customfunction
 * @param {string} address The address of the cell from which to retrieve the value.
 * @returns The value of the cell at the input address.
 **/
async function getRangeValue(address) {
 // Retrieve the context object. 
 const context = new Excel.RequestContext();
 
 // Use the context object to access the cell at the input address. 
 const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>通过自定义函数调用 Excel JavaScript API 的限制

不要从更改 Excel 环境的自定义函数调用 Excel JavaScript API。 这意味着自定义函数不应执行以下任一操作：

- 在电子表格上插入、删除或设置单元格的格式。
- 更改另一个单元格的值。
- 将工作表移动、重命名、删除或添加到工作簿。
- 更改任何环境选项，例如计算模式或屏幕视图。
- 将名称添加到工作簿。
- 设置属性或执行大多数方法。

更改 Excel 可能会导致性能不佳、超时和无限循环。 执行 Excel 重新计算时不应运行自定义函数计算，因为它将导致不可预知的结果。

而是从功能区按钮或任务窗格的上下文对 Excel 进行更改。

## <a name="next-steps"></a>后续步骤

- [Excel JavaScript API 基本编程概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>另请参阅

- [在 Excel 自定义函数和任务窗格教程之间共享数据和事件](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [将 Office 外接程序配置为使用共享运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
