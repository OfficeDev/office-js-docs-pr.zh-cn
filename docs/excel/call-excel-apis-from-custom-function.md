---
title: 从自定义函数调用 Microsoft Excel Api
description: 了解可以从自定义函数调用的 Microsoft Excel Api。
ms.date: 02/06/2020
localization_priority: Normal
ms.openlocfilehash: 2f24f8fc27db65466cb586307d7f4bc8f8eefe20
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284118"
---
# <a name="call-microsoft-excel-apis-from-a-custom-function"></a>从自定义函数调用 Microsoft Excel Api

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

从自定义函数中调用 node.js Excel Api，以获取范围数据并获取更多用于计算的上下文。

在以下情况中，通过自定义函数调用 node.js Api 可能很有用：

- 自定义函数需要在计算之前从 Excel 中获取信息。 此信息可能包括文档属性、范围格式、自定义 XML 部件、工作簿名称或其他特定于 Excel 的信息。
- 自定义函数将在计算后设置单元格的返回值的数字格式。

[!include[Excel shared runtime note](../includes/note-requires-shared-runtime.md)]

## <a name="code-sample"></a>代码示例

若要调入到 node.js Api，首先需要一个上下文。 使用`Excel.RequestContext`对象获取上下文。 然后，使用上下文调用工作簿中所需的 Api。

下面的代码示例演示如何从工作簿中获取值的范围。

```JavaScript
/**
 * @customfunction
 * @param address range's address
 **/
async function getRangeValue (address) {
 var context = new Excel.RequestContext();
 var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
 range.load();
 await context.sync();
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-officejs-through-a-custom-function"></a>通过自定义函数调用 node.js 的限制

请勿从更改 Excel 环境的自定义函数中调用 node.js Api。 这意味着您的自定义函数不应执行以下任一操作：

- 插入、删除或格式化电子表格中的单元格。
- 更改其他单元格的值。
- 将工作表移动、重命名、删除或添加到工作簿中。
- 更改任何环境选项，如计算模式或屏幕视图。
- 将名称添加到工作簿中。
- 设置属性或执行大多数方法。

更改 Excel 可能导致性能下降、超时和无限循环。 在 Excel 重新计算发生时，不应运行自定义函数计算，因为这会导致不可预测的结果。

而是在功能区按钮或任务窗格的上下文中对 Excel 进行更改。

## <a name="next-steps"></a>后续步骤

- [Excel JavaScript API 基本编程概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>另请参阅

- [在 Excel 自定义函数和任务窗格教程之间共享数据和事件教程](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)