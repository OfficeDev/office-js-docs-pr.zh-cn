---
title: 从Excel调用 JavaScript API
description: 了解Excel函数调用哪些 JavaScript API。
ms.date: 08/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7b60f3fbdeb317169800c688b77982580dfbf8c4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744396"
---
# <a name="call-excel-javascript-apis-from-a-custom-function"></a>从Excel调用 JavaScript API

从Excel函数调用 JavaScript API 以获取区域数据，并获取用于计算的更多上下文。 在Excel函数调用 JavaScript API 可能会有所帮助：

- 自定义函数需要在计算之前从Excel信息。 此信息可能包括文档属性、范围格式、自定义 XML 部件、工作簿名称或其他Excel特定的信息。
- 自定义函数将在计算后设置返回值的单元格编号格式。

> [!IMPORTANT]
> 若要从Excel调用 JavaScript API，你需要使用共享的 JavaScript 运行时。 查看 [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md) 以了解更多信息。

## <a name="code-sample"></a>代码示例

若要从Excel调用 JavaScript API，首先需要上下文。 使用[Excel。获取上下文的 RequestContext](/javascript/api/excel/excel.requestcontext) 对象。 然后，使用上下文调用工作簿中所需的 API。

下面的代码示例演示如何使用 从 `Excel.RequestContext` 工作簿的单元格获取值。 在此示例中，参数`address`将传递到 JavaScript API [Worksheet.getRange Excel中](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getrange-member(1))，并且必须以字符串形式输入。 例如，在用户界面中输入的Excel函数`=CONTOSO.GETRANGEVALUE("A1")``"A1"`必须遵循 模式，其中 是从中检索值的单元格的地址。

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
 range.load("values");
 await context.sync();
 
 // Return the value of the cell at the input address.
 return range.values[0][0];
}
```

## <a name="limitations-of-calling-excel-javascript-apis-through-a-custom-function"></a>通过自定义函数Excel JavaScript API 的限制

不要从更改Excel环境自定义函数调用 JavaScript Excel。 这意味着自定义函数不应执行下列任何操作：

- 在电子表格中插入、删除或设置单元格的格式。
- 更改另一个单元格的值。
- 移动、重命名、删除工作表或向工作簿添加工作表。
- 更改任何环境选项，如计算模式或屏幕视图。
- 向工作簿添加名称。
- 设置属性或执行大多数方法。

更改Excel可能会导致性能不佳、时间不足和无限循环。 自定义函数计算不应在重新计算Excel运行，因为它会导致不可预知的结果。

相反，请Excel功能区按钮或任务窗格的上下文中对自定义项进行更改。

## <a name="next-steps"></a>后续步骤

- [Excel JavaScript API 基本编程概念](../reference/overview/excel-add-ins-reference-overview.md)

## <a name="see-also"></a>另请参阅

- [在自定义函数和任务Excel之间共享数据和事件教程](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
