---
title: 将数据验证添加到 Excel 范围
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: fd40cab045da0472a060752651a27f0b26028b4b
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944875"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a>将数据验证添加到 Excel 范围（预览）

> [!NOTE]
> 虽然数据验证 API 处于预览状态，但你必须加载 Office JavaScript 库的 Beta 版才能使用它们。 URL 是 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 。 如果正在使用 TypeScript，或者代码编辑器使用 TypeScript 类型定义文件实现智能感知，请使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts 。

> [!NOTE]
> 虽然数据验证API处于预览状态，但本文中 API 引用的链接将不起作用。 在此期间，你可以使用 [草稿 Excel API 引用](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel)。

Excel JavaScript 库提供的 API 可让外接程序将自动数据验证添加到工作簿中的表、列、行和其他范围。 要了解数据验证的概念和术语，请参阅以下关于用户如何通过 Excel UI 添加数据验证的文章：

- [将数据验证应用于单元格](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [有关数据验证的更多信息](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Excel 中数据验证的说明和示例](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a>数据验证的程序控制

属性需要一个 [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) 对象，是 Excel 中数据验证的编程控制入口点。`Range.dataValidation` 对象有五个属性：`DataValidation`

- `rule` - 定义构成范围的有效数据。 请参阅 [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule)。
- `errorAlert` - 指定用户输入无效数据时是否弹出错误，并定义警报文本，标题和样式;例如，**信息**、**警告**和**停止**。 请参阅 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)。
- `prompt` - 指定当用户将光标悬停在范围上时是否显示提示并且定义提示消息。 请参阅 [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)。
- `ignoreBlanks` - 指定数据验证规则是否适用于范围内的空白单元格。 默认为 `true` 。
- `type` - 验证类型的只读标识，例如 WholeNumber、Date、TextLength 等。在设置 `rule` 属性时间接设置。

> [!NOTE]
> 通过编程添加的数据验证其工作方式等同于手动添加的数据验证。 尤其要注意的是，仅当用户直接将值输入到单元格中或从工作簿中的其他位置复制和粘贴单元格并选择**值**粘贴选项时，才会触发数据验证。 如果用户复制一个单元格并将其粘贴到包含数据验证的范围中，则不会触发验证。

### <a name="creating-validation-rules"></a>创建验证规则

要将数据验证添加到范围，代码必须在 `Range.dataValidation` 中设置 `DataValidation` 对象的 `rule` 属性。 这需要一个 [DataValidationRule](https://docs.microsoft.com/javascript/api/excel?view=office-js) 对象， 它具有七个可选属性。 *任何 `DataValidationRule` 对象中的这些属性均不得超过一个。* 所包含的属性决定了验证的类型。

#### <a name="basic-and-datetime-validation-rule-types"></a>Basic 和 DateTime 验证规则类型

前三个 `DataValidationRule` 属性（即验证规则类型）需要一个 [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel?view=office-js) 对象作为它们的值。

- `wholeNumber` - 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个整数。
- `decimal` - 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个十进制数。
- `textLength` - 在 `BasicDataValidation` 对象中针对单元格值的*长度*应用验证细节。

以下是创建验证规则的示例。 关于此代码，请注意以下几点：

- 是二元运算符 "GreaterThan"。`operator` 无论何时使用二元运算符，用户尝试输入到单元格的值都是左侧的操作数，并且在 `formula1` 中指定的值是右侧操作数。 所以这条规则说只有大于 0 的整数才有效。 
- 是一个硬编码的数字。`formula1` 如果在写代码时不知道该值应该是多少，还可以使用 Excel 公式（作为字符串）来计算该值。 例如，`formula1` 的值也可以是 "= A3" 和 "= SUM(A4,B5)"。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: "GreaterThan"
            }
        };

    return context.sync();
})
```

有关其他二元运算符的列表，请参阅 [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation)。 

还有两个三元运算符："Between" 和 "NotBetween"。 要使用这些运算符，必须指定可选的 `formula2` 属性。 和 `formula2` 值是边界操作数。`formula1` 用户尝试输入单元格的值是第三个（评估）的操作数。 以下是使用 "Between" 运算符的示例：

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
                operator: "Between"
            }
        };

    return context.sync();
})
```

接下来的两个规则属性需要一个 [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) 对象作为它们的值。

- `date`
- `time`

对象的结构类似于 `BasicDataValidation`：它有属性 `formula1`、`formula2` 和 `operator`，并以相同的方式使用。`DateTimeDataValidation` 不同之处在于，不能在公式属性中使用数字，但可以输入一个 [ISO 8606 日期时间](https://www.iso.org/iso-8601-date-and-time-format.html)字符串（或 Excel 公式）。 以下是将有效值定义为 2018 年 4 月第一周日期的示例。 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.rule = {
            date: {
                formula1: "2018-04-01",
                formula2: "2018-04-08",
                operator: "Between"
            }
        };

    return context.sync();
})
```

#### <a name="list-validation-rule-type"></a>列表验证规则类型

使用 `DataValidationRule` 对象的 `list` 属性来指定唯一有效值是来自有限列表中的那些值。 示例如下。 关于此代码，请注意以下几点：

- 它假定有一个名为 "Names" 的工作表，并且 "A1:A3" 范围中的值是名称。
- 属性指定有效值的列表。`source` 具有名称的范围已分配给它。 也可分配以逗号分隔的列表，例如："Sue, Ricky, Liz"。 
- 属性指定用户选择单元格时是否出现下拉控件。`inCellDropDown` 如果设置为 `true`，将显示下拉列表，其中带有来自 `source` 中的值。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: nameSourceRange
        }
    };

    return context.sync();
})
```

#### <a name="custom-validation-rule-type"></a>自定义验证规则类型

使用 `DataValidationRule` 对象中的 `custom` 属性来指定自定义验证公式。 示例如下。 关于此代码，请注意以下几点：

- 它假定有一个包含 **Athlete Name** 和 **Comments** 两列的表，它们分别位于工作表的 A 和 B 列。
- 为了减少 **Comment** 列中的冗余，它会认定包含运动员姓名的数据无效。
- `SEARCH(A2,B2)` 返回 A2 字符串在 B2 字符串中的起始位置。 如果 A2 不包含在 B2 中，它不会返回一个数字。 `ISNUMBER()` 返回一个布尔值。 所以 `formula` 属性表示，**Comment** 列的有效数据是不包含 **Athlete Name** 列字符串的数据。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    return context.sync();
})
```

### <a name="create-validation-error-alerts"></a>创建验证错误警报

你可以创建用户尝试在单元格中输入无效数据时出现的自定义错误警报。 下面展示了一个简单的示例。 关于此代码，请注意以下几点：

- 属性确定用户是否收到信息提示、警告或“停止”警报。`style` 只有 `Stop` 可以实际防止用户添加无效数据。 和 `Information` 的弹出窗口具有允许用户仍然输入无效数据的选项。`Warning`
- 属性默认为 `true`。`showAlert` 这意味着 Excel 主机将弹出一个通用警报（类型为 `Stop`），除非你通过创建自定义警报将 `showAlert` 设置为 `false` 或设置自定义消息、标题和样式。 此代码设置自定义消息和标题。


```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // default is 'true'
            style: "Stop", // other possible values: Warning, Information
            title: "Negative or Decimal Number Entered"
        };
    
    // Set range.dataValidation.rule and optionally .prompt here.

    return context.sync();
})
```

有关更多信息，请参阅 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert) 。

### <a name="create-validation-prompts"></a>创建验证提示

可以创建一个在用户悬停指针或选择应用了数据验证的单元格时出现的指示性提示。 示例如下：

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
   
    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // default is 'false'
            title: "Positive Whole Numbers Only."
        };
    
    // Set range.dataValidation.rule and optionally .errorAlert here.

    return context.sync();
})
```

有关更多信息，请参阅 [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt) 。

### <a name="remove-data-validation-from-a-range"></a>从范围中删除数据验证

要从范围中删除数据验证，请调用 [Range.dataValidation.clear（）](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) 方法。

```js
myrange.dataValidation.clear()
```

清除的范围与添加数据验证的范围不一定需要完全相同。 如果两者不相同，则只清除两个范围中重叠的单元格（如果有的话）。 

> [!NOTE]
> 清除范围内的数据验证也将清除用户手动添加到范围内的任何数据验证。

## <a name="see-also"></a>另请参阅

- [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)
- [DataValidation 对象 (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [Range 对象 (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
