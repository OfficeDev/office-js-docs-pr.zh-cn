---
title: 向特定 Excel 范围添加数据验证
description: 了解 Excel JavaScript API 如何允许外接程序向工作簿中的表、列、行和其他区域添加自动数据验证。
ms.date: 03/19/2019
ms.localizationpriority: medium
ms.openlocfilehash: 83f7f21621b6ddffa3cb7e51134a3b4cd1cc2aaa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149530"
---
# <a name="add-data-validation-to-excel-ranges"></a>向特定 Excel 范围添加数据验证

Excel JavaScript 库提供的 API 可支持使用外接程序来向表格、列、行及工作簿中的其他范围添加自动数据验证。 若要了解数据验证的概念和术语，请参阅以下文章，这些文章介绍了用户如何通过 Excel UI 添加数据验证。

- [向单元格应用数据验证](https://support.microsoft.com/office/29fecbcc-d1b9-42c1-9d76-eff3ce5f7249)
- [有关数据验证的更多信息](https://support.microsoft.com/office/f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [Excel 中的数据验证的说明和示例](https://support.microsoft.com/help/211485)

## <a name="programmatic-control-of-data-validation"></a>数据验证的编程控制

`Range.dataValidation` 属性（使用 [DataValidation](/javascript/api/excel/excel.datavalidation) 对象）是在 Excel 中对数据验证进行编程控制的切入点。 `DataValidation` 对象有到五个属性：

- `rule` &#8212; 为相应范围定义构成有效数据的条件。 请参阅 [DataValidationRule](/javascript/api/excel/excel.datavalidationrule)。
- `errorAlert` &#8212; 指定如果用户输入无效数据，是否弹出错误，并定义警报文本、标题和样式，例如，**Informational**、**Warning** 和 **Stop**。 请参阅 [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)。
- `prompt` &#8212; 指定当用户将鼠标悬停在相应范围上时是否显示提示语并定义提示语消息。 请参阅 [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)。
- `ignoreBlanks` &#8212; 指定数据验证规则是否应用于相应范围内的空白单元格。 默认为 `true`。
- `type` &#8212; 验证类型的只读标识，例如 WholeNumber、Date、TextLength 等。在设置 `rule` 属性时会间接设置该属性。

> [!NOTE]
> 以编程方式添加的数据验证与手动添加的数据验证的行为方式类似。 特别要注意的是，只有当用户直接将值输入单元格或从工作簿的其他地方复制单元格并选择“**数值**”粘贴选项来进行粘贴时，才会触发数据验证。 如果用户复制单元格并将纯文本粘贴到具有数据验证的范围内，则不会触发验证。

## <a name="creating-validation-rules"></a>创建验证规则

若要为某个范围添加数据验证，你的代码必须在 `Range.dataValidation` 中设置 `DataValidation` 对象的 `rule` 属性。 这会用到 [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) 对象，该对象具有七个可选属性。 *任何 `DataValidationRule` 对象中都最多只能有一个上述属性。* 其中包括的属性将决定验证的类型。

### <a name="basic-and-datetime-validation-rule-types"></a>基本和日期/时间验证规则类型

前三个 `DataValidationRule` 属性（即验证规则类型）会取 [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) 对象作为其值。

- `wholeNumber` &#8212; 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个整数。
- `decimal` &#8212; 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个小数。
- `textLength` &#8212; 将 `BasicDataValidation` 对象中的验证细节应用于单元格中值的 *长度*。

以下是一个创建验证规则的示例。 对于此代码，请注意以下事项。

- `operator` 是二进制运算符“GreaterThan”。 每当使用二进制运算符时，用户试图在单元格中输入的值是左操作数，`formula1` 中指定的值是右操作数。 所以这个规则的含义是，只有大于 0 的整数才是有效的。
- `formula1` 是一个硬编码数字。 如果在编码时不知道该值是什么，也可以使用 Excel 公式（作为字符串）来计算该值。 例如，“=A3”和“=SUM(A4,B5)”也可以是 `formula1` 的值。

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

如需其他二进制运算符的列表，请参阅 [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)。 

此外，还有两个三元运算符：“Between”和“NotBetween”。 若要使用这些运算符，则必须指定可选的 `formula2` 属性。 `formula1` 和 `formula2` 值为边界操作数。 用户试图在单元格中输入的值是第三个（被评估）操作数。 下面是使用"Between"运算符的示例。

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

接下来的两个规则属性均取 [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) 对象作为其值。

- `date`
- `time`

`DateTimeDataValidation` 对象的结构与 `BasicDataValidation` 类似：具有属性 `formula1`、`formula2` 和 `operator`，而且使用方式相同。 不同之处在于，你不能在公式属性中使用数字，但是可以输入 [ISO 8606 日期/时间](https://www.iso.org/iso-8601-date-and-time-format.html)字符串（或 Excel 公式）。 下面的示例将有效值定义为 2018 年 4 月第一周内的日期。 

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

### <a name="list-validation-rule-type"></a>列表验证规则类型

使用 `DataValidationRule` 对象中的 `list` 属性来指定只有来自某个有限列表的值才是有效值。 示例如下。 对于此代码，请注意以下事项。

- 它假定有一个名为“"Names”的工作表，且范围“A1:A3”内的值均为姓名。
- `source` 属性指定一个有效值列表。 该字符串参数会指向一个包含姓名的范围。 此外，也可以分配一个逗号分隔的列表；例如：“Sue, Ricky, Liz”。
- `inCellDropDown` 属性指定当用户选择某个单元格时是否在该单元格中显示下拉控件。 如果设置为 `true`，则显示下拉控件并附带来自 `source` 的值的列表。

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a>自定义验证规则类型

使用 `DataValidationRule` 对象中的 `custom` 属性来指定自定义验证公式。 示例如下。 对于此代码，请注意以下事项。

- 它假定有一个两列的表格，**Athlete Name** 和 **Comments** 列分别为工作表中的 A 列和 B 列。
- 为了缩短 **Comments** 列的长度，它使包含运动员姓名的数据变为无效。
- `SEARCH(A2,B2)` 返回 A2 中字符串在 B2 中字符串的起始位置。 如果 A2 不包含在 B2 中，则不返回数字。 `ISNUMBER()` 会返回布尔值。 因此，`formula` 属性表示的是，**Comment** 列的有效数据是不包含 **Athlete Name** 列中的字符串的数据。

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

## <a name="create-validation-error-alerts"></a>创建验证错误警报

你可以创建一个在用户试图在单元格中输入无效数据时显示的自定义错误警报。 下面展示了一个非常简单的示例。 对于此代码，请注意以下事项。

- `style` 属性决定用户是会收到信息警报、警告还是“停止”警报。 实际上，只有 `Stop` 会阻止用户添加无效数据。 `Warning` 和 `Information` 弹出窗口都具有允许用户输入无效数据的选项。
- `showAlert` 属性默认为 `true`。 这意味着Excel将弹出一个常规警报 (类型为) ，除非您创建自定义警报来设置或设置自定义消息、标题和 `Stop` `showAlert` `false` 样式。 以下代码设置了自定义消息和标题。

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

有关详细信息，请参阅 [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)。

## <a name="create-validation-prompts"></a>创建验证提示语

你可以创建一个在以下情况下显示的说明性提示语：当用户将鼠标悬停在某个应用了数据验证的单元格上或选择该单元格时。 示例如下。

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

有关详细信息，请参阅 [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)。

## <a name="remove-data-validation-from-a-range"></a>从某个范围删除数据验证

若要从某个范围删除数据验证，请调用 [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear__) 方法。

```js
myrange.dataValidation.clear()
```

清除数据验证的范围不必与当初添加数据验证的范围完全相同。 如果二者不相同，则仅清除二者中重叠的单元格（如果存在）。 

> [!NOTE]
> 从某个范围清除数据验证还会清除用户手动添加至该范围的任何数据验证。

## <a name="see-also"></a>另请参阅

- [Excel 加载项中的 Word JavaScript 对象模型](excel-add-ins-core-concepts.md)
- [DataValidation 对象（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.datavalidation)
- [Range 对象（适用于 Excel 的 JavaScript API）](/javascript/api/excel/excel.range)
