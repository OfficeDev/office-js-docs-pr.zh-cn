---
title: 将数据验证添加到 Excel 范围
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 9e3aba8d87e84405bb3e1ae35a8d35d60ce8e2b6
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459152"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="504a1-102">将数据验证添加到 Excel 范围</span><span class="sxs-lookup"><span data-stu-id="504a1-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="504a1-p101">Excel JavaScript 库提供的 API 可让外接程序将自动数据验证添加到工作簿中的表、列、行和其他范围。 要了解数据验证的概念和术语，请参阅以下关于用户如何通过 Excel UI 添加数据验证的文章：</span><span class="sxs-lookup"><span data-stu-id="504a1-p101">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook. To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="504a1-105">将数据验证应用于单元格</span><span class="sxs-lookup"><span data-stu-id="504a1-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="504a1-106">有关数据验证的更多信息</span><span class="sxs-lookup"><span data-stu-id="504a1-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="504a1-107">Excel 中数据验证的说明和示例</span><span class="sxs-lookup"><span data-stu-id="504a1-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="504a1-108">数据验证的程序控制</span><span class="sxs-lookup"><span data-stu-id="504a1-108">Programmatic control of data validation</span></span>

<span data-ttu-id="504a1-p102">属性`Range.dataValidation` 需要一个 [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) 对象，是 Excel 中数据验证的编程控制入口点。 `DataValidation` 对象有五个属性：</span><span class="sxs-lookup"><span data-stu-id="504a1-p102">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel. There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="504a1-p103">`rule` — 定义构成范围的有效数据。请参阅[DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule)。</span><span class="sxs-lookup"><span data-stu-id="504a1-p103">`rule` &#8212; Defines what constitutes valid data for the range. See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="504a1-p104">`errorAlert` — 指定用户输入无效数据时是否弹出错误，并定义警报文本、标题和样式；例如，**信息**、**警告**和**停止**。请参阅[DataValidationErrorAlert ](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="504a1-p104">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**. See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="504a1-p105">`prompt` — 指定当用户将光标悬停在范围上时是否显示提示并且定义提示消息。请参阅[DataValidationPrompt ](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="504a1-p105">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message. See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="504a1-p106">`ignoreBlanks`  — 指定数据验证规则是否适用于范围内的空白单元格。默认为`true`。</span><span class="sxs-lookup"><span data-stu-id="504a1-p106">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range. Defaults to `true`.</span></span>
- <span data-ttu-id="504a1-119">`type` — 验证类型的只读标识，例如 WholeNumber、Date、TextLength 等。在设置 `rule` 属性时间接设置。</span><span class="sxs-lookup"><span data-stu-id="504a1-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="504a1-p107">通过编程添加的数据验证其工作方式等同于手动添加的数据验证。尤其要注意的是，仅当用户直接将值输入到单元格中或从工作簿中的其他位置复制和粘贴单元格并选择**值**粘贴选项时，才会触发数据验证。如果用户复制一个单元格并将其粘贴到包含数据验证的范围中，则不会触发验证。</span><span class="sxs-lookup"><span data-stu-id="504a1-p107">Data validation added programmatically behaves just like manually added data validation. In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option. If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="504a1-123">创建验证规则</span><span class="sxs-lookup"><span data-stu-id="504a1-123">Creating validation rules</span></span>

<span data-ttu-id="504a1-p108">要将数据验证添加到范围，代码必须在 `Range.dataValidation` 中设置 `DataValidation` 对象的 `rule` 属性。 这需要一个具有七个可选属性的[DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) 对象。在任何 `DataValidationRule`  对象中，*这些属性不得出现一个以上。* 所包含的属性决定了验证的类型。 </span><span class="sxs-lookup"><span data-stu-id="504a1-p108">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`. This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties. *No more than one of these properties may be present in any `DataValidationRule` object.* The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="504a1-128">Basic 和 DateTime 验证规则类型</span><span class="sxs-lookup"><span data-stu-id="504a1-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="504a1-129">前三个 `DataValidationRule` 属性（即验证规则类型）需要一个 [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) 对象作为它们的值。</span><span class="sxs-lookup"><span data-stu-id="504a1-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="504a1-130">`wholeNumber` — 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个整数。</span><span class="sxs-lookup"><span data-stu-id="504a1-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="504a1-131">`decimal` — 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个十进制数。</span><span class="sxs-lookup"><span data-stu-id="504a1-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="504a1-132">`textLength` — 在 `BasicDataValidation` 对象中针对单元格值的*长度*应用验证细节。</span><span class="sxs-lookup"><span data-stu-id="504a1-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="504a1-p109">下面是创建验证规则的一个示例。 关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="504a1-p109">Here is an example of creating a validation rule. Note the following about this code:</span></span>

- <span data-ttu-id="504a1-p110"> `operator` 是二元运算符"GreaterThan"。无论何时使用二元运算符，用户尝试输入到单元格的值都是左侧的操作数，并且在 `formula1`  中指定的值是右侧操作数。所以这条规则说只有大于 0 的整数才有效。</span><span class="sxs-lookup"><span data-stu-id="504a1-p110">The `operator` is the binary operator "GreaterThan". Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand. So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="504a1-p111"> `formula1` 是一个硬编码的数字。如果在写代码时不知道该值应该是多少，还可以使用 Excel 公式（作为字符串）来计算该值。例如，`formula1` 的值也可以是 "= A3" 和 "= SUM(A4,B5)"。</span><span class="sxs-lookup"><span data-stu-id="504a1-p111">The `formula1` is a hard-coded number. If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value. For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="504a1-141">有关其他二元运算符的列表，请参阅 [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation)。</span><span class="sxs-lookup"><span data-stu-id="504a1-141">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="504a1-p112">还有两个三元运算符："Between" 和 "NotBetween"。 要使用这些运算符，必须指定可选的  `formula2` 属性。`formula1` 和 `formula2` 值是边界操作数。用户尝试输入单元格的值是第三个（评估）的操作数。 下面是使用 "Between" 运算符的一个示例：</span><span class="sxs-lookup"><span data-stu-id="504a1-p112">There are also two ternary operators: "Between" and "NotBetween". To use these, you must specify the optional `formula2` property. The `formula1` and `formula2` values are the bounding operands. The value that the user tries to enter in the cell is the third (evaluated) operand. The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="504a1-147">接下来的两个规则属性需要一个 [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) 对象作为它们的值。</span><span class="sxs-lookup"><span data-stu-id="504a1-147">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="504a1-p113"> `DateTimeDataValidation` 对象的结构类似于`BasicDataValidation\`：它具有属性 `formula1\`、`formula2` 和 `operator\`，并以相同的方式使用。区别在于，不能在公式属性中使用数字，但可以输入一个 [ISO 8606 日期时间](https://www.iso.org/iso-8601-date-and-time-format.html) 字符串 （或 Excel 公式）。下面是将日期有效值定义为 2018 年 4 月第一周的示例。</span><span class="sxs-lookup"><span data-stu-id="504a1-p113">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way. The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula). The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="504a1-151">列表验证规则类型</span><span class="sxs-lookup"><span data-stu-id="504a1-151">List validation rule type</span></span>

<span data-ttu-id="504a1-p114">使用 `DataValidationRule` 对象的 `list` 属性来指定唯一有效值是来自有限列表中的那些值。示例如下，关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="504a1-p114">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="504a1-155">它假定有一个名为 "Names" 的工作表，并且 "A1:A3" 范围中的值是名称。</span><span class="sxs-lookup"><span data-stu-id="504a1-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="504a1-p115"> `source` 属性指定有效值的列表。已给它分配具有名称的范围。也可分配以逗号分隔的列表，例如："Sue, Ricky, Liz"。</span><span class="sxs-lookup"><span data-stu-id="504a1-p115">The `source` property specifies the list of valid values. The range with the names has been assigned to it. You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="504a1-p116"> `inCellDropDown` 属性指定用户选择单元格时是否出现下拉控件。如果设置为 `true\`，则显示带有来自 `source` 之值的下拉列表。</span><span class="sxs-lookup"><span data-stu-id="504a1-p116">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it. If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="504a1-161">自定义验证规则类型</span><span class="sxs-lookup"><span data-stu-id="504a1-161">Custom validation rule type</span></span>

<span data-ttu-id="504a1-p117">使用 `DataValidationRule` 对象的 `custom` 属性来指定自定义验证公式。示例如下，关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="504a1-p117">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="504a1-165">它假定有一个包含 **Athlete Name** 和 **Comments** 两列的表，它们分别位于工作表的 A 和 B 列。</span><span class="sxs-lookup"><span data-stu-id="504a1-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="504a1-166">为了减少 **Comment** 列中的冗余，它会认定包含运动员姓名的数据无效。</span><span class="sxs-lookup"><span data-stu-id="504a1-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="504a1-p118">`SEARCH(A2,B2)` 返回 A2 字符串在 B2 字符串中的起始位置。 如果 A2 不包含在 B2 中，它不会返回一个数字。 `ISNUMBER()` 返回一个布尔值。因此， `formula` 属性表示，**Comment** 列的有效数据是不包含字符串 **Athlete Name** 列字符串的数据。</span><span class="sxs-lookup"><span data-stu-id="504a1-p118">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2. If A2 is not contained in B2, it does not return a number. `ISNUMBER()` returns a boolean. So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="504a1-171">创建验证错误警报</span><span class="sxs-lookup"><span data-stu-id="504a1-171">Create validation error alerts</span></span>

<span data-ttu-id="504a1-p119">你可以创建用户尝试在单元格中输入无效数据时出现的自定义错误警报。下面展示了一个简单的示例。 关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="504a1-p119">You can a create custom error alert that appears when a user tries to enter invalid data in a cell. The following is a simple example. Note the following about this code:</span></span>

- <span data-ttu-id="504a1-p120"> `style` 属性确定用户是否收到信息提示、警告或“停止”警报。只有 `Stop` 能实际防止用户添加无效数据。 `Warning` 和 `Information` 的弹出窗口具有允许用户仍然输入无效数据的选项。</span><span class="sxs-lookup"><span data-stu-id="504a1-p120">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert. Only `Stop` actually prevents the user from adding invalid data. The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="504a1-p121"> `showAlert` 属性默认为 `true\`。这意味着 Excel 主机将弹出一个通用警报（类型为 `Stop\`），除非创建一个自定义警报将 `showAlert` 设置为 `false` 或设置自定义消息、标题和样式。此代码将设置自定义消息和标题。</span><span class="sxs-lookup"><span data-stu-id="504a1-p121">The `showAlert` property defaults to `true`. This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style. This code sets a custom message and title.</span></span>


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

<span data-ttu-id="504a1-181">有关更多信息，请参阅 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="504a1-181">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="504a1-182">创建验证提示</span><span class="sxs-lookup"><span data-stu-id="504a1-182">Create validation prompts</span></span>

<span data-ttu-id="504a1-p122">可以创建一个在用户悬停指针或选择应用了数据验证的单元格时出现的指示性提示。下面是一个示例：</span><span class="sxs-lookup"><span data-stu-id="504a1-p122">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied. The following is an example:</span></span>

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

<span data-ttu-id="504a1-185">有关更多信息，请参阅 [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="504a1-185">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="504a1-186">从范围中删除数据验证</span><span class="sxs-lookup"><span data-stu-id="504a1-186">Remove data validation from a range</span></span>

<span data-ttu-id="504a1-187">要从范围中删除数据验证，请调用 [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) 方法。</span><span class="sxs-lookup"><span data-stu-id="504a1-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="504a1-p123">清除的范围与添加数据验证的范围不一定需要完全相同。 如果两者不相同，则只清除两个范围中重叠的单元格（如果有的话）。</span><span class="sxs-lookup"><span data-stu-id="504a1-p123">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation. If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="504a1-190">清除范围内的数据验证也将清除用户手动添加到范围内的任何数据验证。</span><span class="sxs-lookup"><span data-stu-id="504a1-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="504a1-191">另请参阅</span><span class="sxs-lookup"><span data-stu-id="504a1-191">See also</span></span>

- [<span data-ttu-id="504a1-192">使用 Excel JavaScript API 的基本编程概念</span><span class="sxs-lookup"><span data-stu-id="504a1-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="504a1-193">DataValidation 对象 (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="504a1-193">Chart Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="504a1-194">Range 对象 (JavaScript API for Excel)</span><span class="sxs-lookup"><span data-stu-id="504a1-194">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
