---
title: 向特定 Excel 范围添加数据验证
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: b0b2d886ceb9026ebe41414fed4ef8be1b59cc95
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449209"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="11f68-102">向特定 Excel 范围添加数据验证</span><span class="sxs-lookup"><span data-stu-id="11f68-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="11f68-103">Excel JavaScript 库提供的 API 可支持使用外接程序来向表格、列、行及工作簿中的其他范围添加自动数据验证。</span><span class="sxs-lookup"><span data-stu-id="11f68-103">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="11f68-104">若要了解数据验证的概念以及术语，请参阅介绍用户如何通过 Excel UI 来添加数据验证的以下文章：</span><span class="sxs-lookup"><span data-stu-id="11f68-104">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="11f68-105">向单元格应用数据验证</span><span class="sxs-lookup"><span data-stu-id="11f68-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="11f68-106">有关数据验证的更多信息</span><span class="sxs-lookup"><span data-stu-id="11f68-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="11f68-107">Excel 中的数据验证的说明和示例</span><span class="sxs-lookup"><span data-stu-id="11f68-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="11f68-108">数据验证的编程控制</span><span class="sxs-lookup"><span data-stu-id="11f68-108">Programmatic control of data validation</span></span>

<span data-ttu-id="11f68-109">`Range.dataValidation` 属性（使用 [DataValidation](/javascript/api/excel/excel.datavalidation) 对象）是在 Excel 中对数据验证进行编程控制的切入点。</span><span class="sxs-lookup"><span data-stu-id="11f68-109">The `Range.dataValidation` property, which takes a [DataValidation](/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="11f68-110">`DataValidation` 对象有到五个属性：</span><span class="sxs-lookup"><span data-stu-id="11f68-110">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="11f68-111">`rule` &#8212; 为相应范围定义构成有效数据的条件。</span><span class="sxs-lookup"><span data-stu-id="11f68-111">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="11f68-112">请参阅 [DataValidationRule](/javascript/api/excel/excel.datavalidationrule)。</span><span class="sxs-lookup"><span data-stu-id="11f68-112">See [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="11f68-113">`errorAlert` &#8212; 指定如果用户输入无效数据，是否弹出错误，并定义警报文本、标题和样式，例如，**Informational**、**Warning** 和 **Stop**。</span><span class="sxs-lookup"><span data-stu-id="11f68-113">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="11f68-114">请参阅 [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="11f68-114">See [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="11f68-115">`prompt` &#8212; 指定当用户将鼠标悬停在相应范围上时是否显示提示语并定义提示语消息。</span><span class="sxs-lookup"><span data-stu-id="11f68-115">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="11f68-116">请参阅 [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="11f68-116">See [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="11f68-117">`ignoreBlanks` &#8212; 指定数据验证规则是否应用于相应范围内的空白单元格。</span><span class="sxs-lookup"><span data-stu-id="11f68-117">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="11f68-118">默认为 `true`。</span><span class="sxs-lookup"><span data-stu-id="11f68-118">Defaults to `true`.</span></span>
- <span data-ttu-id="11f68-119">`type` &#8212; 验证类型的只读标识，例如 WholeNumber、Date、TextLength 等。在设置 `rule` 属性时会间接设置该属性。</span><span class="sxs-lookup"><span data-stu-id="11f68-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="11f68-120">以编程方式添加的数据验证与手动添加的数据验证的行为方式类似。</span><span class="sxs-lookup"><span data-stu-id="11f68-120">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="11f68-121">特别要注意的是，只有当用户直接将值输入单元格或从工作簿的其他地方复制单元格并选择“**数值**”粘贴选项来进行粘贴时，才会触发数据验证。</span><span class="sxs-lookup"><span data-stu-id="11f68-121">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="11f68-122">如果用户复制单元格并将纯文本粘贴到具有数据验证的范围内，则不会触发验证。</span><span class="sxs-lookup"><span data-stu-id="11f68-122">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="11f68-123">创建验证规则</span><span class="sxs-lookup"><span data-stu-id="11f68-123">Creating validation rules</span></span>

<span data-ttu-id="11f68-124">若要为某个范围添加数据验证，你的代码必须在 `Range.dataValidation` 中设置 `DataValidation` 对象的 `rule` 属性。</span><span class="sxs-lookup"><span data-stu-id="11f68-124">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="11f68-125">这会用到 [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) 对象，该对象具有七个可选属性。</span><span class="sxs-lookup"><span data-stu-id="11f68-125">This takes a [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="11f68-126">*任何 `DataValidationRule` 对象中都最多只能有一个上述属性。*</span><span class="sxs-lookup"><span data-stu-id="11f68-126">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="11f68-127">其中包括的属性将决定验证的类型。</span><span class="sxs-lookup"><span data-stu-id="11f68-127">The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="11f68-128">基本和日期/时间验证规则类型</span><span class="sxs-lookup"><span data-stu-id="11f68-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="11f68-129">前三个 `DataValidationRule` 属性（即验证规则类型）会取 [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) 对象作为其值。</span><span class="sxs-lookup"><span data-stu-id="11f68-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="11f68-130">`wholeNumber` &#8212; 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个整数。</span><span class="sxs-lookup"><span data-stu-id="11f68-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="11f68-131">`decimal` &#8212; 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个小数。</span><span class="sxs-lookup"><span data-stu-id="11f68-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="11f68-132">`textLength` &#8212; 将 `BasicDataValidation` 对象中的验证细节应用于单元格中值的*长度*。</span><span class="sxs-lookup"><span data-stu-id="11f68-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="11f68-133">以下是一个创建验证规则的示例。</span><span class="sxs-lookup"><span data-stu-id="11f68-133">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="11f68-134">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="11f68-134">Note the following about this code:</span></span>

- <span data-ttu-id="11f68-135">`operator` 是二进制运算符“GreaterThan”。</span><span class="sxs-lookup"><span data-stu-id="11f68-135">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="11f68-136">每当使用二进制运算符时，用户试图在单元格中输入的值是左操作数，`formula1` 中指定的值是右操作数。</span><span class="sxs-lookup"><span data-stu-id="11f68-136">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="11f68-137">所以这个规则的含义是，只有大于 0 的整数才是有效的。</span><span class="sxs-lookup"><span data-stu-id="11f68-137">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="11f68-138">`formula1` 是一个硬编码数字。</span><span class="sxs-lookup"><span data-stu-id="11f68-138">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="11f68-139">如果在编码时不知道该值是什么，也可以使用 Excel 公式（作为字符串）来计算该值。</span><span class="sxs-lookup"><span data-stu-id="11f68-139">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="11f68-140">例如，“=A3”和“=SUM(A4,B5)”也可以是 `formula1` 的值。</span><span class="sxs-lookup"><span data-stu-id="11f68-140">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="11f68-141">如需其他二进制运算符的列表，请参阅 [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation)。</span><span class="sxs-lookup"><span data-stu-id="11f68-141">See [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="11f68-142">此外，还有两个三元运算符：“Between”和“NotBetween”。</span><span class="sxs-lookup"><span data-stu-id="11f68-142">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="11f68-143">若要使用这些运算符，则必须指定可选的 `formula2` 属性。</span><span class="sxs-lookup"><span data-stu-id="11f68-143">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="11f68-144">`formula1` 和 `formula2` 值为边界操作数。</span><span class="sxs-lookup"><span data-stu-id="11f68-144">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="11f68-145">用户试图在单元格中输入的值是第三个（被评估）操作数。</span><span class="sxs-lookup"><span data-stu-id="11f68-145">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="11f68-146">以下是使用“Between”运算符的示例：</span><span class="sxs-lookup"><span data-stu-id="11f68-146">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="11f68-147">接下来的两个规则属性均取 [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) 对象作为其值。</span><span class="sxs-lookup"><span data-stu-id="11f68-147">The next two rule properties take a [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="11f68-148">`DateTimeDataValidation` 对象的结构与 `BasicDataValidation` 类似：具有属性 `formula1`、`formula2` 和 `operator`，而且使用方式相同。</span><span class="sxs-lookup"><span data-stu-id="11f68-148">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="11f68-149">不同之处在于，你不能在公式属性中使用数字，但是可以输入 [ISO 8606 日期/时间](https://www.iso.org/iso-8601-date-and-time-format.html)字符串（或 Excel 公式）。</span><span class="sxs-lookup"><span data-stu-id="11f68-149">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="11f68-150">下面的示例将有效值定义为 2018 年 4 月第一周内的日期。</span><span class="sxs-lookup"><span data-stu-id="11f68-150">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="11f68-151">列表验证规则类型</span><span class="sxs-lookup"><span data-stu-id="11f68-151">List validation rule type</span></span>

<span data-ttu-id="11f68-152">使用 `DataValidationRule` 对象中的 `list` 属性来指定只有来自某个有限列表的值才是有效值。</span><span class="sxs-lookup"><span data-stu-id="11f68-152">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="11f68-153">示例如下。</span><span class="sxs-lookup"><span data-stu-id="11f68-153">The following is an example.</span></span> <span data-ttu-id="11f68-154">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="11f68-154">Note the following about this code:</span></span>

- <span data-ttu-id="11f68-155">它假定有一个名为“"Names”的工作表，且范围“A1:A3”内的值均为姓名。</span><span class="sxs-lookup"><span data-stu-id="11f68-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="11f68-156">`source` 属性指定一个有效值列表。</span><span class="sxs-lookup"><span data-stu-id="11f68-156">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="11f68-157">该字符串参数会指向一个包含姓名的范围。</span><span class="sxs-lookup"><span data-stu-id="11f68-157">The string argument refers to a range containing the names.</span></span> <span data-ttu-id="11f68-158">此外，也可以分配一个逗号分隔的列表；例如：“Sue, Ricky, Liz”。</span><span class="sxs-lookup"><span data-stu-id="11f68-158">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="11f68-159">`inCellDropDown` 属性指定当用户选择某个单元格时是否在该单元格中显示下拉控件。</span><span class="sxs-lookup"><span data-stu-id="11f68-159">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="11f68-160">如果设置为 `true`，则显示下拉控件并附带来自 `source` 的值的列表。</span><span class="sxs-lookup"><span data-stu-id="11f68-160">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

### <a name="custom-validation-rule-type"></a><span data-ttu-id="11f68-161">自定义验证规则类型</span><span class="sxs-lookup"><span data-stu-id="11f68-161">Custom validation rule type</span></span>

<span data-ttu-id="11f68-162">使用 `DataValidationRule` 对象中的 `custom` 属性来指定自定义验证公式。</span><span class="sxs-lookup"><span data-stu-id="11f68-162">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="11f68-163">示例如下。</span><span class="sxs-lookup"><span data-stu-id="11f68-163">The following is an example.</span></span> <span data-ttu-id="11f68-164">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="11f68-164">Note the following about this code:</span></span>

- <span data-ttu-id="11f68-165">它假定有一个两列的表格，**Athlete Name** 和 **Comments** 列分别为工作表中的 A 列和 B 列。</span><span class="sxs-lookup"><span data-stu-id="11f68-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="11f68-166">为了缩短 **Comments** 列的长度，它使包含运动员姓名的数据变为无效。</span><span class="sxs-lookup"><span data-stu-id="11f68-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="11f68-167">`SEARCH(A2,B2)` 返回 A2 中字符串在 B2 中字符串的起始位置。</span><span class="sxs-lookup"><span data-stu-id="11f68-167">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="11f68-168">如果 A2 不包含在 B2 中，则不返回数字。</span><span class="sxs-lookup"><span data-stu-id="11f68-168">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="11f68-169">`ISNUMBER()` 会返回布尔值。</span><span class="sxs-lookup"><span data-stu-id="11f68-169">`ISNUMBER()` returns a boolean.</span></span> <span data-ttu-id="11f68-170">因此，`formula` 属性表示的是，**Comment** 列的有效数据是不包含 **Athlete Name** 列中的字符串的数据。</span><span class="sxs-lookup"><span data-stu-id="11f68-170">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="11f68-171">创建验证错误警报</span><span class="sxs-lookup"><span data-stu-id="11f68-171">Create validation error alerts</span></span>

<span data-ttu-id="11f68-172">你可以创建一个在用户试图在单元格中输入无效数据时显示的自定义错误警报。</span><span class="sxs-lookup"><span data-stu-id="11f68-172">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="11f68-173">下面展示了一个非常简单的示例。</span><span class="sxs-lookup"><span data-stu-id="11f68-173">The following is a simple example.</span></span> <span data-ttu-id="11f68-174">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="11f68-174">Note the following about this code:</span></span>

- <span data-ttu-id="11f68-175">`style` 属性决定用户是会收到信息警报、警告还是“停止”警报。</span><span class="sxs-lookup"><span data-stu-id="11f68-175">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="11f68-176">实际上，只有 `Stop` 会阻止用户添加无效数据。</span><span class="sxs-lookup"><span data-stu-id="11f68-176">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="11f68-177">`Warning` 和 `Information` 弹出窗口都具有允许用户输入无效数据的选项。</span><span class="sxs-lookup"><span data-stu-id="11f68-177">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="11f68-178">`showAlert` 属性默认为 `true`。</span><span class="sxs-lookup"><span data-stu-id="11f68-178">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="11f68-179">这意味着除非创建自定义警报，在其中将 `showAlert` 设置为 `false` 或者设置自定义消息、标题和样式，否则 Excel 主机将会弹出（类型 `Stop` 的）一般性警报。</span><span class="sxs-lookup"><span data-stu-id="11f68-179">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="11f68-180">以下代码设置了自定义消息和标题。</span><span class="sxs-lookup"><span data-stu-id="11f68-180">This code sets a custom message and title.</span></span>

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

<span data-ttu-id="11f68-181">有关详细信息，请参阅 [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="11f68-181">For more information, see [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="11f68-182">创建验证提示语</span><span class="sxs-lookup"><span data-stu-id="11f68-182">Create validation prompts</span></span>

<span data-ttu-id="11f68-183">你可以创建一个在以下情况下显示的说明性提示语：当用户将鼠标悬停在某个应用了数据验证的单元格上或选择该单元格时。</span><span class="sxs-lookup"><span data-stu-id="11f68-183">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="11f68-184">示例如下：</span><span class="sxs-lookup"><span data-stu-id="11f68-184">The following is an example:</span></span>

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

<span data-ttu-id="11f68-185">有关详细信息，请参阅 [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="11f68-185">For more information, see [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="11f68-186">从某个范围删除数据验证</span><span class="sxs-lookup"><span data-stu-id="11f68-186">Remove data validation from a range</span></span>

<span data-ttu-id="11f68-187">若要从某个范围删除数据验证，请调用 [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) 方法。</span><span class="sxs-lookup"><span data-stu-id="11f68-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="11f68-188">清除数据验证的范围不必与当初添加数据验证的范围完全相同。</span><span class="sxs-lookup"><span data-stu-id="11f68-188">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="11f68-189">如果二者不相同，则仅清除二者中重叠的单元格（如果存在）。</span><span class="sxs-lookup"><span data-stu-id="11f68-189">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="11f68-190">从某个范围清除数据验证还会清除用户手动添加至该范围的任何数据验证。</span><span class="sxs-lookup"><span data-stu-id="11f68-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="11f68-191">另请参阅</span><span class="sxs-lookup"><span data-stu-id="11f68-191">See also</span></span>

- [<span data-ttu-id="11f68-192">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="11f68-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="11f68-193">DataValidation 对象（适用于 Excel 的 JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="11f68-193">DataValidation Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="11f68-194">Range 对象（适用于 Excel 的 JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="11f68-194">Range Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.range)
