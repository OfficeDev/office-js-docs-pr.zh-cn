---
title: 将数据验证添加到 Excel 范围
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: af965df4a1aece5b7f8d5ea89664519b576a4850
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925309"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="7be06-102">将数据验证添加到 Excel 范围（预览）</span><span class="sxs-lookup"><span data-stu-id="7be06-102">Add data validation to Excel ranges (Preview)</span></span>

> [!NOTE]
> <span data-ttu-id="7be06-103">虽然数据验证 API 处于预览状态，但你必须加载 Office JavaScript 库的 Beta 版才能使用它们。</span><span class="sxs-lookup"><span data-stu-id="7be06-103">While the data validation APIs are in preview, you must load the beta version of the Office JavaScript library to use them.</span></span> <span data-ttu-id="7be06-104">URL 是 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="7be06-104">The URL is https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span> <span data-ttu-id="7be06-105">如果你正在使用 TypeScript，或者你的代码编辑器使用 TypeScript 类型定义文件实现 IntelliSense，请使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。</span><span class="sxs-lookup"><span data-stu-id="7be06-105">If you are using TypeScript or your code editor uses a TypeScript type definition file for intellisense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

> [!NOTE]
> <span data-ttu-id="7be06-106">虽然数据验证API处于预览状态，但本文中 API 引用的链接将不起作用。</span><span class="sxs-lookup"><span data-stu-id="7be06-106">While the data validation APIs are in preview, the links in this article to API reference will not work.</span></span> <span data-ttu-id="7be06-107">在此期间，你可以使用 [草稿 Excel API 引用](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel)。</span><span class="sxs-lookup"><span data-stu-id="7be06-107">In the meantime, you can use the [draft Excel API reference](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel).</span></span>

<span data-ttu-id="7be06-108">Excel JavaScript 库提供的 API 可让外接程序将自动数据验证添加到工作簿中的表、列、行和其他范围。</span><span class="sxs-lookup"><span data-stu-id="7be06-108">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="7be06-109">要了解数据验证的概念和术语，请参阅以下关于用户如何通过 Excel UI 添加数据验证的文章：</span><span class="sxs-lookup"><span data-stu-id="7be06-109">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="7be06-110">将数据验证应用于单元格</span><span class="sxs-lookup"><span data-stu-id="7be06-110">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="7be06-111">有关数据验证的更多信息</span><span class="sxs-lookup"><span data-stu-id="7be06-111">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="7be06-112">Excel 中数据验证的说明和示例</span><span class="sxs-lookup"><span data-stu-id="7be06-112">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="7be06-113">数据验证的程序控制</span><span class="sxs-lookup"><span data-stu-id="7be06-113">Programmatic control of data validation</span></span>

<span data-ttu-id="7be06-114">属性需要一个 [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) 对象，是 Excel 中数据验证的编程控制入口点。`Range.dataValidation`</span><span class="sxs-lookup"><span data-stu-id="7be06-114">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="7be06-115">对象有五个属性：`DataValidation`</span><span class="sxs-lookup"><span data-stu-id="7be06-115">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="7be06-116">`rule` - 定义构成范围的有效数据。</span><span class="sxs-lookup"><span data-stu-id="7be06-116">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="7be06-117">请参阅 [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule)。</span><span class="sxs-lookup"><span data-stu-id="7be06-117">See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="7be06-118">`errorAlert` - 指定用户输入无效数据时是否弹出错误，并定义警报文本，标题和样式;例如，**信息**、**警告**和**停止**。</span><span class="sxs-lookup"><span data-stu-id="7be06-118">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="7be06-119">请参阅 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="7be06-119">See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="7be06-120">`prompt` - 指定当用户将光标悬停在范围上时是否显示提示并且定义提示消息。</span><span class="sxs-lookup"><span data-stu-id="7be06-120">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="7be06-121">请参阅 [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="7be06-121">See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="7be06-122">`ignoreBlanks` - 指定数据验证规则是否适用于范围内的空白单元格。</span><span class="sxs-lookup"><span data-stu-id="7be06-122">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="7be06-123">默认为 `true`。</span><span class="sxs-lookup"><span data-stu-id="7be06-123">Defaults to `true`.</span></span>
- <span data-ttu-id="7be06-124">`type` - 验证类型的只读标识，例如 WholeNumber、Date、TextLength 等。在设置 `rule` 属性时间接设置。</span><span class="sxs-lookup"><span data-stu-id="7be06-124">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="7be06-125">通过编程添加的数据验证其工作方式等同于手动添加的数据验证。</span><span class="sxs-lookup"><span data-stu-id="7be06-125">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="7be06-126">尤其要注意的是，仅当用户直接将值输入到单元格中或从工作簿中的其他位置复制和粘贴单元格并选择**值**粘贴选项时，才会触发数据验证。</span><span class="sxs-lookup"><span data-stu-id="7be06-126">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="7be06-127">如果用户复制一个单元格并将其粘贴到包含数据验证的范围中，则不会触发验证。</span><span class="sxs-lookup"><span data-stu-id="7be06-127">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="7be06-128">创建验证规则</span><span class="sxs-lookup"><span data-stu-id="7be06-128">Creating validation rules</span></span>

<span data-ttu-id="7be06-129">要将数据验证添加到范围，代码必须在 `Range.dataValidation` 中设置 `DataValidation` 对象的 `rule` 属性。</span><span class="sxs-lookup"><span data-stu-id="7be06-129">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="7be06-130">这需要一个 [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) 对象， 它具有七个可选属性。</span><span class="sxs-lookup"><span data-stu-id="7be06-130">This takes a [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="7be06-131">*任何 `DataValidationRule` 对象中的这些属性均不得超过一个。*</span><span class="sxs-lookup"><span data-stu-id="7be06-131">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="7be06-132">所包含的属性决定了验证的类型。</span><span class="sxs-lookup"><span data-stu-id="7be06-132">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="7be06-133">Basic 和 DateTime 验证规则类型</span><span class="sxs-lookup"><span data-stu-id="7be06-133">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="7be06-134">前三个 `DataValidationRule` 属性（即验证规则类型）需要一个 [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) 对象作为它们的值。</span><span class="sxs-lookup"><span data-stu-id="7be06-134">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="7be06-135">`wholeNumber` - 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个整数。</span><span class="sxs-lookup"><span data-stu-id="7be06-135">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="7be06-136">`decimal` - 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个十进制数。</span><span class="sxs-lookup"><span data-stu-id="7be06-136">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="7be06-137">`textLength` - 在 `BasicDataValidation` 对象中针对单元格值的*长度*应用验证细节。</span><span class="sxs-lookup"><span data-stu-id="7be06-137">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="7be06-138">以下是创建验证规则的示例。</span><span class="sxs-lookup"><span data-stu-id="7be06-138">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="7be06-139">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7be06-139">Note the following about this code:</span></span>

- <span data-ttu-id="7be06-140">是二元运算符 "GreaterThan"。`operator`</span><span class="sxs-lookup"><span data-stu-id="7be06-140">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="7be06-141">无论何时使用二元运算符，用户尝试输入到单元格的值都是左侧的操作数，并且在 `formula1` 中指定的值是右侧操作数。</span><span class="sxs-lookup"><span data-stu-id="7be06-141">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="7be06-142">所以这条规则说只有大于 0 的整数才有效。</span><span class="sxs-lookup"><span data-stu-id="7be06-142">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="7be06-143">是一个硬编码的数字。`formula1`</span><span class="sxs-lookup"><span data-stu-id="7be06-143">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="7be06-144">如果在写代码时不知道该值应该是多少，还可以使用 Excel 公式（作为字符串）来计算该值。</span><span class="sxs-lookup"><span data-stu-id="7be06-144">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="7be06-145">例如，`formula1` 的值也可以是 "= A3" 和 "= SUM(A4,B5)"。</span><span class="sxs-lookup"><span data-stu-id="7be06-145">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="7be06-146">有关其他二元运算符的列表，请参阅 [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation)。</span><span class="sxs-lookup"><span data-stu-id="7be06-146">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="7be06-147">还有两个三元运算符："Between" 和 "NotBetween"。</span><span class="sxs-lookup"><span data-stu-id="7be06-147">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="7be06-148">要使用这些运算符，必须指定可选的 `formula2` 属性。</span><span class="sxs-lookup"><span data-stu-id="7be06-148">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="7be06-149">和 `formula2` 值是边界操作数。`formula1`</span><span class="sxs-lookup"><span data-stu-id="7be06-149">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="7be06-150">用户尝试输入单元格的值是第三个（评估）的操作数。</span><span class="sxs-lookup"><span data-stu-id="7be06-150">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="7be06-151">以下是使用 "Between" 运算符的示例：</span><span class="sxs-lookup"><span data-stu-id="7be06-151">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="7be06-152">接下来的两个规则属性需要一个 [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) 对象作为它们的值。</span><span class="sxs-lookup"><span data-stu-id="7be06-152">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="7be06-153">对象的结构类似于 `BasicDataValidation`：它有属性 `formula1`、`formula2` 和 `operator`，并以相同的方式使用。`DateTimeDataValidation`</span><span class="sxs-lookup"><span data-stu-id="7be06-153">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="7be06-154">不同之处在于，不能在公式属性中使用数字，但可以输入一个 [ISO 8606 日期时间](https://www.iso.org/iso-8601-date-and-time-format.html)字符串（或 Excel 公式）。</span><span class="sxs-lookup"><span data-stu-id="7be06-154">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="7be06-155">以下是将有效值定义为 2018 年 4 月第一周日期的示例。</span><span class="sxs-lookup"><span data-stu-id="7be06-155">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

#### <a name="list-validation-rule-type"></a><span data-ttu-id="7be06-156">列表验证规则类型</span><span class="sxs-lookup"><span data-stu-id="7be06-156">List validation rule type</span></span>

<span data-ttu-id="7be06-157">使用 `DataValidationRule` 对象的 `list` 属性来指定唯一有效值是来自有限列表中的那些值。</span><span class="sxs-lookup"><span data-stu-id="7be06-157">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="7be06-158">示例如下。</span><span class="sxs-lookup"><span data-stu-id="7be06-158">The following is an example.</span></span> <span data-ttu-id="7be06-159">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7be06-159">Note the following about this code:</span></span>

- <span data-ttu-id="7be06-160">它假定有一个名为 "Names" 的工作表，并且 "A1:A3" 范围中的值是名称。</span><span class="sxs-lookup"><span data-stu-id="7be06-160">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="7be06-161">属性指定有效值的列表。`source`</span><span class="sxs-lookup"><span data-stu-id="7be06-161">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="7be06-162">具有名称的范围已分配给它。</span><span class="sxs-lookup"><span data-stu-id="7be06-162">The range with the names has been assigned to it.</span></span> <span data-ttu-id="7be06-163">也可分配以逗号分隔的列表，例如："Sue, Ricky, Liz"。</span><span class="sxs-lookup"><span data-stu-id="7be06-163">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="7be06-164">属性指定用户选择单元格时是否出现下拉控件。`inCellDropDown`</span><span class="sxs-lookup"><span data-stu-id="7be06-164">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="7be06-165">如果设置为 `true`，将显示下拉列表，其中带有来自 `source` 中的值。</span><span class="sxs-lookup"><span data-stu-id="7be06-165">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="7be06-166">自定义验证规则类型</span><span class="sxs-lookup"><span data-stu-id="7be06-166">Custom validation rule type</span></span>

<span data-ttu-id="7be06-167">使用 `DataValidationRule` 对象中的 `custom` 属性来指定自定义验证公式。</span><span class="sxs-lookup"><span data-stu-id="7be06-167">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="7be06-168">示例如下。</span><span class="sxs-lookup"><span data-stu-id="7be06-168">The following is an example.</span></span> <span data-ttu-id="7be06-169">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7be06-169">Note the following about this code:</span></span>

- <span data-ttu-id="7be06-170">它假定有一个包含 **Athlete Name** 和 **Comments** 两列的表，它们分别位于工作表的 A 和 B 列。</span><span class="sxs-lookup"><span data-stu-id="7be06-170">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="7be06-171">为了减少 **Comment** 列中的冗余，它会认定包含运动员姓名的数据无效。</span><span class="sxs-lookup"><span data-stu-id="7be06-171">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="7be06-172">`SEARCH(A2,B2)` 返回 A2 字符串在 B2 字符串中的起始位置。</span><span class="sxs-lookup"><span data-stu-id="7be06-172">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="7be06-173">如果 A2 不包含在 B2 中，它不会返回一个数字。</span><span class="sxs-lookup"><span data-stu-id="7be06-173">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="7be06-174">`ISNUMBER()` 返回一个布尔值。</span><span class="sxs-lookup"><span data-stu-id="7be06-174">Returns a `ISNUMBER()`.</span></span> <span data-ttu-id="7be06-175">所以 `formula` 属性表示，**Comment** 列的有效数据是不包含 **Athlete Name** 列字符串的数据。</span><span class="sxs-lookup"><span data-stu-id="7be06-175">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

### <a name="create-validation-error-alerts"></a><span data-ttu-id="7be06-176">创建验证错误警报</span><span class="sxs-lookup"><span data-stu-id="7be06-176">Create validation error alerts</span></span>

<span data-ttu-id="7be06-177">你可以创建用户尝试在单元格中输入无效数据时出现的自定义错误警报。</span><span class="sxs-lookup"><span data-stu-id="7be06-177">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="7be06-178">下面展示了一个非常简单的示例。</span><span class="sxs-lookup"><span data-stu-id="7be06-178">The following is a simple example:</span></span> <span data-ttu-id="7be06-179">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="7be06-179">Note the following about this code:</span></span>

- <span data-ttu-id="7be06-180">属性确定用户是否收到信息提示、警告或“停止”警报。`style`</span><span class="sxs-lookup"><span data-stu-id="7be06-180">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="7be06-181">只有 `Stop` 可以实际防止用户添加无效数据。</span><span class="sxs-lookup"><span data-stu-id="7be06-181">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="7be06-182">和 `Information` 的弹出窗口具有允许用户仍然输入无效数据的选项。`Warning`</span><span class="sxs-lookup"><span data-stu-id="7be06-182">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="7be06-183">属性默认为 `true`。`showAlert`</span><span class="sxs-lookup"><span data-stu-id="7be06-183">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="7be06-184">这意味着 Excel 主机将弹出一个通用警报（类型为 `Stop`），除非你通过创建自定义警报将 `showAlert` 设置为 `false` 或设置自定义消息、标题和样式。</span><span class="sxs-lookup"><span data-stu-id="7be06-184">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="7be06-185">此代码设置自定义消息和标题。</span><span class="sxs-lookup"><span data-stu-id="7be06-185">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="7be06-186">有关更多信息，请参阅 [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="7be06-186">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="7be06-187">创建验证提示</span><span class="sxs-lookup"><span data-stu-id="7be06-187">Create validation prompts</span></span>

<span data-ttu-id="7be06-188">可以创建一个在用户悬停指针或选择应用了数据验证的单元格时出现的指示性提示。</span><span class="sxs-lookup"><span data-stu-id="7be06-188">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="7be06-189">示例如下：</span><span class="sxs-lookup"><span data-stu-id="7be06-189">The following is an example:</span></span>

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

<span data-ttu-id="7be06-190">有关更多信息，请参阅 [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="7be06-190">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="7be06-191">从范围中删除数据验证</span><span class="sxs-lookup"><span data-stu-id="7be06-191">Remove data validation from a range</span></span>

<span data-ttu-id="7be06-192">要从范围中删除数据验证，请调用 [Range.dataValidation.clear（）](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) 方法。</span><span class="sxs-lookup"><span data-stu-id="7be06-192">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="7be06-193">清除的范围与添加数据验证的范围不一定需要完全相同。</span><span class="sxs-lookup"><span data-stu-id="7be06-193">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="7be06-194">如果两者不相同，则只清除两个范围中重叠的单元格（如果有的话）。</span><span class="sxs-lookup"><span data-stu-id="7be06-194">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="7be06-195">清除范围内的数据验证也将清除用户手动添加到范围内的任何数据验证。</span><span class="sxs-lookup"><span data-stu-id="7be06-195">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="7be06-196">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7be06-196">See also</span></span>

- [<span data-ttu-id="7be06-197">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="7be06-197">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="7be06-198">DataValidation 对象 (Excel JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="7be06-198">Chart Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="7be06-199">Range 对象 (Excel JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="7be06-199">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
