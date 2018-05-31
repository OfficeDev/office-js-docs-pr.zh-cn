---
title: 将数据验证添加到 Excel 范围
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: 8e5f09f1c566103f34ad584885769229c17ab1f7
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437526"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="58aca-102">将数据验证添加到 Excel 范围（预览）</span><span class="sxs-lookup"><span data-stu-id="58aca-102">Add data validation to Excel ranges (Preview)</span></span>

> [!NOTE]
> <span data-ttu-id="58aca-103">虽然数据验证 API 处于预览状态，但你必须加载 Office JavaScript 库的 Beta 版才能使用它们。</span><span class="sxs-lookup"><span data-stu-id="58aca-103">While the data validation APIs are in preview, you must load the beta version of the Office JavaScript library to use them.</span></span> <span data-ttu-id="58aca-104">URL 是 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="58aca-104">The URL is https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span> <span data-ttu-id="58aca-105">如果你正在使用 TypeScript，或者你的代码编辑器使用 TypeScript 类型定义文件实现 IntelliSense，请使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。</span><span class="sxs-lookup"><span data-stu-id="58aca-105">If you are using TypeScript or your code editor uses a TypeScript type definition file for intellisense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="58aca-106">Excel JavaScript 库提供的 API 可让外接程序将自动数据验证添加到工作簿中的表、列、行和其他范围。</span><span class="sxs-lookup"><span data-stu-id="58aca-106">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="58aca-107">要了解数据验证的概念和术语，请参阅以下关于用户如何通过 Excel UI 添加数据验证的文章：</span><span class="sxs-lookup"><span data-stu-id="58aca-107">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="58aca-108">将数据验证应用于单元格</span><span class="sxs-lookup"><span data-stu-id="58aca-108">Apply data validation to cells</span></span>](https://support.office.com/en-us/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="58aca-109">有关数据验证的更多信息</span><span class="sxs-lookup"><span data-stu-id="58aca-109">More on data validation</span></span>](https://microsoft.sharepoint.com/:p:/r/teams/oext/_layouts/15/Doc.aspx?sourcedoc=%7B51143964-d52c-429d-bfac-c7495473d536%7D&action=edit)
- [<span data-ttu-id="58aca-110">Excel 中数据验证的说明和示例</span><span class="sxs-lookup"><span data-stu-id="58aca-110">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/en-us/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="58aca-111">数据验证的程序控制</span><span class="sxs-lookup"><span data-stu-id="58aca-111">Programmatic control of data validation</span></span>

<span data-ttu-id="58aca-112">`Range.dataValidation` 属性需要一个 [DataValidation](https://dev.office.com/reference/add-ins/excel/datavalidation) 对象，是 Excel 中数据验证的编程控制入口点。</span><span class="sxs-lookup"><span data-stu-id="58aca-112">The `Range.dataValidation` property, which takes a [DataValidation](https://dev.office.com/reference/add-ins/excel/datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="58aca-113">`DataValidation` 对象有五个属性：</span><span class="sxs-lookup"><span data-stu-id="58aca-113">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="58aca-114">`rule` - 定义构成范围的有效数据。</span><span class="sxs-lookup"><span data-stu-id="58aca-114">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="58aca-115">请参阅 [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule)。</span><span class="sxs-lookup"><span data-stu-id="58aca-115">See [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule).</span></span>
- <span data-ttu-id="58aca-116">`errorAlert` - 指定用户输入无效数据时是否弹出错误，并定义警报文本，标题和样式;例如，**信息**、**警告**和**停止**。</span><span class="sxs-lookup"><span data-stu-id="58aca-116">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="58aca-117">请参阅 [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="58aca-117">See [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span></span>
- <span data-ttu-id="58aca-118">`prompt` - 指定当用户将光标悬停在范围上时是否显示提示并且定义提示消息。</span><span class="sxs-lookup"><span data-stu-id="58aca-118">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="58aca-119">请参阅 [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="58aca-119">See [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span></span>
- <span data-ttu-id="58aca-120">`ignoreBlanks` - 指定数据验证规则是否适用于范围内的空白单元格。</span><span class="sxs-lookup"><span data-stu-id="58aca-120">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="58aca-121">默认为 `true`。</span><span class="sxs-lookup"><span data-stu-id="58aca-121">Defaults to `true`.</span></span>
- <span data-ttu-id="58aca-122">`type` - 验证类型的只读标识，例如 WholeNumber、Date、TextLength 等。在设置 `rule` 属性时间接设置。</span><span class="sxs-lookup"><span data-stu-id="58aca-122">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="58aca-123">通过编程添加的数据验证其工作方式等同于手动添加的数据验证。</span><span class="sxs-lookup"><span data-stu-id="58aca-123">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="58aca-124">尤其要注意的是，仅当用户直接将值输入到单元格中或从工作簿中的其他位置复制和粘贴单元格并选择**值**粘贴选项时，才会触发数据验证。</span><span class="sxs-lookup"><span data-stu-id="58aca-124">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="58aca-125">如果用户复制一个单元格并将其粘贴到包含数据验证的范围中，则不会触发验证。</span><span class="sxs-lookup"><span data-stu-id="58aca-125">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="58aca-126">创建验证规则</span><span class="sxs-lookup"><span data-stu-id="58aca-126">Creating validation rules</span></span>

<span data-ttu-id="58aca-127">要将数据验证添加到范围，代码必须在 `Range.dataValidation` 中设置 `DataValidation` 对象的 `rule` 属性。</span><span class="sxs-lookup"><span data-stu-id="58aca-127">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="58aca-128">这需要一个 [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) 对象， 它具有七个可选属性。</span><span class="sxs-lookup"><span data-stu-id="58aca-128">This takes a [DataValidationRule](https://dev.office.com/reference/add-ins/excel/datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="58aca-129">*任何 `DataValidationRule` 对象中的这些属性均不得超过一个。*</span><span class="sxs-lookup"><span data-stu-id="58aca-129">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="58aca-130">所包含的属性决定了验证的类型。</span><span class="sxs-lookup"><span data-stu-id="58aca-130">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="58aca-131">Basic 和 DateTime 验证规则类型</span><span class="sxs-lookup"><span data-stu-id="58aca-131">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="58aca-132">前三个 `DataValidationRule` 属性（即验证规则类型）需要一个 [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) 对象作为它们的值。</span><span class="sxs-lookup"><span data-stu-id="58aca-132">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="58aca-133">`wholeNumber` - 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个整数。</span><span class="sxs-lookup"><span data-stu-id="58aca-133">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="58aca-134">`decimal` - 除了 `BasicDataValidation` 对象指定的任何其他验证之外，还需要一个十进制数。</span><span class="sxs-lookup"><span data-stu-id="58aca-134">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="58aca-135">`textLength` - 在 `BasicDataValidation` 对象中针对单元格值的*长度*应用验证细节。</span><span class="sxs-lookup"><span data-stu-id="58aca-135">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="58aca-136">以下是创建验证规则的示例。</span><span class="sxs-lookup"><span data-stu-id="58aca-136">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="58aca-137">注意有关这段代码的以下方面：
</span><span class="sxs-lookup"><span data-stu-id="58aca-137">Note the following about this code:</span></span>

- <span data-ttu-id="58aca-138"> `operator` 是二元运算符 "GreaterThan"。</span><span class="sxs-lookup"><span data-stu-id="58aca-138">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="58aca-139">无论何时使用二元运算符，用户尝试输入到单元格的值都是左侧的操作数，并且在 `formula1` 中指定的值是右侧操作数。</span><span class="sxs-lookup"><span data-stu-id="58aca-139">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="58aca-140">所以这条规则说只有大于 0 的整数才有效。</span><span class="sxs-lookup"><span data-stu-id="58aca-140">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="58aca-141">`formula1` 是一个硬编码的数字。</span><span class="sxs-lookup"><span data-stu-id="58aca-141">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="58aca-142">如果在写代码时不知道该值应该是多少，还可以使用 Excel 公式（作为字符串）来计算该值。</span><span class="sxs-lookup"><span data-stu-id="58aca-142">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="58aca-143">例如，`formula1` 的值也可以是 "= A3" 和 "= SUM(A4,B5)"。</span><span class="sxs-lookup"><span data-stu-id="58aca-143">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="58aca-144">有关其他二元运算符的列表，请参阅 [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation)。</span><span class="sxs-lookup"><span data-stu-id="58aca-144">See [BasicDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="58aca-145">还有两个三元运算符："Between" 和 "NotBetween"。</span><span class="sxs-lookup"><span data-stu-id="58aca-145">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="58aca-146">要使用这些运算符，必须指定可选的 `formula2` 属性。</span><span class="sxs-lookup"><span data-stu-id="58aca-146">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="58aca-147">`formula1` 和 `formula2` 值是边界操作数。</span><span class="sxs-lookup"><span data-stu-id="58aca-147">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="58aca-148">用户尝试输入单元格的值是第三个（评估）的操作数。</span><span class="sxs-lookup"><span data-stu-id="58aca-148">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="58aca-149">以下是使用 "Between" 运算符的示例：</span><span class="sxs-lookup"><span data-stu-id="58aca-149">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="58aca-150">接下来的两个规则属性需要一个 [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) 对象作为它们的值。</span><span class="sxs-lookup"><span data-stu-id="58aca-150">The next two rule properties take a [DateTimeDataValidation](https://dev.office.com/reference/add-ins/excel/basicdatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="58aca-151">`DateTimeDataValidation` 对象的结构类似于 `BasicDataValidation`：它有属性 `formula1`、`formula2` 和 `operator`，并以相同的方式使用。</span><span class="sxs-lookup"><span data-stu-id="58aca-151">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="58aca-152">不同之处在于，不能在公式属性中使用数字，但可以输入一个 [ISO 8606 日期时间](https://www.iso.org/iso-8601-date-and-time-format.html)字符串（或 Excel 公式）。</span><span class="sxs-lookup"><span data-stu-id="58aca-152">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="58aca-153">以下是将有效值定义为 2018 年 4 月第一周日期的示例。</span><span class="sxs-lookup"><span data-stu-id="58aca-153">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

#### <a name="list-validation-rule-type"></a><span data-ttu-id="58aca-154">列表验证规则类型</span><span class="sxs-lookup"><span data-stu-id="58aca-154">List validation rule type</span></span>

<span data-ttu-id="58aca-155">使用 `DataValidationRule` 对象的 `list` 属性来指定唯一有效值是来自有限列表中的那些值。</span><span class="sxs-lookup"><span data-stu-id="58aca-155">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="58aca-156">示例如下。</span><span class="sxs-lookup"><span data-stu-id="58aca-156">The following is an example.</span></span> <span data-ttu-id="58aca-157">注意有关这段代码的以下方面：
</span><span class="sxs-lookup"><span data-stu-id="58aca-157">Note the following about this code:</span></span>

- <span data-ttu-id="58aca-158">它假定有一个名为 "Names" 的工作表，并且 "A1:A3" 范围中的值是名称。</span><span class="sxs-lookup"><span data-stu-id="58aca-158">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="58aca-159">`source` 属性指定有效值的列表。</span><span class="sxs-lookup"><span data-stu-id="58aca-159">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="58aca-160">具有名称的范围已分配给它。</span><span class="sxs-lookup"><span data-stu-id="58aca-160">The range with the names has been assigned to it.</span></span> <span data-ttu-id="58aca-161">也可分配以逗号分隔的列表，例如："Sue, Ricky, Liz"。</span><span class="sxs-lookup"><span data-stu-id="58aca-161">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="58aca-162">`inCellDropDown` 属性指定用户选择单元格时是否出现下拉控件。</span><span class="sxs-lookup"><span data-stu-id="58aca-162">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="58aca-163">如果设置为 `true`，将显示下拉列表，其中带有来自 `source` 中的值。</span><span class="sxs-lookup"><span data-stu-id="58aca-163">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="58aca-164">自定义验证规则类型</span><span class="sxs-lookup"><span data-stu-id="58aca-164">Custom validation rule type</span></span>

<span data-ttu-id="58aca-165">使用 `DataValidationRule` 对象中的 `custom` 属性来指定自定义验证公式。</span><span class="sxs-lookup"><span data-stu-id="58aca-165">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="58aca-166">示例如下。</span><span class="sxs-lookup"><span data-stu-id="58aca-166">The following is an example.</span></span> <span data-ttu-id="58aca-167">注意有关这段代码的以下方面：
</span><span class="sxs-lookup"><span data-stu-id="58aca-167">Note the following about this code:</span></span>

- <span data-ttu-id="58aca-168">它假定有一个包含 **Athlete Name** 和 **Comments** 两列的表，它们分别位于工作表的 A 和 B 列。</span><span class="sxs-lookup"><span data-stu-id="58aca-168">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="58aca-169">为了减少 **Comment** 列中的冗余，它会认定包含运动员姓名的数据无效。</span><span class="sxs-lookup"><span data-stu-id="58aca-169">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="58aca-170">`SEARCH(A2,B2)` 返回 A2 字符串在 B2 字符串中的起始位置。</span><span class="sxs-lookup"><span data-stu-id="58aca-170">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="58aca-171">如果 A2 不包含在 B2 中，它不会返回一个数字。</span><span class="sxs-lookup"><span data-stu-id="58aca-171">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="58aca-172">`ISNUMBER()` 返回一个布尔值。</span><span class="sxs-lookup"><span data-stu-id="58aca-172">Returns a `ISNUMBER()`.</span></span> <span data-ttu-id="58aca-173">所以 `formula` 属性表示，**Comment** 列的有效数据是不包含 **Athlete Name** 列字符串的数据。</span><span class="sxs-lookup"><span data-stu-id="58aca-173">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

### <a name="create-validation-error-alerts"></a><span data-ttu-id="58aca-174">创建验证错误警报</span><span class="sxs-lookup"><span data-stu-id="58aca-174">Create validation error alerts</span></span>

<span data-ttu-id="58aca-175">你可以创建用户尝试在单元格中输入无效数据时出现的自定义错误警报。</span><span class="sxs-lookup"><span data-stu-id="58aca-175">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="58aca-176">下面展示了一个非常简单的示例。</span><span class="sxs-lookup"><span data-stu-id="58aca-176">The following is a simple example:</span></span> <span data-ttu-id="58aca-177">注意有关这段代码的以下方面：
</span><span class="sxs-lookup"><span data-stu-id="58aca-177">Note the following about this code:</span></span>

- <span data-ttu-id="58aca-178">`style` 属性确定用户是否收到信息提示、警告或“停止”警报。</span><span class="sxs-lookup"><span data-stu-id="58aca-178">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="58aca-179">只有 `Stop` 可以实际防止用户添加无效数据。</span><span class="sxs-lookup"><span data-stu-id="58aca-179">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="58aca-180">`Warning` 和 `Information` 的弹出窗口具有允许用户仍然输入无效数据的选项。</span><span class="sxs-lookup"><span data-stu-id="58aca-180">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="58aca-181">`showAlert` 属性默认为 `true`。</span><span class="sxs-lookup"><span data-stu-id="58aca-181">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="58aca-182">这意味着 Excel 主机将弹出一个通用警报（类型为 `Stop`），除非你通过创建自定义警报将 `showAlert` 设置为 `false` 或设置自定义消息、标题和样式。</span><span class="sxs-lookup"><span data-stu-id="58aca-182">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="58aca-183">此代码设置自定义消息和标题。</span><span class="sxs-lookup"><span data-stu-id="58aca-183">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="58aca-184">有关更多信息，请参阅 [DataValidationErrorAlert](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert)。</span><span class="sxs-lookup"><span data-stu-id="58aca-184">For more information, see [NextRecordset](https://dev.office.com/reference/add-ins/excel/datavalidationerroralert).</span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="58aca-185">创建验证提示</span><span class="sxs-lookup"><span data-stu-id="58aca-185">Create validation prompts</span></span>

<span data-ttu-id="58aca-186">可以创建一个在用户悬停指针或选择应用了数据验证的单元格时出现的指示性提示。</span><span class="sxs-lookup"><span data-stu-id="58aca-186">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="58aca-187">示例如下：</span><span class="sxs-lookup"><span data-stu-id="58aca-187">The following is an example:</span></span>

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

<span data-ttu-id="58aca-188">有关更多信息，请参阅 [DataValidationPrompt](https://dev.office.com/reference/add-ins/excel/datavalidationprompt)。</span><span class="sxs-lookup"><span data-stu-id="58aca-188">For more information, see [NextRecordset](https://dev.office.com/reference/add-ins/excel/datavalidationprompt).</span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="58aca-189">从范围中删除数据验证</span><span class="sxs-lookup"><span data-stu-id="58aca-189">Remove data validation from a range</span></span>

<span data-ttu-id="58aca-190">要从范围中删除数据验证，请调用 [Range.dataValidation.clear（）](https://dev.office.com/reference/add-ins/excel/datavalidation#clear) 方法。</span><span class="sxs-lookup"><span data-stu-id="58aca-190">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://dev.office.com/reference/add-ins/excel/datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="58aca-191">清除的范围与添加数据验证的范围不一定需要完全相同。</span><span class="sxs-lookup"><span data-stu-id="58aca-191">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="58aca-192">如果两者不相同，则只清除两个范围中重叠的单元格（如果有的话）。</span><span class="sxs-lookup"><span data-stu-id="58aca-192">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="58aca-193">清除范围内的数据验证也将清除用户手动添加到范围内的任何数据验证。</span><span class="sxs-lookup"><span data-stu-id="58aca-193">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="58aca-194">另请参阅</span><span class="sxs-lookup"><span data-stu-id="58aca-194">See also</span></span>

- [<span data-ttu-id="58aca-195">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="58aca-195">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="58aca-196">DataValidation 对象 (Excel JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="58aca-196">Worksheet Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/datavalidation)
- [<span data-ttu-id="58aca-197">Range 对象 (Excel JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="58aca-197">Range Object (JavaScript API for Excel)</span></span>](https://dev.office.com/reference/add-ins/excel/range)



 
