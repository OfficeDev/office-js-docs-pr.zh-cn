---
title: 使用 Excel JavaScript API 对区域执行操作（高级）
description: ''
ms.date: 12/18/2018
ms.openlocfilehash: 6d3da1e7eff4e61ae1b88213d0b432581d8f6a8a
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383237"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="c19dc-102">使用 Excel JavaScript API 对区域执行操作（高级）</span><span class="sxs-lookup"><span data-stu-id="c19dc-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="c19dc-103">本文基于[使用 Excel JavaScript API 对区域执行操作（基本）](excel-add-ins-ranges.md)中包含的信息，它提供了显示如何使用 Excel JavaScript API 对区域执行更多高级任务的代码示例。</span><span class="sxs-lookup"><span data-stu-id="c19dc-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="c19dc-104">有关 **Range** 对象支持的属性和方法的完整列表，请参阅 [Range 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="c19dc-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="c19dc-105">使用 Moment-MSDate 插件处理日期</span><span class="sxs-lookup"><span data-stu-id="c19dc-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="c19dc-106">[时刻 JavaScript 库](https://momentjs.com/)提供了使用日期和时间戳的便捷方式。</span><span class="sxs-lookup"><span data-stu-id="c19dc-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="c19dc-107">[Moment-MSDate 插件](https://www.npmjs.com/package/moment-msdate)可将时刻格式转换为 Excel 所需的格式。</span><span class="sxs-lookup"><span data-stu-id="c19dc-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="c19dc-108">这是 [NOW 函数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)返回的相同格式。</span><span class="sxs-lookup"><span data-stu-id="c19dc-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="c19dc-109">以下代码显示如何将 **B4** 处的范围设置为时刻的时间戳：</span><span class="sxs-lookup"><span data-stu-id="c19dc-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c19dc-110">这是一项类似于在单元格之外获取日期并将其转换为时刻或其他格式的技术，如以下代码中所示：</span><span class="sxs-lookup"><span data-stu-id="c19dc-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp 
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c19dc-111">你的加载项将必须对范围进行格式化才能以更可读的形式显示日期。</span><span class="sxs-lookup"><span data-stu-id="c19dc-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="c19dc-112">`"[$-409]m/d/yy h:mm AM/PM;@"` 的示例显示类似“12/3/18 3:57 PM”的时间。</span><span class="sxs-lookup"><span data-stu-id="c19dc-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="c19dc-113">有关日期和时间数字格式的详细信息，请参阅[查看自定义数字格式的准则](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)一文中的“日期和时间格式的准则”。</span><span class="sxs-lookup"><span data-stu-id="c19dc-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously-preview"></a><span data-ttu-id="c19dc-114">同时处理多个区域（预览版）</span><span class="sxs-lookup"><span data-stu-id="c19dc-114">Work with multiple ranges simultaneously in Excel add-ins</span></span>

> [!NOTE]
> <span data-ttu-id="c19dc-115">`RangeAreas` 对象当前仅适用于公共预览版（beta 版本）。</span><span class="sxs-lookup"><span data-stu-id="c19dc-115">The `RangeAreas` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="c19dc-116">若要使用此功能，必须使用 Office.js CDN 的 beta 版库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="c19dc-116">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="c19dc-117">如果使用的是 TypeScript 或代码编辑器将 TypeScript 类型定义文件用于 IntelliSense，则使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。</span><span class="sxs-lookup"><span data-stu-id="c19dc-117">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="c19dc-118">`RangeAreas` 对象允许外接程序每次在多个区域上执行操作。</span><span class="sxs-lookup"><span data-stu-id="c19dc-118">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="c19dc-119">这些区域可能但不必是连续区域。</span><span class="sxs-lookup"><span data-stu-id="c19dc-119">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="c19dc-120">`RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。</span><span class="sxs-lookup"><span data-stu-id="c19dc-120">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="copy-and-paste-preview"></a><span data-ttu-id="c19dc-121">复制和粘贴（预览版）</span><span class="sxs-lookup"><span data-stu-id="c19dc-121">Copy and paste</span></span>

> [!NOTE]
> <span data-ttu-id="c19dc-122">`Range.copyFrom` 函数当前仅适用于公共预览版（beta 版本）。</span><span class="sxs-lookup"><span data-stu-id="c19dc-122">The `Range.copyFrom` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="c19dc-123">若要使用此功能，必须使用 Office.js CDN 的 beta 版库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="c19dc-123">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="c19dc-124">如果使用的是 TypeScript 或代码编辑器将 TypeScript 类型定义文件用于 IntelliSense，则使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。</span><span class="sxs-lookup"><span data-stu-id="c19dc-124">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="c19dc-125">区域的 `copyFrom` 函数将复制 Excel UI 的“复制和粘贴”行为。</span><span class="sxs-lookup"><span data-stu-id="c19dc-125">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="c19dc-126">调用 `copyFrom` 的区域对象是目标。</span><span class="sxs-lookup"><span data-stu-id="c19dc-126">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="c19dc-127">将要复制的源作为一个范围或一个表示范围的字符串地址进行传递。</span><span class="sxs-lookup"><span data-stu-id="c19dc-127">The source to be copied is passed as a range or a string address representing a range.</span></span> <span data-ttu-id="c19dc-128">以下代码示例将数据从“A1:E1”\*\*\*\* 复制到“G1”\*\*\*\* 开始的范围（粘贴到“G1:K1”\*\*\*\* 结束）。</span><span class="sxs-lookup"><span data-stu-id="c19dc-128">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c19dc-129">`Range.copyFrom` 具有三个可选参数。</span><span class="sxs-lookup"><span data-stu-id="c19dc-129">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="c19dc-130">`copyType` 指定将哪些数据从源复制到目标。</span><span class="sxs-lookup"><span data-stu-id="c19dc-130">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="c19dc-131">`"Formulas"` 转换源单元格中的公式，并保留这些公式范围的相对位置。</span><span class="sxs-lookup"><span data-stu-id="c19dc-131">`"Formulas"` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="c19dc-132">将原样复制任何非公式条目。</span><span class="sxs-lookup"><span data-stu-id="c19dc-132">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="c19dc-133">`"Values"` 复制数据值，如果是公式，则复制公式的结果。</span><span class="sxs-lookup"><span data-stu-id="c19dc-133">`"Values"` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="c19dc-134">`"Formats"` 复制范围的格式设置（包括字体、颜色和其他格式），但不会复制任何值。</span><span class="sxs-lookup"><span data-stu-id="c19dc-134">`"Formats"` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="c19dc-135">`"All"`（默认选项）复制数据和格式设置，保留单元格的公式（如果找到）。</span><span class="sxs-lookup"><span data-stu-id="c19dc-135">`"All"` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="c19dc-136">`skipBlanks` 设置是否将空白单元格复制到目标。</span><span class="sxs-lookup"><span data-stu-id="c19dc-136">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="c19dc-137">如果为 true，`copyFrom` 将跳过源范围中的空白单元格。</span><span class="sxs-lookup"><span data-stu-id="c19dc-137">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="c19dc-138">跳过的单元格不会覆盖目标范围中其对应单元格的现有数据。</span><span class="sxs-lookup"><span data-stu-id="c19dc-138">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="c19dc-139">默认值为 false。</span><span class="sxs-lookup"><span data-stu-id="c19dc-139">The default is false.</span></span>

<span data-ttu-id="c19dc-140">`transpose` 确定是否将数据转置（即切换其行和列）到源位置。</span><span class="sxs-lookup"><span data-stu-id="c19dc-140">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="c19dc-141">转置范围沿主对角线翻转，因此，行“1”\*\*\*\*、“2”\*\*\*\* 和“3”\*\*\*\* 将成为列“A”\*\*\*\*、“B”\*\*\*\* 和“C”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="c19dc-141">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="c19dc-142">以下代码示例和图像在一个简单的方案中演示此行为。</span><span class="sxs-lookup"><span data-stu-id="c19dc-142">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c19dc-143">*在上一个函数已运行之前。*</span><span class="sxs-lookup"><span data-stu-id="c19dc-143">*Before the preceding function has been run.*</span></span>

![在区域中运行复制方法之前的 Excel 中的数据](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="c19dc-145">*在上一个函数已运行之后。*</span><span class="sxs-lookup"><span data-stu-id="c19dc-145">*After the preceding function has been run.*</span></span>

![在区域中运行复制方法之后的 Excel 中的数据](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates-preview"></a><span data-ttu-id="c19dc-147">删除重复项（预览版）</span><span class="sxs-lookup"><span data-stu-id="c19dc-147">Remove duplicates (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="c19dc-148">区域对象的 `removeDuplicates` 函数当前仅适用于公共预览版（beta 版本）。</span><span class="sxs-lookup"><span data-stu-id="c19dc-148">The Range object's `removeDuplicates` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="c19dc-149">若要使用此功能，必须使用 Office.js CDN 的 beta 版库：https://appsforoffice.microsoft.com/lib/beta/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="c19dc-149">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="c19dc-150">如果使用的是 TypeScript 或代码编辑器将 TypeScript 类型定义文件用于 IntelliSense，则使用 https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts。</span><span class="sxs-lookup"><span data-stu-id="c19dc-150">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="c19dc-151">区域对象的 `removeDuplicates` 函数将删除在指定列中具有重复条目的行。</span><span class="sxs-lookup"><span data-stu-id="c19dc-151">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="c19dc-152">该函数将从区域最低值索引到最高值索引（从上到下）遍历区域中的每一行。</span><span class="sxs-lookup"><span data-stu-id="c19dc-152">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="c19dc-153">如果指定列中的值之前显示在区域中，则会删除该行。</span><span class="sxs-lookup"><span data-stu-id="c19dc-153">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="c19dc-154">在区域内位于已删除行下方的行将上移。</span><span class="sxs-lookup"><span data-stu-id="c19dc-154">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="c19dc-155">`removeDuplicates` 不影响该区域外的单元格位置。</span><span class="sxs-lookup"><span data-stu-id="c19dc-155">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="c19dc-156">`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。</span><span class="sxs-lookup"><span data-stu-id="c19dc-156">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="c19dc-157">此数组从零开始并且与区域而不是与工作表相关。</span><span class="sxs-lookup"><span data-stu-id="c19dc-157">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="c19dc-158">该函数还使用一个布尔参数来指定第一行是否为标题。</span><span class="sxs-lookup"><span data-stu-id="c19dc-158">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="c19dc-159">如果为 **true**，则在考虑重复项时将忽略顶行。</span><span class="sxs-lookup"><span data-stu-id="c19dc-159">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="c19dc-160">`removeDuplicates` 函数将返回 `RemoveDuplicatesResult` 对象，用于指定已删除的行数和剩余的唯一行数。</span><span class="sxs-lookup"><span data-stu-id="c19dc-160">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="c19dc-161">在使用区域的 `removeDuplicates` 函数时，应记住以下几点：</span><span class="sxs-lookup"><span data-stu-id="c19dc-161">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="c19dc-162">`removeDuplicates` 会考虑单元格值，而不是函数结果。</span><span class="sxs-lookup"><span data-stu-id="c19dc-162">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="c19dc-163">如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。</span><span class="sxs-lookup"><span data-stu-id="c19dc-163">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="c19dc-164">`removeDuplicates` 不会忽略空单元格。</span><span class="sxs-lookup"><span data-stu-id="c19dc-164">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="c19dc-165">空单元格的值与任何其他值具有相同的处理方式。</span><span class="sxs-lookup"><span data-stu-id="c19dc-165">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="c19dc-166">这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。</span><span class="sxs-lookup"><span data-stu-id="c19dc-166">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="c19dc-167">以下示例显示删除第一列中具有重复值的条目。</span><span class="sxs-lookup"><span data-stu-id="c19dc-167">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c19dc-168">*在上一个函数已运行之前。*</span><span class="sxs-lookup"><span data-stu-id="c19dc-168">*Before the preceding function has been run.*</span></span>

![在区域中运行删除重复项方法之前的 Excel 中的数据](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="c19dc-170">*在上一个函数已运行之后。*</span><span class="sxs-lookup"><span data-stu-id="c19dc-170">*After the preceding function has been run.*</span></span>

![在区域中运行删除重复项方法之后的 Excel 中的数据](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="c19dc-172">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c19dc-172">See also</span></span>

- [<span data-ttu-id="c19dc-173">使用 Excel JavaScript API 对区域执行操作</span><span class="sxs-lookup"><span data-stu-id="c19dc-173">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="c19dc-174">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="c19dc-174">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)