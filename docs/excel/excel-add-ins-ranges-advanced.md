---
title: 使用 Excel JavaScript API 对区域执行操作（高级）
description: ''
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: aacbe930e2cf3da4d10b61bfe8f34efe1094c113
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914238"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="f81b5-102">使用 Excel JavaScript API 对区域执行操作（高级）</span><span class="sxs-lookup"><span data-stu-id="f81b5-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="f81b5-103">本文基于[使用 Excel JavaScript API 对区域执行操作（基本）](excel-add-ins-ranges.md)中包含的信息，它提供了显示如何使用 Excel JavaScript API 对区域执行更多高级任务的代码示例。</span><span class="sxs-lookup"><span data-stu-id="f81b5-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="f81b5-104">有关 **Range** 对象支持的属性和方法的完整列表，请参阅 [Range 对象 (Excel JavaScript API)](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="f81b5-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="f81b5-105">使用 Moment-MSDate 插件处理日期</span><span class="sxs-lookup"><span data-stu-id="f81b5-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="f81b5-106">[时刻 JavaScript 库](https://momentjs.com/)提供了使用日期和时间戳的便捷方式。</span><span class="sxs-lookup"><span data-stu-id="f81b5-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="f81b5-107">[Moment-MSDate 插件](https://www.npmjs.com/package/moment-msdate)可将时刻格式转换为 Excel 所需的格式。</span><span class="sxs-lookup"><span data-stu-id="f81b5-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="f81b5-108">这是 [NOW 函数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)返回的相同格式。</span><span class="sxs-lookup"><span data-stu-id="f81b5-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="f81b5-109">以下代码显示如何将 **B4** 处的范围设置为时刻的时间戳：</span><span class="sxs-lookup"><span data-stu-id="f81b5-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="f81b5-110">这是一项类似于在单元格之外获取日期并将其转换为时刻或其他格式的技术，如以下代码中所示：</span><span class="sxs-lookup"><span data-stu-id="f81b5-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="f81b5-111">你的加载项将必须对范围进行格式化才能以更可读的形式显示日期。</span><span class="sxs-lookup"><span data-stu-id="f81b5-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="f81b5-112">`"[$-409]m/d/yy h:mm AM/PM;@"` 的示例显示类似“12/3/18 3:57 PM”的时间。</span><span class="sxs-lookup"><span data-stu-id="f81b5-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="f81b5-113">有关日期和时间数字格式的详细信息，请参阅[查看自定义数字格式的准则](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)一文中的“日期和时间格式的准则”。</span><span class="sxs-lookup"><span data-stu-id="f81b5-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously-preview"></a><span data-ttu-id="f81b5-114">同时处理多个区域（预览版）</span><span class="sxs-lookup"><span data-stu-id="f81b5-114">Work with multiple ranges simultaneously (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f81b5-115">该`RangeAreas`对象当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="f81b5-115">The `RangeAreas` object is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f81b5-116">`RangeAreas` 对象允许外接程序每次在多个区域上执行操作。</span><span class="sxs-lookup"><span data-stu-id="f81b5-116">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="f81b5-117">这些区域可能但不必是连续区域。</span><span class="sxs-lookup"><span data-stu-id="f81b5-117">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="f81b5-118">`RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。</span><span class="sxs-lookup"><span data-stu-id="f81b5-118">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range-preview"></a><span data-ttu-id="f81b5-119">查找区域内特殊单元格（预览）</span><span class="sxs-lookup"><span data-stu-id="f81b5-119">Find special cells within a range (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f81b5-120">`getSpecialCells`和`getSpecialCellsOrNullObject`方法目前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="f81b5-120">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods are currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f81b5-121">`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法根据单元格特征和值类型来查找区域。</span><span class="sxs-lookup"><span data-stu-id="f81b5-121">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="f81b5-122">这两种方法都返回 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="f81b5-122">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="f81b5-123">以下是 TypeScript 数据类型文件中方法的签名：</span><span class="sxs-lookup"><span data-stu-id="f81b5-123">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="f81b5-124">下面示例使用 `getSpecialCells` 方法来查找有公式的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-124">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="f81b5-125">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="f81b5-125">About this code, note:</span></span>

- <span data-ttu-id="f81b5-126">它通过先调用 `Worksheet.getUsedRange` 并仅调用该区域的 `getSpecialCells` 来限制需要搜索的工作表部分。</span><span class="sxs-lookup"><span data-stu-id="f81b5-126">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="f81b5-127">`getSpecialCells` 方法将返回 `RangeAreas` 对象，因此包含公式的单元格都会变成粉色，即使它们并非都是连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-127">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="f81b5-128">如果区域中不存在具有目标特征的单元格，`getSpecialCells` 会引发 **ItemNotFound**错误。</span><span class="sxs-lookup"><span data-stu-id="f81b5-128">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="f81b5-129">这会将控制流转移到 `catch` 信息块（如果存在）。</span><span class="sxs-lookup"><span data-stu-id="f81b5-129">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="f81b5-130">如果不存在 `catch` 信息块，则错误会终止函数。</span><span class="sxs-lookup"><span data-stu-id="f81b5-130">If there isn't a `catch` block, the error halts the function.</span></span>

<span data-ttu-id="f81b5-131">如果你希望具有目标特征的单元格始终存在，则你可能想要代码在没有这些单元格的时候引发错误。</span><span class="sxs-lookup"><span data-stu-id="f81b5-131">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="f81b5-132">若没有匹配单元格是一个有效应用场景，代码应该会检查这种可能的情况并按正常方式处理它，而不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="f81b5-132">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="f81b5-133">可以用此 `getSpecialCellsOrNullObject` 方法及其返回的 `isNullObject` 属性实现此行为。</span><span class="sxs-lookup"><span data-stu-id="f81b5-133">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="f81b5-134">此示例使用此模式。</span><span class="sxs-lookup"><span data-stu-id="f81b5-134">The following example uses this pattern.</span></span> <span data-ttu-id="f81b5-135">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="f81b5-135">About this code, note:</span></span>

- <span data-ttu-id="f81b5-136">`getSpecialCellsOrNullObject` 方法将始终返回代理对象，因此在一般的 JavaScript 认知中，它从不为 `null`。</span><span class="sxs-lookup"><span data-stu-id="f81b5-136">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="f81b5-137">但是，如果没有找到匹配的单元格，则对象的 `isNullObject` 属性将设置为 `true`。</span><span class="sxs-lookup"><span data-stu-id="f81b5-137">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="f81b5-138">在测试 `isNullObject` 属性*之前*，它将调用 `context.sync`。</span><span class="sxs-lookup"><span data-stu-id="f81b5-138">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="f81b5-139">这是所有 `*OrNullObject` 方法和属性的要求，因为你必须始终加载和同步属性才能读取它。</span><span class="sxs-lookup"><span data-stu-id="f81b5-139">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="f81b5-140">但是，不必*明确*加载 `isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="f81b5-140">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="f81b5-141">即使未在对象上调用 `load`，`context.sync` 也会自动加载该属性。</span><span class="sxs-lookup"><span data-stu-id="f81b5-141">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="f81b5-142">有关详细信息，请参阅 [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods)。</span><span class="sxs-lookup"><span data-stu-id="f81b5-142">For more information, see [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span></span>
- <span data-ttu-id="f81b5-143">你可以测试此代码，方法是先选择没有公式单元格的区域并运行它。</span><span class="sxs-lookup"><span data-stu-id="f81b5-143">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="f81b5-144">然后选择至少包含一个带公式的单元格的区域，并再次运行它。</span><span class="sxs-lookup"><span data-stu-id="f81b5-144">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

<span data-ttu-id="f81b5-145">为简单起见，本文中的所有其他示例都使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject`。</span><span class="sxs-lookup"><span data-stu-id="f81b5-145">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="f81b5-146">通过单元格值类型缩小目标单元格的范围</span><span class="sxs-lookup"><span data-stu-id="f81b5-146">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="f81b5-147">`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法接受一个可选第二参数，用于进一步缩小目标单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-147">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="f81b5-148">此第二参数是你用于指定只希望包含特定数值类型单元格的一个 `Excel.SpecialCellValueType`。</span><span class="sxs-lookup"><span data-stu-id="f81b5-148">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="f81b5-149">当且仅当 `Excel.SpecialCellType` 为 `Excel.SpecialCellType.formulas` 或 `Excel.SpecialCellType.constants` 时才使用 `Excel.SpecialCellValueType` 参数。</span><span class="sxs-lookup"><span data-stu-id="f81b5-149">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="f81b5-150">测试单个单元格值类型</span><span class="sxs-lookup"><span data-stu-id="f81b5-150">Test for a single cell value type</span></span>

<span data-ttu-id="f81b5-151">`Excel.SpecialCellValueType` 枚举有四种基本类型 （本节后续部分所述其他组合值除外）：</span><span class="sxs-lookup"><span data-stu-id="f81b5-151">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="f81b5-152">`Excel.SpecialCellValueType.logical`（意味着布尔值）</span><span class="sxs-lookup"><span data-stu-id="f81b5-152">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="f81b5-153">以下示例查找数值常量的特殊单元格，并将这些单元格设置为粉色。</span><span class="sxs-lookup"><span data-stu-id="f81b5-153">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="f81b5-154">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="f81b5-154">About this code, note:</span></span>

- <span data-ttu-id="f81b5-155">它只会突出显示具有文本数值的单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-155">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="f81b5-156">它既不会突出显示具有公式的单元格（即使结果是数字），也不会突出显示布尔、文本或错误状态单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-156">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="f81b5-157">要测试代码，请确保工作表中的某些单元格包含文本数值，某些包含其他类型的文本值，而某些则包含公式。</span><span class="sxs-lookup"><span data-stu-id="f81b5-157">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="f81b5-158">测试多个单元格值类型</span><span class="sxs-lookup"><span data-stu-id="f81b5-158">Test for multiple cell value types</span></span>

<span data-ttu-id="f81b5-159">有时，你需要对多种单元格值类型执行操作，例如所有文本值和所有布尔值（`Excel.SpecialCellValueType.logical`）单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-159">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="f81b5-160">`Excel.SpecialCellValueType` 枚举具有组合类型的值。</span><span class="sxs-lookup"><span data-stu-id="f81b5-160">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="f81b5-161">例如，`Excel.SpecialCellValueType.logicalText` 将定向所有布尔值和所有文本值单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-161">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="f81b5-162">`Excel.SpecialCellValueType.all` 是默认值，并不限制返回的单元格值类型。</span><span class="sxs-lookup"><span data-stu-id="f81b5-162">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="f81b5-163">以下示例设置了包含用于生成数字或布尔值的公式的所有单元格颜色。</span><span class="sxs-lookup"><span data-stu-id="f81b5-163">The following example colors all cells with formulas that produce number or boolean value.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="copy-and-paste-preview"></a><span data-ttu-id="f81b5-164">复制和粘贴（预览版）</span><span class="sxs-lookup"><span data-stu-id="f81b5-164">Copy and paste (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f81b5-165">`Range.copyFrom` 函数当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="f81b5-165">The `Range.copyFrom` function is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f81b5-166">区域的 `copyFrom` 函数将复制 Excel UI 的“复制和粘贴”行为。</span><span class="sxs-lookup"><span data-stu-id="f81b5-166">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="f81b5-167">调用 `copyFrom` 的区域对象是目标。</span><span class="sxs-lookup"><span data-stu-id="f81b5-167">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="f81b5-168">将要复制的源作为一个范围或一个表示范围的字符串地址进行传递。</span><span class="sxs-lookup"><span data-stu-id="f81b5-168">The source to be copied is passed as a range or a string address representing a range.</span></span>
<span data-ttu-id="f81b5-169">以下代码示例将数据从“A1:E1”\*\*\*\* 复制到“G1”\*\*\*\* 开始的范围（粘贴到“G1:K1”\*\*\*\* 结束）。</span><span class="sxs-lookup"><span data-stu-id="f81b5-169">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="f81b5-170">`Range.copyFrom` 具有三个可选参数。</span><span class="sxs-lookup"><span data-stu-id="f81b5-170">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="f81b5-171">`copyType` 指定将哪些数据从源复制到目标。</span><span class="sxs-lookup"><span data-stu-id="f81b5-171">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="f81b5-172">`Excel.RangeCopyType.formulas` 转换源单元格中的公式，并保留这些公式范围的相对位置。</span><span class="sxs-lookup"><span data-stu-id="f81b5-172">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="f81b5-173">将原样复制任何非公式条目。</span><span class="sxs-lookup"><span data-stu-id="f81b5-173">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="f81b5-174">`Excel.RangeCopyType.values` 复制数据值，如果是公式，则复制公式的结果。</span><span class="sxs-lookup"><span data-stu-id="f81b5-174">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="f81b5-175">`Excel.RangeCopyType.formats` 复制范围的格式设置（包括字体、颜色和其他格式），但不会复制任何值。</span><span class="sxs-lookup"><span data-stu-id="f81b5-175">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="f81b5-176">`Excel.RangeCopyType.all`（默认选项）复制数据和格式设置，保留单元格的公式（如果找到）。</span><span class="sxs-lookup"><span data-stu-id="f81b5-176">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="f81b5-177">`skipBlanks` 设置是否将空白单元格复制到目标。</span><span class="sxs-lookup"><span data-stu-id="f81b5-177">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="f81b5-178">如果为 true，`copyFrom` 将跳过源范围中的空白单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-178">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="f81b5-179">跳过的单元格不会覆盖目标范围中其对应单元格的现有数据。</span><span class="sxs-lookup"><span data-stu-id="f81b5-179">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="f81b5-180">默认值为 false。</span><span class="sxs-lookup"><span data-stu-id="f81b5-180">The default is false.</span></span>

<span data-ttu-id="f81b5-181">`transpose` 确定是否将数据转置（即切换其行和列）到源位置。</span><span class="sxs-lookup"><span data-stu-id="f81b5-181">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="f81b5-182">转置范围沿主对角线翻转，因此，行“1”\*\*\*\*、“2”\*\*\*\* 和“3”\*\*\*\* 将成为列“A”\*\*\*\*、“B”\*\*\*\* 和“C”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="f81b5-182">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="f81b5-183">以下代码示例和图像在一个简单的方案中演示此行为。</span><span class="sxs-lookup"><span data-stu-id="f81b5-183">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="f81b5-184">*在上一个函数已运行之前。*</span><span class="sxs-lookup"><span data-stu-id="f81b5-184">*Before the preceding function has been run.*</span></span>

![在区域中运行复制方法之前的 Excel 中的数据](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="f81b5-186">*在上一个函数已运行之后。*</span><span class="sxs-lookup"><span data-stu-id="f81b5-186">*After the preceding function has been run.*</span></span>

![在区域中运行复制方法之后的 Excel 中的数据](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates-preview"></a><span data-ttu-id="f81b5-188">删除重复项（预览版）</span><span class="sxs-lookup"><span data-stu-id="f81b5-188">Remove duplicates (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="f81b5-189">区域对象的 `removeDuplicates` 函数当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="f81b5-189">The Range object's `removeDuplicates` function is currently available only in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="f81b5-190">区域对象的 `removeDuplicates` 函数将删除在指定列中具有重复条目的行。</span><span class="sxs-lookup"><span data-stu-id="f81b5-190">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="f81b5-191">该函数将从区域最低值索引到最高值索引（从上到下）遍历区域中的每一行。</span><span class="sxs-lookup"><span data-stu-id="f81b5-191">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="f81b5-192">如果指定列中的值之前显示在区域中，则会删除该行。</span><span class="sxs-lookup"><span data-stu-id="f81b5-192">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="f81b5-193">在区域内位于已删除行下方的行将上移。</span><span class="sxs-lookup"><span data-stu-id="f81b5-193">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="f81b5-194">`removeDuplicates` 不影响该区域外的单元格位置。</span><span class="sxs-lookup"><span data-stu-id="f81b5-194">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="f81b5-195">`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。</span><span class="sxs-lookup"><span data-stu-id="f81b5-195">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="f81b5-196">此数组从零开始并且与区域而不是与工作表相关。</span><span class="sxs-lookup"><span data-stu-id="f81b5-196">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="f81b5-197">该函数还使用一个布尔参数来指定第一行是否为标题。</span><span class="sxs-lookup"><span data-stu-id="f81b5-197">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="f81b5-198">如果为 **true**，则在考虑重复项时将忽略顶行。</span><span class="sxs-lookup"><span data-stu-id="f81b5-198">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="f81b5-199">`removeDuplicates` 函数将返回 `RemoveDuplicatesResult` 对象，用于指定已删除的行数和剩余的唯一行数。</span><span class="sxs-lookup"><span data-stu-id="f81b5-199">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="f81b5-200">在使用区域的 `removeDuplicates` 函数时，应记住以下几点：</span><span class="sxs-lookup"><span data-stu-id="f81b5-200">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="f81b5-201">`removeDuplicates` 会考虑单元格值，而不是函数结果。</span><span class="sxs-lookup"><span data-stu-id="f81b5-201">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="f81b5-202">如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。</span><span class="sxs-lookup"><span data-stu-id="f81b5-202">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="f81b5-203">`removeDuplicates` 不会忽略空单元格。</span><span class="sxs-lookup"><span data-stu-id="f81b5-203">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="f81b5-204">空单元格的值与任何其他值具有相同的处理方式。</span><span class="sxs-lookup"><span data-stu-id="f81b5-204">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="f81b5-205">这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。</span><span class="sxs-lookup"><span data-stu-id="f81b5-205">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="f81b5-206">以下示例显示删除第一列中具有重复值的条目。</span><span class="sxs-lookup"><span data-stu-id="f81b5-206">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

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

<span data-ttu-id="f81b5-207">*在上一个函数已运行之前。*</span><span class="sxs-lookup"><span data-stu-id="f81b5-207">*Before the preceding function has been run.*</span></span>

![在区域中运行删除重复项方法之前的 Excel 中的数据](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="f81b5-209">*在上一个函数已运行之后。*</span><span class="sxs-lookup"><span data-stu-id="f81b5-209">*After the preceding function has been run.*</span></span>

![在区域中运行删除重复项方法之后的 Excel 中的数据](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="f81b5-211">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f81b5-211">See also</span></span>

- [<span data-ttu-id="f81b5-212">使用 Excel JavaScript API 对区域执行操作</span><span class="sxs-lookup"><span data-stu-id="f81b5-212">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="f81b5-213">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="f81b5-213">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f81b5-214"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="f81b5-214">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
