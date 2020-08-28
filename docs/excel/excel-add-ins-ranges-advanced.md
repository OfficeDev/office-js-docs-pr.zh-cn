---
title: 使用 Excel JavaScript API 对区域执行操作（高级）
description: 高级的 range 对象函数和方案，如特殊单元格、删除重复项以及使用日期。
ms.date: 08/26/2020
localization_priority: Normal
ms.openlocfilehash: 47f154c2bffac2e730aba21204261bc1bd536af2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294155"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="ee419-103">使用 Excel JavaScript API 对区域执行操作（高级）</span><span class="sxs-lookup"><span data-stu-id="ee419-103">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="ee419-104">本文基于[使用 Excel JavaScript API 对区域执行操作（基本）](excel-add-ins-ranges.md)中包含的信息，它提供了显示如何使用 Excel JavaScript API 对区域执行更多高级任务的代码示例。</span><span class="sxs-lookup"><span data-stu-id="ee419-104">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="ee419-105">有关该对象支持的属性和方法的完整列表 `Range` ，请参阅 [Range 对象 (适用于 Excel 的 JavaScript API) ](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="ee419-105">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="ee419-106">使用 Moment-MSDate 插件处理日期</span><span class="sxs-lookup"><span data-stu-id="ee419-106">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="ee419-107">[时刻 JavaScript 库](https://momentjs.com/)提供了使用日期和时间戳的便捷方式。</span><span class="sxs-lookup"><span data-stu-id="ee419-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="ee419-108">[Moment-MSDate 插件](https://www.npmjs.com/package/moment-msdate)可将时刻格式转换为 Excel 所需的格式。</span><span class="sxs-lookup"><span data-stu-id="ee419-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="ee419-109">这是 [NOW 函数](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)返回的相同格式。</span><span class="sxs-lookup"><span data-stu-id="ee419-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="ee419-110">以下代码显示如何将 **B4** 处的范围设置为时刻的时间戳：</span><span class="sxs-lookup"><span data-stu-id="ee419-110">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="ee419-111">这是一项类似于在单元格之外获取日期并将其转换为时刻或其他格式的技术，如以下代码中所示：</span><span class="sxs-lookup"><span data-stu-id="ee419-111">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="ee419-112">你的加载项将必须对范围进行格式化才能以更可读的形式显示日期。</span><span class="sxs-lookup"><span data-stu-id="ee419-112">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="ee419-113">`"[$-409]m/d/yy h:mm AM/PM;@"` 的示例显示类似“12/3/18 3:57 PM”的时间。</span><span class="sxs-lookup"><span data-stu-id="ee419-113">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="ee419-114">有关日期和时间数字格式的详细信息，请参阅[查看自定义数字格式的准则](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5)一文中的“日期和时间格式的准则”。</span><span class="sxs-lookup"><span data-stu-id="ee419-114">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously"></a><span data-ttu-id="ee419-115">同时处理多个区域</span><span class="sxs-lookup"><span data-stu-id="ee419-115">Work with multiple ranges simultaneously</span></span>

<span data-ttu-id="ee419-116">[RangeAreas](/javascript/api/excel/excel.rangeareas)对象允许外接程序一次在多个区域上执行操作。</span><span class="sxs-lookup"><span data-stu-id="ee419-116">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="ee419-117">这些区域可能但不必是连续区域。</span><span class="sxs-lookup"><span data-stu-id="ee419-117">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="ee419-118">`RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。</span><span class="sxs-lookup"><span data-stu-id="ee419-118">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range"></a><span data-ttu-id="ee419-119">查找区域中的特殊单元格</span><span class="sxs-lookup"><span data-stu-id="ee419-119">Find special cells within a range</span></span>

<span data-ttu-id="ee419-120">[GetSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)和[getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)方法根据单元格的特征和单元格的值类型来查找区域。</span><span class="sxs-lookup"><span data-stu-id="ee419-120">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="ee419-121">这两种方法都返回 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="ee419-121">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="ee419-122">以下是 TypeScript 数据类型文件中方法的签名：</span><span class="sxs-lookup"><span data-stu-id="ee419-122">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="ee419-123">下面示例使用 `getSpecialCells` 方法来查找有公式的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-123">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="ee419-124">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ee419-124">About this code, note:</span></span>

- <span data-ttu-id="ee419-125">它通过先调用 `Worksheet.getUsedRange` 并仅调用该区域的 `getSpecialCells` 来限制需要搜索的工作表部分。</span><span class="sxs-lookup"><span data-stu-id="ee419-125">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="ee419-126">`getSpecialCells` 方法将返回 `RangeAreas` 对象，因此包含公式的单元格都会变成粉色，即使它们并非都是连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-126">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="ee419-127">如果区域中不存在具有目标特征的单元格，`getSpecialCells` 会引发 **ItemNotFound**错误。</span><span class="sxs-lookup"><span data-stu-id="ee419-127">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="ee419-128">这会将控制流转移到 `catch` 信息块（如果存在）。</span><span class="sxs-lookup"><span data-stu-id="ee419-128">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="ee419-129">如果没有 `catch` 块，则错误停止方法。</span><span class="sxs-lookup"><span data-stu-id="ee419-129">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="ee419-130">如果你希望具有目标特征的单元格始终存在，则你可能想要代码在没有这些单元格的时候引发错误。</span><span class="sxs-lookup"><span data-stu-id="ee419-130">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="ee419-131">若没有匹配单元格是一个有效应用场景，代码应该会检查这种可能的情况并按正常方式处理它，而不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="ee419-131">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="ee419-132">可以用此 `getSpecialCellsOrNullObject` 方法及其返回的 `isNullObject` 属性实现此行为。</span><span class="sxs-lookup"><span data-stu-id="ee419-132">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="ee419-133">此示例使用此模式。</span><span class="sxs-lookup"><span data-stu-id="ee419-133">The following example uses this pattern.</span></span> <span data-ttu-id="ee419-134">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ee419-134">About this code, note:</span></span>

- <span data-ttu-id="ee419-135">`getSpecialCellsOrNullObject` 方法将始终返回代理对象，因此在一般的 JavaScript 认知中，它从不为 `null`。</span><span class="sxs-lookup"><span data-stu-id="ee419-135">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="ee419-136">但是，如果没有找到匹配的单元格，则对象的 `isNullObject` 属性将设置为 `true`。</span><span class="sxs-lookup"><span data-stu-id="ee419-136">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="ee419-137">在测试 `isNullObject` 属性*之前*，它将调用 `context.sync`。</span><span class="sxs-lookup"><span data-stu-id="ee419-137">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="ee419-138">这是所有 `*OrNullObject` 方法和属性的要求，因为你必须始终加载和同步属性才能读取它。</span><span class="sxs-lookup"><span data-stu-id="ee419-138">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="ee419-139">但是，不必*明确*加载 `isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="ee419-139">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="ee419-140">即使未在对象上调用 `load`，`context.sync` 也会自动加载该属性。</span><span class="sxs-lookup"><span data-stu-id="ee419-140">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="ee419-141">有关详细信息，请参阅[ \* OrNullObject 方法和属性](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)。</span><span class="sxs-lookup"><span data-stu-id="ee419-141">For more information, see [\*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).</span></span>
- <span data-ttu-id="ee419-142">你可以测试此代码，方法是先选择没有公式单元格的区域并运行它。</span><span class="sxs-lookup"><span data-stu-id="ee419-142">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="ee419-143">然后选择至少包含一个带公式的单元格的区域，并再次运行它。</span><span class="sxs-lookup"><span data-stu-id="ee419-143">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="ee419-144">为简单起见，本文中的所有其他示例都使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject`。</span><span class="sxs-lookup"><span data-stu-id="ee419-144">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="ee419-145">通过单元格值类型缩小目标单元格的范围</span><span class="sxs-lookup"><span data-stu-id="ee419-145">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="ee419-146">`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法接受一个可选第二参数，用于进一步缩小目标单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-146">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="ee419-147">此第二参数是你用于指定只希望包含特定数值类型单元格的一个 `Excel.SpecialCellValueType`。</span><span class="sxs-lookup"><span data-stu-id="ee419-147">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="ee419-148">当且仅当 `Excel.SpecialCellType` 为 `Excel.SpecialCellType.formulas` 或 `Excel.SpecialCellType.constants` 时才使用 `Excel.SpecialCellValueType` 参数。</span><span class="sxs-lookup"><span data-stu-id="ee419-148">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="ee419-149">测试单个单元格值类型</span><span class="sxs-lookup"><span data-stu-id="ee419-149">Test for a single cell value type</span></span>

<span data-ttu-id="ee419-150">`Excel.SpecialCellValueType` 枚举有四种基本类型 （本节后续部分所述其他组合值除外）：</span><span class="sxs-lookup"><span data-stu-id="ee419-150">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="ee419-151">`Excel.SpecialCellValueType.logical`（意味着布尔值）</span><span class="sxs-lookup"><span data-stu-id="ee419-151">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="ee419-152">以下示例查找数值常量的特殊单元格，并将这些单元格设置为粉色。</span><span class="sxs-lookup"><span data-stu-id="ee419-152">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="ee419-153">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="ee419-153">About this code, note:</span></span>

- <span data-ttu-id="ee419-154">它只会突出显示具有文本数值的单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-154">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="ee419-155">它既不会突出显示具有公式的单元格（即使结果是数字），也不会突出显示布尔、文本或错误状态单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-155">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="ee419-156">要测试代码，请确保工作表中的某些单元格包含文本数值，某些包含其他类型的文本值，而某些则包含公式。</span><span class="sxs-lookup"><span data-stu-id="ee419-156">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="ee419-157">测试多个单元格值类型</span><span class="sxs-lookup"><span data-stu-id="ee419-157">Test for multiple cell value types</span></span>

<span data-ttu-id="ee419-158">有时，你需要对多种单元格值类型执行操作，例如所有文本值和所有布尔值（`Excel.SpecialCellValueType.logical`）单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-158">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="ee419-159">`Excel.SpecialCellValueType` 枚举具有组合类型的值。</span><span class="sxs-lookup"><span data-stu-id="ee419-159">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="ee419-160">例如，`Excel.SpecialCellValueType.logicalText` 将定向所有布尔值和所有文本值单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-160">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="ee419-161">`Excel.SpecialCellValueType.all` 是默认值，并不限制返回的单元格值类型。</span><span class="sxs-lookup"><span data-stu-id="ee419-161">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="ee419-162">以下示例设置了包含用于生成数字或布尔值的公式的所有单元格颜色。</span><span class="sxs-lookup"><span data-stu-id="ee419-162">The following example colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="cut-copy-and-paste"></a><span data-ttu-id="ee419-163">剪切、复制和粘贴</span><span class="sxs-lookup"><span data-stu-id="ee419-163">Cut, copy, and paste</span></span>

### <a name="copy-and-paste"></a><span data-ttu-id="ee419-164">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="ee419-164">Copy and paste</span></span>

<span data-ttu-id="ee419-165">[CopyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)方法复制 Excel UI 的**复制**和**粘贴**操作。</span><span class="sxs-lookup"><span data-stu-id="ee419-165">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="ee419-166">调用 `copyFrom` 的区域对象是目标。</span><span class="sxs-lookup"><span data-stu-id="ee419-166">The range object that `copyFrom` is called on is the destination.</span></span> <span data-ttu-id="ee419-167">将要复制的源作为一个范围或一个表示范围的字符串地址进行传递。</span><span class="sxs-lookup"><span data-stu-id="ee419-167">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="ee419-168">以下代码示例将数据从“A1:E1”\*\*\*\* 复制到“G1”\*\*\*\* 开始的范围（粘贴到“G1:K1”\*\*\*\* 结束）。</span><span class="sxs-lookup"><span data-stu-id="ee419-168">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ee419-169">`Range.copyFrom` 具有三个可选参数。</span><span class="sxs-lookup"><span data-stu-id="ee419-169">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="ee419-170">`copyType` 指定将哪些数据从源复制到目标。</span><span class="sxs-lookup"><span data-stu-id="ee419-170">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="ee419-171">`Excel.RangeCopyType.formulas` 传输源单元格中的公式，并保留这些公式区域的相对定位。</span><span class="sxs-lookup"><span data-stu-id="ee419-171">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="ee419-172">将原样复制任何非公式条目。</span><span class="sxs-lookup"><span data-stu-id="ee419-172">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="ee419-173">`Excel.RangeCopyType.values` 复制数据值，如果是公式，则复制公式的结果。</span><span class="sxs-lookup"><span data-stu-id="ee419-173">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="ee419-174">`Excel.RangeCopyType.formats` 复制范围的格式设置（包括字体、颜色和其他格式），但不会复制任何值。</span><span class="sxs-lookup"><span data-stu-id="ee419-174">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="ee419-175">`Excel.RangeCopyType.all` (默认选项) 复制数据和格式，并保留单元格的公式（如果找到）。</span><span class="sxs-lookup"><span data-stu-id="ee419-175">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="ee419-176">`skipBlanks` 设置是否将空白单元格复制到目标。</span><span class="sxs-lookup"><span data-stu-id="ee419-176">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="ee419-177">如果为 true，`copyFrom` 将跳过源范围中的空白单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-177">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="ee419-178">跳过的单元格不会覆盖目标范围中其对应单元格的现有数据。</span><span class="sxs-lookup"><span data-stu-id="ee419-178">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="ee419-179">默认值为 false。</span><span class="sxs-lookup"><span data-stu-id="ee419-179">The default is false.</span></span>

<span data-ttu-id="ee419-180">`transpose` 确定是否将数据转置（即切换其行和列）到源位置。</span><span class="sxs-lookup"><span data-stu-id="ee419-180">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="ee419-181">转置范围沿主对角线翻转，因此，行“1”\*\*\*\*、“2”\*\*\*\* 和“3”\*\*\*\* 将成为列“A”\*\*\*\*、“B”\*\*\*\* 和“C”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="ee419-181">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="ee419-182">以下代码示例和图像在一个简单的方案中演示此行为。</span><span class="sxs-lookup"><span data-stu-id="ee419-182">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="ee419-183">*在上一个函数已运行之前。*</span><span class="sxs-lookup"><span data-stu-id="ee419-183">*Before the preceding function has been run.*</span></span>

![Excel 中的数据，在区域的复制方法已运行之前](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="ee419-185">*在上一个函数已运行之后。*</span><span class="sxs-lookup"><span data-stu-id="ee419-185">*After the preceding function has been run.*</span></span>

![Excel 中区域的复制方法已运行后的数据](../images/excel-range-copyfrom-skipblanks-after.png)

### <a name="cut-and-paste-move-cells"></a><span data-ttu-id="ee419-187">剪切并粘贴 (移动) 单元格</span><span class="sxs-lookup"><span data-stu-id="ee419-187">Cut and paste (move) cells</span></span>

<span data-ttu-id="ee419-188">该 [范围的 moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) 方法将单元格移动到工作簿中的新位置。</span><span class="sxs-lookup"><span data-stu-id="ee419-188">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="ee419-189">此单元格移动行为与单元格移动时的工作方式相同， [拖动区域边框](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) 或执行 **剪切** 和 **粘贴** 操作时。</span><span class="sxs-lookup"><span data-stu-id="ee419-189">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="ee419-190">将区域的格式和值移到指定作为参数的位置 `destinationRange` 。</span><span class="sxs-lookup"><span data-stu-id="ee419-190">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="ee419-191">下面的代码示例显示了使用方法移动的范围 `Range.moveTo` 。</span><span class="sxs-lookup"><span data-stu-id="ee419-191">The following code sample shows a range being moved with the `Range.moveTo` method.</span></span> <span data-ttu-id="ee419-192">请注意，如果目标区域小于源，它将被扩展以包含源内容。</span><span class="sxs-lookup"><span data-stu-id="ee419-192">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="remove-duplicates"></a><span data-ttu-id="ee419-193">删除重复项</span><span class="sxs-lookup"><span data-stu-id="ee419-193">Remove duplicates</span></span>

<span data-ttu-id="ee419-194">[RemoveDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)方法删除指定列中具有重复条目的行。</span><span class="sxs-lookup"><span data-stu-id="ee419-194">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="ee419-195">该方法将从最小值索引到范围 (从上到下) 的最高值索引的范围中的每一行进行遍历。</span><span class="sxs-lookup"><span data-stu-id="ee419-195">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="ee419-196">如果指定列中的值之前显示在区域中，则会删除该行。</span><span class="sxs-lookup"><span data-stu-id="ee419-196">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="ee419-197">在区域内位于已删除行下方的行将上移。</span><span class="sxs-lookup"><span data-stu-id="ee419-197">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="ee419-198">`removeDuplicates` 不影响该区域外的单元格位置。</span><span class="sxs-lookup"><span data-stu-id="ee419-198">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="ee419-199">`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。</span><span class="sxs-lookup"><span data-stu-id="ee419-199">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="ee419-200">此数组从零开始并且与区域而不是与工作表相关。</span><span class="sxs-lookup"><span data-stu-id="ee419-200">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="ee419-201">此方法还采用一个布尔参数，用于指定第一行是否为标头。</span><span class="sxs-lookup"><span data-stu-id="ee419-201">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="ee419-202">如果为 **true**，则在考虑重复项时将忽略顶行。</span><span class="sxs-lookup"><span data-stu-id="ee419-202">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="ee419-203">`removeDuplicates`方法返回一个 `RemoveDuplicatesResult` 对象，该对象指定删除的行数和剩余的唯一行数。</span><span class="sxs-lookup"><span data-stu-id="ee419-203">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="ee419-204">使用区域的方法时 `removeDuplicates` ，请记住以下几点：</span><span class="sxs-lookup"><span data-stu-id="ee419-204">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="ee419-205">`removeDuplicates` 会考虑单元格值，而不是函数结果。</span><span class="sxs-lookup"><span data-stu-id="ee419-205">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="ee419-206">如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。</span><span class="sxs-lookup"><span data-stu-id="ee419-206">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="ee419-207">`removeDuplicates` 不会忽略空单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-207">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="ee419-208">空单元格的值与任何其他值具有相同的处理方式。</span><span class="sxs-lookup"><span data-stu-id="ee419-208">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="ee419-209">这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。</span><span class="sxs-lookup"><span data-stu-id="ee419-209">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="ee419-210">以下示例显示删除第一列中具有重复值的条目。</span><span class="sxs-lookup"><span data-stu-id="ee419-210">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
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

<span data-ttu-id="ee419-211">*在上一个函数已运行之前。*</span><span class="sxs-lookup"><span data-stu-id="ee419-211">*Before the preceding function has been run.*</span></span>

![Excel 中的数据之前，已运行区域的 "删除重复方法" 方法](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="ee419-213">*在上一个函数已运行之后。*</span><span class="sxs-lookup"><span data-stu-id="ee419-213">*After the preceding function has been run.*</span></span>

![Excel 中已运行区域的 "删除重复项" 方法后 Excel 中的数据](../images/excel-ranges-remove-duplicates-after.png)

## <a name="group-data-for-an-outline"></a><span data-ttu-id="ee419-215">分级显示的组数据</span><span class="sxs-lookup"><span data-stu-id="ee419-215">Group data for an outline</span></span>

<span data-ttu-id="ee419-216">可以将区域中的行或列组合在一起，以创建 [分级显示](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)。</span><span class="sxs-lookup"><span data-stu-id="ee419-216">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="ee419-217">可以对这些组进行折叠和扩展以隐藏和显示相应的单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-217">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="ee419-218">这样可以更轻松地快速分析顶线数据。</span><span class="sxs-lookup"><span data-stu-id="ee419-218">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="ee419-219">使用 [Range](/javascript/api/excel/excel.range#group-groupoption-) 可以创建这些分级显示组。</span><span class="sxs-lookup"><span data-stu-id="ee419-219">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="ee419-220">大纲可以有层次结构，其中较小的组嵌套在更大的组下。</span><span class="sxs-lookup"><span data-stu-id="ee419-220">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="ee419-221">这样，可以在不同的级别查看大纲。</span><span class="sxs-lookup"><span data-stu-id="ee419-221">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="ee419-222">更改可见大纲级别可以通过 [showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) 方法以编程方式完成。</span><span class="sxs-lookup"><span data-stu-id="ee419-222">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="ee419-223">请注意，Excel 仅支持八种级别的分级显示组。</span><span class="sxs-lookup"><span data-stu-id="ee419-223">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="ee419-224">下面的代码示例演示如何创建一个大纲，其中包含两个级别的行和列的组。</span><span class="sxs-lookup"><span data-stu-id="ee419-224">The following code sample shows how to create an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="ee419-225">随后的图像显示该轮廓的分组。</span><span class="sxs-lookup"><span data-stu-id="ee419-225">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="ee419-226">请注意，在代码示例中，要分组的区域不包括大纲控件的行或列 (本示例的 "总计") 。</span><span class="sxs-lookup"><span data-stu-id="ee419-226">Note that in the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="ee419-227">组定义将折叠的内容，而不是控件的行或列。</span><span class="sxs-lookup"><span data-stu-id="ee419-227">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![具有两个级别的两维轮廓的范围](../images/excel-outline.png)

<span data-ttu-id="ee419-229">若要取消行或列组的分组，请使用 [Range](/javascript/api/excel/excel.range#ungroup-groupoption-) 方法。</span><span class="sxs-lookup"><span data-stu-id="ee419-229">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="ee419-230">这将从大纲中删除最外面的级别。</span><span class="sxs-lookup"><span data-stu-id="ee419-230">This removes the outermost level from the outline.</span></span> <span data-ttu-id="ee419-231">如果同一行或列类型的多个组在指定区域中的同一级别，则所有这些组都将被取消组合。</span><span class="sxs-lookup"><span data-stu-id="ee419-231">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="handle-dynamic-arrays-and-spilling-preview"></a><span data-ttu-id="ee419-232">处理动态数组和 spilling (预览) </span><span class="sxs-lookup"><span data-stu-id="ee419-232">Handle dynamic arrays and spilling (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="ee419-233">动态数组和 range spilling Api 当前处于预览阶段。</span><span class="sxs-lookup"><span data-stu-id="ee419-233">Dynamic array and range spilling APIs are currently in preview.</span></span> [!INCLUDE [Information about using preview Excel APIs](../includes/using-excel-preview-apis.md)]

<span data-ttu-id="ee419-234">有些 Excel 公式返回 [动态数组](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)。</span><span class="sxs-lookup"><span data-stu-id="ee419-234">Some Excel formulas return [Dynamic arrays](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span> <span data-ttu-id="ee419-235">这些值填充公式的原始单元格外部的多个单元格的值。</span><span class="sxs-lookup"><span data-stu-id="ee419-235">These fill the values of multiple cells outside of the formula's orginal cell.</span></span> <span data-ttu-id="ee419-236">此值溢出称为 "溢出"。</span><span class="sxs-lookup"><span data-stu-id="ee419-236">This value overflow is referred to as a "spill".</span></span> <span data-ttu-id="ee419-237">您的外接程序可以使用 getSpillingToRange 方法查找用于溢出的范围[。](/javascript/api/excel/excel.range#getspillingtorange--)</span><span class="sxs-lookup"><span data-stu-id="ee419-237">Your add-in can find the range used for a spill with the [Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) method.</span></span> <span data-ttu-id="ee419-238">此外，还有一个 [\* OrNullObject 版本](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties) `Range.getSpillingToRangeOrNullObject` 。</span><span class="sxs-lookup"><span data-stu-id="ee419-238">There is also a [\*OrNullObject version](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.</span></span>

<span data-ttu-id="ee419-239">下面的示例演示将区域的内容复制到单元格中的基本公式，这些单元格会扩散到相邻的单元格中。</span><span class="sxs-lookup"><span data-stu-id="ee419-239">The following sample shows a basic formula that copies the contents of a range into a cell, which spills into neighboring cells.</span></span> <span data-ttu-id="ee419-240">然后，外接程序会记录包含溢出的范围。</span><span class="sxs-lookup"><span data-stu-id="ee419-240">The add-in then logs the range that contains the spill.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="ee419-241">您还可以使用 getSpillParent 方法，在给定单元格 [内](/javascript/api/excel/excel.range#getspillparent--) 查找负责 spilling 的单元格。</span><span class="sxs-lookup"><span data-stu-id="ee419-241">You can also find the cell responsible for spilling into a given cell by using the [Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--) method.</span></span> <span data-ttu-id="ee419-242">请注意， `getSpillParent` 仅当 range 对象为单个单元格时才起作用。</span><span class="sxs-lookup"><span data-stu-id="ee419-242">Note that `getSpillParent` only works when the range object is a single cell.</span></span> <span data-ttu-id="ee419-243">`getSpillParent`对包含多个单元格的区域进行调用将导致 (引发错误，或返回) 的 null 范围 `Range.getSpillParentOrNullObject` 。</span><span class="sxs-lookup"><span data-stu-id="ee419-243">Calling `getSpillParent` on a range with multiple cells will result in an error being thrown (or a null range being returned for `Range.getSpillParentOrNullObject`).</span></span>

## <a name="see-also"></a><span data-ttu-id="ee419-244">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ee419-244">See also</span></span>

- [<span data-ttu-id="ee419-245">使用 Excel JavaScript API 对区域执行操作</span><span class="sxs-lookup"><span data-stu-id="ee419-245">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="ee419-246">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="ee419-246">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ee419-247"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="ee419-247">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
