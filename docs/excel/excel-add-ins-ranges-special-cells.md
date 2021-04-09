---
title: 使用 Excel JavaScript API 查找区域内的特殊单元格
description: 了解如何使用 Excel JavaScript API 查找特殊单元格，例如包含公式、错误或数字的单元格。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6504873bcd8ab50bd4c03fe4f54b71d0bd920c5b
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652788"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a><span data-ttu-id="086a2-103">使用 Excel JavaScript API 查找区域内的特殊单元格</span><span class="sxs-lookup"><span data-stu-id="086a2-103">Find special cells within a range using the Excel JavaScript API</span></span>

<span data-ttu-id="086a2-104">本文提供的代码示例使用 Excel JavaScript API 查找区域内的特殊单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-104">This article provides code samples that find special cells within a range using the Excel JavaScript API.</span></span> <span data-ttu-id="086a2-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="086a2-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="find-ranges-with-special-cells"></a><span data-ttu-id="086a2-106">查找包含特殊单元格的范围</span><span class="sxs-lookup"><span data-stu-id="086a2-106">Find ranges with special cells</span></span>

<span data-ttu-id="086a2-107">[Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-)和[Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-)方法根据单元格的特征及其单元格的值类型查找区域。</span><span class="sxs-lookup"><span data-stu-id="086a2-107">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="086a2-108">这两种方法都返回 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="086a2-108">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="086a2-109">以下是 TypeScript 数据类型文件中方法的签名：</span><span class="sxs-lookup"><span data-stu-id="086a2-109">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="086a2-110">下面的代码示例使用 `getSpecialCells` 方法查找包含公式的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-110">The following code sample uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="086a2-111">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="086a2-111">About this code, note:</span></span>

- <span data-ttu-id="086a2-112">它通过先调用 `Worksheet.getUsedRange` 并仅调用该区域的 `getSpecialCells` 来限制需要搜索的工作表部分。</span><span class="sxs-lookup"><span data-stu-id="086a2-112">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="086a2-113">`getSpecialCells` 方法将返回 `RangeAreas` 对象，因此包含公式的单元格都会变成粉色，即使它们并非都是连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-113">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="086a2-114">如果区域中不存在具有目标特征的单元格，`getSpecialCells` 会引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="086a2-114">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="086a2-115">这会将控制流转移到 `catch` 信息块（如果存在）。</span><span class="sxs-lookup"><span data-stu-id="086a2-115">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="086a2-116">如果没有块， `catch` 错误将终止方法。</span><span class="sxs-lookup"><span data-stu-id="086a2-116">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="086a2-117">如果你希望具有目标特征的单元格始终存在，则你可能想要代码在没有这些单元格的时候引发错误。</span><span class="sxs-lookup"><span data-stu-id="086a2-117">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="086a2-118">若没有匹配单元格是一个有效应用场景，代码应该会检查这种可能的情况并按正常方式处理它，而不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="086a2-118">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="086a2-119">可以用此 `getSpecialCellsOrNullObject` 方法及其返回的 `isNullObject` 属性实现此行为。</span><span class="sxs-lookup"><span data-stu-id="086a2-119">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="086a2-120">下面的代码示例使用此模式。</span><span class="sxs-lookup"><span data-stu-id="086a2-120">The following code sample uses this pattern.</span></span> <span data-ttu-id="086a2-121">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="086a2-121">About this code, note:</span></span>

- <span data-ttu-id="086a2-122">`getSpecialCellsOrNullObject`方法始终返回代理对象，因此在普通 JavaScript 意义上，它 `null` 永远不会返回。</span><span class="sxs-lookup"><span data-stu-id="086a2-122">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it's never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="086a2-123">但是，如果没有找到匹配的单元格，则对象的 `isNullObject` 属性将设置为 `true`。</span><span class="sxs-lookup"><span data-stu-id="086a2-123">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="086a2-124">在测试 `isNullObject` 属性 *之前*，它将调用 `context.sync`。</span><span class="sxs-lookup"><span data-stu-id="086a2-124">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="086a2-125">这是所有 `*OrNullObject` 方法和属性的要求，因为你必须始终加载和同步属性才能读取它。</span><span class="sxs-lookup"><span data-stu-id="086a2-125">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="086a2-126">但是，不需要显式 *加载* `isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="086a2-126">However, it's not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="086a2-127">即使未在 对象上调用 `context.sync` ，它 `load` 也会自动加载。</span><span class="sxs-lookup"><span data-stu-id="086a2-127">It's automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="086a2-128">有关详细信息，请参阅[ \* OrNullObject 方法和属性](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)。</span><span class="sxs-lookup"><span data-stu-id="086a2-128">For more information, see [\*OrNullObject methods and properties](../develop/application-specific-api-model.md#ornullobject-methods-and-properties).</span></span>
- <span data-ttu-id="086a2-129">你可以测试此代码，方法是先选择没有公式单元格的区域并运行它。</span><span class="sxs-lookup"><span data-stu-id="086a2-129">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="086a2-130">然后选择至少包含一个带公式的单元格的区域，并再次运行它。</span><span class="sxs-lookup"><span data-stu-id="086a2-130">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="086a2-131">为简单起见，本文中的所有其他代码示例都 `getSpecialCells` 使用 方法而不是  `getSpecialCellsOrNullObject` 。</span><span class="sxs-lookup"><span data-stu-id="086a2-131">For simplicity, all other code samples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

## <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="086a2-132">通过单元格值类型缩小目标单元格的范围</span><span class="sxs-lookup"><span data-stu-id="086a2-132">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="086a2-133">`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法接受一个可选第二参数，用于进一步缩小目标单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-133">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="086a2-134">此第二参数是你用于指定只希望包含特定数值类型单元格的一个 `Excel.SpecialCellValueType`。</span><span class="sxs-lookup"><span data-stu-id="086a2-134">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="086a2-135">当且仅当 `Excel.SpecialCellType` 为 `Excel.SpecialCellType.formulas` 或 `Excel.SpecialCellType.constants` 时才使用 `Excel.SpecialCellValueType` 参数。</span><span class="sxs-lookup"><span data-stu-id="086a2-135">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="086a2-136">测试单个单元格值类型</span><span class="sxs-lookup"><span data-stu-id="086a2-136">Test for a single cell value type</span></span>

<span data-ttu-id="086a2-137">`Excel.SpecialCellValueType` 枚举有四种基本类型 （本节后续部分所述其他组合值除外）：</span><span class="sxs-lookup"><span data-stu-id="086a2-137">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="086a2-138">`Excel.SpecialCellValueType.logical`（意味着布尔值）</span><span class="sxs-lookup"><span data-stu-id="086a2-138">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="086a2-139">下面的代码示例查找数值常量的特殊单元格，并设置这些单元格的粉色。</span><span class="sxs-lookup"><span data-stu-id="086a2-139">The following code sample finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="086a2-140">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="086a2-140">About this code, note:</span></span>

- <span data-ttu-id="086a2-141">它只突出显示具有文字数字值的单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-141">It only highlights cells that have a literal number value.</span></span> <span data-ttu-id="086a2-142">它不会突出显示具有公式单元格 (即使结果是数字或布尔) 、文本或错误状态单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-142">It won't highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="086a2-143">要测试代码，请确保工作表中的某些单元格包含文本数值，某些包含其他类型的文本值，而某些则包含公式。</span><span class="sxs-lookup"><span data-stu-id="086a2-143">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="086a2-144">测试多个单元格值类型</span><span class="sxs-lookup"><span data-stu-id="086a2-144">Test for multiple cell value types</span></span>

<span data-ttu-id="086a2-145">有时，你需要对多种单元格值类型执行操作，例如所有文本值和所有布尔值（`Excel.SpecialCellValueType.logical`）单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-145">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="086a2-146">`Excel.SpecialCellValueType` 枚举具有组合类型的值。</span><span class="sxs-lookup"><span data-stu-id="086a2-146">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="086a2-147">例如，`Excel.SpecialCellValueType.logicalText` 将定向所有布尔值和所有文本值单元格。</span><span class="sxs-lookup"><span data-stu-id="086a2-147">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="086a2-148">`Excel.SpecialCellValueType.all` 是默认值，并不限制返回的单元格值类型。</span><span class="sxs-lookup"><span data-stu-id="086a2-148">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="086a2-149">下面的代码示例使用产生数字或布尔值的公式来设置所有单元格的颜色。</span><span class="sxs-lookup"><span data-stu-id="086a2-149">The following code sample colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="086a2-150">另请参阅</span><span class="sxs-lookup"><span data-stu-id="086a2-150">See also</span></span>

- [<span data-ttu-id="086a2-151">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="086a2-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="086a2-152">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="086a2-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="086a2-153">使用 Excel JavaScript API 查找字符串</span><span class="sxs-lookup"><span data-stu-id="086a2-153">Find a string using the Excel JavaScript API</span></span>](excel-add-ins-ranges-string-match.md)
- [<span data-ttu-id="086a2-154"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="086a2-154">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
