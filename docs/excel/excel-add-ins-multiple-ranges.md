---
title: 同时在 Excel 加载项中处理多个区域
description: 了解 Excel JavaScript 库如何使加载项能够同时在多个区域上执行操作和设置属性。
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: 6a508d8481d9851c7f7ae98ec959fcec9663972c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609767"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a><span data-ttu-id="7aa73-103">同时在 Excel 加载项中处理多个区域</span><span class="sxs-lookup"><span data-stu-id="7aa73-103">Work with multiple ranges simultaneously in Excel add-ins</span></span>

<span data-ttu-id="7aa73-104">Excel JavaScript 库允许你使用加载项同时在多个区域上执行操作和设置属性。</span><span class="sxs-lookup"><span data-stu-id="7aa73-104">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="7aa73-105">这些区域不必是连续区域。</span><span class="sxs-lookup"><span data-stu-id="7aa73-105">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="7aa73-106">除了简化代码以外，这种设置属性的方法还比为每个区域单独设置相同的属性要快得多。</span><span class="sxs-lookup"><span data-stu-id="7aa73-106">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

## <a name="rangeareas"></a><span data-ttu-id="7aa73-107">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="7aa73-107">RangeAreas</span></span>

<span data-ttu-id="7aa73-108">一组（可能是不连续的）区域由一个[RangeAreas](/javascript/api/excel/excel.rangeareas)对象表示。</span><span class="sxs-lookup"><span data-stu-id="7aa73-108">A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object.</span></span> <span data-ttu-id="7aa73-109">它具有与 `Range` 类型类似的属性和方法（许多具有相同或相似的名称），但已对以下对象进行了调整：</span><span class="sxs-lookup"><span data-stu-id="7aa73-109">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="7aa73-110">属性和 Setter 及 Getter 行为的数据类型。</span><span class="sxs-lookup"><span data-stu-id="7aa73-110">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="7aa73-111">方法参数和方法行为的数据类型。</span><span class="sxs-lookup"><span data-stu-id="7aa73-111">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="7aa73-112">方法返回值的数据类型。</span><span class="sxs-lookup"><span data-stu-id="7aa73-112">The data types of method return values.</span></span>

<span data-ttu-id="7aa73-113">例如：</span><span class="sxs-lookup"><span data-stu-id="7aa73-113">Some examples:</span></span>

- <span data-ttu-id="7aa73-114">`RangeAreas` 具有 `address` 属性，它将返回一串以逗号分隔的区域地址，而不是像 `Range.address` 属性一样只返回一个地址。</span><span class="sxs-lookup"><span data-stu-id="7aa73-114">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="7aa73-115">`RangeAreas` 具有 `dataValidation` 属性，它将返回一个 `DataValidation` 对象，用来表示 `RangeAreas` 中的所有区域的数据验证（如果保持一致）。</span><span class="sxs-lookup"><span data-stu-id="7aa73-115">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="7aa73-116">如果相同的 `DataValidation` 对象未应用到 `RangeAreas` 中的所有区域，则该属性将为 `null`。</span><span class="sxs-lookup"><span data-stu-id="7aa73-116">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="7aa73-117">对于 `RangeAreas` 对象，这是一般原则，而非通用原则：*如果某个属性在 `RangeAreas` 的所有区域上没有一致的值，则该属性将为 `null`。*</span><span class="sxs-lookup"><span data-stu-id="7aa73-117">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="7aa73-118">请参阅[读取 RangeAreas 的属性](#read-properties-of-rangeareas)，以了解详细信息和某些例外情况。</span><span class="sxs-lookup"><span data-stu-id="7aa73-118">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="7aa73-119">`RangeAreas.cellCount` 将获取 `RangeAreas` 中的所有区域的单元格总数。</span><span class="sxs-lookup"><span data-stu-id="7aa73-119">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="7aa73-120">`RangeAreas.calculate` 将重新计算 `RangeAreas` 中的所有区域的单元格数。</span><span class="sxs-lookup"><span data-stu-id="7aa73-120">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="7aa73-121">`RangeAreas.getEntireColumn` 和 `RangeAreas.getEntireRow` 将返回另一个 `RangeAreas` 对象，用来表示 `RangeAreas` 中的所有区域的列数（或行数）。</span><span class="sxs-lookup"><span data-stu-id="7aa73-121">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="7aa73-122">例如，如果 `RangeAreas` 表示“A1:C4”和“F14:L15”，则 `RangeAreas.getEntireColumn` 将返回一个表示“A:C”和“F:L”的 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="7aa73-122">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="7aa73-123">`RangeAreas.copyFrom` 可以采用 `Range` 或 `RangeAreas` 参数，用来表示复制操作的源区域。</span><span class="sxs-lookup"><span data-stu-id="7aa73-123">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="7aa73-124">RangeAreas 还提供了区域成员的完整列表</span><span class="sxs-lookup"><span data-stu-id="7aa73-124">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="7aa73-125">属性</span><span class="sxs-lookup"><span data-stu-id="7aa73-125">Properties</span></span>

<span data-ttu-id="7aa73-126">在编写用于读取任何所列属性的代码之前，请先熟悉[读取 RangeAreas 的属性](#read-properties-of-rangeareas)。</span><span class="sxs-lookup"><span data-stu-id="7aa73-126">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="7aa73-127">返回的内容存在细微差别。</span><span class="sxs-lookup"><span data-stu-id="7aa73-127">There are subtleties to what gets returned.</span></span>

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

##### <a name="methods"></a><span data-ttu-id="7aa73-128">Methods</span><span class="sxs-lookup"><span data-stu-id="7aa73-128">Methods</span></span>

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- <span data-ttu-id="7aa73-129">`getOffsetRange()`（ `getOffsetRangeAreas` 在对象上命名 `RangeAreas` ）</span><span class="sxs-lookup"><span data-stu-id="7aa73-129">`getOffsetRange()` (named `getOffsetRangeAreas` on the `RangeAreas` object)</span></span>
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- <span data-ttu-id="7aa73-130">`getUsedRange()`（ `getUsedRangeAreas` 在对象上命名 `RangeAreas` ）</span><span class="sxs-lookup"><span data-stu-id="7aa73-130">`getUsedRange()` (named `getUsedRangeAreas` on the `RangeAreas` object)</span></span>
- <span data-ttu-id="7aa73-131">`getUsedRangeOrNullObject()`（ `getUsedRangeAreasOrNullObject` 在对象上命名 `RangeAreas` ）</span><span class="sxs-lookup"><span data-stu-id="7aa73-131">`getUsedRangeOrNullObject()` (named `getUsedRangeAreasOrNullObject` on the `RangeAreas` object)</span></span>
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="7aa73-132">特定于 RangeArea 的属性和方法</span><span class="sxs-lookup"><span data-stu-id="7aa73-132">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="7aa73-133">`RangeAreas` 类型具有一些未包含在 `Range` 对象中的属性和方法。</span><span class="sxs-lookup"><span data-stu-id="7aa73-133">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="7aa73-134">以下是其中的一部分：</span><span class="sxs-lookup"><span data-stu-id="7aa73-134">The following is a selection of them:</span></span>

- <span data-ttu-id="7aa73-135">`areas`：一种 `RangeCollection` 对象，它包含由 `RangeAreas` 对象表示的所有区域。</span><span class="sxs-lookup"><span data-stu-id="7aa73-135">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="7aa73-136">`RangeCollection` 也是新对象，与其他 Excel 集合对象类似。</span><span class="sxs-lookup"><span data-stu-id="7aa73-136">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="7aa73-137">它具有 `items` 属性，它是一组表示区域的 `Range` 对象。</span><span class="sxs-lookup"><span data-stu-id="7aa73-137">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="7aa73-138">`areaCount`：`RangeAreas` 中的区域总数。</span><span class="sxs-lookup"><span data-stu-id="7aa73-138">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="7aa73-139">`getOffsetRangeAreas`：与 [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 的作用类似，不同之处在于，前者将返回 `RangeAreas` 并且包含多个区域，每个区域都是原始 `RangeAreas` 中的区域的偏移。</span><span class="sxs-lookup"><span data-stu-id="7aa73-139">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="7aa73-140">创建 RangeAreas</span><span class="sxs-lookup"><span data-stu-id="7aa73-140">Create RangeAreas</span></span>

<span data-ttu-id="7aa73-141">可以通过两种基本方法创建 `RangeAreas` 对象：</span><span class="sxs-lookup"><span data-stu-id="7aa73-141">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="7aa73-142">调用 `Worksheet.getRanges()` 并向其传递具有以逗号分隔的区域地址的字符串。</span><span class="sxs-lookup"><span data-stu-id="7aa73-142">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="7aa73-143">如果要包含的任何区域已插入到 [NamedItem](/javascript/api/excel/excel.nameditem) 中，则可以在字符串中包含名称而不是地址。</span><span class="sxs-lookup"><span data-stu-id="7aa73-143">If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="7aa73-144">调用 `Workbook.getSelectedRanges()`。</span><span class="sxs-lookup"><span data-stu-id="7aa73-144">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="7aa73-145">此方法将返回 `RangeAreas`，它表示在当前活动工作表上选择的所有区域。</span><span class="sxs-lookup"><span data-stu-id="7aa73-145">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="7aa73-146">获得 `RangeAreas` 对象后，你可以在返回 `RangeAreas` 的对象上使用该方法创建其他对象，例如 `getOffsetRangeAreas` 和 `getIntersection`。</span><span class="sxs-lookup"><span data-stu-id="7aa73-146">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="7aa73-147">你不能直接将其他区域添加到 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="7aa73-147">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="7aa73-148">例如，`RangeAreas.areas` 中的集合不具有 `add` 方法。</span><span class="sxs-lookup"><span data-stu-id="7aa73-148">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="7aa73-149">不要尝试直接添加或删除 `RangeAreas.areas.items` 数组的成员。</span><span class="sxs-lookup"><span data-stu-id="7aa73-149">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="7aa73-150">这将导致代码中出现不需要的行为。</span><span class="sxs-lookup"><span data-stu-id="7aa73-150">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="7aa73-151">例如，可能会将其他 `Range` 对象推送到数组上，但这样做会导致错误，因为 `RangeAreas` 属性和方法将表现为如同新项目并不存在一样。</span><span class="sxs-lookup"><span data-stu-id="7aa73-151">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="7aa73-152">例如，`areaCount` 属性不包含通过这种方法推送的区域，并且如果 `index` 大于 `areasCount-1`，则 `RangeAreas.getItemAt(index)` 将引发错误。</span><span class="sxs-lookup"><span data-stu-id="7aa73-152">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="7aa73-153">同样，删除 `RangeAreas.areas.items` 数组中的 `Range` 对象（通过获取对它的引用并调用其 `Range.delete` 方法）也会导致错误：尽管 `Range` 对象*已被*删除，但父 `RangeAreas` 对象的属性和方法将表现为或尝试表现为如同它仍然存在一样。</span><span class="sxs-lookup"><span data-stu-id="7aa73-153">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="7aa73-154">例如，如果你的代码调用 `RangeAreas.calculate`，Office 将尝试计算区域，但这会引发错误，因为区域对象并不存在。</span><span class="sxs-lookup"><span data-stu-id="7aa73-154">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="7aa73-155">在多个区域设置属性</span><span class="sxs-lookup"><span data-stu-id="7aa73-155">Set properties on multiple ranges</span></span>

<span data-ttu-id="7aa73-156">在 `RangeAreas` 对象上设置属性会在 `RangeAreas.areas` 集合中的所有区域上设置相应的属性。</span><span class="sxs-lookup"><span data-stu-id="7aa73-156">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="7aa73-157">以下是在多个区域上设置属性的示例。</span><span class="sxs-lookup"><span data-stu-id="7aa73-157">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="7aa73-158">函数将突出显示区域 **F3:F5** 和 **H3:H5**。</span><span class="sxs-lookup"><span data-stu-id="7aa73-158">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="7aa73-159">此示例适用于可以对传递给 `getRanges` 的区域地址进行硬编码或在运行时轻松进行计算的应用场景。</span><span class="sxs-lookup"><span data-stu-id="7aa73-159">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="7aa73-160">一些适用的应用场景包括：</span><span class="sxs-lookup"><span data-stu-id="7aa73-160">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="7aa73-161">代码在已知模板的上下文中运行。</span><span class="sxs-lookup"><span data-stu-id="7aa73-161">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="7aa73-162">代码在导入数据的上下文中运行，其中数据架构是已知的。</span><span class="sxs-lookup"><span data-stu-id="7aa73-162">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="7aa73-163">从多个区域获取特殊单元格</span><span class="sxs-lookup"><span data-stu-id="7aa73-163">Get special cells from multiple ranges</span></span>

<span data-ttu-id="7aa73-164">`RangeAreas` 对象上的 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法与 `Range` 对象上的同名方法工作原理类似。</span><span class="sxs-lookup"><span data-stu-id="7aa73-164">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="7aa73-165">这些方法从 `RangeAreas.areas` 集合中所有区域返回包含指定特征的单元格。</span><span class="sxs-lookup"><span data-stu-id="7aa73-165">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="7aa73-166">请参阅 [查找区域内特殊单元格](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) 部分了解特殊单元格更多详细信息。</span><span class="sxs-lookup"><span data-stu-id="7aa73-166">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) section for more details on special cells.</span></span>

<span data-ttu-id="7aa73-167">调用 `RangeAreas` 对象上的 `getSpecialCells` 或 `getSpecialCellsOrNullObject` 方法时：</span><span class="sxs-lookup"><span data-stu-id="7aa73-167">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="7aa73-168">如果传递 `Excel.SpecialCellType.sameConditionalFormat` 作为第一个参数，该方法返回具有相同条件格式的所有单元格作为 `RangeAreas.areas` 集合第一个区域左上角单元格。</span><span class="sxs-lookup"><span data-stu-id="7aa73-168">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="7aa73-169">如果传递 `Excel.SpecialCellType.sameDataValidation` 作为第一个参数，该方法返回具有相同数据验证规则的所有单元格作为 `RangeAreas.areas` 集合第一个区域左上角单元格。</span><span class="sxs-lookup"><span data-stu-id="7aa73-169">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="7aa73-170">读取 RangeAreas 的属性</span><span class="sxs-lookup"><span data-stu-id="7aa73-170">Read properties of RangeAreas</span></span>

<span data-ttu-id="7aa73-171">读取 `RangeAreas` 的属性值时须小心操作，因为对于 `RangeAreas` 内的不同区域，给定的属性可能具有不同的值。</span><span class="sxs-lookup"><span data-stu-id="7aa73-171">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="7aa73-172">一般规则是，如果*可以*返回一致的值，则系统会返回该值。</span><span class="sxs-lookup"><span data-stu-id="7aa73-172">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="7aa73-173">例如，在以下代码中，RGB 粉色代码 (`#FFC0CB`) 和 `true` 将记录到控制台，因为 `RangeAreas` 对象中的两个区域都具有粉色填充，并且都是整列。</span><span class="sxs-lookup"><span data-stu-id="7aa73-173">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

<span data-ttu-id="7aa73-174">如果无法实现一致性，则情况将变得更加复杂。</span><span class="sxs-lookup"><span data-stu-id="7aa73-174">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="7aa73-175">`RangeAreas` 属性的行为遵循以下三个原则：</span><span class="sxs-lookup"><span data-stu-id="7aa73-175">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="7aa73-176">除非所有成员区域的属性均为 true，否则 `RangeAreas` 对象的布尔属性将返回 `false`。</span><span class="sxs-lookup"><span data-stu-id="7aa73-176">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="7aa73-177">除非所有成员区域上的对应属性都具有相同的值，否则非布尔属性（`address` 属性除外）将返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="7aa73-177">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="7aa73-178">`address` 属性将返回一串以逗号分隔的成员区域地址。</span><span class="sxs-lookup"><span data-stu-id="7aa73-178">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="7aa73-179">例如，以下代码将创建 `RangeAreas`，其中只有一个区域是整列，并且只有一个区域具有粉色填充。</span><span class="sxs-lookup"><span data-stu-id="7aa73-179">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="7aa73-180">控制台将为填充颜色显示 `null`，为 `isEntireRow` 属性显示 `false`，并为 `address` 属性显示“Sheet1!F3:F5, Sheet1!H:H”（假设工作表名称为“Sheet1”）。</span><span class="sxs-lookup"><span data-stu-id="7aa73-180">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a><span data-ttu-id="7aa73-181">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7aa73-181">See also</span></span>

- [<span data-ttu-id="7aa73-182">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="7aa73-182">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="7aa73-183">使用 Excel JavaScript API 对区域执行操作（基本）</span><span class="sxs-lookup"><span data-stu-id="7aa73-183">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="7aa73-184">使用 Excel JavaScript API 对区域执行操作（高级）</span><span class="sxs-lookup"><span data-stu-id="7aa73-184">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
