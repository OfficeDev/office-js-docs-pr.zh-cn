---
title: 同时在 Excel 加载项中处理多个区域
description: ''
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: a327b6c379884107f5e00c0663ecfa6c71b8097f
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/21/2019
ms.locfileid: "33620043"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a><span data-ttu-id="4209e-102">同时在 Excel 加载项中处理多个区域</span><span class="sxs-lookup"><span data-stu-id="4209e-102">Work with multiple ranges simultaneously in Excel add-ins</span></span>

<span data-ttu-id="4209e-103">Excel JavaScript 库允许你使用加载项同时在多个区域上执行操作和设置属性。</span><span class="sxs-lookup"><span data-stu-id="4209e-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="4209e-104">这些区域不必是连续区域。</span><span class="sxs-lookup"><span data-stu-id="4209e-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="4209e-105">除了简化代码以外，这种设置属性的方法还比为每个区域单独设置相同的属性要快得多。</span><span class="sxs-lookup"><span data-stu-id="4209e-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

## <a name="rangeareas"></a><span data-ttu-id="4209e-106">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="4209e-106">RangeAreas</span></span>

<span data-ttu-id="4209e-107">一组（可能是不连续的）区域由一个[RangeAreas](/javascript/api/excel/excel.rangeareas)对象表示。</span><span class="sxs-lookup"><span data-stu-id="4209e-107">A set of (possibly discontiguous) ranges is represented by a [RangeAreas](/javascript/api/excel/excel.rangeareas) object.</span></span> <span data-ttu-id="4209e-108">它具有与 `Range` 类型类似的属性和方法（许多具有相同或相似的名称），但已对以下对象进行了调整：</span><span class="sxs-lookup"><span data-stu-id="4209e-108">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="4209e-109">属性和 Setter 及 Getter 行为的数据类型。</span><span class="sxs-lookup"><span data-stu-id="4209e-109">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="4209e-110">方法参数和方法行为的数据类型。</span><span class="sxs-lookup"><span data-stu-id="4209e-110">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="4209e-111">方法返回值的数据类型。</span><span class="sxs-lookup"><span data-stu-id="4209e-111">The data types of method return values.</span></span>

<span data-ttu-id="4209e-112">例如：</span><span class="sxs-lookup"><span data-stu-id="4209e-112">Some examples:</span></span>

- <span data-ttu-id="4209e-113">`RangeAreas` 具有 `address` 属性，它将返回一串以逗号分隔的区域地址，而不是像 `Range.address` 属性一样只返回一个地址。</span><span class="sxs-lookup"><span data-stu-id="4209e-113">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="4209e-114">`RangeAreas` 具有 `dataValidation` 属性，它将返回一个 `DataValidation` 对象，用来表示 `RangeAreas` 中的所有区域的数据验证（如果保持一致）。</span><span class="sxs-lookup"><span data-stu-id="4209e-114">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="4209e-115">如果相同的 `DataValidation` 对象未应用到 `RangeAreas` 中的所有区域，则该属性将为 `null`。</span><span class="sxs-lookup"><span data-stu-id="4209e-115">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="4209e-116">对于 `RangeAreas` 对象，这是一般原则，而非通用原则：*如果某个属性在 `RangeAreas` 的所有区域上没有一致的值，则该属性将为 `null`。*</span><span class="sxs-lookup"><span data-stu-id="4209e-116">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="4209e-117">请参阅[读取 RangeAreas 的属性](#read-properties-of-rangeareas)，以了解详细信息和某些例外情况。</span><span class="sxs-lookup"><span data-stu-id="4209e-117">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="4209e-118">`RangeAreas.cellCount` 将获取 `RangeAreas` 中的所有区域的单元格总数。</span><span class="sxs-lookup"><span data-stu-id="4209e-118">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="4209e-119">`RangeAreas.calculate` 将重新计算 `RangeAreas` 中的所有区域的单元格数。</span><span class="sxs-lookup"><span data-stu-id="4209e-119">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="4209e-120">`RangeAreas.getEntireColumn` 和 `RangeAreas.getEntireRow` 将返回另一个 `RangeAreas` 对象，用来表示 `RangeAreas` 中的所有区域的列数（或行数）。</span><span class="sxs-lookup"><span data-stu-id="4209e-120">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="4209e-121">例如，如果 `RangeAreas` 表示“A1:C4”和“F14:L15”，则 `RangeAreas.getEntireColumn` 将返回一个表示“A:C”和“F:L”的 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="4209e-121">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="4209e-122">`RangeAreas.copyFrom` 可以采用 `Range` 或 `RangeAreas` 参数，用来表示复制操作的源区域。</span><span class="sxs-lookup"><span data-stu-id="4209e-122">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="4209e-123">RangeAreas 还提供了区域成员的完整列表</span><span class="sxs-lookup"><span data-stu-id="4209e-123">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="4209e-124">属性</span><span class="sxs-lookup"><span data-stu-id="4209e-124">Properties</span></span>

<span data-ttu-id="4209e-125">在编写用于读取任何所列属性的代码之前，请先熟悉[读取 RangeAreas 的属性](#read-properties-of-rangeareas)。</span><span class="sxs-lookup"><span data-stu-id="4209e-125">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="4209e-126">返回的内容存在细微差别。</span><span class="sxs-lookup"><span data-stu-id="4209e-126">There are subtleties to what gets returned.</span></span>

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

##### <a name="methods"></a><span data-ttu-id="4209e-127">方法</span><span class="sxs-lookup"><span data-stu-id="4209e-127">Methods</span></span>

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- <span data-ttu-id="4209e-128">`getOffsetRange()`（在`getOffsetRangeAreas` `RangeAreas`对象上命名）</span><span class="sxs-lookup"><span data-stu-id="4209e-128">`getOffsetRange()` (named `getOffsetRangeAreas` on the `RangeAreas` object)</span></span>
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- <span data-ttu-id="4209e-129">`getUsedRange()`（在`getUsedRangeAreas` `RangeAreas`对象上命名）</span><span class="sxs-lookup"><span data-stu-id="4209e-129">`getUsedRange()` (named `getUsedRangeAreas` on the `RangeAreas` object)</span></span>
- <span data-ttu-id="4209e-130">`getUsedRangeOrNullObject()`（在`getUsedRangeAreasOrNullObject` `RangeAreas`对象上命名）</span><span class="sxs-lookup"><span data-stu-id="4209e-130">`getUsedRangeOrNullObject()` (named `getUsedRangeAreasOrNullObject` on the `RangeAreas` object)</span></span>
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="4209e-131">特定于 RangeArea 的属性和方法</span><span class="sxs-lookup"><span data-stu-id="4209e-131">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="4209e-132">`RangeAreas` 类型具有一些未包含在 `Range` 对象中的属性和方法。</span><span class="sxs-lookup"><span data-stu-id="4209e-132">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="4209e-133">以下是其中的一部分：</span><span class="sxs-lookup"><span data-stu-id="4209e-133">The following is a selection of them:</span></span>

- <span data-ttu-id="4209e-134">`areas`：一种 `RangeCollection` 对象，它包含由 `RangeAreas` 对象表示的所有区域。</span><span class="sxs-lookup"><span data-stu-id="4209e-134">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="4209e-135">`RangeCollection` 也是新对象，与其他 Excel 集合对象类似。</span><span class="sxs-lookup"><span data-stu-id="4209e-135">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="4209e-136">它具有 `items` 属性，它是一组表示区域的 `Range` 对象。</span><span class="sxs-lookup"><span data-stu-id="4209e-136">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="4209e-137">`areaCount`：`RangeAreas` 中的区域总数。</span><span class="sxs-lookup"><span data-stu-id="4209e-137">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="4209e-138">`getOffsetRangeAreas`：与 [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 的作用类似，不同之处在于，前者将返回 `RangeAreas` 并且包含多个区域，每个区域都是原始 `RangeAreas` 中的区域的偏移。</span><span class="sxs-lookup"><span data-stu-id="4209e-138">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas"></a><span data-ttu-id="4209e-139">创建 RangeAreas</span><span class="sxs-lookup"><span data-stu-id="4209e-139">Create RangeAreas</span></span>

<span data-ttu-id="4209e-140">可以通过两种基本方法创建 `RangeAreas` 对象：</span><span class="sxs-lookup"><span data-stu-id="4209e-140">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="4209e-141">调用 `Worksheet.getRanges()` 并向其传递具有以逗号分隔的区域地址的字符串。</span><span class="sxs-lookup"><span data-stu-id="4209e-141">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="4209e-142">如果要包含的任何区域已插入到 [NamedItem](/javascript/api/excel/excel.nameditem) 中，则可以在字符串中包含名称而不是地址。</span><span class="sxs-lookup"><span data-stu-id="4209e-142">If any range you want to include has been made into a [NamedItem](/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="4209e-143">调用 `Workbook.getSelectedRanges()`。</span><span class="sxs-lookup"><span data-stu-id="4209e-143">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="4209e-144">此方法将返回 `RangeAreas`，它表示在当前活动工作表上选择的所有区域。</span><span class="sxs-lookup"><span data-stu-id="4209e-144">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="4209e-145">获得 `RangeAreas` 对象后，你可以在返回 `RangeAreas` 的对象上使用该方法创建其他对象，例如 `getOffsetRangeAreas` 和 `getIntersection`。</span><span class="sxs-lookup"><span data-stu-id="4209e-145">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="4209e-146">你不能直接将其他区域添加到 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="4209e-146">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="4209e-147">例如，`RangeAreas.areas` 中的集合不具有 `add` 方法。</span><span class="sxs-lookup"><span data-stu-id="4209e-147">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>

> [!WARNING]
> <span data-ttu-id="4209e-148">不要尝试直接添加或删除 `RangeAreas.areas.items` 数组的成员。</span><span class="sxs-lookup"><span data-stu-id="4209e-148">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="4209e-149">这将导致代码中出现不需要的行为。</span><span class="sxs-lookup"><span data-stu-id="4209e-149">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="4209e-150">例如，可能会将其他 `Range` 对象推送到数组上，但这样做会导致错误，因为 `RangeAreas` 属性和方法将表现为如同新项目并不存在一样。</span><span class="sxs-lookup"><span data-stu-id="4209e-150">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="4209e-151">例如，`areaCount` 属性不包含通过这种方法推送的区域，并且如果 `index` 大于 `areasCount-1`，则 `RangeAreas.getItemAt(index)` 将引发错误。</span><span class="sxs-lookup"><span data-stu-id="4209e-151">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="4209e-152">同样，删除 `RangeAreas.areas.items` 数组中的 `Range` 对象（通过获取对它的引用并调用其 `Range.delete` 方法）也会导致错误：尽管 `Range` 对象*已被*删除，但父 `RangeAreas` 对象的属性和方法将表现为或尝试表现为如同它仍然存在一样。</span><span class="sxs-lookup"><span data-stu-id="4209e-152">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="4209e-153">例如，如果你的代码调用 `RangeAreas.calculate`，Office 将尝试计算区域，但这会引发错误，因为区域对象并不存在。</span><span class="sxs-lookup"><span data-stu-id="4209e-153">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

## <a name="set-properties-on-multiple-ranges"></a><span data-ttu-id="4209e-154">在多个区域设置属性</span><span class="sxs-lookup"><span data-stu-id="4209e-154">Set properties on multiple ranges</span></span>

<span data-ttu-id="4209e-155">在 `RangeAreas` 对象上设置属性会在 `RangeAreas.areas` 集合中的所有区域上设置相应的属性。</span><span class="sxs-lookup"><span data-stu-id="4209e-155">Setting a property on a `RangeAreas` object sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="4209e-156">以下是在多个区域上设置属性的示例。</span><span class="sxs-lookup"><span data-stu-id="4209e-156">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="4209e-157">函数将突出显示区域 **F3:F5** 和 **H3:H5**。</span><span class="sxs-lookup"><span data-stu-id="4209e-157">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="4209e-158">此示例适用于可以对传递给 `getRanges` 的区域地址进行硬编码或在运行时轻松进行计算的应用场景。</span><span class="sxs-lookup"><span data-stu-id="4209e-158">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="4209e-159">一些适用的应用场景包括：</span><span class="sxs-lookup"><span data-stu-id="4209e-159">Some of the scenarios in which this would be true include:</span></span>

- <span data-ttu-id="4209e-160">代码在已知模板的上下文中运行。</span><span class="sxs-lookup"><span data-stu-id="4209e-160">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="4209e-161">代码在导入数据的上下文中运行，其中数据架构是已知的。</span><span class="sxs-lookup"><span data-stu-id="4209e-161">The code runs in the context of imported data where the schema of the data is known.</span></span>

## <a name="get-special-cells-from-multiple-ranges"></a><span data-ttu-id="4209e-162">从多个区域获取特殊单元格</span><span class="sxs-lookup"><span data-stu-id="4209e-162">Get special cells from multiple ranges</span></span>

<span data-ttu-id="4209e-163">`RangeAreas` 对象上的 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法与 `Range` 对象上的同名方法工作原理类似。</span><span class="sxs-lookup"><span data-stu-id="4209e-163">The `getSpecialCells` and `getSpecialCellsOrNullObject` methods on the `RangeAreas` object work analogously to methods of the same name on the `Range` object.</span></span> <span data-ttu-id="4209e-164">这些方法从 `RangeAreas.areas` 集合中所有区域返回包含指定特征的单元格。</span><span class="sxs-lookup"><span data-stu-id="4209e-164">These methods return the cells with the specified characteristic from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="4209e-165">请参阅 [查找区域内特殊单元格](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) 部分了解特殊单元格更多详细信息。</span><span class="sxs-lookup"><span data-stu-id="4209e-165">See the [Find special cells within a range](excel-add-ins-ranges-advanced.md#find-special-cells-within-a-range) section for more details on special cells.</span></span>

<span data-ttu-id="4209e-166">调用 `RangeAreas` 对象上的 `getSpecialCells` 或 `getSpecialCellsOrNullObject` 方法时：</span><span class="sxs-lookup"><span data-stu-id="4209e-166">When calling the `getSpecialCells` or `getSpecialCellsOrNullObject` method on a `RangeAreas` object:</span></span>

- <span data-ttu-id="4209e-167">如果传递 `Excel.SpecialCellType.sameConditionalFormat` 作为第一个参数，该方法返回具有相同条件格式的所有单元格作为 `RangeAreas.areas` 集合第一个区域左上角单元格。</span><span class="sxs-lookup"><span data-stu-id="4209e-167">If you pass `Excel.SpecialCellType.sameConditionalFormat` as the first parameter, the method returns all cells with the same conditional formatting as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>
- <span data-ttu-id="4209e-168">如果传递 `Excel.SpecialCellType.sameDataValidation` 作为第一个参数，该方法返回具有相同数据验证规则的所有单元格作为 `RangeAreas.areas` 集合第一个区域左上角单元格。</span><span class="sxs-lookup"><span data-stu-id="4209e-168">If you pass `Excel.SpecialCellType.sameDataValidation` as the first parameter, the method returns all cells with the same data validation rule as the upper-leftmost cell in the first range in the `RangeAreas.areas` collection.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="4209e-169">读取 RangeAreas 的属性</span><span class="sxs-lookup"><span data-stu-id="4209e-169">Read properties of RangeAreas</span></span>

<span data-ttu-id="4209e-170">读取 `RangeAreas` 的属性值时须小心操作，因为对于 `RangeAreas` 内的不同区域，给定的属性可能具有不同的值。</span><span class="sxs-lookup"><span data-stu-id="4209e-170">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="4209e-171">一般规则是，如果*可以*返回一致的值，则系统会返回该值。</span><span class="sxs-lookup"><span data-stu-id="4209e-171">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="4209e-172">例如，在以下代码中，RGB 粉色代码 (`#FFC0CB`) 和 `true` 将记录到控制台，因为 `RangeAreas` 对象中的两个区域都具有粉色填充，并且都是整列。</span><span class="sxs-lookup"><span data-stu-id="4209e-172">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="4209e-173">如果无法实现一致性，则情况将变得更加复杂。</span><span class="sxs-lookup"><span data-stu-id="4209e-173">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="4209e-174">`RangeAreas` 属性的行为遵循以下三个原则：</span><span class="sxs-lookup"><span data-stu-id="4209e-174">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="4209e-175">除非所有成员区域的属性均为 true，否则 `RangeAreas` 对象的布尔属性将返回 `false`。</span><span class="sxs-lookup"><span data-stu-id="4209e-175">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="4209e-176">除非所有成员区域上的对应属性都具有相同的值，否则非布尔属性（`address` 属性除外）将返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="4209e-176">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="4209e-177">`address` 属性将返回一串以逗号分隔的成员区域地址。</span><span class="sxs-lookup"><span data-stu-id="4209e-177">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="4209e-178">例如，以下代码将创建 `RangeAreas`，其中只有一个区域是整列，并且只有一个区域具有粉色填充。</span><span class="sxs-lookup"><span data-stu-id="4209e-178">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="4209e-179">控制台将为填充颜色显示 `null`，为 `isEntireRow` 属性显示 `false`，并为 `address` 属性显示“Sheet1!F3:F5, Sheet1!H:H”（假设工作表名称为“Sheet1”）。</span><span class="sxs-lookup"><span data-stu-id="4209e-179">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="4209e-180">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4209e-180">See also</span></span>

- [<span data-ttu-id="4209e-181">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="4209e-181">Fundamental programming concepts with the Excel JavaScript API</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
- [<span data-ttu-id="4209e-182">使用 Excel JavaScript API 对区域执行操作（基本）</span><span class="sxs-lookup"><span data-stu-id="4209e-182">Work with ranges using the Excel JavaScript API (fundamental)</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="4209e-183">使用 Excel JavaScript API 对区域执行操作（高级）</span><span class="sxs-lookup"><span data-stu-id="4209e-183">Work with ranges using the Excel JavaScript API (advanced)</span></span>](excel-add-ins-ranges-advanced.md)
