---
title: 同时在 Excel 加载项中处理多个区域
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: 37f9c8a9f3127d78e1cc794aea9e6d1502cdeaf9
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270976"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="27d1b-102">同时在 Excel 加载项中处理多个区域（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="27d1b-103">Excel JavaScript 库允许你使用加载项同时在多个区域上执行操作和设置属性。</span><span class="sxs-lookup"><span data-stu-id="27d1b-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="27d1b-104">这些区域不必是连续区域。</span><span class="sxs-lookup"><span data-stu-id="27d1b-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="27d1b-105">除了简化代码以外，这种设置属性的方法还比为每个区域单独设置相同的属性要快得多。</span><span class="sxs-lookup"><span data-stu-id="27d1b-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="27d1b-106">本文中所述的 API 需要 **Office 2016 即点即用版本 1809 内部版本 10820.20000** 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="27d1b-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="27d1b-107">（您可能需要加入 [Office 预览体验成员计划](https://products.office.com/office-insider)才能获取相应的内部版本。）此外，您还必须从 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 加载 Office JavaScript 库的 Beta 版本。</span><span class="sxs-lookup"><span data-stu-id="27d1b-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="27d1b-108">最后，我们还没有提供这些 API 的参考页面。</span><span class="sxs-lookup"><span data-stu-id="27d1b-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="27d1b-109">但是，以下定义类型文件提供了相关说明：[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="27d1b-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="27d1b-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="27d1b-110">RangeAreas</span></span>

<span data-ttu-id="27d1b-111">一组区域（可能是非连续区域）由 `Excel.RangeAreas` 对象表示。</span><span class="sxs-lookup"><span data-stu-id="27d1b-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="27d1b-112">它具有与 `Range` 类型类似的属性和方法（许多具有相同或相似的名称），但已对以下对象进行了调整：</span><span class="sxs-lookup"><span data-stu-id="27d1b-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="27d1b-113">属性和 Setter 及 Getter 行为的数据类型。</span><span class="sxs-lookup"><span data-stu-id="27d1b-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="27d1b-114">方法参数和方法行为的数据类型。</span><span class="sxs-lookup"><span data-stu-id="27d1b-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="27d1b-115">方法返回值的数据类型。</span><span class="sxs-lookup"><span data-stu-id="27d1b-115">The data types of method return values.</span></span>

<span data-ttu-id="27d1b-116">例如：</span><span class="sxs-lookup"><span data-stu-id="27d1b-116">Some examples:</span></span>

- <span data-ttu-id="27d1b-117">`RangeAreas` 具有 `address` 属性，它将返回一串以逗号分隔的区域地址，而不是像 `Range.address` 属性一样只返回一个地址。</span><span class="sxs-lookup"><span data-stu-id="27d1b-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="27d1b-118">`RangeAreas` 具有 `dataValidation` 属性，它将返回一个 `DataValidation` 对象，用来表示 `RangeAreas` 中的所有区域的数据验证（如果保持一致）。</span><span class="sxs-lookup"><span data-stu-id="27d1b-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="27d1b-119">如果相同的 `DataValidation` 对象未应用到 `RangeAreas` 中的所有区域，则该属性将为 `null`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="27d1b-120">对于 `RangeAreas` 对象，这是一般原则，而非通用原则：*如果某个属性在 `RangeAreas` 的所有区域上没有一致的值，则该属性将为 `null`。*</span><span class="sxs-lookup"><span data-stu-id="27d1b-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="27d1b-121">请参阅[读取 RangeAreas 的属性](#read-properties-of-rangeareas)，以了解详细信息和某些例外情况。</span><span class="sxs-lookup"><span data-stu-id="27d1b-121">See [Read properties of RangeAreas](#read-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="27d1b-122">`RangeAreas.cellCount` 将获取 `RangeAreas` 中的所有区域的单元格总数。</span><span class="sxs-lookup"><span data-stu-id="27d1b-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="27d1b-123">`RangeAreas.calculate` 将重新计算 `RangeAreas` 中的所有区域的单元格数。</span><span class="sxs-lookup"><span data-stu-id="27d1b-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="27d1b-124">`RangeAreas.getEntireColumn` 和 `RangeAreas.getEntireRow` 将返回另一个 `RangeAreas` 对象，用来表示 `RangeAreas` 中的所有区域的列数（或行数）。</span><span class="sxs-lookup"><span data-stu-id="27d1b-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="27d1b-125">例如，如果 `RangeAreas` 表示“A1:C4”和“F14:L15”，则 `RangeAreas.getEntireColumn` 将返回一个表示“A:C”和“F:L”的 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="27d1b-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="27d1b-126">`RangeAreas.copyFrom` 可以采用 `Range` 或 `RangeAreas` 参数，用来表示复制操作的源区域。</span><span class="sxs-lookup"><span data-stu-id="27d1b-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="27d1b-127">RangeAreas 还提供了区域成员的完整列表</span><span class="sxs-lookup"><span data-stu-id="27d1b-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="27d1b-128">属性</span><span class="sxs-lookup"><span data-stu-id="27d1b-128">Properties</span></span>

<span data-ttu-id="27d1b-129">在编写用于读取任何所列属性的代码之前，请先熟悉[读取 RangeAreas 的属性](#read-properties-of-rangeareas)。</span><span class="sxs-lookup"><span data-stu-id="27d1b-129">Be familiar with [Read properties of RangeAreas](#read-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="27d1b-130">返回的内容存在细微差别。</span><span class="sxs-lookup"><span data-stu-id="27d1b-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="27d1b-131">address</span><span class="sxs-lookup"><span data-stu-id="27d1b-131">address</span></span>
- <span data-ttu-id="27d1b-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="27d1b-132">addressLocal</span></span>
- <span data-ttu-id="27d1b-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="27d1b-133">cellCount</span></span>
- <span data-ttu-id="27d1b-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="27d1b-134">conditionalFormats</span></span>
- <span data-ttu-id="27d1b-135">context</span><span class="sxs-lookup"><span data-stu-id="27d1b-135">context</span></span>
- <span data-ttu-id="27d1b-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="27d1b-136">dataValidation</span></span>
- <span data-ttu-id="27d1b-137">format</span><span class="sxs-lookup"><span data-stu-id="27d1b-137">format</span></span>
- <span data-ttu-id="27d1b-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="27d1b-138">isEntireColumn</span></span>
- <span data-ttu-id="27d1b-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="27d1b-139">isEntireRow</span></span>
- <span data-ttu-id="27d1b-140">style</span><span class="sxs-lookup"><span data-stu-id="27d1b-140">style</span></span>
- <span data-ttu-id="27d1b-141">worksheet</span><span class="sxs-lookup"><span data-stu-id="27d1b-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="27d1b-142">Methods</span><span class="sxs-lookup"><span data-stu-id="27d1b-142">Methods</span></span>

<span data-ttu-id="27d1b-143">将标记预览中的区域方法。</span><span class="sxs-lookup"><span data-stu-id="27d1b-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="27d1b-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="27d1b-144">calculate()</span></span>
- <span data-ttu-id="27d1b-145">clear()</span><span class="sxs-lookup"><span data-stu-id="27d1b-145">clear()</span></span>
- <span data-ttu-id="27d1b-146">convertDataTypeToText()（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="27d1b-147">convertToLinkedDataType()（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="27d1b-148">copyFrom()（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="27d1b-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="27d1b-149">getEntireColumn()</span></span>
- <span data-ttu-id="27d1b-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="27d1b-150">getEntireRow()</span></span>
- <span data-ttu-id="27d1b-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="27d1b-151">getIntersection()</span></span>
- <span data-ttu-id="27d1b-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="27d1b-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="27d1b-153">getOffsetRange()（在 RangeAreas 对象上名为 getOffsetRangeAreas）</span><span class="sxs-lookup"><span data-stu-id="27d1b-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="27d1b-154">getSpecialCells()（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="27d1b-155">getSpecialCellsOrNullObject()（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="27d1b-156">getTables()（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-156">getTables() (preview)</span></span>
- <span data-ttu-id="27d1b-157">getUsedRange()（在 RangeAreas 对象上名为 getUsedRangeAreas）</span><span class="sxs-lookup"><span data-stu-id="27d1b-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="27d1b-158">getUsedRangeOrNullObject()（在 RangeAreas 对象上名为 getUsedRangeAreasOrNullObject）</span><span class="sxs-lookup"><span data-stu-id="27d1b-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="27d1b-159">load()</span><span class="sxs-lookup"><span data-stu-id="27d1b-159">load()</span></span>
- <span data-ttu-id="27d1b-160">set()</span><span class="sxs-lookup"><span data-stu-id="27d1b-160">Set</span></span>
- <span data-ttu-id="27d1b-161">setDirty()（预览）</span><span class="sxs-lookup"><span data-stu-id="27d1b-161">setDirty() (preview)</span></span>
- <span data-ttu-id="27d1b-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="27d1b-162">toJSON()</span></span>
- <span data-ttu-id="27d1b-163">track()</span><span class="sxs-lookup"><span data-stu-id="27d1b-163">track</span></span>
- <span data-ttu-id="27d1b-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="27d1b-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="27d1b-165">特定于 RangeArea 的属性和方法</span><span class="sxs-lookup"><span data-stu-id="27d1b-165">Language-specific properties and methods</span></span>

<span data-ttu-id="27d1b-166">`RangeAreas` 类型具有一些未包含在 `Range` 对象中的属性和方法。</span><span class="sxs-lookup"><span data-stu-id="27d1b-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object.</span></span> <span data-ttu-id="27d1b-167">以下是其中的一部分：</span><span class="sxs-lookup"><span data-stu-id="27d1b-167">The following is a selection of them:</span></span>

- <span data-ttu-id="27d1b-168">`areas`：一种 `RangeCollection` 对象，它包含由 `RangeAreas` 对象表示的所有区域。</span><span class="sxs-lookup"><span data-stu-id="27d1b-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="27d1b-169">`RangeCollection` 也是新对象，与其他 Excel 集合对象类似。</span><span class="sxs-lookup"><span data-stu-id="27d1b-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="27d1b-170">它具有 `items` 属性，它是一组表示区域的 `Range` 对象。</span><span class="sxs-lookup"><span data-stu-id="27d1b-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="27d1b-171">`areaCount`：`RangeAreas` 中的区域总数。</span><span class="sxs-lookup"><span data-stu-id="27d1b-171">`areaCount`: The total number of ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="27d1b-172">`getOffsetRangeAreas`：与 [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 的作用类似，不同之处在于，前者将返回 `RangeAreas` 并且包含多个区域，每个区域都是原始 `RangeAreas` 中的区域的偏移。</span><span class="sxs-lookup"><span data-stu-id="27d1b-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="27d1b-173">创建 RangeAreas 并设置属性</span><span class="sxs-lookup"><span data-stu-id="27d1b-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="27d1b-174">可以通过两种基本方法创建 `RangeAreas` 对象：</span><span class="sxs-lookup"><span data-stu-id="27d1b-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="27d1b-175">调用 `Worksheet.getRanges()` 并向其传递具有以逗号分隔的区域地址的字符串。</span><span class="sxs-lookup"><span data-stu-id="27d1b-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="27d1b-176">如果要包含的任何区域已插入到 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) 中，则可以在字符串中包含名称而不是地址。</span><span class="sxs-lookup"><span data-stu-id="27d1b-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="27d1b-177">调用 `Workbook.getSelectedRanges()`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="27d1b-178">此方法将返回 `RangeAreas`，它表示在当前活动工作表上选择的所有区域。</span><span class="sxs-lookup"><span data-stu-id="27d1b-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="27d1b-179">获得 `RangeAreas` 对象后，你可以在返回 `RangeAreas` 的对象上使用该方法创建其他对象，例如 `getOffsetRangeAreas` 和 `getIntersection`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="27d1b-180">你不能直接将其他区域添加到 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="27d1b-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="27d1b-181">例如，`RangeAreas.areas` 中的集合不具有 `add` 方法。</span><span class="sxs-lookup"><span data-stu-id="27d1b-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="27d1b-182">不要尝试直接添加或删除 `RangeAreas.areas.items` 数组的成员。</span><span class="sxs-lookup"><span data-stu-id="27d1b-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="27d1b-183">这将导致代码中出现不需要的行为。</span><span class="sxs-lookup"><span data-stu-id="27d1b-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="27d1b-184">例如，可能会将其他 `Range` 对象推送到数组上，但这样做会导致错误，因为 `RangeAreas` 属性和方法将表现为如同新项目并不存在一样。</span><span class="sxs-lookup"><span data-stu-id="27d1b-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="27d1b-185">例如，`areaCount` 属性不包含通过这种方法推送的区域，并且如果 `index` 大于 `areasCount-1`，则 `RangeAreas.getItemAt(index)` 将引发错误。</span><span class="sxs-lookup"><span data-stu-id="27d1b-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="27d1b-186">同样，删除 `RangeAreas.areas.items` 数组中的 `Range` 对象（通过获取对它的引用并调用其 `Range.delete` 方法）也会导致错误：尽管 `Range` 对象*已被*删除，但父 `RangeAreas` 对象的属性和方法将表现为或尝试表现为如同它仍然存在一样。</span><span class="sxs-lookup"><span data-stu-id="27d1b-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="27d1b-187">例如，如果你的代码调用 `RangeAreas.calculate`，Office 将尝试计算区域，但这会引发错误，因为区域对象并不存在。</span><span class="sxs-lookup"><span data-stu-id="27d1b-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="27d1b-188">在 `RangeAreas` 上设置属性会在 `RangeAreas.areas` 集合中的所有区域上设置相应的属性。</span><span class="sxs-lookup"><span data-stu-id="27d1b-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="27d1b-189">以下是在多个区域上设置属性的示例。</span><span class="sxs-lookup"><span data-stu-id="27d1b-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="27d1b-190">函数将突出显示区域 **F3:F5** 和 **H3:H5**。</span><span class="sxs-lookup"><span data-stu-id="27d1b-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="27d1b-191">此示例适用于可以对传递给 `getRanges` 的区域地址进行硬编码或在运行时轻松进行计算的应用场景。</span><span class="sxs-lookup"><span data-stu-id="27d1b-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="27d1b-192">一些适用的应用场景包括：</span><span class="sxs-lookup"><span data-stu-id="27d1b-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="27d1b-193">代码在已知模板的上下文中运行。</span><span class="sxs-lookup"><span data-stu-id="27d1b-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="27d1b-194">代码在导入数据的上下文中运行，其中数据架构是已知的。</span><span class="sxs-lookup"><span data-stu-id="27d1b-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="27d1b-195">如果你在编码时不知道要对哪个区域执行操作，则必须在运行时发现它们。</span><span class="sxs-lookup"><span data-stu-id="27d1b-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="27d1b-196">下一节将讨论这些应用场景。</span><span class="sxs-lookup"><span data-stu-id="27d1b-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="27d1b-197">以编程方式发现区域</span><span class="sxs-lookup"><span data-stu-id="27d1b-197">Discover range areas programmatically</span></span>

<span data-ttu-id="27d1b-198">`Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法使你能够在运行时根据单元格特征和单元格的值类型查找要对其执行操作的区域。</span><span class="sxs-lookup"><span data-stu-id="27d1b-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="27d1b-199">以下是 TypeScript 数据类型文件中的方法签名：</span><span class="sxs-lookup"><span data-stu-id="27d1b-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="27d1b-200">以下是使用第一种方法的示例。</span><span class="sxs-lookup"><span data-stu-id="27d1b-200">The following is an example of using the "Between" operator:</span></span> <span data-ttu-id="27d1b-201">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="27d1b-201">About this code, note:</span></span>

- <span data-ttu-id="27d1b-202">它通过先调用 `Worksheet.getUsedRange` 并仅调用该区域的 `getSpecialCells` 来限制需要搜索的工作表部分。</span><span class="sxs-lookup"><span data-stu-id="27d1b-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="27d1b-203">它以参数形式传递给 `getSpecialCells`，即 `Excel.SpecialCellType` 枚举中的值的字符串版本。</span><span class="sxs-lookup"><span data-stu-id="27d1b-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="27d1b-204">某些可以传递的其他值包括：适用于空单元格的“空白”，适用于包含文本值而不是公式的单元格的“常量”，以及适用于与 `usedRange` 中的第一个单元格具有相同条件格式的单元格的“SameConditionalFormat”。</span><span class="sxs-lookup"><span data-stu-id="27d1b-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="27d1b-205">第一个单元格是指最左上角的单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="27d1b-206">有关枚举中的值的完整列表，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="27d1b-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="27d1b-207">`getSpecialCells` 方法将返回 `RangeAreas` 对象，因此包含公式的单元格都会变成粉色，即使它们并非都是连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="27d1b-208">有时，区域未包含*任何*具有目标特征的单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="27d1b-209">如果 `getSpecialCells` 未找到任何单元格，它将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="27d1b-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="27d1b-210">这会将控制流转移到 `catch` 信息块/方法（如果存在）。</span><span class="sxs-lookup"><span data-stu-id="27d1b-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="27d1b-211">如果不存在，则该错误会终止函数。</span><span class="sxs-lookup"><span data-stu-id="27d1b-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="27d1b-212">可能存在这样的应用场景，即你正好希望当不存在具有目标特征的单元格时引发错误。</span><span class="sxs-lookup"><span data-stu-id="27d1b-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="27d1b-213">但在没有匹配单元格（这是正常现象，但可能并不常见）的应用场景中，你的代码应该会检查这种可能的情况并按正常方式处理它，而不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="27d1b-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="27d1b-214">对于这些应用场景，请使用 `getSpecialCellsOrNullObject` 方法并测试 `RangeAreas.isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="27d1b-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="27d1b-215">示例如下。</span><span class="sxs-lookup"><span data-stu-id="27d1b-215">The following is an example.</span></span> <span data-ttu-id="27d1b-216">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="27d1b-216">Note about this code:</span></span>

- <span data-ttu-id="27d1b-217">`getSpecialCellsOrNullObject` 方法将始终返回代理对象，因此在一般的 JavaScript 认知中，它从不为 `null`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="27d1b-218">但是，如果没有找到匹配的单元格，则对象的 `isNullObject` 属性将设置为 `true`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="27d1b-219">在测试 `isNullObject` 属性*之前*，它将调用 `context.sync`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="27d1b-220">这是所有 `*OrNullObject` 方法和属性的要求，因为你必须始终加载和同步属性才能读取它。</span><span class="sxs-lookup"><span data-stu-id="27d1b-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="27d1b-221">但是，不必*明确*加载 `isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="27d1b-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="27d1b-222">即使未在对象上调用 `load`，`context.sync` 也会自动加载该属性。</span><span class="sxs-lookup"><span data-stu-id="27d1b-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="27d1b-223">有关详细信息，请参阅 [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)。</span><span class="sxs-lookup"><span data-stu-id="27d1b-223">For more information, see [\*](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods) object .</span></span>
- <span data-ttu-id="27d1b-224">你可以测试此代码，方法是先选择没有公式单元格的区域并运行它。</span><span class="sxs-lookup"><span data-stu-id="27d1b-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="27d1b-225">然后选择至少包含一个带公式的单元格的区域，并再次运行它。</span><span class="sxs-lookup"><span data-stu-id="27d1b-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
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

<span data-ttu-id="27d1b-226">为简单起见，本文中的所有其他示例都使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="27d1b-227">通过单元格值类型缩小目标单元格的范围</span><span class="sxs-lookup"><span data-stu-id="27d1b-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="27d1b-228">有一个枚举类型为 `Excel.SpecialCellValueType` 的第二个可选参数，它可以进一步缩小要定位的单元格范围。</span><span class="sxs-lookup"><span data-stu-id="27d1b-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="27d1b-229">仅当将“公式”或“常量”传递给 `getSpecialCells` 或 `getSpecialCellsOrNullObject` 时，才能使用它。</span><span class="sxs-lookup"><span data-stu-id="27d1b-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="27d1b-230">该参数指定你只需要具有特定值类型的单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="27d1b-231">有四种基本类型：“错误”、“逻辑”（它表示布尔）、“数字”和“文本”。</span><span class="sxs-lookup"><span data-stu-id="27d1b-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="27d1b-232">（除了这四种类型以外，枚举还具有其他值，将在下文对此展开讨论。）以下是一个示例。</span><span class="sxs-lookup"><span data-stu-id="27d1b-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="27d1b-233">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="27d1b-233">About this code, note:</span></span>

- <span data-ttu-id="27d1b-234">它只会突出显示具有文本数值的单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="27d1b-235">它既不会突出显示具有公式的单元格（即使结果是数字），也不会突出显示布尔、文本或错误状态单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="27d1b-236">要测试代码，请确保工作表中的某些单元格包含文本数值，某些包含其他类型的文本值，而某些则包含公式。</span><span class="sxs-lookup"><span data-stu-id="27d1b-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="27d1b-237">有时，你需要对多种单元格值类型执行操作，例如所有文本值和所有布尔值（“逻辑”）单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="27d1b-238">`Excel.SpecialCellValueType` 枚举提供的值允许你组合不同的类型。</span><span class="sxs-lookup"><span data-stu-id="27d1b-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="27d1b-239">例如，“LogicalText”将定向所有布尔值和所有文本值单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="27d1b-240">你可以组合四种基本类型中的任意两种或任意三种。</span><span class="sxs-lookup"><span data-stu-id="27d1b-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="27d1b-241">这些用于组合基本类型的枚举值的名称始终按字母顺序排列。</span><span class="sxs-lookup"><span data-stu-id="27d1b-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="27d1b-242">因此，要组合错误值、文本值和布尔值单元格，请使用“ErrorLogicalText”，而不是“LogicalErrorText”或“TextErrorLogical”。</span><span class="sxs-lookup"><span data-stu-id="27d1b-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="27d1b-243">默认参数“全部”将组合所有四种类型。</span><span class="sxs-lookup"><span data-stu-id="27d1b-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="27d1b-244">以下示例突出显示包含用于生成数字或布尔值的公式的所有单元格：</span><span class="sxs-lookup"><span data-stu-id="27d1b-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> <span data-ttu-id="27d1b-245">仅当 `Excel.SpecialCellType` 参数为“公式”或“常量”时，才能使用 `Excel.SpecialCellValueType` 参数。</span><span class="sxs-lookup"><span data-stu-id="27d1b-245">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` parameter is "Formulas" or "Constants".</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="27d1b-246">获取 RangeAreas 内的 RangeAreas</span><span class="sxs-lookup"><span data-stu-id="27d1b-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="27d1b-247">`RangeAreas` 类型本身也具有 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法，它们采用两个相同的参数。</span><span class="sxs-lookup"><span data-stu-id="27d1b-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="27d1b-248">这些方法将返回 `RangeAreas.areas` 集合中所有区域的所有目标单元格。</span><span class="sxs-lookup"><span data-stu-id="27d1b-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="27d1b-249">在 `RangeAreas` 对象而不是 `Range` 对象上调用时，这些方法的行为存在一个小差异：如果将“SameConditionalFormat”作为第一个参数进行传递，则该方法将返回与 `RangeAreas.areas` 集合中第一个区域*最左上角的单元格具有相同条件格式的所有单元格*。</span><span class="sxs-lookup"><span data-stu-id="27d1b-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="27d1b-250">这一点也适用于“SameDataValidation”：如果传递给 `Range.getSpecialCells`，它将返回与区域最左上角的单元格*具有相同数据验证规则的所有单元格*。</span><span class="sxs-lookup"><span data-stu-id="27d1b-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="27d1b-251">但是，如果传递给 `RangeAreas.getSpecialCells`，它将返回与 `RangeAreas.areas` 集合中第一个区域*最左上角的单元格具有相同数据验证规则的所有单元格*。</span><span class="sxs-lookup"><span data-stu-id="27d1b-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="27d1b-252">读取 RangeAreas 的属性</span><span class="sxs-lookup"><span data-stu-id="27d1b-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="27d1b-253">读取 `RangeAreas` 的属性值时须小心操作，因为对于 `RangeAreas` 内的不同区域，给定的属性可能具有不同的值。</span><span class="sxs-lookup"><span data-stu-id="27d1b-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="27d1b-254">一般规则是，如果*可以*返回一致的值，则系统会返回该值。</span><span class="sxs-lookup"><span data-stu-id="27d1b-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="27d1b-255">例如，在以下代码中，RGB 粉色代码 (`#FFC0CB`) 和 `true` 将记录到控制台，因为 `RangeAreas` 对象中的两个区域都具有粉色填充，并且都是整列。</span><span class="sxs-lookup"><span data-stu-id="27d1b-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
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

<span data-ttu-id="27d1b-256">如果无法实现一致性，则情况将变得更加复杂。</span><span class="sxs-lookup"><span data-stu-id="27d1b-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="27d1b-257">`RangeAreas` 属性的行为遵循以下三个原则：</span><span class="sxs-lookup"><span data-stu-id="27d1b-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="27d1b-258">除非所有成员区域的属性均为 true，否则 `RangeAreas` 对象的布尔属性将返回 `false`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="27d1b-259">除非所有成员区域上的对应属性都具有相同的值，否则非布尔属性（`address` 属性除外）将返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="27d1b-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="27d1b-260">`address` 属性将返回一串以逗号分隔的成员区域地址。</span><span class="sxs-lookup"><span data-stu-id="27d1b-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="27d1b-261">例如，以下代码将创建 `RangeAreas`，其中只有一个区域是整列，并且只有一个区域具有粉色填充。</span><span class="sxs-lookup"><span data-stu-id="27d1b-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="27d1b-262">控制台将为填充颜色显示 `null`，为 `isEntireRow` 属性显示 `false`，并为 `address` 属性显示“Sheet1!F3:F5, Sheet1!H:H”（假设工作表名称为“Sheet1”）。</span><span class="sxs-lookup"><span data-stu-id="27d1b-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
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

## <a name="see-also"></a><span data-ttu-id="27d1b-263">另请参阅</span><span class="sxs-lookup"><span data-stu-id="27d1b-263">See also</span></span>

- [<span data-ttu-id="27d1b-264">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="27d1b-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="27d1b-265">Range 对象（适用于 Excel 的 JavaScript API）</span><span class="sxs-lookup"><span data-stu-id="27d1b-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="27d1b-266">[Range 对象（适用于 Excel 的 JavaScript API）](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas)（当 API 处于预览状态时，此链接可能无效。</span><span class="sxs-lookup"><span data-stu-id="27d1b-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="27d1b-267">作为替代方法，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。)</span><span class="sxs-lookup"><span data-stu-id="27d1b-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>