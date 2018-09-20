---
title: 在 Excel 加载项中同时处理多个范围
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: bcb14d1f4c015fe675c2d65cb5f1198d485dd4c5
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016456"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="c8f56-102">在 Excel 加载项同时处理多个范围（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="c8f56-103">Excel JavaScript 库使加载项能够执行操作，并同时在多个范围设置属性。</span><span class="sxs-lookup"><span data-stu-id="c8f56-103">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously.</span></span> <span data-ttu-id="c8f56-104">这些范围不需要连续。</span><span class="sxs-lookup"><span data-stu-id="c8f56-104">The ranges do not have to be contiguous.</span></span> <span data-ttu-id="c8f56-105">这样的设置属性方式除了能简化您的代码，其运行速度也快于单独地为每个范围设置相同属性。</span><span class="sxs-lookup"><span data-stu-id="c8f56-105">In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="c8f56-106">本文中介绍的 API 要求 **Office 2016 即点即用版本 1809 内部版本 10820.20000** 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="c8f56-106">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later.</span></span> <span data-ttu-id="c8f56-107">（您可能需要加入 [Office 预览体验计划](https://products.office.com/office-insider) 以获取适当的版本。）此外，您必须从 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)加载 beta 版本的 Office JavaScript 库。</span><span class="sxs-lookup"><span data-stu-id="c8f56-107">(You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="c8f56-108">最后，我们尚未 为这些 API 提供参考页。</span><span class="sxs-lookup"><span data-stu-id="c8f56-108">Finally, we don't have reference pages for these APIs yet.</span></span> <span data-ttu-id="c8f56-109">但下面的定义类型文件提供 了相关说明：[beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="c8f56-109">But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="c8f56-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c8f56-110">RangeAreas</span></span>

<span data-ttu-id="c8f56-111">一组（可能不连续的）范围以 `Excel.RangeAreas` 对象表示。</span><span class="sxs-lookup"><span data-stu-id="c8f56-111">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object.</span></span> <span data-ttu-id="c8f56-112">其属性和方法类似于 `Range` 类型 （许多拥有相同或类似的名称），但也存在以下调整：</span><span class="sxs-lookup"><span data-stu-id="c8f56-112">It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="c8f56-113">属性的数据类型以及 getter 和 setter的行为。</span><span class="sxs-lookup"><span data-stu-id="c8f56-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="c8f56-114">方法参数的数据类型和方法的行为。</span><span class="sxs-lookup"><span data-stu-id="c8f56-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="c8f56-115">返回值的数据类型。</span><span class="sxs-lookup"><span data-stu-id="c8f56-115">The data types of method return values.</span></span>

<span data-ttu-id="c8f56-116">例如：</span><span class="sxs-lookup"><span data-stu-id="c8f56-116">Some examples:</span></span>

- <span data-ttu-id="c8f56-117">`RangeAreas` 具有 `address` 属性，返回使用逗号分隔的范围地址字符串，而不是像 `Range.address` 属性那样返回一个地址。</span><span class="sxs-lookup"><span data-stu-id="c8f56-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="c8f56-118">`RangeAreas` 具有 `dataValidation` 属性，返回的 `DataValidation` 对象代表 `RangeAreas` 中所有范围的数据验证（如果一致）。</span><span class="sxs-lookup"><span data-stu-id="c8f56-118">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent.</span></span> <span data-ttu-id="c8f56-119">如果相同的 `DataValidation` 对象不应用于 `RangeAreas` 中的所有范围，则属性为 `null`。</span><span class="sxs-lookup"><span data-stu-id="c8f56-119">The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="c8f56-120">对象有一项常规、但不是通用的原则：*如果属性在 `RangeAreas` 的所有区域上具有一致的值，则它是 `null`。*`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="c8f56-120">This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.*</span></span> <span data-ttu-id="c8f56-121">要了解有关详细信息和一些例外，请参阅[读取 RangeAreas 的属性](#reading-properties-of-rangeareas)。</span><span class="sxs-lookup"><span data-stu-id="c8f56-121">See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="c8f56-122">`RangeAreas.cellCount` 获取 `RangeAreas` 中所有范围的单元格总数目。</span><span class="sxs-lookup"><span data-stu-id="c8f56-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c8f56-123">`RangeAreas.calculate` 重新计算 `RangeAreas` 中所有范围的单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="c8f56-124">`RangeAreas.getEntireColumn` 和 `RangeAreas.getEntireRow` 返回另一个 `RangeAreas` 对象，表示 `RangeAreas` 中所有范围的所有列 （或行）。</span><span class="sxs-lookup"><span data-stu-id="c8f56-124">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`.</span></span> <span data-ttu-id="c8f56-125">例如，如果 `RangeAreas` 表示"A:C" 和 "F:L"，则 `RangeAreas.getEntireColumn` 返回表示 "A1:C4" 和 "F14:L15" 的 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="c8f56-125">For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="c8f56-126">`RangeAreas.copyFrom` 可以是 `Range` 或 `RangeAreas` 参数，表示复制操作的源范围。</span><span class="sxs-lookup"><span data-stu-id="c8f56-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="c8f56-127">在 RangeAreas 上也可用的“范围”成员的完整列表</span><span class="sxs-lookup"><span data-stu-id="c8f56-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="c8f56-128">属性</span><span class="sxs-lookup"><span data-stu-id="c8f56-128">Properties</span></span>

<span data-ttu-id="c8f56-129">请先熟悉[读取 RangeAreas 的属性](#reading-properties-of-rangeareas)，然后再编写可读取列出的所有属性的代码。</span><span class="sxs-lookup"><span data-stu-id="c8f56-129">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed.</span></span> <span data-ttu-id="c8f56-130">返回什么有不少微妙之处。</span><span class="sxs-lookup"><span data-stu-id="c8f56-130">There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="c8f56-131">address</span><span class="sxs-lookup"><span data-stu-id="c8f56-131">address</span></span>
- <span data-ttu-id="c8f56-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="c8f56-132">addressLocal</span></span>
- <span data-ttu-id="c8f56-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="c8f56-133">cellCount</span></span>
- <span data-ttu-id="c8f56-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="c8f56-134">conditionalFormats</span></span>
- <span data-ttu-id="c8f56-135">context</span><span class="sxs-lookup"><span data-stu-id="c8f56-135">context</span></span>
- <span data-ttu-id="c8f56-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="c8f56-136">dataValidation</span></span>
- <span data-ttu-id="c8f56-137">格式</span><span class="sxs-lookup"><span data-stu-id="c8f56-137">format</span></span>
- <span data-ttu-id="c8f56-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="c8f56-138">isEntireColumn</span></span>
- <span data-ttu-id="c8f56-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="c8f56-139">isEntireRow</span></span>
- <span data-ttu-id="c8f56-140">样式</span><span class="sxs-lookup"><span data-stu-id="c8f56-140">style</span></span>
- <span data-ttu-id="c8f56-141">工作表</span><span class="sxs-lookup"><span data-stu-id="c8f56-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="c8f56-142">方法</span><span class="sxs-lookup"><span data-stu-id="c8f56-142">Methods</span></span>

<span data-ttu-id="c8f56-143">预览版中的范围方法是做了标记的。</span><span class="sxs-lookup"><span data-stu-id="c8f56-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="c8f56-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="c8f56-144">calculate()</span></span>
- <span data-ttu-id="c8f56-145">clear()</span><span class="sxs-lookup"><span data-stu-id="c8f56-145">clear()</span></span>
- <span data-ttu-id="c8f56-146">convertDataTypeToText()（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="c8f56-147">convertToLinkedDataType()（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="c8f56-148">copyFrom()（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="c8f56-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="c8f56-149">getEntireColumn()</span></span>
- <span data-ttu-id="c8f56-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="c8f56-150">getEntireRow()</span></span>
- <span data-ttu-id="c8f56-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="c8f56-151">getIntersection()</span></span>
- <span data-ttu-id="c8f56-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="c8f56-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="c8f56-153">getOffsetRange()（名为 RangeAreas 对象上的 getOffsetRangeAreas）</span><span class="sxs-lookup"><span data-stu-id="c8f56-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="c8f56-154">getSpecialCells()（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="c8f56-155">getSpecialCellsOrNullObject()（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="c8f56-156">getTables()（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-156">getTables() (preview)</span></span>
- <span data-ttu-id="c8f56-157">getUsedRange()（名为 RangeAreas 对象上的 getUsedRangeAreas）</span><span class="sxs-lookup"><span data-stu-id="c8f56-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="c8f56-158">getUsedRangeOrNullObject() （名为 RangeAreas 对象上的 getUsedRangeAreasOrNullObject）</span><span class="sxs-lookup"><span data-stu-id="c8f56-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="c8f56-159">load()</span><span class="sxs-lookup"><span data-stu-id="c8f56-159">load()</span></span>
- <span data-ttu-id="c8f56-160">set()</span><span class="sxs-lookup"><span data-stu-id="c8f56-160">set\*</span></span>
- <span data-ttu-id="c8f56-161">setDirty()（预览）</span><span class="sxs-lookup"><span data-stu-id="c8f56-161">setDirty() (preview)</span></span>
- <span data-ttu-id="c8f56-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="c8f56-162">toJSON()</span></span>
- <span data-ttu-id="c8f56-163">track()</span><span class="sxs-lookup"><span data-stu-id="c8f56-163">track</span></span>
- <span data-ttu-id="c8f56-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="c8f56-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="c8f56-165">特定于 RangeArea 的属性和方法</span><span class="sxs-lookup"><span data-stu-id="c8f56-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="c8f56-166">`RangeAreas` 类型具有某些不在 `Range` 对象上的属性和方法。</span><span class="sxs-lookup"><span data-stu-id="c8f56-166">The `RangeAreas` type has some properties and methods that are not on the `Range` object:</span></span> <span data-ttu-id="c8f56-167">以下是其中的选定内容：</span><span class="sxs-lookup"><span data-stu-id="c8f56-167">The following is a selection of them:</span></span>

- <span data-ttu-id="c8f56-168">`areas`：`RangeCollection` 对象，其中包含 `RangeAreas` 对象所表示的所有范围。</span><span class="sxs-lookup"><span data-stu-id="c8f56-168">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object.</span></span> <span data-ttu-id="c8f56-169">对象也是新增，并类似于其他 Excel 集合对象。`RangeCollection`</span><span class="sxs-lookup"><span data-stu-id="c8f56-169">The `RangeCollection` object is also new and is similar to other Excel collection objects.</span></span> <span data-ttu-id="c8f56-170">它具有 `items` 属性，是表示范围的 `Range` 对象数组。</span><span class="sxs-lookup"><span data-stu-id="c8f56-170">It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="c8f56-171">`areaCount`：`RangeAreas` 中范围的总数目。</span><span class="sxs-lookup"><span data-stu-id="c8f56-171">The total number of recipients in the message.</span></span>
- <span data-ttu-id="c8f56-172">`getOffsetRangeAreas`：工作原理和 [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 一样，只不过返回一个 `RangeAreas`，它包含的范围与原 `RangeAreas` 中的某一范围产生偏移。</span><span class="sxs-lookup"><span data-stu-id="c8f56-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="c8f56-173">创建 RangeAreas 并设置属性</span><span class="sxs-lookup"><span data-stu-id="c8f56-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="c8f56-174">您可以通过两种基本方式创建 `RangeAreas` 对象：</span><span class="sxs-lookup"><span data-stu-id="c8f56-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="c8f56-175">调用 `Worksheet.getRanges()` 并向它传递以逗号分隔的范围地址的字符串。</span><span class="sxs-lookup"><span data-stu-id="c8f56-175">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses.</span></span> <span data-ttu-id="c8f56-176">如果您想要包括的任何区域已变成 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem)，可以在字符串中包括该名称，而不是该地址。</span><span class="sxs-lookup"><span data-stu-id="c8f56-176">If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="c8f56-177">调用 `Workbook.getSelectedRanges()`。</span><span class="sxs-lookup"><span data-stu-id="c8f56-177">Call `Workbook.getSelectedRanges()`.</span></span> <span data-ttu-id="c8f56-178">此方法返回一个 `RangeAreas`，它表示当前活动工作表的所有选定范围。</span><span class="sxs-lookup"><span data-stu-id="c8f56-178">This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="c8f56-179">拥有 `RangeAreas` 对象后，您可以使用返回如 `getOffsetRangeAreas` 和 `getIntersection` 等 `RangeAreas` 的对象创建其他对象。</span><span class="sxs-lookup"><span data-stu-id="c8f56-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="c8f56-180">不能直接添加其他范围到 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="c8f56-180">You cannot directly add additional ranges to a `RangeAreas` object.</span></span> <span data-ttu-id="c8f56-181">例如，`RangeAreas.areas` 中的集合中没有 `add` 方法。</span><span class="sxs-lookup"><span data-stu-id="c8f56-181">For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="c8f56-182">切勿尝试直接添加或删除 `RangeAreas.areas.items` 数组的成员。</span><span class="sxs-lookup"><span data-stu-id="c8f56-182">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array.</span></span> <span data-ttu-id="c8f56-183">这将导致代码中的不正常行为。</span><span class="sxs-lookup"><span data-stu-id="c8f56-183">This will lead to undesirable behavior in your code.</span></span> <span data-ttu-id="c8f56-184">例如，可以将额外的 `Range` 对象推送到数组，但这样做将导致错误，因为 `RangeAreas` 属性和方法的行为不会因新项目的添加而有所不同。</span><span class="sxs-lookup"><span data-stu-id="c8f56-184">For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there.</span></span> <span data-ttu-id="c8f56-185">例如，`areaCount` 属性不包括以这种方式推送的范围，并且如果 `index` 大于 `areasCount-1`，`RangeAreas.getItemAt(index)` 将抛出错误。</span><span class="sxs-lookup"><span data-stu-id="c8f56-185">For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`.</span></span> <span data-ttu-id="c8f56-186">同样，通过获取引用并调用其 方法删除  数组中的  对象会导致错误：虽然  对象已删除，父  对象仍将视其为存在。`Range``RangeAreas.areas.items``Range.delete``Range`\*\*`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="c8f56-186">Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence.</span></span> <span data-ttu-id="c8f56-187">例如，如果您的代码调用 `RangeAreas.calculate`，Office 会尝试计算该范围，但将发生错误，因为 range 对象不存在。</span><span class="sxs-lookup"><span data-stu-id="c8f56-187">For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="c8f56-188">在 `RangeAreas` 上设置属性会在 `RangeAreas.areas` 集合的所有范围上设置相应的属性。</span><span class="sxs-lookup"><span data-stu-id="c8f56-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="c8f56-189">以下是在多个范围上设置属性的示例。</span><span class="sxs-lookup"><span data-stu-id="c8f56-189">The following is an example of setting a property on multiple ranges.</span></span> <span data-ttu-id="c8f56-190">该函数高亮显示 **F3:F5** 至 **H3:H5** 之间的范围。</span><span class="sxs-lookup"><span data-stu-id="c8f56-190">The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="c8f56-191">本示例适用于可以硬编码或运行时轻松计算传递给 `getRanges` 的范围地址的场景。</span><span class="sxs-lookup"><span data-stu-id="c8f56-191">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime.</span></span> <span data-ttu-id="c8f56-192">一些符合条件的方案包括：</span><span class="sxs-lookup"><span data-stu-id="c8f56-192">Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="c8f56-193">代码在已知模板的上下文中运行。</span><span class="sxs-lookup"><span data-stu-id="c8f56-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="c8f56-194">代码在导入数据的上下文中运行，其中数据架构为已知。</span><span class="sxs-lookup"><span data-stu-id="c8f56-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="c8f56-195">当你不能在编码时确定需要执行操作的范围时，必须在运行时确定。</span><span class="sxs-lookup"><span data-stu-id="c8f56-195">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime.</span></span> <span data-ttu-id="c8f56-196">下一节讨论了这些方案。</span><span class="sxs-lookup"><span data-stu-id="c8f56-196">The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="c8f56-197">以编程方式发现范围</span><span class="sxs-lookup"><span data-stu-id="c8f56-197">Discover range areas programmatically</span></span>

<span data-ttu-id="c8f56-198">和 `Range.getSpecialCellsOrNullObject()` 方法使你能够在运行时根据单元格和单元格值类型的特征而确定对其执行操作的范围。`Range.getSpecialCells()`</span><span class="sxs-lookup"><span data-stu-id="c8f56-198">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells.</span></span> <span data-ttu-id="c8f56-199">下面是 TypeScript 数据类型文件中的方法签名：</span><span class="sxs-lookup"><span data-stu-id="c8f56-199">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="c8f56-200">下面是使用第一个方法的示例。</span><span class="sxs-lookup"><span data-stu-id="c8f56-200">The following is an example of using the "Between" operator:</span></span> <span data-ttu-id="c8f56-201">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="c8f56-201">About this code, note:</span></span>

- <span data-ttu-id="c8f56-202">它通过首先调用 `Worksheet.getUsedRange` 并针对仅该范围调用 `getSpecialCells`，将搜索限制于工作表的所需部分。</span><span class="sxs-lookup"><span data-stu-id="c8f56-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="c8f56-203">它将 `Excel.SpecialCellType` 枚举中值的字符串版本作为参数传递到 `getSpecialCells`。</span><span class="sxs-lookup"><span data-stu-id="c8f56-203">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum.</span></span> <span data-ttu-id="c8f56-204">某些无法传递的其他值 包括： 空白单元格 为 "Blanks" ，具有文字值而不是公式的单元格为 "Constants"，与 `usedRange` 中第一个单元格具有相同条件格式的单元格 为 "SameConditionalFormat"。</span><span class="sxs-lookup"><span data-stu-id="c8f56-204">Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`.</span></span> <span data-ttu-id="c8f56-205">第一个单元格为左上角的单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-205">The first cell is the upper leftmost cell.</span></span> <span data-ttu-id="c8f56-206">要了解枚举中值的完整列表，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="c8f56-206">For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="c8f56-207">`getSpecialCells` 方法返回 `RangeAreas` 对象，所以所有带公式的单元格都将以粉红色标示，即使它们并非全都相邻。</span><span class="sxs-lookup"><span data-stu-id="c8f56-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="c8f56-208">有时，范围不具有*任何*带目标特征的单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-208">Sometimes the range doesn't have *any* cells with the targeted characteristic.</span></span> <span data-ttu-id="c8f56-209">如果 `getSpecialCells` 找不到任何单元格，将引发 **ItemNotFound** 错误。</span><span class="sxs-lookup"><span data-stu-id="c8f56-209">If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error.</span></span> <span data-ttu-id="c8f56-210">如果存在 `catch` 块/方法，控制流将重定向到该块/方法。</span><span class="sxs-lookup"><span data-stu-id="c8f56-210">This would divert the flow of control to a `catch` block/method, if there is one.</span></span> <span data-ttu-id="c8f56-211">如果不存在，错误将暂停函数的执行。</span><span class="sxs-lookup"><span data-stu-id="c8f56-211">If there isn't, the error halts the function.</span></span> <span data-ttu-id="c8f56-212">在某些时候，如果不存在具有目标特征的单元格，你可能希望抛出错误。</span><span class="sxs-lookup"><span data-stu-id="c8f56-212">There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="c8f56-213">但在其他时候，不存在匹配单元格的情况属于正常但可能不常见；你的代码应检查这种可能性并恰当处理而不抛出错误。</span><span class="sxs-lookup"><span data-stu-id="c8f56-213">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="c8f56-214">对于这些方案，请使用 `getSpecialCellsOrNullObject` 方法并测试 `RangeAreas.isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="c8f56-214">For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property.</span></span> <span data-ttu-id="c8f56-215">示例如下。</span><span class="sxs-lookup"><span data-stu-id="c8f56-215">The following is an example.</span></span> <span data-ttu-id="c8f56-216">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="c8f56-216">Note about this code:</span></span>

- <span data-ttu-id="c8f56-217">方法总是返回代理对象，因此永远不会成为一般 JavaScript 意义上的  `null`。`getSpecialCellsOrNullObject`</span><span class="sxs-lookup"><span data-stu-id="c8f56-217">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="c8f56-218">但是，如果找到匹配的单元格，对象的 `isNullObject` 属性将设置为 `true`。</span><span class="sxs-lookup"><span data-stu-id="c8f56-218">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="c8f56-219">它将在测试 `isNullObject` 属性*之前*调用 `context.sync`。</span><span class="sxs-lookup"><span data-stu-id="c8f56-219">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="c8f56-220">这是所有 `*OrNullObject` 方法和属性的要求，因为你始终需要加载和同步属性才能读取它。</span><span class="sxs-lookup"><span data-stu-id="c8f56-220">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="c8f56-221">但是，不需要*显式*加载 `isNullObject` 属性。</span><span class="sxs-lookup"><span data-stu-id="c8f56-221">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="c8f56-222">它会由 `context.sync` 自动加载，即使对象没有调用 `load`。</span><span class="sxs-lookup"><span data-stu-id="c8f56-222">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="c8f56-223">有关详细信息，请参阅 [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods)。</span><span class="sxs-lookup"><span data-stu-id="c8f56-223">For more information, see [\*](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods) object.</span></span>
- <span data-ttu-id="c8f56-224">你可以通过首先选择没有公式的单元格范围然后运行它来测试此代码。</span><span class="sxs-lookup"><span data-stu-id="c8f56-224">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="c8f56-225">然后选择其中至少具有一个带公式的单元格的范围并再次运行。</span><span class="sxs-lookup"><span data-stu-id="c8f56-225">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="c8f56-226">为了简单起见，在此文章中所有其他示例均使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject`。</span><span class="sxs-lookup"><span data-stu-id="c8f56-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="c8f56-227">通过单元格的值类型缩小目标单元格范围</span><span class="sxs-lookup"><span data-stu-id="c8f56-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="c8f56-228">可选的第二个参数为枚举类型的 `Excel.SpecialCellValueType`，它进一步缩小目标单元格的范围。</span><span class="sxs-lookup"><span data-stu-id="c8f56-228">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target.</span></span> <span data-ttu-id="c8f56-229">仅当你向 `getSpecialCells` 或 `getSpecialCellsOrNullObject`传递 "Formulas" 或 "Constants" 时可以使用。</span><span class="sxs-lookup"><span data-stu-id="c8f56-229">You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`.</span></span> <span data-ttu-id="c8f56-230">该参数指定单元格必须拥有特定类型的值。</span><span class="sxs-lookup"><span data-stu-id="c8f56-230">The parameter specifies that you only want cells with certain types of values.</span></span> <span data-ttu-id="c8f56-231">有四个基本类型："Error"、"Logical"（ 即布尔值）、"Numbers" 和 "Text"。</span><span class="sxs-lookup"><span data-stu-id="c8f56-231">There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text".</span></span> <span data-ttu-id="c8f56-232">（除了这四个值，该枚举还有其他值，下文将会讨论。）下面是一个示例。</span><span class="sxs-lookup"><span data-stu-id="c8f56-232">(The enum has other values besides these four which are discussed below.) The following is an example.</span></span> <span data-ttu-id="c8f56-233">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="c8f56-233">About this code, note:</span></span>

- <span data-ttu-id="c8f56-234">它仅将突出显示包含文本数字值的单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-234">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="c8f56-235">它不会突出显示包含公式 （即使结果是数字）的单元格或布尔值、文本或错误状态的单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-235">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="c8f56-236">若要测试代码，请确保工作表的一些单元格拥有文本数字值、一些单元格带有其他类型的文本值，另有一些含有公式。</span><span class="sxs-lookup"><span data-stu-id="c8f56-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="c8f56-237">有时，你需要对多个单元格值类型执行操作—— 例如所有文本值和所有布尔值 ("Logical")。</span><span class="sxs-lookup"><span data-stu-id="c8f56-237">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells.</span></span> <span data-ttu-id="c8f56-238">枚举具有让你组合多个类型的值。`Excel.SpecialCellValueType`</span><span class="sxs-lookup"><span data-stu-id="c8f56-238">The `Excel.SpecialCellValueType` enum has values that let you combine types.</span></span> <span data-ttu-id="c8f56-239">例如，"LogicalText" 将匹配所有布尔值和所有文本值单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-239">For example, "LogicalText" will target all boolean and all text-valued cells.</span></span> <span data-ttu-id="c8f56-240">可以将这四种基本类型中的任何两种或任何三种类型组合使用。</span><span class="sxs-lookup"><span data-stu-id="c8f56-240">You can combine any two or any three of the four basic types.</span></span> <span data-ttu-id="c8f56-241">这些枚举值的基本类型组合名称通常以字母顺序列出。</span><span class="sxs-lookup"><span data-stu-id="c8f56-241">The names of these enum values that combine basic types are always in alphabetical order.</span></span> <span data-ttu-id="c8f56-242">因此，若要合并错误值、文本值和布尔值的单元格，请使用 "ErrorLogicalText"，而不是 "LogicalErrorText" 或 "TextErrorLogical"。</span><span class="sxs-lookup"><span data-stu-id="c8f56-242">So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical".</span></span> <span data-ttu-id="c8f56-243">默认参数 "All" 合并所有四种类型。</span><span class="sxs-lookup"><span data-stu-id="c8f56-243">The default parameter of "All" combines all four types.</span></span> <span data-ttu-id="c8f56-244">下面的示例突出显示带有生成数字或布尔值的公式的所有单元格：</span><span class="sxs-lookup"><span data-stu-id="c8f56-244">The following example highlights all cells with formulas that produce number or boolean values:</span></span>

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
> <span data-ttu-id="c8f56-245">`Excel.SpecialCellValueType` 参数仅在 `Excel.SpecialCellType` 参数为 "Formulas" 或 "Constants" 时可以使用。</span><span class="sxs-lookup"><span data-stu-id="c8f56-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="c8f56-246">获取 RangeAreas 中的 RangeAreas</span><span class="sxs-lookup"><span data-stu-id="c8f56-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="c8f56-247">类型本身也有 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法，它们采用相同的两个参数。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="c8f56-247">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters.</span></span> <span data-ttu-id="c8f56-248">这些方法返回 `RangeAreas.areas` 集合所有范围的所有目标单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-248">These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection.</span></span> <span data-ttu-id="c8f56-249">方法调用于  对象而不是  对象的行为中存在一个细小区别：将 "SameConditionalFormat" 作为第一个参数传递时，方法会返回  集合第一个范围中与左上角单元格具有同一条件格式的所有单元格。`RangeAreas``Range`*`RangeAreas.areas`*</span><span class="sxs-lookup"><span data-stu-id="c8f56-249">There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span> <span data-ttu-id="c8f56-250">这一点同样适用于 "SameDataValidation"：传递给 `Range.getSpecialCells` 时，它将返回与\* 范围中\*左上角单元格具有相同数据验证规则的的单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-250">The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*.</span></span> <span data-ttu-id="c8f56-251">但当传递给 `RangeAreas.getSpecialCells` 时，将返回与 *`RangeAreas.areas` 集合第一个范围中*的左上角单元格具有相同数据验证规则的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="c8f56-251">But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="c8f56-252">读取 RangeAreas 的属性</span><span class="sxs-lookup"><span data-stu-id="c8f56-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="c8f56-253">读取 `RangeAreas` 的属性值时需要小心，因为给定的属性可能在 `RangeAreas` 的不同范围内具有不同的值。</span><span class="sxs-lookup"><span data-stu-id="c8f56-253">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`.</span></span> <span data-ttu-id="c8f56-254">一般规则是，如果*可以*返回一致的值，它就会返回。</span><span class="sxs-lookup"><span data-stu-id="c8f56-254">The general rule is that if a consistent value *can* be returned it will be returned.</span></span> <span data-ttu-id="c8f56-255">例如，在下面的代码中，粉红色的 RGB 代码 (`#FFC0CB`) 和 `true` 将记录到控制台中，因为 `RangeAreas` 对象中的这两个范围中具有粉红色填充且两者均为整个列。</span><span class="sxs-lookup"><span data-stu-id="c8f56-255">For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="c8f56-256">当不可能达成一致性时，事情会变得更为复杂。</span><span class="sxs-lookup"><span data-stu-id="c8f56-256">Things get more complicated when consistency isn't possible.</span></span> <span data-ttu-id="c8f56-257">属性的行为遵循以下三个原则：`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="c8f56-257">The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="c8f56-258">对象的布尔值属性返回 `false`，除非该属性对于所有成员范围均为 true。`RangeAreas`</span><span class="sxs-lookup"><span data-stu-id="c8f56-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="c8f56-259">除 `address` 属性外的非布尔值属性将返回 `null`，除非所有成员范围的相应属性具有相同的值。</span><span class="sxs-lookup"><span data-stu-id="c8f56-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="c8f56-260">属性返回成员范围地址的以逗号分隔 字符串。`address`</span><span class="sxs-lookup"><span data-stu-id="c8f56-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="c8f56-261">例如，下面的代码生成一个 `RangeAreas`，其中只有一个范围是整列，也只有一个 范围填充粉红色。</span><span class="sxs-lookup"><span data-stu-id="c8f56-261">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink.</span></span> <span data-ttu-id="c8f56-262">控制台将显示填充颜色为 `null`，`isEntireRow` 属性为 `false` 且 `address` 属性为 "Sheet1!F3:F5, Sheet1!H:H"（假定工作表名称为 "Sheet1"）。</span><span class="sxs-lookup"><span data-stu-id="c8f56-262">The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="c8f56-263">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c8f56-263">See also</span></span>

- [<span data-ttu-id="c8f56-264">Excel JavaScript API 核心概念</span><span class="sxs-lookup"><span data-stu-id="c8f56-264">Excel JavaScript API core concepts</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="c8f56-265">Range 对象 (Excel JavaScript API)</span><span class="sxs-lookup"><span data-stu-id="c8f56-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="c8f56-266">[RangeAreas 对象 (Excel JavaScript API)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas)（该链接可能无法在预览版 API 中生效）。</span><span class="sxs-lookup"><span data-stu-id="c8f56-266">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview.</span></span> <span data-ttu-id="c8f56-267">作为替代方法，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)。）</span><span class="sxs-lookup"><span data-stu-id="c8f56-267">As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>