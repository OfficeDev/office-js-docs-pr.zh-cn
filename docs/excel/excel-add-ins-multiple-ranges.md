---
title: 在 Excel 加载项中同时处理多个范围
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: 2387be8dc17d85028b1d086cb192ac1accf167d5
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459194"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a><span data-ttu-id="fab7b-102">在 Excel 加载项中同时处理多个范围（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-102">Work with multiple ranges simultaneously in Excel add-ins (Preview)</span></span>

<span data-ttu-id="fab7b-p101">Excel JavaScript 库同时在多个范围启用加载项完成操作，并且设置属性 。范围不必是相邻的。除了使代码更加简单以外，这种设置属性的方式比单独为每个范围设置相同属性的方式运行速度要快很多。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p101">The Excel JavaScript library enables your add-in to perform operations, and set properties, on multiple ranges simultaneously. The ranges do not have to be contiguous. In addition to making your code simpler, this way of setting a property runs much faster than setting the same property individually for each of the ranges.</span></span>

> [!NOTE]
> <span data-ttu-id="fab7b-p102">本文介绍的 API 要求 **Office 2016 Click-to-Run version 1809 Build 10820.20000** 或更高版本。（可能需要加入 [Office Insider 程序](https://products.office.com/office-insider) 来获取适当的内部版本。）此外，必须从 [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) 加载 beta 版本的 Office JavaScript 库。最后，我们没有这些 API 的参考页。但以下定义类型文件具有它们的说明： [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p102">The APIs described in this article require **Office 2016 Click-to-Run version 1809 Build 10820.20000** or later. (You may need to join the [Office Insider program](https://products.office.com/office-insider) to get an appropriate build.) Also, you must load the beta version of the Office JavaScript library from [Office.js CDN](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Finally, we don't have reference pages for these APIs yet. But the following definition type file has descriptions for them: [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>

## <a name="rangeareas"></a><span data-ttu-id="fab7b-110">RangeAreas</span><span class="sxs-lookup"><span data-stu-id="fab7b-110">RangeAreas</span></span>

<span data-ttu-id="fab7b-p103">`Excel.RangeAreas` 对象表示一组范围（可能不相邻）。它具有类似于 `Range` 类型的属性和方法 （很多具有相同或类似的名称），但已经进行了调整：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p103">A set of (possibly discontiguous) ranges is represented by an `Excel.RangeAreas` object. It has properties and methods similar to the `Range` type (many with the same, or similar, names), but adjustments have been made to:</span></span>

- <span data-ttu-id="fab7b-113">属性的数据类型以及 getter 和 setter 的行为。</span><span class="sxs-lookup"><span data-stu-id="fab7b-113">The data types for properties and the behavior of the setters and getters.</span></span>
- <span data-ttu-id="fab7b-114">方法参数的数据类型和方法的行为。</span><span class="sxs-lookup"><span data-stu-id="fab7b-114">The data types of method parameters and the method behaviors.</span></span>
- <span data-ttu-id="fab7b-115">返回值的数据类型。</span><span class="sxs-lookup"><span data-stu-id="fab7b-115">The data types of method return values.</span></span>

<span data-ttu-id="fab7b-116">例如：</span><span class="sxs-lookup"><span data-stu-id="fab7b-116">Some examples:</span></span>

- <span data-ttu-id="fab7b-117">`RangeAreas` 具有返回范围地址的逗号分隔字符串的 `address` 属性，而不是像 `Range.address` 属性那样只是一个地址。</span><span class="sxs-lookup"><span data-stu-id="fab7b-117">`RangeAreas` has an `address` property that returns a comma-delimited string of range addresses, instead of just one address as with the `Range.address` property.</span></span>
- <span data-ttu-id="fab7b-p104">`RangeAreas` 如果是一致的，具有返回 `DataValidation` 对象的 `dataValidation` 属性，该对象表示 `RangeAreas` 内所有范围的数据有效性。如果相同的 `DataValidation` 对象没有应用于 `RangeAreas` 内的所有范围则对象为 `null` 。这是 `RangeAreas` 对象的一个常规但非通用原则： *如果属性在 `RangeAreas` 内所有范围上没有一致的值，它便是 `null` 。* 有关详细信息和一些例外，请参阅 [读取 RangeAreas 的属性](#reading-properties-of-rangeareas) 。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p104">`RangeAreas` has a `dataValidation` property that returns a `DataValidation` object that represents the data validation of all the ranges in the `RangeAreas`, if it is consistent. The property is `null` if identical `DataValidation` objects are not applied to all the all the ranges in the `RangeAreas`. This is a general, but not universal, principle with the `RangeAreas` object: *If a property does not have consistent values on all the all the ranges in the `RangeAreas`, then it is `null`.* See [Reading properties of RangeAreas](#reading-properties-of-rangeareas) for more information and some exceptions.</span></span>
- <span data-ttu-id="fab7b-122">`RangeAreas.cellCount` 获取 `RangeAreas` 内所有范围的单元格总数。</span><span class="sxs-lookup"><span data-stu-id="fab7b-122">`RangeAreas.cellCount` gets the total number of cells in all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="fab7b-123">`RangeAreas.calculate` 重新计算 `RangeAreas` 内所有范围的单元格。</span><span class="sxs-lookup"><span data-stu-id="fab7b-123">`RangeAreas.calculate` recalculates the cells of all the ranges in the `RangeAreas`.</span></span>
- <span data-ttu-id="fab7b-p105">`RangeAreas.getEntireColumn` `RangeAreas.getEntireRow` 返回另一个 `RangeAreas` 对象，该对象表示 `RangeAreas` 内所有范围的所有列 （或行）。例如，如果 `RangeAreas` 表示“ A1:C4 ”和“ F14:L15 ”，则 `RangeAreas.getEntireColumn` 返回表示“ A:C ”和“ F:L ”的 `RangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p105">`RangeAreas.getEntireColumn` and `RangeAreas.getEntireRow` return another `RangeAreas` object that represents all of the columns (or rows) in all the ranges in the `RangeAreas`. For example, if the `RangeAreas` represents "A1:C4" and "F14:L15", then `RangeAreas.getEntireColumn` returns a `RangeAreas` object that represents "A:C" and "F:L".</span></span>
- <span data-ttu-id="fab7b-126">`RangeAreas.copyFrom` 可以采用 `Range` 或 `RangeAreas` 参数，表示复制操作的源范围。</span><span class="sxs-lookup"><span data-stu-id="fab7b-126">`RangeAreas.copyFrom` can take either a `Range` or a `RangeAreas` parameter representing the source range(s) of the copy operation.</span></span>

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a><span data-ttu-id="fab7b-127">在 RangeAreas 上也可用的“范围”成员的完整列表</span><span class="sxs-lookup"><span data-stu-id="fab7b-127">Complete list of Range members that are also available on RangeAreas</span></span>

##### <a name="properties"></a><span data-ttu-id="fab7b-128">属性</span><span class="sxs-lookup"><span data-stu-id="fab7b-128">Properties</span></span>

<span data-ttu-id="fab7b-p106">在编写读取任何所列属性的代码之前应熟知 [读取 RangeAreas 的属性](#reading-properties-of-rangeareas) 。返回内容存在细微差异。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p106">Be familiar with [Reading properties of RangeAreas](#reading-properties-of-rangeareas) before you write code that reads any properties listed. There are subtleties to what gets returned.</span></span>

- <span data-ttu-id="fab7b-131">地址</span><span class="sxs-lookup"><span data-stu-id="fab7b-131">address</span></span>
- <span data-ttu-id="fab7b-132">addressLocal</span><span class="sxs-lookup"><span data-stu-id="fab7b-132">addressLocal</span></span>
- <span data-ttu-id="fab7b-133">cellCount</span><span class="sxs-lookup"><span data-stu-id="fab7b-133">cellCount</span></span>
- <span data-ttu-id="fab7b-134">conditionalFormats</span><span class="sxs-lookup"><span data-stu-id="fab7b-134">conditionalFormats</span></span>
- <span data-ttu-id="fab7b-135">context</span><span class="sxs-lookup"><span data-stu-id="fab7b-135">context</span></span>
- <span data-ttu-id="fab7b-136">dataValidation</span><span class="sxs-lookup"><span data-stu-id="fab7b-136">dataValidation</span></span>
- <span data-ttu-id="fab7b-137">格式</span><span class="sxs-lookup"><span data-stu-id="fab7b-137">format</span></span>
- <span data-ttu-id="fab7b-138">isEntireColumn</span><span class="sxs-lookup"><span data-stu-id="fab7b-138">isEntireColumn</span></span>
- <span data-ttu-id="fab7b-139">isEntireRow</span><span class="sxs-lookup"><span data-stu-id="fab7b-139">isEntireRow</span></span>
- <span data-ttu-id="fab7b-140">样式</span><span class="sxs-lookup"><span data-stu-id="fab7b-140">style</span></span>
- <span data-ttu-id="fab7b-141">工作表</span><span class="sxs-lookup"><span data-stu-id="fab7b-141">worksheet</span></span>

##### <a name="methods"></a><span data-ttu-id="fab7b-142">方法</span><span class="sxs-lookup"><span data-stu-id="fab7b-142">Methods</span></span>

<span data-ttu-id="fab7b-143">预览版中的范围方法是做了标记的。</span><span class="sxs-lookup"><span data-stu-id="fab7b-143">Range methods in preview are marked.</span></span>

- <span data-ttu-id="fab7b-144">calculate()</span><span class="sxs-lookup"><span data-stu-id="fab7b-144">calculate()</span></span>
- <span data-ttu-id="fab7b-145">clear()</span><span class="sxs-lookup"><span data-stu-id="fab7b-145">clear()</span></span>
- <span data-ttu-id="fab7b-146">convertDataTypeToText()（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-146">convertDataTypeToText() (preview)</span></span>
- <span data-ttu-id="fab7b-147">convertToLinkedDataType()（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-147">convertToLinkedDataType() (preview)</span></span>
- <span data-ttu-id="fab7b-148">copyFrom()（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-148">copyFrom() (preview)</span></span>
- <span data-ttu-id="fab7b-149">getEntireColumn()</span><span class="sxs-lookup"><span data-stu-id="fab7b-149">getEntireColumn()</span></span>
- <span data-ttu-id="fab7b-150">getEntireRow()</span><span class="sxs-lookup"><span data-stu-id="fab7b-150">getEntireRow()</span></span>
- <span data-ttu-id="fab7b-151">getIntersection()</span><span class="sxs-lookup"><span data-stu-id="fab7b-151">getIntersection()</span></span>
- <span data-ttu-id="fab7b-152">getIntersectionOrNullObject()</span><span class="sxs-lookup"><span data-stu-id="fab7b-152">getIntersectionOrNullObject()</span></span>
- <span data-ttu-id="fab7b-153">getOffsetRange()（名为 RangeAreas 对象上的 getOffsetRangeAreas）</span><span class="sxs-lookup"><span data-stu-id="fab7b-153">getOffsetRange() (named getOffsetRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="fab7b-154">getSpecialCells()（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-154">getSpecialCells() (preview)</span></span>
- <span data-ttu-id="fab7b-155">getSpecialCellsOrNullObject()（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-155">getSpecialCellsOrNullObject() (preview)</span></span>
- <span data-ttu-id="fab7b-156">getTables()（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-156">getTables() (preview)</span></span>
- <span data-ttu-id="fab7b-157">getUsedRange()（名为 RangeAreas 对象上的 getUsedRangeAreas）</span><span class="sxs-lookup"><span data-stu-id="fab7b-157">getUsedRange() (named getUsedRangeAreas on the RangeAreas object)</span></span>
- <span data-ttu-id="fab7b-158">getUsedRangeOrNullObject() （名为 RangeAreas 对象上的 getUsedRangeAreasOrNullObject）</span><span class="sxs-lookup"><span data-stu-id="fab7b-158">getUsedRangeOrNullObject() (named getUsedRangeAreasOrNullObject on the RangeAreas object)</span></span>
- <span data-ttu-id="fab7b-159">load()</span><span class="sxs-lookup"><span data-stu-id="fab7b-159">load()</span></span>
- <span data-ttu-id="fab7b-160">set()</span><span class="sxs-lookup"><span data-stu-id="fab7b-160">set\*</span></span>
- <span data-ttu-id="fab7b-161">setDirty()（预览）</span><span class="sxs-lookup"><span data-stu-id="fab7b-161">setDirty() (preview)</span></span>
- <span data-ttu-id="fab7b-162">toJSON()</span><span class="sxs-lookup"><span data-stu-id="fab7b-162">toJSON()</span></span>
- <span data-ttu-id="fab7b-163">track()</span><span class="sxs-lookup"><span data-stu-id="fab7b-163">track</span></span>
- <span data-ttu-id="fab7b-164">untrack()</span><span class="sxs-lookup"><span data-stu-id="fab7b-164">untrack()</span></span>

### <a name="rangearea-specific-properties-and-methods"></a><span data-ttu-id="fab7b-165">RangeArea 特定属性和方法</span><span class="sxs-lookup"><span data-stu-id="fab7b-165">RangeArea-specific properties and methods</span></span>

<span data-ttu-id="fab7b-p107"> `RangeAreas` 类型具有某些不在 `Range` 对象上的属性和方法。以下是它们的选择：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p107">The `RangeAreas` type has some properties and methods that are not on the `Range` object. The following is a selection of them:</span></span>

- <span data-ttu-id="fab7b-p108">`areas`: `RangeCollection` 对象，其中包含所有 `RangeAreas` 对象表示的范围。 `RangeCollection` 对象也是新对象并且类似于其他 Excel 集合对象。它具有 `items` 属性，是一个 `Range` 对象的数组，表示范围。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p108">`areas`: A `RangeCollection` object that contains all of the ranges represented by the `RangeAreas` object. The `RangeCollection` object is also new and is similar to other Excel collection objects. It has an `items` property which is an array of `Range` objects representing the ranges.</span></span>
- <span data-ttu-id="fab7b-171">`areaCount`： `RangeAreas` 内范围的总数。</span><span class="sxs-lookup"><span data-stu-id="fab7b-171">The total number of recipients in the message.</span></span>
- <span data-ttu-id="fab7b-172">`getOffsetRangeAreas`：工作原理和 [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-) 一样，只不过返回一个 `RangeAreas`，它包含的范围与原 `RangeAreas` 中的某一范围产生偏移。</span><span class="sxs-lookup"><span data-stu-id="fab7b-172">`getOffsetRangeAreas`: Works just like [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), except that a `RangeAreas` is returned and it contains ranges that are each offset from one of the ranges in the original `RangeAreas`.</span></span>

## <a name="create-rangeareas-and-set-properties"></a><span data-ttu-id="fab7b-173">创建 RangeAreas 并设置属性</span><span class="sxs-lookup"><span data-stu-id="fab7b-173">Create RangeAreas and set properties</span></span>

<span data-ttu-id="fab7b-174">可以通过两种基本方式创建 `RangeAreas` 对象：</span><span class="sxs-lookup"><span data-stu-id="fab7b-174">You can create `RangeAreas` object in two basic ways:</span></span>

- <span data-ttu-id="fab7b-p109">调用 `Worksheet.getRanges()` ，将带逗号分隔范围地址的字符串传递给它。如果想要包含的任何范围已变成 [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem) ，可以在字符串中包含名称，而不是地址。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p109">Call `Worksheet.getRanges()` and pass it a string with comma-delimited range addresses. If any range you want to include has been made into a [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), you can include the name, instead of the address, in the string.</span></span>
- <span data-ttu-id="fab7b-p110">调用 `Workbook.getSelectedRanges()` 。此方法返回一个 `RangeAreas` ，表示在当前活动工作表上选择的所有范围。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p110">Call `Workbook.getSelectedRanges()`. This method returns a `RangeAreas` representing all the ranges that are selected on the currently active worksheet.</span></span>

<span data-ttu-id="fab7b-179">拥有 `RangeAreas` 对象后，可以使用返回诸如 `getOffsetRangeAreas` 和 `getIntersection` 等的 `RangeAreas` 对象创建其他对象。</span><span class="sxs-lookup"><span data-stu-id="fab7b-179">Once you have a `RangeAreas` object, you can create others using the methods on the object that return `RangeAreas` such as `getOffsetRangeAreas` and `getIntersection`.</span></span>

> [!NOTE]
> <span data-ttu-id="fab7b-p111">不能向 `RangeAreas` 对象直接添加其他范围。例如， `RangeAreas.areas` 中的集合不具有 `add` 方法。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p111">You cannot directly add additional ranges to a `RangeAreas` object. For example, the collection in `RangeAreas.areas` does not have an `add` method.</span></span>


> [!WARNING] 
> <span data-ttu-id="fab7b-p112">不要尝试直接添加或者删除 `RangeAreas.areas.items` 数组的成员。这样会引起代码出现异常。例如，可以将其他 `Range` 对象推送给数组，但是此举会引发错误，因为 `RangeAreas` 属性和方法并不表现出数组中存在新的项目。例如， `areaCount` 属性不包含以这种方式推送的范围，如果 `index` 大于 `areasCount-1` ， `RangeAreas.getItemAt(index)` 会引发一个错误。同样，通过获取引用或者调用其 `Range.delete`  方法来删除 `RangeAreas.areas.items` 数组中的 `Range` 对象会引发 bug ： 尽管 `Range` 对象 *被* 删除，但父级 `RangeAreas` 对象的属性和方法表现出，或者试图表现出，对象仍然存在。例如，如果代码调用 `RangeAreas.calculate` ， Office 便会对该范围进行计算，但会出错，因为该范围对象已经不存在。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p112">Do not attempt to directly add or delete members of the the `RangeAreas.areas.items` array. This will lead to undesirable behavior in your code. For example, it is possible to push an additional `Range` object onto the array, but doing so will cause errors because `RangeAreas` properties and methods behave as if the new item isn't there. For example, the `areaCount` property does not include ranges pushed in this way, and the `RangeAreas.getItemAt(index)` throws an error if `index` is larger than `areasCount-1`. Similarly, deleting a `Range` object in the `RangeAreas.areas.items` array by getting a reference to it and calling its `Range.delete` method causes bugs: although the `Range` object *is* deleted, the properties and methods of the parent `RangeAreas` object behave, or try to, as if it is still in existence. For example, if your code calls `RangeAreas.calculate`, Office will try to calculate the range, but will error because the range object is gone.</span></span>

<span data-ttu-id="fab7b-188">在 `RangeAreas` 上设置属性会在 `RangeAreas.areas` 集合内的所有范围上设置相应的属性。</span><span class="sxs-lookup"><span data-stu-id="fab7b-188">Setting a property on a `RangeAreas` sets the corresponding property on all the ranges in the `RangeAreas.areas` collection.</span></span>

<span data-ttu-id="fab7b-p113">以下是在多个范围设置属性的示例。该功能突出显示 **F3:F5** 和 **H3:H5** 范围。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p113">The following is an example of setting a property on multiple ranges. The function highlights the ranges **F3:F5** and **H3:H5**.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="fab7b-p114">本示例应用于可以硬编码传递给 `getRanges` 的范围地址或在运行时对其进行轻松计算的方案。一些可以真正将此包含在其中的方案：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p114">This example applies to scenarios in which you can hard code the range addresses that you pass to `getRanges` or easily calculate them at runtime. Some of the scenarios in which this would be true include:</span></span> 

- <span data-ttu-id="fab7b-193">代码在已知模板的上下文中运行。</span><span class="sxs-lookup"><span data-stu-id="fab7b-193">The code runs in the context of a known template.</span></span>
- <span data-ttu-id="fab7b-194">代码在导入数据的上下文中运行，其中数据架构为已知。</span><span class="sxs-lookup"><span data-stu-id="fab7b-194">The code runs in the context of imported data where the schema of the data is known.</span></span>

<span data-ttu-id="fab7b-p115">当不能在编码时确定需要执行操作的范围时，必须在运行时发现它们。下一节讨论这些方案。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p115">When you can't know at coding-time which ranges you need to operate on, you must discover them at runtime. The next section discusses these scenarios.</span></span>

### <a name="discover-range-areas-programmatically"></a><span data-ttu-id="fab7b-197">以编程方式发现范围</span><span class="sxs-lookup"><span data-stu-id="fab7b-197">Discover range areas programmatically</span></span>

<span data-ttu-id="fab7b-p116"> `Range.getSpecialCells()` 和 `Range.getSpecialCellsOrNullObject()` 方法能够根据单元格的特征和单元格值的类型，实现在运行时找到需要进行操作的范围。下面是 TypeScript 数据类型文件中的方法签名：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p116">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods enable you to find at runtime the ranges that you want to operate on based on the characteristics of the cells and the type of the values of the cells. Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="fab7b-p117">下面是使用第一个的示例。有关此代码，请注意：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p117">The following is an example of using the first one. About this code, note:</span></span>

- <span data-ttu-id="fab7b-202">它通过首先调用 `Worksheet.getUsedRange` 并针对仅该范围调用 `getSpecialCells` ，将搜索限制于工作表的所需部分。</span><span class="sxs-lookup"><span data-stu-id="fab7b-202">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="fab7b-p118">它将作为参数传递给 `getSpecialCells` 来自 `Excel.SpecialCellType` 枚举值的字符串版本。某些可以传递的其他值可以是空单元格的“空白”、非公式文本值单元格的"常量"以及拥有与 `usedRange` 中第一单元格相同条件格式的单元格的“ SameConditionalFormat ”。第一个单元格是最左上角单元格。有关枚举中值的完整列表，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p118">It passes as a parameter to `getSpecialCells` the string version of a value from the `Excel.SpecialCellType` enum. Some of the other values that could be passed instead are "Blanks" for empty cells, "Constants" for cells with literal values instead of formulas, and "SameConditionalFormat" for cells that have the same conditional formatting as the first cell in the `usedRange`. The first cell is the upper leftmost cell. For a complete list of the values in the enum, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).</span></span>
- <span data-ttu-id="fab7b-207"> `getSpecialCells` 方法返回 `RangeAreas` 对象，因此所有带公式的单元格都将以粉红色标示，即使它们并非全部连续。</span><span class="sxs-lookup"><span data-stu-id="fab7b-207">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span> 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="fab7b-p119">有时候范围没有 *任何* 具有目标特征的单元格。如果 `getSpecialCells` 找不到任何这样的单元格，它将引发 **ItemNotFound** 错误。如果有，这会将控制流转移至 `catch` block/ 方法。如果没有，错误暂停函数。可以有在不存在具有目标特征的单元格时如愿引发错误的方案。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p119">Sometimes the range doesn't have *any* cells with the targeted characteristic. If `getSpecialCells` doesn't find any, it throws an **ItemNotFound** error. This would divert the flow of control to a `catch` block/method, if there is one. If there isn't, the error halts the function. There may be scenarios in which throwing the error is exactly what you want to happen when there are no cells with the targeted characteristic.</span></span> 

<span data-ttu-id="fab7b-p120">但在一些它是正常的方案中，但可能不常见，因为有不匹配的单元格；代码应检查这种可能性并从容地对其进行处理，而不引发错误。对于这些方案，使用 `getSpecialCellsOrNullObject` 方法并测试 `RangeAreas.isNullObject` 属性。下面是一个示例。请注意此代码：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p120">But in scenarios in which it is normal, but perhaps uncommon, for there to be no matching cells; your code should check for this possibility and handle it gracefully without throwing an error. For these scenarios, use the `getSpecialCellsOrNullObject` method and test the `RangeAreas.isNullObject` property. The following is an example. Note about this code:</span></span>

- <span data-ttu-id="fab7b-p121"> `getSpecialCellsOrNullObject` 方法总是返回代理对象，因此从一般 JavaScript 意义上说，它不会是 `null` 。但是，如果找不到匹配的单元格， 对象的 `isNullObject` 属性设置为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p121">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense. But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="fab7b-p122">它将调用 `context.sync` ， *之前* 它测试 `isNullObject` 属性。这是所有 `*OrNullObject` 方法和属性的要求，因为始终要加载和同步属性以便读取它。但是，不需要 *明确地* 加载 `isNullObject` 属性。它是由 `context.sync` 自动加载的，即便在对象上不会对 `load` 进行调用。有关详细信息，请参阅 [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods) 。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p122">It calls `context.sync` *before* it tests the `isNullObject` property. This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it. However, it is not necessary to *explicitly* load the `isNullObject` property. It is automatically loaded by the `context.sync` even if `load` is not called on the object. For more information, see [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).</span></span>
- <span data-ttu-id="fab7b-p123">可以通过首先选择无公式单元格的范围来测试此代码并运行它。然后选择至少具有一个带公式单元格的范围并再次运行它。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p123">You can test this code by first selecting a range that has no formula cells and running it. Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="fab7b-226">为了简单起见，此文中所有其他示例均使用 `getSpecialCells` 方法，而不是 `getSpecialCellsOrNullObject` 。</span><span class="sxs-lookup"><span data-stu-id="fab7b-226">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

#### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="fab7b-227">通过单元格的值类型缩小目标单元格范围</span><span class="sxs-lookup"><span data-stu-id="fab7b-227">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="fab7b-p124">存在一个枚举类型 `Excel.SpecialCellValueType` 的可选第二参数，这进一步缩小了目标单元格范围。仅在将“公式”或者“常量”传递给 `getSpecialCells` 或者 `getSpecialCellsOrNullObject` 时才能使用它。此参数仅指定所要的具有某些值类型的单元格。有四种基本类型：“错误”、“逻辑”（即布尔值）、“数字”和“文本”。（枚举拥有这四种类型以外其他值，以下将对其进行讨论）下面是个示例。有关此代码，请注意：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p124">There is an optional second parameter, of enum type `Excel.SpecialCellValueType`, that further narrows down the cells to target. You can use it only when you pass either "Formulas" or "Constants" to `getSpecialCells` or `getSpecialCellsOrNullObject`. The parameter specifies that you only want cells with certain types of values. There are four basic types: "Error", "Logical" (which means boolean), "Numbers", and "Text". (The enum has other values besides these four which are discussed below.) The following is an example. About this code, note:</span></span>

- <span data-ttu-id="fab7b-p125">它仅突出显示具有文本数字值的单元格。它不会突出显示具有公式（即便结果是数字）或布尔值、 文本或错误状态的单元格。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p125">It will only highlight cells that have a literal number value. It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="fab7b-236">若要测试代码，请确保工作表拥有一些带文本数字值的单元格、一些带其他种类文本值的单元格，以及一些带公式的单元格。</span><span class="sxs-lookup"><span data-stu-id="fab7b-236">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="fab7b-p126">有时候需要对一个以上的单元格值类型进行操作，例如所有文本值和所有布尔值（“逻辑”）单元格。 `Excel.SpecialCellValueType` 枚举具有允许合并类型的值。例如，“逻辑文本”将针对所有布尔值和所有文本值。可以合并四种基本类型中的任意两种或任意三种类型。这些合并基本类型的枚举值的名称总是按照字母排序。因此要合并错误值、文本值以及布尔值单元格，应使用“ ErrorLogicalText ”，而不是“ LogicalErrorText ”或“ TextErrorLogical ”。"所有"默认参数合并全部四种类型。下面的示例突出显示所有生成数字或者布尔值的带公式单元格。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p126">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued ("Logical") cells. The `Excel.SpecialCellValueType` enum has values that let you combine types. For example, "LogicalText" will target all boolean and all text-valued cells. You can combine any two or any three of the four basic types. The names of these enum values that combine basic types are always in alphabetical order. So to combine error-valued, text-valued, and boolean-valued cells, use "ErrorLogicalText", not "LogicalErrorText" or "TextErrorLogical". The default parameter of "All" combines all four types. The following example highlights all cells with formulas that produce number or boolean values:</span></span>

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
> <span data-ttu-id="fab7b-245">`Excel.SpecialCellValueType` 参数仅在 `Excel.SpecialCellType` 参数为“公式”或“常量”时才可以使用。</span><span class="sxs-lookup"><span data-stu-id="fab7b-245">The ChildObjectTypes parameter can only be used if the AccessRights parameter is set to CreateChild or DeleteChild.</span></span>

### <a name="get-rangeareas-within-rangeareas"></a><span data-ttu-id="fab7b-246">获取 RangeAreas 中的 RangeAreas</span><span class="sxs-lookup"><span data-stu-id="fab7b-246">Get RangeAreas within RangeAreas</span></span>

<span data-ttu-id="fab7b-p127"> `RangeAreas` 类型本身也有 `getSpecialCells` 和 `getSpecialCellsOrNullObject` 方法，它们采用相同的两个参数。这些方法从 `RangeAreas.areas` 集合内的所有范围返回全部目标单元格。当在 `RangeAreas` 对象而不是 `Range` 对象上调用时，这些方法的行为存在一个小差异： 当“ SameConditionalFormat ”作为第一参数传递时，此方法会返回具有与 *`RangeAreas.areas` 集合内第一个范围中* 最左上角单元格相同条件格式的所有单元格。同一个点适用于“ SameDataValidation ”：在传递给 `Range.getSpecialCells` 时，它将返回具有与 \*此范围内* 最左上角单元格相同数据验证规则的全部单元格。但当传递给 `RangeAreas.getSpecialCells` 时，它将返回具与 * `RangeAreas.areas` 集合内第一个范围中\* 最左上角单元格相同数据验证规则的全部单元格。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p127">The `RangeAreas` type itself also has `getSpecialCells` and `getSpecialCellsOrNullObject` methods which take the same two parameters. These methods return all the targeted cells from all of the ranges in the `RangeAreas.areas` collection. There is one small difference in the behavior of the methods when called on a `RangeAreas` object instead of a `Range` object: when you pass "SameConditionalFormat" as the first parameter, the method returns all cells that have the same conditional formatting as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*. The same point applies to "SameDataValidation": when passed to `Range.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the range*. But when it is passed to `RangeAreas.getSpecialCells`, it returns all cells that have the same data validation rule as the upper leftmost cell *in the first range in the `RangeAreas.areas` collection*.</span></span>

## <a name="read-properties-of-rangeareas"></a><span data-ttu-id="fab7b-252">读取 RangeAreas 的属性</span><span class="sxs-lookup"><span data-stu-id="fab7b-252">Read properties of RangeAreas</span></span>

<span data-ttu-id="fab7b-p128">读取 `RangeAreas` 属性值需要小心，因为对于 `RangeAreas` 内的不同范围，给定的属性可能具有不同的值。一般规则是，如果 *可以* 返回一个一致的值则会将其返回。例如，在下面的代码中，粉红色 (`#FFC0CB`) 和 `true` 的 RGB 代码将被记录到控制台，因为 `RangeAreas` 对象中的两个范围均被充填粉红色而且两者均是整列。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p128">Reading property values of `RangeAreas` requires care, because a given property may have different values for different ranges within the `RangeAreas`. The general rule is that if a consistent value *can* be returned it will be returned. For example, in the following code, The RGB code for pink (`#FFC0CB`) and `true` will be logged to the console because both the ranges in the `RangeAreas` object have a pink fill and both are entire columns.</span></span>

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

<span data-ttu-id="fab7b-p129">当一致性为不可能时，情况会变得更加复杂。 `RangeAreas` 属性的行为遵循以下三个原则：</span><span class="sxs-lookup"><span data-stu-id="fab7b-p129">Things get more complicated when consistency isn't possible. The behavior of `RangeAreas` properties follows these three principles:</span></span>

- <span data-ttu-id="fab7b-258"> `RangeAreas` 对象的布尔值属性返回 `false` ，除非该属性对于所有成员范围均为真。</span><span class="sxs-lookup"><span data-stu-id="fab7b-258">A boolean property of a `RangeAreas` object returns `false` unless the property is true for all the member ranges.</span></span>
- <span data-ttu-id="fab7b-259">除 `address` 属性外的非布尔值属性将返回 `null`，除非所有成员范围的相应属性具有相同的值。</span><span class="sxs-lookup"><span data-stu-id="fab7b-259">Non-boolean properties, with the exception of the `address` property, return `null` unless the corresponding property on all the member ranges has the same value.</span></span>
- <span data-ttu-id="fab7b-260"> `address` 属性返回成员范围地址的逗号分隔字符串。</span><span class="sxs-lookup"><span data-stu-id="fab7b-260">The `address` property returns a comma-delimited string of the addresses of the member ranges.</span></span>

<span data-ttu-id="fab7b-p130">例如，下面的代码创建 `RangeAreas` ，其中只有一个范围是整个列，只有一个填充粉红色。控制台将显示填充颜色的 `null` 、 `isEntireRow` 属性的 `false` 和 `address`  属性的“ Sheet1 ！F3:F5，Sheet1 ！H:H ”（假定工作表名称为“ Sheet1 ”）。</span><span class="sxs-lookup"><span data-stu-id="fab7b-p130">For example, the following code creates a `RangeAreas` in which only one range is an entire column and only one is filled with pink. The console will show `null` for the fill color, `false` for the `isEntireRow` property, and "Sheet1!F3:F5, Sheet1!H:H" (assuming the sheet name is "Sheet1") for the `address` property.</span></span> 

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

## <a name="see-also"></a><span data-ttu-id="fab7b-263">另请参阅</span><span class="sxs-lookup"><span data-stu-id="fab7b-263">See also</span></span>

- [<span data-ttu-id="fab7b-264">使用 Excel JavaScript API 的基本编程概念</span><span class="sxs-lookup"><span data-stu-id="fab7b-264">Fundamental programming concepts with the Excel JavaScript API</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview)
- [<span data-ttu-id="fab7b-265">Range 对象（ JavaScript API for Excel ）</span><span class="sxs-lookup"><span data-stu-id="fab7b-265">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)
- <span data-ttu-id="fab7b-p131">[RangeAreas 对象（ JavaScript API for Excel ）](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) （预览 API 时该链接可能无法使用。作为替代方法，请参阅 [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 。）</span><span class="sxs-lookup"><span data-stu-id="fab7b-p131">[RangeAreas Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (This link may not work while the API is in preview. As an alternative, see [beta office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)</span></span>