---
title: Excel 加载项疑难解答
description: 了解如何解决 Excel 加载项中的开发错误。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: 0efc8b4d25d9d748975146e187104972e4ad58a9
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270726"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="c7694-103">Excel 加载项疑难解答</span><span class="sxs-lookup"><span data-stu-id="c7694-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="c7694-104">本文讨论 Excel 特有的疑难解答问题。</span><span class="sxs-lookup"><span data-stu-id="c7694-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="c7694-105">请使用页面底部的反馈工具建议可添加到文章中的其他问题。</span><span class="sxs-lookup"><span data-stu-id="c7694-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="c7694-106">活动工作簿切换时的 API 限制</span><span class="sxs-lookup"><span data-stu-id="c7694-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="c7694-107">Excel 外接程序旨在一次对一个工作簿进行操作。</span><span class="sxs-lookup"><span data-stu-id="c7694-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="c7694-108">当与运行加载项的工作簿分开的工作簿获得焦点时，可能会出现错误。</span><span class="sxs-lookup"><span data-stu-id="c7694-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="c7694-109">只有在焦点更改时调用特定方法时，才会发生此情况。</span><span class="sxs-lookup"><span data-stu-id="c7694-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="c7694-110">以下 API 受此工作簿开关的影响：</span><span class="sxs-lookup"><span data-stu-id="c7694-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="c7694-111">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="c7694-111">Excel JavaScript API</span></span> | <span data-ttu-id="c7694-112">引发错误</span><span class="sxs-lookup"><span data-stu-id="c7694-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="c7694-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="c7694-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="c7694-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="c7694-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c7694-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="c7694-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c7694-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="c7694-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c7694-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="c7694-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="c7694-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="c7694-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="c7694-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="c7694-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="c7694-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="c7694-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="c7694-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="c7694-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="c7694-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="c7694-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="c7694-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="c7694-129">这仅适用于在 Windows 或 Mac 上打开的多个 Excel 工作簿。</span><span class="sxs-lookup"><span data-stu-id="c7694-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="c7694-130">共同创作</span><span class="sxs-lookup"><span data-stu-id="c7694-130">Coauthoring</span></span>

<span data-ttu-id="c7694-131">请参阅 [Excel 加载项中的](co-authoring-in-excel-add-ins.md) 共同授权，了解用于共同授权环境中事件的模式。</span><span class="sxs-lookup"><span data-stu-id="c7694-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="c7694-132">本文还讨论使用某些 API 时的潜在合并冲突，例如 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。</span><span class="sxs-lookup"><span data-stu-id="c7694-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="known-issues"></a><span data-ttu-id="c7694-133">已知问题</span><span class="sxs-lookup"><span data-stu-id="c7694-133">Known Issues</span></span>

### <a name="binding-events-return-temporary-binding-obects"></a><span data-ttu-id="c7694-134">绑定事件返回临时 `Binding` 对象</span><span class="sxs-lookup"><span data-stu-id="c7694-134">Binding events return temporary `Binding` obects</span></span>

<span data-ttu-id="c7694-135">[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding)和[BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding)都返回一个临时对象，该对象包含引发该事件 `Binding` 的对象的 `Binding` ID。</span><span class="sxs-lookup"><span data-stu-id="c7694-135">Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event.</span></span> <span data-ttu-id="c7694-136">使用此 ID `BindingCollection.getItem(id)` 检索 `Binding` 引发事件的对象。</span><span class="sxs-lookup"><span data-stu-id="c7694-136">Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.</span></span>

<span data-ttu-id="c7694-137">下面的代码示例演示如何使用此临时绑定 ID 检索相关 `Binding` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7694-137">The following code sample shows how to use this temporary binding ID to retrieve the related `Binding` object.</span></span> <span data-ttu-id="c7694-138">在示例中，将事件侦听器分配给绑定。</span><span class="sxs-lookup"><span data-stu-id="c7694-138">In the sample, an event listener is assigned to a binding.</span></span> <span data-ttu-id="c7694-139">当触发 `getBindingId` 事件时，侦听器 `onDataChanged` 将调用该方法。</span><span class="sxs-lookup"><span data-stu-id="c7694-139">The listener calls the `getBindingId` method when the `onDataChanged` event is triggered.</span></span> <span data-ttu-id="c7694-140">`getBindingId`该方法使用临时对象的 ID 检索 `Binding` `Binding` 引发事件的对象。</span><span class="sxs-lookup"><span data-stu-id="c7694-140">The `getBindingId` method uses the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.</span></span>

```js
Excel.run(function (context) {
    // Retrieve your binding.
    var binding = context.workbook.bindings.getItemAt(0);

    return context.sync().then(function () {
        // Register an event listener to detect changes to your binding
        // and then trigger the `getBindingId` method when the data changes. 
        binding.onDataChanged.add(getBindingId);

        return context.sync();
    });
});

function getBindingId(eventArgs) {
    return Excel.run(function (context) {
        // Get the temporary binding object and load its ID. 
        var tempBindingObject = eventArgs.binding;
        tempBindingObject.load("id");

        // Use the temporary binding object's ID to retrieve the original binding object. 
        var originalBindingObject = context.workbook.bindings.getItem(tempBindingObject.id);

        // You now have the binding object that raised the event: `originalBindingObject`. 
    });
}
```

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a><span data-ttu-id="c7694-141">单元格格式 `useStandardHeight` `useStandardWidth` 和问题</span><span class="sxs-lookup"><span data-stu-id="c7694-141">Cell format `useStandardHeight` and `useStandardWidth` issues</span></span>

<span data-ttu-id="c7694-142">[useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)属性在 `CellPropertiesFormat` Excel 网页中无法正常工作。</span><span class="sxs-lookup"><span data-stu-id="c7694-142">The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web.</span></span> <span data-ttu-id="c7694-143">由于 Excel 网页 UI 中的问题，将该属性设置为在此平台上计算高度不 `useStandardHeight` `true` 精确。</span><span class="sxs-lookup"><span data-stu-id="c7694-143">Due to an issue in the Excel on the web UI, setting the `useStandardHeight` property to `true` calculates height imprecisely on this platform.</span></span> <span data-ttu-id="c7694-144">例如，在 Excel 网页版中，标准高度 **14** 修改为 **14.25。**</span><span class="sxs-lookup"><span data-stu-id="c7694-144">For example, a standard height of **14** is modified to **14.25** in Excel on the web.</span></span>

<span data-ttu-id="c7694-145">在所有平台上 [，useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) 和 [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) 属性仅用于 `CellPropertiesFormat` 设置为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="c7694-145">On all platforms, the [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) and [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) properties of `CellPropertiesFormat` are only intended to be set to `true`.</span></span> <span data-ttu-id="c7694-146">将这些属性设置为 `false` 不起作用。</span><span class="sxs-lookup"><span data-stu-id="c7694-146">Setting these properties to `false` has no effect.</span></span> 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a><span data-ttu-id="c7694-147">Excel `getImage` for Mac 不支持 Range 方法</span><span class="sxs-lookup"><span data-stu-id="c7694-147">Range `getImage` method unsupported on Excel for Mac</span></span>

<span data-ttu-id="c7694-148">Excel for Mac 当前不支持 Range [getImage](/javascript/api/excel/excel.range#getImage__) 方法。</span><span class="sxs-lookup"><span data-stu-id="c7694-148">The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac.</span></span> <span data-ttu-id="c7694-149">请参阅 [OfficeDev/office-js #235](https://github.com/OfficeDev/office-js/issues/235) 当前状态。</span><span class="sxs-lookup"><span data-stu-id="c7694-149">See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.</span></span>

### <a name="range-return-character-limit"></a><span data-ttu-id="c7694-150">区域返回字符限制</span><span class="sxs-lookup"><span data-stu-id="c7694-150">Range return character limit</span></span>

<span data-ttu-id="c7694-151">[Worksheet.getRange (address) ](/javascript/api/excel/excel.worksheet#getRange_address_) [和 Worksheet.getRanges (address) ](/javascript/api/excel/excel.worksheet#getRanges_address_)方法的地址字符串限制为 8192 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7694-151">The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters.</span></span> <span data-ttu-id="c7694-152">超过此限制时，地址字符串将被截断为 8192 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7694-152">When this limit is exceeded, the address string is truncated to 8192 characters.</span></span>

## <a name="see-also"></a><span data-ttu-id="c7694-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c7694-153">See also</span></span>

- [<span data-ttu-id="c7694-154">Office 加载项的开发错误疑难解答</span><span class="sxs-lookup"><span data-stu-id="c7694-154">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="c7694-155">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="c7694-155">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
