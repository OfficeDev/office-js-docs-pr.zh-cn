---
title: 加载项Excel疑难解答
description: 了解如何解决加载项中的Excel错误。
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: cb622a1805be7bec61168ab37a41709a57075788
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349439"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="b7d6c-103">加载项Excel疑难解答</span><span class="sxs-lookup"><span data-stu-id="b7d6c-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="b7d6c-104">本文讨论对解决方案唯一的Excel。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="b7d6c-105">请使用页面底部的反馈工具，建议可添加到文章中的其他问题。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="b7d6c-106">活动工作簿切换时的 API 限制</span><span class="sxs-lookup"><span data-stu-id="b7d6c-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="b7d6c-107">加载项Excel一次对一个工作簿进行操作。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="b7d6c-108">与运行加载项的工作簿分开的工作簿获得焦点时，可能会出现错误。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="b7d6c-109">只有在焦点更改时调用特定方法时，才会发生此情况。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="b7d6c-110">以下 API 受此工作簿开关的影响。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-110">The following APIs are affected by this workbook switch.</span></span>

|<span data-ttu-id="b7d6c-111">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="b7d6c-111">Excel JavaScript API</span></span> | <span data-ttu-id="b7d6c-112">抛出的错误</span><span class="sxs-lookup"><span data-stu-id="b7d6c-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="b7d6c-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="b7d6c-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="b7d6c-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="b7d6c-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="b7d6c-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="b7d6c-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="b7d6c-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="b7d6c-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="b7d6c-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="b7d6c-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="b7d6c-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="b7d6c-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="b7d6c-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="b7d6c-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="b7d6c-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="b7d6c-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="b7d6c-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="b7d6c-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="b7d6c-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="b7d6c-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="b7d6c-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="b7d6c-129">这仅适用于在 Excel Mac 上打开的多个Windows工作簿。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="b7d6c-130">共同创作</span><span class="sxs-lookup"><span data-stu-id="b7d6c-130">Coauthoring</span></span>

<span data-ttu-id="b7d6c-131">有关[用于共同Excel](co-authoring-in-excel-add-ins.md)中的事件的模式，请参阅在加载项中共同授权。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="b7d6c-132">本文还讨论了使用某些 API（如 ）时的潜在合并冲突 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="known-issues"></a><span data-ttu-id="b7d6c-133">已知问题</span><span class="sxs-lookup"><span data-stu-id="b7d6c-133">Known Issues</span></span>

### <a name="binding-events-return-temporary-binding-obects"></a><span data-ttu-id="b7d6c-134">绑定事件返回 `Binding` 临时对象</span><span class="sxs-lookup"><span data-stu-id="b7d6c-134">Binding events return temporary `Binding` obects</span></span>

<span data-ttu-id="b7d6c-135">[BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding)和[BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding)都返回一个临时对象，其中包含引发 `Binding` `Binding` 该事件的对象的 ID。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-135">Both [BindingDataChangedEventArgs.binding](/javascript/api/excel/excel.bindingdatachangedeventargs#binding) and [BindingSelectionChangedEventArgs.binding](/javascript/api/excel/excel.bindingselectionchangedeventargs#binding) return a temporary `Binding` object that contains the ID of the `Binding` object that raised the event.</span></span> <span data-ttu-id="b7d6c-136">使用此 ID 检索 `BindingCollection.getItem(id)` `Binding` 引发事件的对象。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-136">Use this ID with `BindingCollection.getItem(id)` to retrieve the `Binding` object that raised the event.</span></span>

<span data-ttu-id="b7d6c-137">下面的代码示例演示如何使用此临时绑定 ID 检索相关 `Binding` 对象。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-137">The following code sample shows how to use this temporary binding ID to retrieve the related `Binding` object.</span></span> <span data-ttu-id="b7d6c-138">在示例中，将事件侦听器分配给绑定。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-138">In the sample, an event listener is assigned to a binding.</span></span> <span data-ttu-id="b7d6c-139">侦听器在 `getBindingId` 触发事件 `onDataChanged` 时调用 方法。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-139">The listener calls the `getBindingId` method when the `onDataChanged` event is triggered.</span></span> <span data-ttu-id="b7d6c-140">`getBindingId`方法使用临时对象的 ID `Binding` 检索 `Binding` 引发事件的对象。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-140">The `getBindingId` method uses the ID of the temporary `Binding` object to retrieve the `Binding` object that raised the event.</span></span>

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

### <a name="cell-format-usestandardheight-and-usestandardwidth-issues"></a><span data-ttu-id="b7d6c-141">单元格格式 `useStandardHeight` 和 `useStandardWidth` 问题</span><span class="sxs-lookup"><span data-stu-id="b7d6c-141">Cell format `useStandardHeight` and `useStandardWidth` issues</span></span>

<span data-ttu-id="b7d6c-142">[的 useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight)属性在属性中 `CellPropertiesFormat` Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-142">The [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) property of `CellPropertiesFormat` doesn't work properly in Excel on the web.</span></span> <span data-ttu-id="b7d6c-143">由于用户界面中Excel web 版问题，因此将 属性设置为不精确地在此平台上 `useStandardHeight` `true` 计算高度。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-143">Due to an issue in the Excel on the web UI, setting the `useStandardHeight` property to `true` calculates height imprecisely on this platform.</span></span> <span data-ttu-id="b7d6c-144">例如，标准高度 **14** 在 Excel web 版 中修改为 **14.25。**</span><span class="sxs-lookup"><span data-stu-id="b7d6c-144">For example, a standard height of **14** is modified to **14.25** in Excel on the web.</span></span>

<span data-ttu-id="b7d6c-145">在所有平台上 [，useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) 和 [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) 属性仅旨在 `CellPropertiesFormat` 设置为 `true` 。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-145">On all platforms, the [useStandardHeight](/javascript/api/excel/excel.cellpropertiesformat#useStandardHeight) and [useStandardWidth](/javascript/api/excel/excel.cellpropertiesformat#useStandardWidth) properties of `CellPropertiesFormat` are only intended to be set to `true`.</span></span> <span data-ttu-id="b7d6c-146">将这些属性设置为 `false` 无效。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-146">Setting these properties to `false` has no effect.</span></span> 

### <a name="range-getimage-method-unsupported-on-excel-for-mac"></a><span data-ttu-id="b7d6c-147">区域 `getImage` 方法不受支持Excel for Mac</span><span class="sxs-lookup"><span data-stu-id="b7d6c-147">Range `getImage` method unsupported on Excel for Mac</span></span>

<span data-ttu-id="b7d6c-148">Range [getImage](/javascript/api/excel/excel.range#getImage__)方法当前在 Excel for Mac。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-148">The Range [getImage](/javascript/api/excel/excel.range#getImage__) method isn't currently supported in Excel for Mac.</span></span> <span data-ttu-id="b7d6c-149">请参阅 [OfficeDev/office-js issue #235](https://github.com/OfficeDev/office-js/issues/235) 了解当前状态。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-149">See [OfficeDev/office-js Issue #235](https://github.com/OfficeDev/office-js/issues/235) for the current status.</span></span>

### <a name="range-return-character-limit"></a><span data-ttu-id="b7d6c-150">区域返回字符限制</span><span class="sxs-lookup"><span data-stu-id="b7d6c-150">Range return character limit</span></span>

<span data-ttu-id="b7d6c-151">[Worksheet.getRange (address) ](/javascript/api/excel/excel.worksheet#getRange_address_) [和 Worksheet.getRanges](/javascript/api/excel/excel.worksheet#getRanges_address_) (address) 方法的地址字符串限制为 8192 个字符。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-151">The [Worksheet.getRange(address)](/javascript/api/excel/excel.worksheet#getRange_address_) and [Worksheet.getRanges(address)](/javascript/api/excel/excel.worksheet#getRanges_address_) methods have an address string limit of 8192 characters.</span></span> <span data-ttu-id="b7d6c-152">超出此限制时，地址字符串将被截断为 8192 个字符。</span><span class="sxs-lookup"><span data-stu-id="b7d6c-152">When this limit is exceeded, the address string is truncated to 8192 characters.</span></span>

## <a name="see-also"></a><span data-ttu-id="b7d6c-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b7d6c-153">See also</span></span>

- [<span data-ttu-id="b7d6c-154">排查Office加载项的开发错误</span><span class="sxs-lookup"><span data-stu-id="b7d6c-154">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="b7d6c-155">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="b7d6c-155">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
