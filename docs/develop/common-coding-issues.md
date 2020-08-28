---
title: 常见问题和意外平台行为的编码指南
description: 开发人员经常遇到的 Office JavaScript API 平台问题的列表。
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: f6d6a31059b32550e3176ed278d7da4c2c7a6c68
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292909"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="6a50b-103">常见问题和意外平台行为的编码指南</span><span class="sxs-lookup"><span data-stu-id="6a50b-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="6a50b-104">本文重点介绍了 Office JavaScript API 的各个方面，这些方面可能导致意外行为或需要特定编码模式来实现所需的结果。</span><span class="sxs-lookup"><span data-stu-id="6a50b-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="6a50b-105">如果遇到此列表中的问题，请使用文章底部的反馈表单告知我们。</span><span class="sxs-lookup"><span data-stu-id="6a50b-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="6a50b-106">通用 Api 和 Outlook Api 不基于承诺</span><span class="sxs-lookup"><span data-stu-id="6a50b-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="6a50b-107">[公共 api](/javascript/api/office) (那些与特定 Office 应用程序不关联的 api) 并且[Outlook api](/javascript/api/outlook)使用基于回调的编程模型。</span><span class="sxs-lookup"><span data-stu-id="6a50b-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office application) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="6a50b-108">与基础 Office 文档进行交互需要进行异步读取或写入调用，以指定在操作完成时运行的回调。</span><span class="sxs-lookup"><span data-stu-id="6a50b-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be run when the operation completes.</span></span> <span data-ttu-id="6a50b-109">有关此模式的示例，请参阅 [document.getfileasync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="6a50b-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="6a50b-110">这些常见 API 和 Outlook API 方法不会返回 [承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。</span><span class="sxs-lookup"><span data-stu-id="6a50b-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="6a50b-111">因此，在异步操作完成之前，不能使用 [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) 暂停执行。</span><span class="sxs-lookup"><span data-stu-id="6a50b-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="6a50b-112">如果需要 `await` 行为，可以在显式创建的承诺中包装方法调用。</span><span class="sxs-lookup"><span data-stu-id="6a50b-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

```js
readDocumentFileAsync(): Promise<any> {
    return new Promise((resolve, reject) => {
        const chunkSize = 65536;
        const self = this;

        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: chunkSize }, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                reject(asyncResult.error);
            } else {
                // `getAllSlices` is a Promise-wrapped implementation of File.getSliceAsync.
                self.getAllSlices(asyncResult.value).then(result => {
                    if (result.IsSuccess) {
                        resolve(result.Data);
                    } else {
                        reject(asyncResult.error);
                    }
                });
            }
        });
    });
}
```

> [!NOTE]
> <span data-ttu-id="6a50b-113">参考文档包含 [getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)的承诺包装实现。</span><span class="sxs-lookup"><span data-stu-id="6a50b-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="6a50b-114">某些属性不能直接设置</span><span class="sxs-lookup"><span data-stu-id="6a50b-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="6a50b-115">本部分仅适用于适用于 Excel 和 Word 的应用程序特定的 Api。</span><span class="sxs-lookup"><span data-stu-id="6a50b-115">This section only applies to the application-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="6a50b-116">某些属性虽然是可写的，但不能设置。</span><span class="sxs-lookup"><span data-stu-id="6a50b-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="6a50b-117">这些属性是父属性的一部分，必须将其设置为单个对象。</span><span class="sxs-lookup"><span data-stu-id="6a50b-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="6a50b-118">这是因为该父属性依赖具有特定逻辑关系的子属性。</span><span class="sxs-lookup"><span data-stu-id="6a50b-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="6a50b-119">必须使用对象文本表示法设置这些父属性，以设置整个对象，而不是设置该对象的单个子属性。</span><span class="sxs-lookup"><span data-stu-id="6a50b-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="6a50b-120">在 [页面布局](/javascript/api/excel/excel.pagelayout)中找到此示例的一个示例。</span><span class="sxs-lookup"><span data-stu-id="6a50b-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="6a50b-121">`zoom`必须使用单个[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)对象设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="6a50b-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="6a50b-122">在上面的示例中，您 ***将无法*** 直接分配 `zoom` 值： `sheet.pageLayout.zoom.scale = 200;` 。</span><span class="sxs-lookup"><span data-stu-id="6a50b-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="6a50b-123">由于未加载，该语句 `zoom` 会引发错误。</span><span class="sxs-lookup"><span data-stu-id="6a50b-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="6a50b-124">即使 `zoom` 要加载，该扩展集也不会生效。</span><span class="sxs-lookup"><span data-stu-id="6a50b-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="6a50b-125">发生所有上下文操作 `zoom` ，刷新加载项中的代理对象并覆盖本地设置的值。</span><span class="sxs-lookup"><span data-stu-id="6a50b-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="6a50b-126">此行为不同于 [导航属性](application-specific-api-model.md#scalar-and-navigation-properties) ，如 [Range. 格式](/javascript/api/excel/excel.range#format)。</span><span class="sxs-lookup"><span data-stu-id="6a50b-126">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="6a50b-127">`format`可以使用对象导航设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="6a50b-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="6a50b-128">您可以通过检查属性的只读修饰符来标识无法直接设置其子属性的属性。</span><span class="sxs-lookup"><span data-stu-id="6a50b-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="6a50b-129">所有只读属性都可以直接设置其非只读的子属性。</span><span class="sxs-lookup"><span data-stu-id="6a50b-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="6a50b-130">`PageLayout.zoom`必须使用该级别的对象设置可写属性（如必须设置）。</span><span class="sxs-lookup"><span data-stu-id="6a50b-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="6a50b-131">摘要：</span><span class="sxs-lookup"><span data-stu-id="6a50b-131">In summary:</span></span>

- <span data-ttu-id="6a50b-132">只读属性：可通过导航设置子属性。</span><span class="sxs-lookup"><span data-stu-id="6a50b-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="6a50b-133">可写属性：不能通过导航 (设置子属性，而必须将初始父对象分配) 的一部分。</span><span class="sxs-lookup"><span data-stu-id="6a50b-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="6a50b-134">设置只读属性</span><span class="sxs-lookup"><span data-stu-id="6a50b-134">Setting read-only properties</span></span>

<span data-ttu-id="6a50b-135">Office JS 的 [TypeScript 定义](referencing-the-javascript-api-for-office-library-from-its-cdn.md) 指定哪些对象属性是只读的。</span><span class="sxs-lookup"><span data-stu-id="6a50b-135">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="6a50b-136">如果尝试设置只读属性，写入操作将无提示地失败，且不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="6a50b-136">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="6a50b-137">下面的示例错误地尝试设置只读属性 [Chart.id](/javascript/api/excel/excel.chart#id)。</span><span class="sxs-lookup"><span data-stu-id="6a50b-137">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="6a50b-138">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="6a50b-138">Removing event handlers</span></span>

<span data-ttu-id="6a50b-139">必须使用在其中添加事件处理程序的相同项将其删除 `RequestContext` 。</span><span class="sxs-lookup"><span data-stu-id="6a50b-139">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="6a50b-140">如果需要加载项在运行时删除事件处理程序，则需要存储用于添加处理程序的 context 对象。</span><span class="sxs-lookup"><span data-stu-id="6a50b-140">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

```js
Excel.run(async (context) => {
    [...]

    // To later remove an event handler, store the context somewhere accessible to the handler removal function.
    // You may find it helpful to also store the event handler object and associate it with the context.
    selectionChangedHandler = myWorksheet.onSelectionChanged.add(callback);
    savedContext = currentContext;
    return context.sync();
}
```

## <a name="supporting-internet-explorer"></a><span data-ttu-id="6a50b-141">支持 Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="6a50b-141">Supporting Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="excel-specific-issues"></a><span data-ttu-id="6a50b-142">特定于 Excel 的问题</span><span class="sxs-lookup"><span data-stu-id="6a50b-142">Excel-specific issues</span></span>

### <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="6a50b-143">活动工作簿切换时的 API 限制</span><span class="sxs-lookup"><span data-stu-id="6a50b-143">API limitations when the active workbook switches</span></span>

<span data-ttu-id="6a50b-144">Excel 相关外接程序用于一次运行单个工作簿。</span><span class="sxs-lookup"><span data-stu-id="6a50b-144">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="6a50b-145">当运行加载项的工作簿获得焦点时，可能会出现错误。</span><span class="sxs-lookup"><span data-stu-id="6a50b-145">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="6a50b-146">仅当焦点更改时要调用的特定方法时，才会发生这种情况。</span><span class="sxs-lookup"><span data-stu-id="6a50b-146">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="6a50b-147">此工作簿开关会影响以下 Api：</span><span class="sxs-lookup"><span data-stu-id="6a50b-147">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="6a50b-148">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="6a50b-148">Excel JavaScript API</span></span> | <span data-ttu-id="6a50b-149">引发的错误</span><span class="sxs-lookup"><span data-stu-id="6a50b-149">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="6a50b-150">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-150">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="6a50b-151">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-151">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="6a50b-152">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-152">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="6a50b-153">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="6a50b-153">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="6a50b-154">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="6a50b-154">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="6a50b-155">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="6a50b-155">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="6a50b-156">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-156">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="6a50b-157">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="6a50b-157">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="6a50b-158">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-158">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="6a50b-159">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-159">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="6a50b-160">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-160">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="6a50b-161">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-161">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="6a50b-162">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-162">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="6a50b-163">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-163">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="6a50b-164">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-164">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="6a50b-165">GeneralException</span><span class="sxs-lookup"><span data-stu-id="6a50b-165">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="6a50b-166">这仅适用于在 Windows 或 Mac 上打开的多个 Excel 工作簿。</span><span class="sxs-lookup"><span data-stu-id="6a50b-166">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

### <a name="coauthoring"></a><span data-ttu-id="6a50b-167">共同创作</span><span class="sxs-lookup"><span data-stu-id="6a50b-167">Coauthoring</span></span>

<span data-ttu-id="6a50b-168">请参阅 [Excel 外接程序中](../excel/co-authoring-in-excel-add-ins.md) 用于共同创作环境中事件的模式的合著。</span><span class="sxs-lookup"><span data-stu-id="6a50b-168">See [Coauthoring in Excel add-ins](../excel/co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="6a50b-169">本文还讨论了使用某些 Api （例如）时的潜在合并冲突 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。</span><span class="sxs-lookup"><span data-stu-id="6a50b-169">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="6a50b-170">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6a50b-170">See also</span></span>

- [<span data-ttu-id="6a50b-171">Office 外接程序的资源限制和性能优化</span><span class="sxs-lookup"><span data-stu-id="6a50b-171">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
- <span data-ttu-id="6a50b-172">[OfficeDev/？ js](https://github.com/OfficeDev/office-js/issues)：报告和查看 office 外接程序平台和 JavaScript api 中的问题的位置。</span><span class="sxs-lookup"><span data-stu-id="6a50b-172">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="6a50b-173">[堆栈溢出](https://stackoverflow.com/questions/tagged/office-js)：询问并查看有关 Office JavaScript api 的编程问题的位置。</span><span class="sxs-lookup"><span data-stu-id="6a50b-173">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="6a50b-174">在发布到堆栈溢出时，请务必对您的问题应用 "office-js" 标记。</span><span class="sxs-lookup"><span data-stu-id="6a50b-174">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="6a50b-175">[UserVoice](https://officespdev.uservoice.com/)：建议 Office 外接程序平台和 Office JavaScript api 的新功能的位置。</span><span class="sxs-lookup"><span data-stu-id="6a50b-175">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
