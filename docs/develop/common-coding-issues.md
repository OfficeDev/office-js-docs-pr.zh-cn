---
title: 常见问题和意外平台行为的编码指南
description: 开发人员经常遇到的 Office JavaScript API 平台问题的列表。
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: dea879899dce2e957d34f2eb8e7498d4fdb868c0
ms.sourcegitcommit: 0fdb78cefa669b727b817614a4147a46d249a0ed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/28/2020
ms.locfileid: "43930314"
---
# <a name="coding-guidance-for-common-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="80eed-103">常见问题和意外平台行为的编码指南</span><span class="sxs-lookup"><span data-stu-id="80eed-103">Coding guidance for common issues and unexpected platform behaviors</span></span>

<span data-ttu-id="80eed-104">本文重点介绍了 Office JavaScript API 的各个方面，这些方面可能导致意外行为或需要特定编码模式来实现所需的结果。</span><span class="sxs-lookup"><span data-stu-id="80eed-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="80eed-105">如果遇到此列表中的问题，请使用文章底部的反馈表单告知我们。</span><span class="sxs-lookup"><span data-stu-id="80eed-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="80eed-106">通用 Api 和 Outlook Api 不基于承诺</span><span class="sxs-lookup"><span data-stu-id="80eed-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="80eed-107">[通用 api](/javascript/api/office) （那些未绑定到特定 Office 主机的 api）和[Outlook api](/javascript/api/outlook)使用基于回调的编程模型。</span><span class="sxs-lookup"><span data-stu-id="80eed-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="80eed-108">与基础 Office 文档进行交互需要进行异步读取或写入调用，以指定在操作完成时要运行的回调。</span><span class="sxs-lookup"><span data-stu-id="80eed-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="80eed-109">有关此模式的示例，请参阅[document.getfileasync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="80eed-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="80eed-110">这些常见 API 和 Outlook API 方法不会返回[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。</span><span class="sxs-lookup"><span data-stu-id="80eed-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="80eed-111">因此，在异步操作完成之前，不能使用[await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await)暂停执行。</span><span class="sxs-lookup"><span data-stu-id="80eed-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="80eed-112">如果需要`await`行为，可以在显式创建的承诺中包装方法调用。</span><span class="sxs-lookup"><span data-stu-id="80eed-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="80eed-113">参考文档包含[getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)的承诺包装实现。</span><span class="sxs-lookup"><span data-stu-id="80eed-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="80eed-114">某些属性不能直接设置</span><span class="sxs-lookup"><span data-stu-id="80eed-114">Some properties cannot be set directly</span></span>

> [!NOTE]
> <span data-ttu-id="80eed-115">本部分仅适用于 Excel 和 Word 的特定于主机的 Api。</span><span class="sxs-lookup"><span data-stu-id="80eed-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="80eed-116">某些属性虽然是可写的，但不能设置。</span><span class="sxs-lookup"><span data-stu-id="80eed-116">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="80eed-117">这些属性是父属性的一部分，必须将其设置为单个对象。</span><span class="sxs-lookup"><span data-stu-id="80eed-117">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="80eed-118">这是因为该父属性依赖具有特定逻辑关系的子属性。</span><span class="sxs-lookup"><span data-stu-id="80eed-118">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="80eed-119">必须使用对象文本表示法设置这些父属性，以设置整个对象，而不是设置该对象的单个子属性。</span><span class="sxs-lookup"><span data-stu-id="80eed-119">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="80eed-120">在[页面布局](/javascript/api/excel/excel.pagelayout)中找到此示例的一个示例。</span><span class="sxs-lookup"><span data-stu-id="80eed-120">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="80eed-121">必须`zoom`使用单个[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)对象设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="80eed-121">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="80eed-122">在上面的示例中，您***将无法***直接分配`zoom`值： `sheet.pageLayout.zoom.scale = 200;`。</span><span class="sxs-lookup"><span data-stu-id="80eed-122">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="80eed-123">由于`zoom`未加载，该语句会引发错误。</span><span class="sxs-lookup"><span data-stu-id="80eed-123">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="80eed-124">`zoom`即使要加载，该扩展集也不会生效。</span><span class="sxs-lookup"><span data-stu-id="80eed-124">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="80eed-125">发生所有上下文操作`zoom`，刷新加载项中的代理对象并覆盖本地设置的值。</span><span class="sxs-lookup"><span data-stu-id="80eed-125">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="80eed-126">此行为不同于[导航属性](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)，如[Range. 格式](/javascript/api/excel/excel.range#format)。</span><span class="sxs-lookup"><span data-stu-id="80eed-126">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="80eed-127">`format`可以使用对象导航设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="80eed-127">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="80eed-128">您可以通过检查属性的只读修饰符来标识无法直接设置其子属性的属性。</span><span class="sxs-lookup"><span data-stu-id="80eed-128">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="80eed-129">所有只读属性都可以直接设置其非只读的子属性。</span><span class="sxs-lookup"><span data-stu-id="80eed-129">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="80eed-130">必须使用该`PageLayout.zoom`级别的对象设置可写属性（如必须设置）。</span><span class="sxs-lookup"><span data-stu-id="80eed-130">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="80eed-131">摘要：</span><span class="sxs-lookup"><span data-stu-id="80eed-131">In summary:</span></span>

- <span data-ttu-id="80eed-132">只读属性：可通过导航设置子属性。</span><span class="sxs-lookup"><span data-stu-id="80eed-132">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="80eed-133">可写属性：子属性不能通过导航设置（必须设置为初始父对象分配的一部分）。</span><span class="sxs-lookup"><span data-stu-id="80eed-133">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>

## <a name="excel-data-transfer-limits"></a><span data-ttu-id="80eed-134">Excel 数据传输限制</span><span class="sxs-lookup"><span data-stu-id="80eed-134">Excel data transfer limits</span></span>

<span data-ttu-id="80eed-135">如果您正在构建 Excel 外接程序，请注意与工作簿交互时的以下大小限制：</span><span class="sxs-lookup"><span data-stu-id="80eed-135">If you're building an Excel add-in, be aware of the following size limitations when interacting with the workbook:</span></span>

- <span data-ttu-id="80eed-136">Excel 网页版将请求和响应的有效负载大小限制为 5MB。</span><span class="sxs-lookup"><span data-stu-id="80eed-136">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="80eed-137">如果超过该限制，将引发 `RichAPI.Error`。</span><span class="sxs-lookup"><span data-stu-id="80eed-137">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="80eed-138">对于 get 操作，范围限制为5000000个单元格。</span><span class="sxs-lookup"><span data-stu-id="80eed-138">A range is limited to five million cells for get operations.</span></span>

<span data-ttu-id="80eed-139">如果您希望用户输入超出这些限制，请务必先检查数据，然后再调用`context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="80eed-139">If you expect user input to exceed these limits, be sure to check the data before calling `context.sync()`.</span></span> <span data-ttu-id="80eed-140">根据需要将操作拆分为较小的部分。</span><span class="sxs-lookup"><span data-stu-id="80eed-140">Split the operation into smaller pieces as needed.</span></span> <span data-ttu-id="80eed-141">请务必为每`context.sync()`个子操作调用，以避免这些操作再次成批组合。</span><span class="sxs-lookup"><span data-stu-id="80eed-141">Be sure to call `context.sync()` for each sub-operation to avoid those operations getting batched together again.</span></span>

<span data-ttu-id="80eed-142">这些限制通常由大型区域所超过。</span><span class="sxs-lookup"><span data-stu-id="80eed-142">These limitations are typically exceeded by large ranges.</span></span> <span data-ttu-id="80eed-143">您的外接程序可能能够使用[RangeAreas](/javascript/api/excel/excel.rangeareas)对较大范围内的单元格进行战略更新。</span><span class="sxs-lookup"><span data-stu-id="80eed-143">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="80eed-144">有关详细信息，请参阅[在 Excel 外接程序中同时处理多个区域](../excel/excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="80eed-144">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="80eed-145">设置只读属性</span><span class="sxs-lookup"><span data-stu-id="80eed-145">Setting read-only properties</span></span>

<span data-ttu-id="80eed-146">Office JS 的[TypeScript 定义](referencing-the-javascript-api-for-office-library-from-its-cdn.md)指定哪些对象属性是只读的。</span><span class="sxs-lookup"><span data-stu-id="80eed-146">The [TypeScript definitions](referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="80eed-147">如果尝试设置只读属性，写入操作将无提示地失败，且不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="80eed-147">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="80eed-148">下面的示例错误地尝试设置只读属性[Chart.id](/javascript/api/excel/excel.chart#id)。</span><span class="sxs-lookup"><span data-stu-id="80eed-148">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="removing-event-handlers"></a><span data-ttu-id="80eed-149">删除事件处理程序</span><span class="sxs-lookup"><span data-stu-id="80eed-149">Removing event handlers</span></span>

<span data-ttu-id="80eed-150">必须使用在其中添加事件处理程序`RequestContext`的相同项将其删除。</span><span class="sxs-lookup"><span data-stu-id="80eed-150">Event handlers must be removed using the same `RequestContext` in which they were added.</span></span> <span data-ttu-id="80eed-151">如果需要加载项在运行时删除事件处理程序，则需要存储用于添加处理程序的 context 对象。</span><span class="sxs-lookup"><span data-stu-id="80eed-151">If you need your add-in to remove an event handler while running, you'll need to store the context object used to add the handler.</span></span>

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

## <a name="supporting-internet-explorer"></a><span data-ttu-id="80eed-152">支持 Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="80eed-152">Supporting Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="80eed-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="80eed-153">See also</span></span>

- <span data-ttu-id="80eed-154">[OfficeDev/？ js](https://github.com/OfficeDev/office-js/issues)：报告和查看 office 外接程序平台和 JavaScript api 中的问题的位置。</span><span class="sxs-lookup"><span data-stu-id="80eed-154">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="80eed-155">[堆栈溢出](https://stackoverflow.com/questions/tagged/office-js)：询问并查看有关 Office JavaScript api 的编程问题的位置。</span><span class="sxs-lookup"><span data-stu-id="80eed-155">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="80eed-156">在发布到堆栈溢出时，请务必对您的问题应用 "office-js" 标记。</span><span class="sxs-lookup"><span data-stu-id="80eed-156">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="80eed-157">[UserVoice](https://officespdev.uservoice.com/)：建议 Office 外接程序平台和 Office JavaScript api 的新功能的位置。</span><span class="sxs-lookup"><span data-stu-id="80eed-157">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
