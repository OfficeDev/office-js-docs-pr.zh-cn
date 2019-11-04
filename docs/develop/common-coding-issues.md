---
title: 常见的编码问题和意外的平台行为
description: 开发人员经常遇到的 Office JavaScript API 平台问题的列表。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d39c379961833cdb924628becf2c2da3f7e271b9
ms.sourcegitcommit: 59d29d01bce7543ebebf86e5a86db00cf54ca14a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/01/2019
ms.locfileid: "37924792"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="f8994-103">常见的编码问题和意外的平台行为</span><span class="sxs-lookup"><span data-stu-id="f8994-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="f8994-104">本文重点介绍了 Office JavaScript API 的各个方面，这些方面可能导致意外行为或需要特定编码模式来实现所需的结果。</span><span class="sxs-lookup"><span data-stu-id="f8994-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="f8994-105">如果遇到此列表中的问题，请使用文章底部的反馈表单告知我们。</span><span class="sxs-lookup"><span data-stu-id="f8994-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="common-apis-and-outlook-apis-are-not-promise-based"></a><span data-ttu-id="f8994-106">通用 Api 和 Outlook Api 不基于承诺</span><span class="sxs-lookup"><span data-stu-id="f8994-106">Common APIs and Outlook APIs are not promise-based</span></span>

<span data-ttu-id="f8994-107">[通用 api](/javascript/api/office) （那些未绑定到特定 Office 主机的 api）和[Outlook api](/javascript/api/outlook)使用基于回调的编程模型。</span><span class="sxs-lookup"><span data-stu-id="f8994-107">The [Common APIs](/javascript/api/office) (those that are not tied to a particular Office host) and [Outlook APIs](/javascript/api/outlook) use a callback-based programming model.</span></span> <span data-ttu-id="f8994-108">与基础 Office 文档进行交互需要进行异步读取或写入调用，以指定在操作完成时要运行的回调。</span><span class="sxs-lookup"><span data-stu-id="f8994-108">Interacting with the underlying Office document requires an asynchronous read or write call that specifies a callback to be ran when the operation completes.</span></span> <span data-ttu-id="f8994-109">有关此模式的示例，请参阅[document.getfileasync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-)。</span><span class="sxs-lookup"><span data-stu-id="f8994-109">For an example of this pattern, see [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-).</span></span>

<span data-ttu-id="f8994-110">这些常见 API 和 Outlook API 方法不会返回[承诺](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)。</span><span class="sxs-lookup"><span data-stu-id="f8994-110">These Common API and Outlook API methods do not return [Promises](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise).</span></span> <span data-ttu-id="f8994-111">因此，在异步操作完成之前，不能使用[await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await)暂停执行。</span><span class="sxs-lookup"><span data-stu-id="f8994-111">Therefore, you cannot use [await](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Operators/await) to pause the execution until the asynchronous operation completes.</span></span> <span data-ttu-id="f8994-112">如果需要`await`行为，可以在显式创建的承诺中包装方法调用。</span><span class="sxs-lookup"><span data-stu-id="f8994-112">If you need `await` behavior, you can wrap the method call in an explicitly created Promise.</span></span>

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
> <span data-ttu-id="f8994-113">参考文档包含[getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-)的承诺包装实现。</span><span class="sxs-lookup"><span data-stu-id="f8994-113">The reference documentation contains the Promise-wrapped implementation of [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-).</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="f8994-114">某些属性必须使用 JSON 结构进行设置</span><span class="sxs-lookup"><span data-stu-id="f8994-114">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="f8994-115">本部分仅适用于 Excel 和 Word 的特定于主机的 Api。</span><span class="sxs-lookup"><span data-stu-id="f8994-115">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="f8994-116">某些属性必须设置为 JSON 结构，而不是设置其单个子属性。</span><span class="sxs-lookup"><span data-stu-id="f8994-116">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="f8994-117">在[页面布局](/javascript/api/excel/excel.pagelayout)中找到此示例的一个示例。</span><span class="sxs-lookup"><span data-stu-id="f8994-117">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="f8994-118">必须`zoom`使用单个[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)对象设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="f8994-118">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="f8994-119">在上面的示例中，您***将无法***直接分配`zoom`值： `sheet.pageLayout.zoom.scale = 200;`。</span><span class="sxs-lookup"><span data-stu-id="f8994-119">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="f8994-120">由于`zoom`未加载，该语句会引发错误。</span><span class="sxs-lookup"><span data-stu-id="f8994-120">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="f8994-121">`zoom`即使要加载，该扩展集也不会生效。</span><span class="sxs-lookup"><span data-stu-id="f8994-121">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="f8994-122">发生所有上下文操作`zoom`，刷新加载项中的代理对象并覆盖本地设置的值。</span><span class="sxs-lookup"><span data-stu-id="f8994-122">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="f8994-123">此行为不同于[导航属性](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)，如[Range. 格式](/javascript/api/excel/excel.range#format)。</span><span class="sxs-lookup"><span data-stu-id="f8994-123">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="f8994-124">`format`可以使用对象导航设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="f8994-124">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="f8994-125">您可以通过检查其只读修饰符来标识必须将其子属性设置为 JSON 结构的属性。</span><span class="sxs-lookup"><span data-stu-id="f8994-125">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="f8994-126">所有只读属性都可以直接设置其非只读的子属性。</span><span class="sxs-lookup"><span data-stu-id="f8994-126">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="f8994-127">必须使用 JSON `PageLayout.zoom`结构设置可写属性（如必须设置）。</span><span class="sxs-lookup"><span data-stu-id="f8994-127">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="f8994-128">摘要：</span><span class="sxs-lookup"><span data-stu-id="f8994-128">In summary:</span></span>

- <span data-ttu-id="f8994-129">只读属性：可通过导航设置子属性。</span><span class="sxs-lookup"><span data-stu-id="f8994-129">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="f8994-130">可写属性：必须使用 JSON 结构设置子属性（且不能通过导航进行设置）。</span><span class="sxs-lookup"><span data-stu-id="f8994-130">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="excel-range-limits"></a><span data-ttu-id="f8994-131">Excel 区域限制</span><span class="sxs-lookup"><span data-stu-id="f8994-131">Excel Range limits</span></span>

<span data-ttu-id="f8994-132">如果您正在构建使用区域的 Excel 加载项，请注意以下大小限制：</span><span class="sxs-lookup"><span data-stu-id="f8994-132">If you're building an Excel add-in that uses ranges, be aware of the following size limitations:</span></span>

- <span data-ttu-id="f8994-133">Excel 网页版将请求和响应的有效负载大小限制为 5MB。</span><span class="sxs-lookup"><span data-stu-id="f8994-133">Excel on the web has a payload size limit for requests and responses of 5MB.</span></span> <span data-ttu-id="f8994-134">如果超过该限制，将引发 `RichAPI.Error`。</span><span class="sxs-lookup"><span data-stu-id="f8994-134">`RichAPI.Error` will be thrown if that limit is exceeded.</span></span>
- <span data-ttu-id="f8994-135">对于 set 操作，范围限制为5000000个单元格。</span><span class="sxs-lookup"><span data-stu-id="f8994-135">A range is limited to five million cells for set operations.</span></span>

<span data-ttu-id="f8994-136">如果您希望用户输入超出这些限制，请务必检查数据并将区域拆分为多个对象。</span><span class="sxs-lookup"><span data-stu-id="f8994-136">If you expect user input to exceed these limits, be sure to check the data and split the ranges into multiple objects.</span></span> <span data-ttu-id="f8994-137">您还需要提交多个`context.sync()`呼叫，以避免将较小的范围操作再次成批组合在一起。</span><span class="sxs-lookup"><span data-stu-id="f8994-137">You'll also need to submit multiple `context.sync()` calls to avoid the smaller range operations getting batched together again.</span></span>

<span data-ttu-id="f8994-138">您的外接程序可能能够使用[RangeAreas](/javascript/api/excel/excel.rangeareas)对较大范围内的单元格进行战略更新。</span><span class="sxs-lookup"><span data-stu-id="f8994-138">Your add-in might be able to use [RangeAreas](/javascript/api/excel/excel.rangeareas) to strategically update cells within a larger range.</span></span> <span data-ttu-id="f8994-139">有关详细信息，请参阅[在 Excel 外接程序中同时处理多个区域](../excel/excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="f8994-139">See [Work with multiple ranges simultaneously in Excel add-ins](../excel/excel-add-ins-multiple-ranges.md) for more information.</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="f8994-140">设置只读属性</span><span class="sxs-lookup"><span data-stu-id="f8994-140">Setting read-only properties</span></span>

<span data-ttu-id="f8994-141">Office JS 的[TypeScript 定义](/referencing-the-javascript-api-for-office-library-from-its-cdn.md)指定哪些对象属性是只读的。</span><span class="sxs-lookup"><span data-stu-id="f8994-141">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="f8994-142">如果尝试设置只读属性，写入操作将无提示地失败，且不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="f8994-142">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="f8994-143">下面的示例错误地尝试设置只读属性[Chart.id](/javascript/api/excel/excel.chart#id)。</span><span class="sxs-lookup"><span data-stu-id="f8994-143">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="f8994-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f8994-144">See also</span></span>

- <span data-ttu-id="f8994-145">[OfficeDev/？ js](https://github.com/OfficeDev/office-js/issues)：报告和查看 office 外接程序平台和 JavaScript api 中的问题的位置。</span><span class="sxs-lookup"><span data-stu-id="f8994-145">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="f8994-146">[堆栈溢出](https://stackoverflow.com/questions/tagged/office-js)：询问并查看有关 Office JavaScript api 的编程问题的位置。</span><span class="sxs-lookup"><span data-stu-id="f8994-146">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="f8994-147">在发布到堆栈溢出时，请务必对您的问题应用 "office-js" 标记。</span><span class="sxs-lookup"><span data-stu-id="f8994-147">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="f8994-148">[UserVoice](https://officespdev.uservoice.com/)：建议 Office 外接程序平台和 Office JavaScript api 的新功能的位置。</span><span class="sxs-lookup"><span data-stu-id="f8994-148">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
