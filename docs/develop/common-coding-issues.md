---
title: 常见的编码问题和意外的平台行为
description: 开发人员经常遇到的 Office JavaScript API 平台问题的列表。
ms.date: 10/29/2019
localization_priority: Normal
ms.openlocfilehash: 8cea95e3214585ba8e0b77535916f9c564dde9df
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902136"
---
# <a name="common-coding-issues-and-unexpected-platform-behaviors"></a><span data-ttu-id="b4864-103">常见的编码问题和意外的平台行为</span><span class="sxs-lookup"><span data-stu-id="b4864-103">Common coding issues and unexpected platform behaviors</span></span>

<span data-ttu-id="b4864-104">本文重点介绍了 Office JavaScript API 的各个方面，这些方面可能导致意外行为或需要特定编码模式来实现所需的结果。</span><span class="sxs-lookup"><span data-stu-id="b4864-104">This article highlights aspects of the Office JavaScript API that may result in unexpected behavior or require specific coding patterns to achieve the desired outcome.</span></span> <span data-ttu-id="b4864-105">如果遇到此列表中的问题，请使用文章底部的反馈表单告知我们。</span><span class="sxs-lookup"><span data-stu-id="b4864-105">If you encounter an issue that belongs in this list, please let us know by using the feedback form at the bottom of the article.</span></span>

## <a name="some-properties-must-be-set-with-json-structs"></a><span data-ttu-id="b4864-106">某些属性必须使用 JSON 结构进行设置</span><span class="sxs-lookup"><span data-stu-id="b4864-106">Some properties must be set with JSON structs</span></span>

> [!NOTE]
> <span data-ttu-id="b4864-107">本部分仅适用于 Excel 和 Word 的特定于主机的 Api。</span><span class="sxs-lookup"><span data-stu-id="b4864-107">This section only applies to the host-specific APIs for Excel and Word.</span></span>

<span data-ttu-id="b4864-108">某些属性必须设置为 JSON 结构，而不是设置其单个子属性。</span><span class="sxs-lookup"><span data-stu-id="b4864-108">Some properties must be set as JSON structs, instead of setting their individual subproperties.</span></span> <span data-ttu-id="b4864-109">在[页面布局](/javascript/api/excel/excel.pagelayout)中找到此示例的一个示例。</span><span class="sxs-lookup"><span data-stu-id="b4864-109">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="b4864-110">必须`zoom`使用单个[PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions)对象设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="b4864-110">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom must be set with JSON struct representing the PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="b4864-111">在上面的示例中，您***将无法***直接分配`zoom`值： `sheet.pageLayout.zoom.scale = 200;`。</span><span class="sxs-lookup"><span data-stu-id="b4864-111">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="b4864-112">由于`zoom`未加载，该语句会引发错误。</span><span class="sxs-lookup"><span data-stu-id="b4864-112">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="b4864-113">`zoom`即使要加载，该扩展集也不会生效。</span><span class="sxs-lookup"><span data-stu-id="b4864-113">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="b4864-114">发生所有上下文操作`zoom`，刷新加载项中的代理对象并覆盖本地设置的值。</span><span class="sxs-lookup"><span data-stu-id="b4864-114">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="b4864-115">此行为不同于[导航属性](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties)，如[Range. 格式](/javascript/api/excel/excel.range#format)。</span><span class="sxs-lookup"><span data-stu-id="b4864-115">This behavior differs from [navigational properties](../excel/excel-add-ins-advanced-concepts.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="b4864-116">`format`可以使用对象导航设置属性，如下所示：</span><span class="sxs-lookup"><span data-stu-id="b4864-116">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="b4864-117">您可以通过检查其只读修饰符来标识必须将其子属性设置为 JSON 结构的属性。</span><span class="sxs-lookup"><span data-stu-id="b4864-117">You can identify a property that must have its subproperties set with a JSON struct by checking its read-only modifier.</span></span> <span data-ttu-id="b4864-118">所有只读属性都可以直接设置其非只读的子属性。</span><span class="sxs-lookup"><span data-stu-id="b4864-118">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="b4864-119">必须使用 JSON `PageLayout.zoom`结构设置可写属性（如必须设置）。</span><span class="sxs-lookup"><span data-stu-id="b4864-119">Writeable properties like `PageLayout.zoom` must be set with a JSON struct.</span></span> <span data-ttu-id="b4864-120">摘要：</span><span class="sxs-lookup"><span data-stu-id="b4864-120">In summary:</span></span>

- <span data-ttu-id="b4864-121">只读属性：可通过导航设置子属性。</span><span class="sxs-lookup"><span data-stu-id="b4864-121">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="b4864-122">可写属性：必须使用 JSON 结构设置子属性（且不能通过导航进行设置）。</span><span class="sxs-lookup"><span data-stu-id="b4864-122">Writable property: Subproperties must be set with a JSON struct (and cannot be set through navigation).</span></span>

## <a name="setting-read-only-properties"></a><span data-ttu-id="b4864-123">设置只读属性</span><span class="sxs-lookup"><span data-stu-id="b4864-123">Setting read-only properties</span></span>

<span data-ttu-id="b4864-124">Office JS 的[TypeScript 定义](/referencing-the-javascript-api-for-office-library-from-its-cdn.md)指定哪些对象属性是只读的。</span><span class="sxs-lookup"><span data-stu-id="b4864-124">The [TypeScript definitions](/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="b4864-125">如果尝试设置只读属性，写入操作将无提示地失败，且不会引发错误。</span><span class="sxs-lookup"><span data-stu-id="b4864-125">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="b4864-126">下面的示例错误地尝试设置只读属性[Chart.id](/javascript/api/excel/excel.chart#id)。</span><span class="sxs-lookup"><span data-stu-id="b4864-126">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="see-also"></a><span data-ttu-id="b4864-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="b4864-127">See also</span></span>

- <span data-ttu-id="b4864-128">[OfficeDev/？ js](https://github.com/OfficeDev/office-js/issues)：报告和查看 office 外接程序平台和 JavaScript api 中的问题的位置。</span><span class="sxs-lookup"><span data-stu-id="b4864-128">[OfficeDev/office-js](https://github.com/OfficeDev/office-js/issues): The place to report and view issues with the Office Add-ins platform and JavaScript APIs.</span></span>
- <span data-ttu-id="b4864-129">[堆栈溢出](https://stackoverflow.com/questions/tagged/office-js)：询问并查看有关 Office JavaScript api 的编程问题的位置。</span><span class="sxs-lookup"><span data-stu-id="b4864-129">[Stack Overflow](https://stackoverflow.com/questions/tagged/office-js): The place to ask and view programming questions about the Office JavaScript APIs.</span></span> <span data-ttu-id="b4864-130">在发布到堆栈溢出时，请务必对您的问题应用 "office-js" 标记。</span><span class="sxs-lookup"><span data-stu-id="b4864-130">Be sure to apply the "office-js" tag to your question when posting to Stack Overflow.</span></span>
- <span data-ttu-id="b4864-131">[UserVoice](https://officespdev.uservoice.com/)：建议 Office 外接程序平台和 Office JavaScript api 的新功能的位置。</span><span class="sxs-lookup"><span data-stu-id="b4864-131">[UserVoice](https://officespdev.uservoice.com/): The place to suggest new features for the Office Add-ins platform and Office JavaScript APIs.</span></span>
