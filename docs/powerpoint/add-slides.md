---
title: 在幻灯片中添加和删除PowerPoint
description: 了解如何添加和删除幻灯片，并指定新幻灯片的主控母版和版式。
ms.date: 06/02/2021
localization_priority: Normal
ms.openlocfilehash: fd1f3c805483050776cc5b71c9e7a9fb61610b07
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348410"
---
# <a name="add-and-delete-slides-in-powerpoint"></a><span data-ttu-id="cbde5-103">在幻灯片中添加和删除PowerPoint</span><span class="sxs-lookup"><span data-stu-id="cbde5-103">Add and delete slides in PowerPoint</span></span>

<span data-ttu-id="cbde5-104">加载项PowerPoint向演示文稿添加幻灯片，并可以选择指定新幻灯片使用哪个幻灯片母版和哪个母版版式。</span><span class="sxs-lookup"><span data-stu-id="cbde5-104">A PowerPoint add-in can add slides to the presentation and optionally specify which slide master, and which layout of the master, is used for the new slide.</span></span> <span data-ttu-id="cbde5-105">加载项还可以删除幻灯片。</span><span class="sxs-lookup"><span data-stu-id="cbde5-105">The add-in can also delete slides.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="cbde5-106">用于添加幻灯片的 API 为预览 [版](../reference/requirement-sets/powerpoint-preview-apis.md) ，不适用于生产加载项。用于删除 *幻灯片的* API 已发布。</span><span class="sxs-lookup"><span data-stu-id="cbde5-106">The APIs for adding slides are in [preview](../reference/requirement-sets/powerpoint-preview-apis.md) and not available for production add-ins. The API for *deleting* slides has been released.</span></span>

<span data-ttu-id="cbde5-107">添加幻灯片的 API 主要用于以下方案：演示文稿中幻灯片母版和版式的标识在编码时已知，或在运行时可在数据源中找到。</span><span class="sxs-lookup"><span data-stu-id="cbde5-107">The APIs for adding slides are primarily used in scenarios where the IDs of the slide masters and layouts in the presentation are known at coding time or can be found in a data source at runtime.</span></span> <span data-ttu-id="cbde5-108">在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (如幻灯片母版和版式的名称或图像与幻灯片母版和版式) 的 ID 相关联。</span><span class="sxs-lookup"><span data-stu-id="cbde5-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as the names or images of slide masters and layouts) with the IDs of the slide masters and layouts.</span></span> <span data-ttu-id="cbde5-109">这些 API 还可用于以下方案：用户可以插入使用默认幻灯片母版和母版的默认版式的幻灯片，以及用户可以选择现有幻灯片并使用同一幻灯片母版和版式创建新幻灯片 (但内容不相同) 。</span><span class="sxs-lookup"><span data-stu-id="cbde5-109">The APIs can also be used in scenarios where the user can insert slides that use the default slide master and the master's default layout, and in scenarios where the user can select an existing slide and create a new one with the same slide master and layout (but not the same content).</span></span> <span data-ttu-id="cbde5-110">有关详细信息 [，](#selecting-which-slide-master-and-layout-to-use) 请参阅选择使用哪个幻灯片母版和版式。</span><span class="sxs-lookup"><span data-stu-id="cbde5-110">See [Selecting which slide master and layout to use](#selecting-which-slide-master-and-layout-to-use) for more information about this.</span></span>

## <a name="add-a-slide-with-slidecollectionadd-preview"></a><span data-ttu-id="cbde5-111">使用 SlideCollection.add 添加幻灯片 (预览) </span><span class="sxs-lookup"><span data-stu-id="cbde5-111">Add a slide with SlideCollection.add (preview)</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

<span data-ttu-id="cbde5-112">使用 [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) 方法添加幻灯片。</span><span class="sxs-lookup"><span data-stu-id="cbde5-112">Add slides with the [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) method.</span></span> <span data-ttu-id="cbde5-113">下面是一个简单的示例，其中添加了使用演示文稿的默认幻灯片母版和该母版的第一个版式的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="cbde5-113">The following is a simple example in which a slide that uses the presentation's default slide master and the first layout of that master is added.</span></span> <span data-ttu-id="cbde5-114">方法始终将新幻灯片添加到演示文稿的末尾。</span><span class="sxs-lookup"><span data-stu-id="cbde5-114">The method always adds new slides to the end of the presentation.</span></span> <span data-ttu-id="cbde5-115">示例如下。</span><span class="sxs-lookup"><span data-stu-id="cbde5-115">The following is an example.</span></span>

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="selecting-which-slide-master-and-layout-to-use"></a><span data-ttu-id="cbde5-116">选择要使用的幻灯片母版和版式</span><span class="sxs-lookup"><span data-stu-id="cbde5-116">Selecting which slide master and layout to use</span></span>

<span data-ttu-id="cbde5-117">使用 [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) 参数可控制新幻灯片使用哪个幻灯片母版以及使用母版中的哪个版式。</span><span class="sxs-lookup"><span data-stu-id="cbde5-117">Use the [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) parameter to control which slide master is used for the new slide and which layout within the master is used.</span></span> <span data-ttu-id="cbde5-118">示例如下。</span><span class="sxs-lookup"><span data-stu-id="cbde5-118">The following is an example.</span></span> <span data-ttu-id="cbde5-119">对于此代码，请注意以下事项。</span><span class="sxs-lookup"><span data-stu-id="cbde5-119">Note the following about this code.</span></span>

- <span data-ttu-id="cbde5-120">可以包括 对象的一个或两个 `AddSlideOptions` 属性。</span><span class="sxs-lookup"><span data-stu-id="cbde5-120">You can include either or both the properties of the `AddSlideOptions` object.</span></span>
- <span data-ttu-id="cbde5-121">如果同时使用这两个属性，则指定的布局必须属于指定的母版，否则将引发错误。</span><span class="sxs-lookup"><span data-stu-id="cbde5-121">If both properties are used, then the specified layout must belong to the specified master or an error is thrown.</span></span>
- <span data-ttu-id="cbde5-122">如果属性不存在 (或者其值为空字符串) ，则使用默认幻灯片母版，并且 必须是该幻灯片母版 `masterId` `layoutId` 的版式。</span><span class="sxs-lookup"><span data-stu-id="cbde5-122">If the `masterId` property isn't present (or its value is an empty string), then the default slide master is used and the `layoutId` must be a layout of that slide master.</span></span>
- <span data-ttu-id="cbde5-123">默认幻灯片母版是演示文稿中最后一张幻灯片使用的幻灯片母版。</span><span class="sxs-lookup"><span data-stu-id="cbde5-123">The default slide master is the slide master used by the last slide in the presentation.</span></span> <span data-ttu-id="cbde5-124"> (在演示文稿中当前没有幻灯片的异常情况下，默认幻灯片母版是演示文稿的第一个幻灯片母版。) </span><span class="sxs-lookup"><span data-stu-id="cbde5-124">(In the unusual case where there are currently no slides in the presentation, then the default slide master is the first slide master in the presentation.)</span></span>
- <span data-ttu-id="cbde5-125">如果属性不存在 (或者其值为空字符串) ，则使用 指定的主 `layoutId` 控母版 `masterId` 的第一个布局。</span><span class="sxs-lookup"><span data-stu-id="cbde5-125">If the `layoutId` property isn't present (or its value is an empty string), then the first layout of the master that is specified by the `masterId` is used.</span></span>
- <span data-ttu-id="cbde5-126">这两个属性都是三种可能形式之一的字符串：\***nnnnnnnnnn\*#**、\* *#* mmmmmmmmm\*\*\*、 或 \**_nnnnnnnnnn_ #* mmmmmmmmm\*\*\*，其中 *nnnnnnnnnn* 是主控位置或布局的 ID (通常为 10 个数字) *而 mmmmmmmmm* 是主控母版或布局的创建 ID (通常为 6 - 10 个数字) 。</span><span class="sxs-lookup"><span data-stu-id="cbde5-126">Both properties are strings of one of three possible forms: \***nnnnnnnnnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnnnnnnnnn_#* mmmmmmmmm\*\*\*, where *nnnnnnnnnn* is the master's or layout's ID (typically 10 digits) and *mmmmmmmmm* is the master's or layout's creation ID (typically 6 - 10 digits).</span></span> <span data-ttu-id="cbde5-127">一些示例包括 `2147483690#2908289500` 、 `2147483690#` 和 `#2908289500` 。</span><span class="sxs-lookup"><span data-stu-id="cbde5-127">Some examples are `2147483690#2908289500`, `2147483690#`, and `#2908289500`.</span></span>

```javascript
async function addSlide() {
    await PowerPoint.run(async function(context) {
        context.presentation.slides.add({
            slideMasterId: "2147483690#2908289500",
            layoutId: "2147483691#2499880"
        });
    
        await context.sync();
    });
}
```

<span data-ttu-id="cbde5-128">用户无法找到幻灯片母版或版式 ID 或创建 ID。</span><span class="sxs-lookup"><span data-stu-id="cbde5-128">There is no practical way that users can discover the ID or creation ID of a slide master or layout.</span></span> <span data-ttu-id="cbde5-129">因此，实际上，只有当在编码时知道这些标识或加载项可以在运行时发现这些标识时，才能 `AddSlideOptions` 真正使用 参数。</span><span class="sxs-lookup"><span data-stu-id="cbde5-129">For this reason, you can really only use the `AddSlideOptions` parameter when either you know the IDs at coding time or your add-in can discover them at runtime.</span></span> <span data-ttu-id="cbde5-130">因为无法预期用户记住 ID，所以还需要一种方法让用户按名称或图像选择幻灯片，然后将每个标题或图像与幻灯片 ID 关联。</span><span class="sxs-lookup"><span data-stu-id="cbde5-130">Because users can't be expected to memorize the IDs, you also need a way to enable the user to select slides, perhaps by name or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="cbde5-131">因此，参数主要用于外接程序设计为使用一组特定的幻灯片母版和布局（其 ID 已知） `AddSlideOptions` 的方案。</span><span class="sxs-lookup"><span data-stu-id="cbde5-131">Accordingly, the `AddSlideOptions` parameter is primarily used in scenarios in which the add-in is designed to work with a specific set of slide masters and layouts whose IDs are known.</span></span> <span data-ttu-id="cbde5-132">在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (例如幻灯片母版和版式名称或图像) 与相应的 ID 或创建 ID 关联。</span><span class="sxs-lookup"><span data-stu-id="cbde5-132">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as slide master and layout names or images) with the corresponding IDs or creation IDs.</span></span>

#### <a name="have-the-user-choose-a-matching-slide"></a><span data-ttu-id="cbde5-133">让用户选择匹配的幻灯片</span><span class="sxs-lookup"><span data-stu-id="cbde5-133">Have the user choose a matching slide</span></span>

<span data-ttu-id="cbde5-134">如果外接程序可用于新幻灯片应使用现有幻灯片使用的幻灯片母版和版式的组合的方案，则外接程序可以 (1) 提示用户选择幻灯片， (2) 读取幻灯片母版和版式 ID。</span><span class="sxs-lookup"><span data-stu-id="cbde5-134">If your add-in can be used in scenarios where the new slide should use the same combination of slide master and layout that is used by an *existing* slide, then your add-in can (1) prompt the user to select a slide and (2) read the IDs of the slide master and layout.</span></span> <span data-ttu-id="cbde5-135">以下步骤演示了如何读取这些 ID 并添加具有匹配母版和布局的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="cbde5-135">The following steps show how to read the IDs and add a slide with a matching master and layout.</span></span>

1. <span data-ttu-id="cbde5-136">创建一个方法，获取选定幻灯片的索引。</span><span class="sxs-lookup"><span data-stu-id="cbde5-136">Create a method to get the index of the selected slide.</span></span> <span data-ttu-id="cbde5-137">示例如下。</span><span class="sxs-lookup"><span data-stu-id="cbde5-137">The following is an example.</span></span> <span data-ttu-id="cbde5-138">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="cbde5-138">Note about this code:</span></span>

    - <span data-ttu-id="cbde5-139">它使用Office.context.docJavaScript API 的 [ ument.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法。</span><span class="sxs-lookup"><span data-stu-id="cbde5-139">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="cbde5-140">对 的 `getSelectedDataAsync` 调用嵌入 Promise 返回函数中。</span><span class="sxs-lookup"><span data-stu-id="cbde5-140">The call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="cbde5-141">有关这样做的原因和如何操作，请参阅在承诺返回函数中[包装通用 API。](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="cbde5-141">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="cbde5-142">`getSelectedDataAsync` 返回一个数组，因为可以选择多个幻灯片。</span><span class="sxs-lookup"><span data-stu-id="cbde5-142">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="cbde5-143">在此方案中，用户只选择了一个，因此代码获取第一张 (第) 张幻灯片，该幻灯片是唯一一个选定的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="cbde5-143">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="cbde5-144">幻灯片的值是用户在缩略图窗格中的幻灯片旁边看到的从 `index` 1 开始的值。</span><span class="sxs-lookup"><span data-stu-id="cbde5-144">The `index` value of the slide is the 1-based value the user sees beside the slide in the thumbnails pane.</span></span>

    ```javascript
    function getSelectedSlideIndex() {
        return new OfficeExtension.Promise<number>(function(resolve, reject) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
                try {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(console.error(asyncResult.error.message));
                    } else {
                        resolve(asyncResult.value.slides[0].index);
                    }
                } 
                catch (error) {
                    reject(console.log(error));
                }
            });
        });
    }
    ```

2. <span data-ttu-id="cbde5-145">在添加幻灯片的主函数的[PowerPoint.run () ](/javascript/api/powerpoint#PowerPoint_run_batch_)中调用新函数。</span><span class="sxs-lookup"><span data-stu-id="cbde5-145">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function that adds the slide.</span></span> <span data-ttu-id="cbde5-146">示例如下。</span><span class="sxs-lookup"><span data-stu-id="cbde5-146">The following is an example.</span></span>

    ```javascript
    async function addSlideWithMatchingLayout() {
        await PowerPoint.run(async function(context) {
    
            let selectedSlideIndex = await getSelectedSlideIndex();
        
            // Decrement the index because the value returned by getSelectedSlideIndex()
            // is 1-based, but SlideCollection.getItemAt() is 0-based.
            const realSlideIndex = selectedSlideIndex - 1;
            const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster/id, layout/id");
        
            await context.sync();
        
            context.presentation.slides.add({
                slideMasterId: selectedSlide.slideMaster.id,
                layoutId: selectedSlide.layout.id
            });
        
            await context.sync();
        });
    }
    ```

## <a name="delete-slides"></a><span data-ttu-id="cbde5-147">删除幻灯片</span><span class="sxs-lookup"><span data-stu-id="cbde5-147">Delete slides</span></span>

<span data-ttu-id="cbde5-148">通过获取对代表幻灯片的 [Slide](/javascript/api/powerpoint/powerpoint.slide) 对象的引用来删除幻灯片并调用 `Slide.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="cbde5-148">Delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="cbde5-149">下面是删除第 4 张幻灯片的示例。</span><span class="sxs-lookup"><span data-stu-id="cbde5-149">The following is an example in which the 4th slide is deleted.</span></span>

```javascript
async function deleteSlide() {
    await PowerPoint.run(async function(context) {

        // The slide index is zero-based. 
        const slide = context.presentation.slides.getItemAt(3);
        slide.delete();

        await context.sync();
    });
}
```
