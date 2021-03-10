---
title: 在 PowerPoint 中添加和删除幻灯片
description: 了解如何添加和删除幻灯片以及指定新幻灯片的主控母版和版式。
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 5c1b9750acb905fd8e92484bb960c70ba39a7ca9
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613941"
---
# <a name="add-and-delete-slides-in-powerpoint-preview"></a><span data-ttu-id="38582-103">在 PowerPoint (预览版中添加和) </span><span class="sxs-lookup"><span data-stu-id="38582-103">Add and delete slides in PowerPoint (preview)</span></span>

<span data-ttu-id="38582-104">PowerPoint 加载项可以将幻灯片添加到演示文稿中，并可以选择指定用于新幻灯片的幻灯片母版和母版的版式。</span><span class="sxs-lookup"><span data-stu-id="38582-104">A PowerPoint add-in can add slides to the presentation and optionally specify which slide master, and which layout of the master, is used for the new slide.</span></span> <span data-ttu-id="38582-105">加载项还可以删除幻灯片。</span><span class="sxs-lookup"><span data-stu-id="38582-105">The add-in can also delete slides.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="38582-106">用于添加幻灯片的 API 为预览版。</span><span class="sxs-lookup"><span data-stu-id="38582-106">The APIs for adding slides are in preview.</span></span> <span data-ttu-id="38582-107">请在开发或测试环境中试验它们，但不要将其添加到生产外接程序。</span><span class="sxs-lookup"><span data-stu-id="38582-107">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span> <span data-ttu-id="38582-108">用于删除 *幻灯片的* API 已发布。</span><span class="sxs-lookup"><span data-stu-id="38582-108">The API for *deleting* slides has been released.</span></span>

<span data-ttu-id="38582-109">用于添加幻灯片的 API 主要用于在编码时知道演示文稿中幻灯片母版和版式中的标识或在运行时可在数据源中找到的方案中。</span><span class="sxs-lookup"><span data-stu-id="38582-109">The APIs for adding slides are primarily used in scenarios where the IDs of the slide masters and layouts in the presentation are known at coding time or can be found in a data source at runtime.</span></span> <span data-ttu-id="38582-110">在这种情况下，您或客户必须创建和维护一个将选择条件 (（如幻灯片母版和版式) 的名称或图像）与幻灯片母版和版式 ID 关联的数据源。</span><span class="sxs-lookup"><span data-stu-id="38582-110">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as the names or images of slide masters and layouts) with the IDs of the slide masters and layouts.</span></span> <span data-ttu-id="38582-111">这些 API 还可用于以下方案：用户可以插入使用默认幻灯片母版和母版的默认版式幻灯片，以及用户可以选择现有幻灯片并创建具有相同幻灯片母版和版式的新幻灯片 (但不使用相同的内容) 。</span><span class="sxs-lookup"><span data-stu-id="38582-111">The APIs can also be used in scenarios where the user can insert slides that use the default slide master and the master's default layout, and in scenarios where the user can select an existing slide and create a new one with the same slide master and layout (but not the same content).</span></span> <span data-ttu-id="38582-112">有关详细信息 [，请参阅](#selecting-which-slide-master-and-layout-to-use) 选择使用哪个幻灯片母版和版式。</span><span class="sxs-lookup"><span data-stu-id="38582-112">See [Selecting which slide master and layout to use](#selecting-which-slide-master-and-layout-to-use) for more information about this.</span></span>

## <a name="add-a-slide-with-slidecollectionadd"></a><span data-ttu-id="38582-113">使用 SlideCollection.add 添加幻灯片</span><span class="sxs-lookup"><span data-stu-id="38582-113">Add a slide with SlideCollection.add</span></span>

<span data-ttu-id="38582-114">使用 [SlideCollection.add 方法添加](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) 幻灯片。</span><span class="sxs-lookup"><span data-stu-id="38582-114">Add slides with the [SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) method.</span></span> <span data-ttu-id="38582-115">下面是一个简单的示例，其中添加了使用演示文稿的默认幻灯片母版和该母版的第一个版式的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="38582-115">The following is a simple example in which a slide that uses the presentation's default slide master and the first layout of that master is added.</span></span> <span data-ttu-id="38582-116">该方法始终将新幻灯片添加到演示文稿的末尾。</span><span class="sxs-lookup"><span data-stu-id="38582-116">The method always adds new slides to the end of the presentation.</span></span> <span data-ttu-id="38582-117">示例如下：</span><span class="sxs-lookup"><span data-stu-id="38582-117">The following is an example:</span></span>

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="selecting-which-slide-master-and-layout-to-use"></a><span data-ttu-id="38582-118">选择要使用的幻灯片母版和版式</span><span class="sxs-lookup"><span data-stu-id="38582-118">Selecting which slide master and layout to use</span></span>

<span data-ttu-id="38582-119">使用 [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) 参数可控制用于新幻灯片的幻灯片母版以及使用母版中的哪个版式。</span><span class="sxs-lookup"><span data-stu-id="38582-119">Use the [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) parameter to control which slide master is used for the new slide and which layout within the master is used.</span></span> <span data-ttu-id="38582-120">示例如下。</span><span class="sxs-lookup"><span data-stu-id="38582-120">The following is an example.</span></span> <span data-ttu-id="38582-121">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="38582-121">Note the following about this code:</span></span>

- <span data-ttu-id="38582-122">可以包括对象的任一属性或同时包含这两 `AddSlideOptions` 个属性。</span><span class="sxs-lookup"><span data-stu-id="38582-122">You can include either or both the properties of the `AddSlideOptions` object.</span></span>
- <span data-ttu-id="38582-123">如果使用这两个属性，则指定的布局必须属于指定的主控点，否则将引发错误。</span><span class="sxs-lookup"><span data-stu-id="38582-123">If both properties are used, then the specified layout must belong to the specified master or an error is thrown.</span></span>
- <span data-ttu-id="38582-124">如果该属性不存在， (其值为空字符串) ，则使用默认幻灯片母版，并且该幻灯片母版的版 `masterId` `layoutId` 式必须为该幻灯片母版。</span><span class="sxs-lookup"><span data-stu-id="38582-124">If the `masterId` property isn't present (or its value is an empty string), then the default slide master is used and the `layoutId` must be a layout of that slide master.</span></span>
- <span data-ttu-id="38582-125">默认幻灯片母版是演示文稿中最后一张幻灯片使用的幻灯片母版。</span><span class="sxs-lookup"><span data-stu-id="38582-125">The default slide master is the slide master used by the last slide in the presentation.</span></span> <span data-ttu-id="38582-126"> (在演示文稿中当前没有幻灯片的异常情况下，默认幻灯片母版是演示文稿的第一个幻灯片母版。) </span><span class="sxs-lookup"><span data-stu-id="38582-126">(In the unusual case where there are currently no slides in the presentation, then the default slide master is the first slide master in the presentation.)</span></span>
- <span data-ttu-id="38582-127">如果该属性不存在， (其值为空字符串) ，则使用由指定的主控母版 `layoutId` `masterId` 的第一个布局。</span><span class="sxs-lookup"><span data-stu-id="38582-127">If the `layoutId` property isn't present (or its value is an empty string), then the first layout of the master that is specified by the `masterId` is used.</span></span>
- <span data-ttu-id="38582-128">这两个属性都是三种可能形式之一的字符串：***nnnnnnnnnn\*#\*\*、 \* *#* mmmmmmmmm***、 或 \**_nnnnnnnnnn_ #* mmmmmmmmm\*\*\*，其中 *nnnnnnnnnnnn 是* 主控母版或布局的 ID (通常为 10 个数字) 而 *mmmmmmmmm* 是主控或布局的创建 ID (通常为 6 - 10 个数字) 。</span><span class="sxs-lookup"><span data-stu-id="38582-128">Both properties are strings of one of three possible forms: \***nnnnnnnnnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnnnnnnnnn_#* mmmmmmmmm\*\*\*, where *nnnnnnnnnn* is the master's or layout's ID (typically 10 digits) and *mmmmmmmmm* is the master's or layout's creation ID (typically 6 - 10 digits).</span></span> <span data-ttu-id="38582-129">一些示例包括 `2147483690#2908289500` `2147483690#` ， 和 `#2908289500` 。</span><span class="sxs-lookup"><span data-stu-id="38582-129">Some examples are `2147483690#2908289500`, `2147483690#`, and `#2908289500`.</span></span>

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

<span data-ttu-id="38582-130">用户无法实际发现幻灯片母版或版式 ID 或创建 ID。</span><span class="sxs-lookup"><span data-stu-id="38582-130">There is no practical way that users can discover the ID or creation ID of a slide master or layout.</span></span> <span data-ttu-id="38582-131">因此，你实际上只能仅在编码时知道这些标识，或者你的外接程序可以在运行时发现 `AddSlideOptions` 它们时，才使用参数。</span><span class="sxs-lookup"><span data-stu-id="38582-131">For this reason, you can really only use the `AddSlideOptions` parameter when either you know the IDs at coding time or your add-in can discover them at runtime.</span></span> <span data-ttu-id="38582-132">由于用户无法记住 ID，因此还需要一种方法使用户能够选择幻灯片（可能按名称或图像选择）然后将每个标题或图像与幻灯片 ID 关联。</span><span class="sxs-lookup"><span data-stu-id="38582-132">Because users can't be expected to memorize the IDs, you also need a way to enable the user to select slides, perhaps by name or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="38582-133">因此，此参数主要用于外接程序设计为与一组特定的幻灯片母版和布局（其 ID 已知）一 `AddSlideOptions` 起使用的方案。</span><span class="sxs-lookup"><span data-stu-id="38582-133">Accordingly, the `AddSlideOptions` parameter is primarily used in scenarios in which the add-in is designed to work with a specific set of slide masters and layouts whose IDs are known.</span></span> <span data-ttu-id="38582-134">在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (（如幻灯片母版和版式名称或图像) ）与相应的 ID 或创建 ID 关联。</span><span class="sxs-lookup"><span data-stu-id="38582-134">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as slide master and layout names or images) with the corresponding IDs or creation IDs.</span></span>

#### <a name="have-the-user-choose-a-matching-slide"></a><span data-ttu-id="38582-135">让用户选择匹配的幻灯片</span><span class="sxs-lookup"><span data-stu-id="38582-135">Have the user choose a matching slide</span></span>

<span data-ttu-id="38582-136">如果新幻灯片应该使用现有幻灯片使用的幻灯片母版和版式的组合，则外接程序可以使用 (1) 提示用户选择幻灯片， (2) 读取幻灯片母版和版式 ID。</span><span class="sxs-lookup"><span data-stu-id="38582-136">If your add-in can be used in scenarios where the new slide should use the same combination of slide master and layout that is used by an *existing* slide, then your add-in can (1) prompt the user to select a slide and (2) read the IDs of the slide master and layout.</span></span> <span data-ttu-id="38582-137">以下步骤演示了如何读取这些 ID 并添加具有匹配母版和布局的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="38582-137">The following steps show how to read the IDs and add a slide with a matching master and layout.</span></span>

1. <span data-ttu-id="38582-138">创建一个方法来获取选定幻灯片的索引。</span><span class="sxs-lookup"><span data-stu-id="38582-138">Create a method to get the index of the selected slide.</span></span> <span data-ttu-id="38582-139">示例如下。</span><span class="sxs-lookup"><span data-stu-id="38582-139">The following is an example.</span></span> <span data-ttu-id="38582-140">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="38582-140">Note about this code:</span></span>

    - <span data-ttu-id="38582-141">它使用Office.context.docJavaScript API 的 [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法。</span><span class="sxs-lookup"><span data-stu-id="38582-141">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="38582-142">对的 `getSelectedDataAsync` 调用嵌入 Promise 返回函数中。</span><span class="sxs-lookup"><span data-stu-id="38582-142">The call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="38582-143">有关这样做的原因和如何操作，请参阅承诺返回函数中的[封装通用 API。](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="38582-143">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="38582-144">`getSelectedDataAsync` 返回一个数组，因为可以选择多个幻灯片。</span><span class="sxs-lookup"><span data-stu-id="38582-144">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="38582-145">在此方案中，用户仅选择了一个，因此代码获取第一个 (0) 幻灯片，这是唯一一个选定的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="38582-145">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="38582-146">幻灯片的值是用户在缩略图窗格中的幻灯片旁边看到的基于 `index` 1 的值。</span><span class="sxs-lookup"><span data-stu-id="38582-146">The `index` value of the slide is the 1-based value the user sees beside the slide in the thumbnails pane.</span></span>

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

2. <span data-ttu-id="38582-147">在添加幻灯片的主函数的 [PowerPoint.run () ](/javascript/api/powerpoint#PowerPoint_run_batch_) 调用新函数。</span><span class="sxs-lookup"><span data-stu-id="38582-147">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function that adds the slide.</span></span> <span data-ttu-id="38582-148">示例如下：</span><span class="sxs-lookup"><span data-stu-id="38582-148">The following is an example:</span></span>

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

## <a name="delete-slides"></a><span data-ttu-id="38582-149">删除幻灯片</span><span class="sxs-lookup"><span data-stu-id="38582-149">Delete slides</span></span>

<span data-ttu-id="38582-150">通过获取对代表幻灯片的 [Slide](/javascript/api/powerpoint/powerpoint.slide) 对象的引用来删除幻灯片并调用 `Slide.delete` 该方法。</span><span class="sxs-lookup"><span data-stu-id="38582-150">Delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="38582-151">下面是删除第 4 张幻灯片的示例：</span><span class="sxs-lookup"><span data-stu-id="38582-151">The following is an example in which the 4th slide is deleted:</span></span>

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
