---
title: 在 PowerPoint 中对演示文稿、幻灯片和形状使用自定义标记
description: 了解如何将标记用于有关演示文稿、幻灯片和形状的自定义元数据。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: fbb13e67da1f7962fc2c0b8d45689f259b015014
ms.sourcegitcommit: 58d394fa49308ecf93cd53f7d3fb6e316ff56209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/16/2021
ms.locfileid: "51876856"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a><span data-ttu-id="c8058-103">在 PowerPoint 中对演示文稿、幻灯片和形状使用自定义标记</span><span class="sxs-lookup"><span data-stu-id="c8058-103">Use custom tags for presentations, slides, and shapes in PowerPoint</span></span>

<span data-ttu-id="c8058-104">加载项可以将自定义元数据（称为"标记"键值对）附加到幻灯片上的演示文稿、特定幻灯片和特定形状。</span><span class="sxs-lookup"><span data-stu-id="c8058-104">An add-in can attach custom metadata, in the form of key-value pairs, called "tags", to presentations, specific slides, and specific shapes on a slide.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c8058-105">标记 API 在预览阶段。</span><span class="sxs-lookup"><span data-stu-id="c8058-105">The APIs for tags are in preview.</span></span> <span data-ttu-id="c8058-106">请在开发或测试环境中试验它们，但不要将它们添加到生产外接程序。</span><span class="sxs-lookup"><span data-stu-id="c8058-106">Please experiment with them in a development or testing environment but don't add them to a production add-in.</span></span>

<span data-ttu-id="c8058-107">使用标记有两种主要方案：</span><span class="sxs-lookup"><span data-stu-id="c8058-107">There are two main scenarios for using tags:</span></span>

- <span data-ttu-id="c8058-108">应用于幻灯片或形状时，标记允许对对象进行分类以便进行批处理。</span><span class="sxs-lookup"><span data-stu-id="c8058-108">When applied to a slide or a shape, a tag enables the object to be categorized for batch processing.</span></span> <span data-ttu-id="c8058-109">例如，假设演示文稿包含一些幻灯片，这些幻灯片应包含在向东部区域而不是西地区的演示文稿中。</span><span class="sxs-lookup"><span data-stu-id="c8058-109">For example, suppose a presentation has some slides that should be included in presentations to the East region but not the West region.</span></span> <span data-ttu-id="c8058-110">同样，还有一些备用幻灯片应只向西显示。</span><span class="sxs-lookup"><span data-stu-id="c8058-110">Similarly, there are alternative slides that should be shown only to the West.</span></span> <span data-ttu-id="c8058-111">您的外接程序可以创建一个包含键和值的标记，并应用于只应在东部使用的 `REGION` `East` 幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-111">Your add-in can create a tag with the key `REGION` and the value `East` and apply it to the slides that should only be used in the East.</span></span> <span data-ttu-id="c8058-112">对于应该只向"西"区域显示的幻灯片，该标记 `West` 的值设置为 。</span><span class="sxs-lookup"><span data-stu-id="c8058-112">The tag's value is set to `West` for the slides that should only be shown to the West region.</span></span> <span data-ttu-id="c8058-113">在向"东部"显示演示文稿之前，加载项中的按钮将运行代码，该代码将循环访问检查标记值的所有 `REGION` 幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-113">Just before a presentation to the East, a button in the add-in runs code that loops through all the slides checking the value of the `REGION` tag.</span></span> <span data-ttu-id="c8058-114">删除区域 `West` 位置的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-114">Slides where the region is `West` are deleted.</span></span> <span data-ttu-id="c8058-115">然后，用户关闭外接程序并启动幻灯片放映。</span><span class="sxs-lookup"><span data-stu-id="c8058-115">The user then closes the add-in and starts the slide show.</span></span>
- <span data-ttu-id="c8058-116">应用于演示文稿时，标记实际上是演示文稿文档中的自定义 (类似于 Word 文档中的 [CustomProperty](/javascript/api/word/word.customproperty)) 。</span><span class="sxs-lookup"><span data-stu-id="c8058-116">When applied to a presentation, a tag is effectively a custom property in the presentation document (similar to a [CustomProperty](/javascript/api/word/word.customproperty) in Word).</span></span>

## <a name="tag-slides-and-shapes"></a><span data-ttu-id="c8058-117">标记幻灯片和形状</span><span class="sxs-lookup"><span data-stu-id="c8058-117">Tag slides and shapes</span></span>

<span data-ttu-id="c8058-118">标记是键值对，其中值始终为类型， `string` 由 [Tag](/javascript/api/powerpoint/powerpoint.tag) 对象表示。</span><span class="sxs-lookup"><span data-stu-id="c8058-118">A tag is a key-value pair, where the value is always of type `string` and is represented by a [Tag](/javascript/api/powerpoint/powerpoint.tag) object.</span></span> <span data-ttu-id="c8058-119">每种类型的父对象（如[Presentation、Slide](/javascript/api/powerpoint/powerpoint.presentation)[](/javascript/api/powerpoint/powerpoint.slide)或[Shape](/javascript/api/powerpoint/powerpoint.shape)对象）都有一个 `tags` [类型为 TagsCollection 的属性](/javascript/api/powerpoint/powerpoint.tagcollection)。</span><span class="sxs-lookup"><span data-stu-id="c8058-119">Each type of parent object, such as a [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide), or [Shape](/javascript/api/powerpoint/powerpoint.shape) object, has a `tags` property of type [TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).</span></span>

### <a name="add-update-and-delete-tags"></a><span data-ttu-id="c8058-120">添加、更新和删除标记</span><span class="sxs-lookup"><span data-stu-id="c8058-120">Add, update, and delete tags</span></span>

<span data-ttu-id="c8058-121">若要向对象添加标记，请调用父对象的属性的 [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) `tags` 方法。</span><span class="sxs-lookup"><span data-stu-id="c8058-121">To add a tag to an object, call the [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) method of the parent object's `tags` property.</span></span> <span data-ttu-id="c8058-122">下面的代码将两个标记添加到演示文稿的第一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-122">The following code adds two tags to the first slide of a presentation.</span></span> <span data-ttu-id="c8058-123">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="c8058-123">About this code, note:</span></span>

- <span data-ttu-id="c8058-124">方法的第一 `add` 个参数是键值对中的键。</span><span class="sxs-lookup"><span data-stu-id="c8058-124">The first parameter of the `add` method is the key in the key-value pair.</span></span> 
- <span data-ttu-id="c8058-125">第二个参数是值。</span><span class="sxs-lookup"><span data-stu-id="c8058-125">The second parameter is the value.</span></span>
- <span data-ttu-id="c8058-126">键为大写字母。</span><span class="sxs-lookup"><span data-stu-id="c8058-126">The key is in uppercase letters.</span></span> <span data-ttu-id="c8058-127">此方法并非严格强制要求;但是 `add` ，PowerPoint 始终将键存储为大写，并且某些与标记相关的方法要求键使用大写形式表示，因此建议最佳做法是始终在代码中对标记键使用大写形式。</span><span class="sxs-lookup"><span data-stu-id="c8058-127">This isn't strictly mandatory for the `add` method; however, the key is always stored by PowerPoint as uppercase, and *some tag-related methods do require that the key be expressed in uppercase*, so we recommend as a best practice that you always use uppercase in your code for a tag key.</span></span>

```javascript
async function addMultipleSlideTags() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}
```

<span data-ttu-id="c8058-128">`add`方法还用于更新标记。</span><span class="sxs-lookup"><span data-stu-id="c8058-128">The `add` method is also used to update a tag.</span></span> <span data-ttu-id="c8058-129">以下代码更改标记 `PLANET` 的值。</span><span class="sxs-lookup"><span data-stu-id="c8058-129">The following code changes the value of the `PLANET` tag.</span></span>

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

<span data-ttu-id="c8058-130">若要删除标记，请对它的父对象调用 方法，将 标记的键作为 参数 `delete` `TagsCollection` 传递。</span><span class="sxs-lookup"><span data-stu-id="c8058-130">To delete a tag, call the `delete` method on it's parent `TagsCollection` object and pass the key of the tag as the parameter.</span></span> <span data-ttu-id="c8058-131">有关示例，请参阅 [在演示文稿上设置自定义元数据](#set-custom-metadata-on-the-presentation)。</span><span class="sxs-lookup"><span data-stu-id="c8058-131">For an example, see [Set custom metadata on the presentation](#set-custom-metadata-on-the-presentation).</span></span>

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a><span data-ttu-id="c8058-132">使用标记选择性地处理幻灯片和形状</span><span class="sxs-lookup"><span data-stu-id="c8058-132">Use tags to selectively process slides and shapes</span></span>

<span data-ttu-id="c8058-133">请考虑以下方案：Contoso Consulting 有一个向所有新客户演示的演示文稿。</span><span class="sxs-lookup"><span data-stu-id="c8058-133">Consider the following scenario: Contoso Consulting has a presentation they show to all new customers.</span></span> <span data-ttu-id="c8058-134">但某些幻灯片应只向已支付"高级"状态费用的客户显示。</span><span class="sxs-lookup"><span data-stu-id="c8058-134">But some slides should only be shown to customers that have paid for "premium" status.</span></span> <span data-ttu-id="c8058-135">在向非高级客户显示演示文稿之前，他们可以复制演示文稿并删除仅高级客户应该看到的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-135">Before showing the presentation to non-premium customers, they make a copy of it and delete the slides that only premium customers should see.</span></span> <span data-ttu-id="c8058-136">通过外接程序，Contoso 可以标记适合高级客户的幻灯片并根据需要删除这些幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-136">An add-in enables Contoso to tag which slides are for premium customers and to delete these slides when needed.</span></span> <span data-ttu-id="c8058-137">下面的列表概述了创建此功能的主要编码步骤。</span><span class="sxs-lookup"><span data-stu-id="c8058-137">The following list outlines the major coding steps to create this functionality.</span></span>

1. <span data-ttu-id="c8058-138">创建一个方法，将当前选定的幻灯片标记为适合 `Premium` 客户。</span><span class="sxs-lookup"><span data-stu-id="c8058-138">Create a method that tags the currently selected slide as intended for `Premium` customers.</span></span> <span data-ttu-id="c8058-139">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="c8058-139">About this code, note:</span></span>

    - <span data-ttu-id="c8058-140">`getSelectedSlideIndex`函数在下一步中定义。</span><span class="sxs-lookup"><span data-stu-id="c8058-140">The `getSelectedSlideIndex` function is defined in the next step.</span></span> <span data-ttu-id="c8058-141">它返回当前选定幻灯片的从 1 开始索引。</span><span class="sxs-lookup"><span data-stu-id="c8058-141">It returns the 1-based index of the currently selected slide.</span></span>
    - <span data-ttu-id="c8058-142">由于 `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) 方法基于 0，因此函数返回的值必须缩小。</span><span class="sxs-lookup"><span data-stu-id="c8058-142">The value returned by the `getSelectedSlideIndex` function has to be decremented because the [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) method is 0-based.</span></span>

    ```javascript
    async function addTagToSelectedSlide() {
      await PowerPoint.run(async function(context) {
        let selectedSlideIndex = await getSelectedSlideIndex();
        selectedSlideIndex = selectedSlideIndex - 1;
        const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
        slide.tags.add("CUSTOMER_TYPE", "Premium");
    
        await context.sync();
      });
    }
    ```

2. <span data-ttu-id="c8058-143">下面的代码创建一个方法，用于获取选定幻灯片的索引。</span><span class="sxs-lookup"><span data-stu-id="c8058-143">The following code creates a method to get the index of the selected slide.</span></span> <span data-ttu-id="c8058-144">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="c8058-144">About this code, note:</span></span>

    - <span data-ttu-id="c8058-145">它使用Office.context.docJavaScript API 的 [ ument.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法。</span><span class="sxs-lookup"><span data-stu-id="c8058-145">It uses the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span>
    - <span data-ttu-id="c8058-146">对 的 `getSelectedDataAsync` 调用嵌入承诺返回函数中。</span><span class="sxs-lookup"><span data-stu-id="c8058-146">The call to `getSelectedDataAsync` is embedded in a promise-returning function.</span></span> <span data-ttu-id="c8058-147">有关这样做的原因和如何操作，请参阅在承诺返回函数中[包装通用 API。](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)</span><span class="sxs-lookup"><span data-stu-id="c8058-147">For more information about why and how to do this, see [Wrap Common APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>
    - <span data-ttu-id="c8058-148">`getSelectedDataAsync` 返回一个数组，因为可以选择多个幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-148">`getSelectedDataAsync` returns an array because multiple slides can be selected.</span></span> <span data-ttu-id="c8058-149">在此方案中，用户只选择了一个，因此代码获取第一张 (第) 张幻灯片，该幻灯片是唯一一个选定的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-149">In this scenario, the user has selected just one, so the code gets the first (0th) slide, which is the only one selected.</span></span>
    - <span data-ttu-id="c8058-150">幻灯片的值是用户在 PowerPoint UI 缩略图窗格中的幻灯片旁边看到的从 `index` 1 开始的值。</span><span class="sxs-lookup"><span data-stu-id="c8058-150">The `index` value of the slide is the 1-based value the user sees beside the slide in the PowerPoint UI thumbnails pane.</span></span>

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

3. <span data-ttu-id="c8058-151">以下代码创建一种方法来删除针对高级客户标记的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="c8058-151">The following code creates a method to delete slides that are tagged for premium customers.</span></span> <span data-ttu-id="c8058-152">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="c8058-152">About this code, note:</span></span>

    - <span data-ttu-id="c8058-153">由于 `key` 标记 `value` 的 和 属性将在 之后读取， `context.sync` 因此必须先加载它们。</span><span class="sxs-lookup"><span data-stu-id="c8058-153">Because the `key` and `value` properties of the tags are going to be read after the `context.sync`, they must be loaded first.</span></span>

    ```javascript
    async function deleteSlidesByAudience() {
      await PowerPoint.run(async function(context) {
        const slides = context.presentation.slides;
        slides.load("tags/key, tags/value");
    
        await context.sync();
    
        for (let i = 0; i < slides.items.length; i++) {
          let currentSlide = slides.items[i];
          for (let j = 0; j < currentSlide.tags.items.length; j++) {
            let currentTag = currentSlide.tags.items[j];
            if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
              currentSlide.delete();
            }
          }
        }
    
        await context.sync();
      });
    }
    ```

## <a name="set-custom-metadata-on-the-presentation"></a><span data-ttu-id="c8058-154">在演示文稿上设置自定义元数据</span><span class="sxs-lookup"><span data-stu-id="c8058-154">Set custom metadata on the presentation</span></span>

<span data-ttu-id="c8058-155">加载项还可以将标记作为一个整体应用于演示文稿。</span><span class="sxs-lookup"><span data-stu-id="c8058-155">Add-ins can also apply tags to the presentation as a whole.</span></span> <span data-ttu-id="c8058-156">这样，您能够将标记用于文档级元数据，类似于 [CustomProperty](/javascript/api/word/word.customproperty)类在 Word 中的使用方式。</span><span class="sxs-lookup"><span data-stu-id="c8058-156">This enables you to use tags for document-level metadata similar to how the [CustomProperty](/javascript/api/word/word.customproperty)class is used in Word.</span></span> <span data-ttu-id="c8058-157">但与 Word `CustomProperty` 类不同，PowerPoint 标记的值只能是类型 `string` 。</span><span class="sxs-lookup"><span data-stu-id="c8058-157">But unlike the Word `CustomProperty` class, the value of a PowerPoint tag can only be of type `string`.</span></span>

<span data-ttu-id="c8058-158">以下代码是向演示文稿添加标记的示例。</span><span class="sxs-lookup"><span data-stu-id="c8058-158">The following code is an example of adding a tag to a presentation.</span></span> 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

<span data-ttu-id="c8058-159">以下代码是一个从演示文稿中删除标记的示例。</span><span class="sxs-lookup"><span data-stu-id="c8058-159">The following code is an example of deleting a tag from a presentation.</span></span> <span data-ttu-id="c8058-160">请注意，标记的键将传递给 `delete` 父对象的 `TagsCollection` 方法。</span><span class="sxs-lookup"><span data-stu-id="c8058-160">Note that the key of the tag is passed to the `delete` method of the parent `TagsCollection` object.</span></span>

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
