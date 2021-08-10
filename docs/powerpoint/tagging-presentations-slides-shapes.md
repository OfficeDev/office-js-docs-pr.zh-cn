---
title: 对演示文稿、幻灯片和演示文稿中的形状使用自定义PowerPoint
description: 了解如何将标记用于有关演示文稿、幻灯片和形状的自定义元数据。
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: 9ae86906a2ac69cb79adac34fa4e923a9bc218a7dc8a7e5bdefd63300b589da5
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093652"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>对演示文稿、幻灯片和演示文稿中的形状使用自定义PowerPoint

加载项可以将自定义元数据（称为"标记"键值对）附加到幻灯片上的演示文稿、特定幻灯片和特定形状。

> [!IMPORTANT]
> 标记 API 在预览阶段。 请在开发或测试环境中试验它们，但不要将它们添加到生产外接程序。

使用标记有两种主要方案：

- 应用于幻灯片或形状时，标记允许对对象进行分类以便进行批处理。 例如，假设演示文稿包含一些幻灯片，这些幻灯片应包含在向东部区域而不是西地区的演示文稿中。 同样，还有一些备用幻灯片应只向西显示。 您的外接程序可以创建一个包含键和值的标记，并应用于只应在东部使用的 `REGION` `East` 幻灯片。 对于应该只向"西"区域显示的幻灯片，该标记 `West` 的值设置为 。 在向"东部"显示演示文稿之前，加载项中的按钮将运行代码，该代码将循环访问检查标记值的所有 `REGION` 幻灯片。 删除区域 `West` 位置的幻灯片。 然后，用户关闭外接程序并启动幻灯片放映。
- 应用于演示文稿时，标记实际上是演示文稿文档中的自定义 (类似于 Word 文档中的 [CustomProperty](/javascript/api/word/word.customproperty)) 。

## <a name="tag-slides-and-shapes"></a>标记幻灯片和形状

标记是键值对，其中值始终为类型， `string` 由 [Tag](/javascript/api/powerpoint/powerpoint.tag) 对象表示。 每种类型的父对象（如[Presentation、Slide](/javascript/api/powerpoint/powerpoint.presentation)[](/javascript/api/powerpoint/powerpoint.slide)或[Shape](/javascript/api/powerpoint/powerpoint.shape)对象）都有一个 `tags` [类型为 TagsCollection 的属性](/javascript/api/powerpoint/powerpoint.tagcollection)。

### <a name="add-update-and-delete-tags"></a>添加、更新和删除标记

若要向对象添加标记，请调用父对象的属性的 [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) `tags` 方法。 下面的代码将两个标记添加到演示文稿的第一张幻灯片。 关于此代码，请注意以下几点：

- 方法的第一 `add` 个参数是键值对中的键。 
- 第二个参数是值。
- 键为大写字母。 此方法并非严格强制要求;但是，键始终由 PowerPoint 存储为大写，并且某些与标记相关的方法要求键以大写形式表示，因此我们建议始终在代码中对标记键使用大写形式。 `add` 

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

`add`方法还用于更新标记。 以下代码更改标记 `PLANET` 的值。

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

若要删除标记，请对它的父对象调用 方法，将 标记的键作为 参数 `delete` `TagsCollection` 传递。 有关示例，请参阅 [在演示文稿上设置自定义元数据](#set-custom-metadata-on-the-presentation)。

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>使用标记选择性地处理幻灯片和形状

请考虑以下方案：Contoso Consulting 有一个向所有新客户演示的演示文稿。 但某些幻灯片应只向已支付"高级"状态费用的客户显示。 在向非高级客户显示演示文稿之前，他们可以复制演示文稿并删除仅高级客户应该看到的幻灯片。 通过外接程序，Contoso 可以标记适合高级客户的幻灯片并根据需要删除这些幻灯片。 下面的列表概述了创建此功能的主要编码步骤。

1. 创建一个方法，将当前选定的幻灯片标记为适合 `Premium` 客户。 关于此代码，请注意以下几点：

    - `getSelectedSlideIndex`函数在下一步中定义。 它返回当前选定幻灯片的从 1 开始索引。
    - 由于 `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) 方法基于 0，因此函数返回的值必须缩小。

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

2. 下面的代码创建一个方法，用于获取选定幻灯片的索引。 关于此代码，请注意以下几点：

    - 它使用Office.context.docJavaScript API 的 [ ument.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法。
    - 对 的 `getSelectedDataAsync` 调用嵌入承诺返回函数中。 有关这样做的原因和如何操作，请参阅在承诺返回函数中[包装通用 API。](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)
    - `getSelectedDataAsync` 返回一个数组，因为可以选择多个幻灯片。 在此方案中，用户只选择了一个，因此代码获取第一张 (第) 张幻灯片，该幻灯片是唯一一个选定的幻灯片。
    - 幻灯片的值是用户在 UI 缩略图窗格中的幻灯片旁边看到的 `index` PowerPoint 1 的值。

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

3. 以下代码创建一种方法来删除针对高级客户标记的幻灯片。 关于此代码，请注意以下几点：

    - 由于 `key` 标记 `value` 的 和 属性将在 之后读取， `context.sync` 因此必须先加载它们。

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

## <a name="set-custom-metadata-on-the-presentation"></a>在演示文稿上设置自定义元数据

加载项还可以将标记作为一个整体应用于演示文稿。 这样，您能够将标记用于文档级元数据，类似于 [CustomProperty](/javascript/api/word/word.customproperty)类在 Word 中的使用方式。 但与 Word `CustomProperty` 类不同，PowerPoint标记的值只能是 类型 `string` 。

以下代码是向演示文稿添加标记的示例。 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

以下代码是一个从演示文稿中删除标记的示例。 请注意，标记的键将传递给 `delete` 父对象的 `TagsCollection` 方法。

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
