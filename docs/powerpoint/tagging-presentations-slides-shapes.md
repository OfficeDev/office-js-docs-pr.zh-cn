---
title: 在 PowerPoint 中的演示文稿、幻灯片和形状上使用自定义标记
description: 了解如何对演示文稿、幻灯片和形状的自定义元数据使用标记。
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: a30beea56286437b1c69461534ca13912107cecf
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958900"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>在 PowerPoint 中对演示文稿、幻灯片和形状使用自定义标记

外接程序可以以键值对的形式将自定义元数据（称为“标记”）附加到幻灯片上的演示文稿、特定幻灯片和特定形状。

使用标记的主要方案有两种：

- 应用于幻灯片或形状时，标记允许对对象进行批处理分类。 例如，假设演示文稿包含一些幻灯片，这些幻灯片应包含在东部区域的演示文稿中，但不应包含在西部区域。 同样，还有一些其他幻灯片应该只向西方显示。 外接程序可以创建包含键 `REGION` 和值 `East` 的标记，并将其应用到应仅在东部使用的幻灯片。 标记的值设置 `West` 为只应向西部区域显示的幻灯片。 在向东部演示文稿之前，外接程序中的按钮运行代码，该代码遍历检查标记值 `REGION` 的所有幻灯片。 删除区域 `West` 的幻灯片。 然后，用户关闭加载项并开始幻灯片放映。
- 应用于演示文稿时，标记实际上是演示文档中的自定义属性 (类似于 Word) 中的 [CustomProperty](/javascript/api/word/word.customproperty) 。

## <a name="tag-slides-and-shapes"></a>标记幻灯片和形状

标记是键值对，其中值始终为类型 `string` ，由 [Tag](/javascript/api/powerpoint/powerpoint.tag) 对象表示。 每种类型的父对象（例如 [演示文稿](/javascript/api/powerpoint/powerpoint.presentation)、 [幻灯片](/javascript/api/powerpoint/powerpoint.slide)或 [Shape](/javascript/api/powerpoint/powerpoint.shape) 对象）都有一个 `tags` 类型 [为 TagsCollection 的](/javascript/api/powerpoint/powerpoint.tagcollection)属性。

### <a name="add-update-and-delete-tags"></a>添加、更新和删除标记

若要向对象添加标记，请调用父对象属性的 `tags` [TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) 方法。 以下代码将两个标记添加到演示文稿的第一张幻灯片中。 关于此代码，请注意以下几点：

- 该方法的第一个 `add` 参数是键值对中的键。
- 第二个参数是值。
- 键以大写字母表示。 此方法不是严格必需 `add` 的;但是，PowerPoint 始终以大写形式存储密钥， *并且某些与标记相关的方法确实要求用大写形式表示密钥*，因此我们建议在代码中始终使用大写标记键作为最佳做法。

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

该 `add` 方法还用于更新标记。 以下代码更改标记的 `PLANET` 值。

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

若要删除标记，请调用 `delete` 其父 `TagsCollection` 对象上的方法，并将标记的键作为参数传递。 有关示例，请参阅 [演示文稿上的“设置自定义元数据](#set-custom-metadata-on-the-presentation)”。

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>使用标记有选择地处理幻灯片和形状

请考虑以下方案：Contoso 咨询具有向所有新客户展示的演示文稿。 但某些幻灯片应仅向已支付“高级”状态的客户显示。 在向非高级客户显示演示文稿之前，他们会制作一份该演示文稿的副本，并删除只有高级客户才应该看到的幻灯片。 加载项使 Contoso 能够标记哪些幻灯片适用于高级客户，并在需要时删除这些幻灯片。 以下列表概述了创建此功能的主要编码步骤。

1. 创建一个函数，根据客户 `Premium` 的预期标记当前选定的幻灯片。 关于此代码，请注意以下几点：

    - 该 `getSelectedSlideIndex` 函数在下一步中定义。 它返回当前所选幻灯片的基于 1 的索引。
    - 函数返回的 `getSelectedSlideIndex` 值必须递减，因为 [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) 方法基于 0。

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

2. 以下代码创建一个方法来获取所选幻灯片的索引。 关于此代码，请注意以下几点：

    - 它使用 Common JavaScript API 的 [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) 方法。
    - 调 `getSelectedDataAsync` 用嵌入到承诺返回函数中。 有关为什么以及如何执行此操作的详细信息，请参阅 [承诺返回函数中的包装通用 API](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。
    - `getSelectedDataAsync` 返回一个数组，因为可以选择多个幻灯片。 在此方案中，用户只选择了一张，因此代码将获取第一张 (第 0 张) 幻灯片，这是唯一选择的幻灯片。
    - 幻 `index` 灯片的值是用户在 PowerPoint UI 缩略图窗格中的幻灯片旁边看到的基于 1 的值。

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

3. 以下代码创建一个函数来删除为高级客户标记的幻灯片。 关于此代码，请注意以下几点：

    - `key`由于标记和`value`属性将在标记之后`context.sync`读取，因此必须先加载它们。

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

加载项还可以将标记作为一个整体应用到演示文稿。 这使你能够对文档级元数据使用标记，类似于 Word 中使用 [CustomProperty](/javascript/api/word/word.customproperty)类的方式。 但与 Word `CustomProperty` 类不同，PowerPoint 标记的值只能是类型 `string`。

下面的代码是向演示文稿添加标记的示例。 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

下面的代码是从演示文稿中删除标记的示例。 请注意，标记的键将传递给 `delete` 父 `TagsCollection` 对象的方法。

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
