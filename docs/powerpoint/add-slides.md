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
# <a name="add-and-delete-slides-in-powerpoint-preview"></a>在 PowerPoint (预览版中添加和) 

PowerPoint 加载项可以将幻灯片添加到演示文稿中，并可以选择指定用于新幻灯片的幻灯片母版和母版的版式。 加载项还可以删除幻灯片。

> [!IMPORTANT]
> 用于添加幻灯片的 API 为预览版。 请在开发或测试环境中试验它们，但不要将其添加到生产外接程序。 用于删除 *幻灯片的* API 已发布。

用于添加幻灯片的 API 主要用于在编码时知道演示文稿中幻灯片母版和版式中的标识或在运行时可在数据源中找到的方案中。 在这种情况下，您或客户必须创建和维护一个将选择条件 (（如幻灯片母版和版式) 的名称或图像）与幻灯片母版和版式 ID 关联的数据源。 这些 API 还可用于以下方案：用户可以插入使用默认幻灯片母版和母版的默认版式幻灯片，以及用户可以选择现有幻灯片并创建具有相同幻灯片母版和版式的新幻灯片 (但不使用相同的内容) 。 有关详细信息 [，请参阅](#selecting-which-slide-master-and-layout-to-use) 选择使用哪个幻灯片母版和版式。

## <a name="add-a-slide-with-slidecollectionadd"></a>使用 SlideCollection.add 添加幻灯片

使用 [SlideCollection.add 方法添加](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_) 幻灯片。 下面是一个简单的示例，其中添加了使用演示文稿的默认幻灯片母版和该母版的第一个版式的幻灯片。 该方法始终将新幻灯片添加到演示文稿的末尾。 示例如下：

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="selecting-which-slide-master-and-layout-to-use"></a>选择要使用的幻灯片母版和版式

使用 [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) 参数可控制用于新幻灯片的幻灯片母版以及使用母版中的哪个版式。 示例如下。 关于此代码，请注意以下几点：

- 可以包括对象的任一属性或同时包含这两 `AddSlideOptions` 个属性。
- 如果使用这两个属性，则指定的布局必须属于指定的主控点，否则将引发错误。
- 如果该属性不存在， (其值为空字符串) ，则使用默认幻灯片母版，并且该幻灯片母版的版 `masterId` `layoutId` 式必须为该幻灯片母版。
- 默认幻灯片母版是演示文稿中最后一张幻灯片使用的幻灯片母版。  (在演示文稿中当前没有幻灯片的异常情况下，默认幻灯片母版是演示文稿的第一个幻灯片母版。) 
- 如果该属性不存在， (其值为空字符串) ，则使用由指定的主控母版 `layoutId` `masterId` 的第一个布局。
- 这两个属性都是三种可能形式之一的字符串：***nnnnnnnnnn*#**、 * *#* mmmmmmmmm***、 或 **_nnnnnnnnnn_ #* mmmmmmmmm***，其中 *nnnnnnnnnnnn 是* 主控母版或布局的 ID (通常为 10 个数字) 而 *mmmmmmmmm* 是主控或布局的创建 ID (通常为 6 - 10 个数字) 。 一些示例包括 `2147483690#2908289500` `2147483690#` ， 和 `#2908289500` 。

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

用户无法实际发现幻灯片母版或版式 ID 或创建 ID。 因此，你实际上只能仅在编码时知道这些标识，或者你的外接程序可以在运行时发现 `AddSlideOptions` 它们时，才使用参数。 由于用户无法记住 ID，因此还需要一种方法使用户能够选择幻灯片（可能按名称或图像选择）然后将每个标题或图像与幻灯片 ID 关联。

因此，此参数主要用于外接程序设计为与一组特定的幻灯片母版和布局（其 ID 已知）一 `AddSlideOptions` 起使用的方案。 在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (（如幻灯片母版和版式名称或图像) ）与相应的 ID 或创建 ID 关联。

#### <a name="have-the-user-choose-a-matching-slide"></a>让用户选择匹配的幻灯片

如果新幻灯片应该使用现有幻灯片使用的幻灯片母版和版式的组合，则外接程序可以使用 (1) 提示用户选择幻灯片， (2) 读取幻灯片母版和版式 ID。 以下步骤演示了如何读取这些 ID 并添加具有匹配母版和布局的幻灯片。

1. 创建一个方法来获取选定幻灯片的索引。 示例如下。 关于此代码，请注意以下几点：

    - 它使用Office.context.docJavaScript API 的 [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法。
    - 对的 `getSelectedDataAsync` 调用嵌入 Promise 返回函数中。 有关这样做的原因和如何操作，请参阅承诺返回函数中的[封装通用 API。](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)
    - `getSelectedDataAsync` 返回一个数组，因为可以选择多个幻灯片。 在此方案中，用户仅选择了一个，因此代码获取第一个 (0) 幻灯片，这是唯一一个选定的幻灯片。
    - 幻灯片的值是用户在缩略图窗格中的幻灯片旁边看到的基于 `index` 1 的值。

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

2. 在添加幻灯片的主函数的 [PowerPoint.run () ](/javascript/api/powerpoint#PowerPoint_run_batch_) 调用新函数。 示例如下：

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

## <a name="delete-slides"></a>删除幻灯片

通过获取对代表幻灯片的 [Slide](/javascript/api/powerpoint/powerpoint.slide) 对象的引用来删除幻灯片并调用 `Slide.delete` 该方法。 下面是删除第 4 张幻灯片的示例：

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
