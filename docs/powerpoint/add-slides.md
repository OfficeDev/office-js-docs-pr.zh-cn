---
title: 在 PowerPoint 中添加和删除幻灯片
description: 了解如何添加和删除幻灯片，以及如何指定新幻灯片的母版和版式。
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cf22c18cf4089bab9091be3f4274f67974662a3
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958311"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>在 PowerPoint 中添加和删除幻灯片

PowerPoint 外接程序可以将幻灯片添加到演示文稿中，并可以选择性地指定用于新幻灯片的幻灯片母版和母版版式。 外接程序还可以删除幻灯片。

用于添加幻灯片的 API 主要用于演示文稿中幻灯片母版和版式的 ID 在编码时已知或可以在运行时的数据源中找到的情况。 在这种情况下，您或客户必须创建和维护与选择条件相关联的数据源 (，例如幻灯片母版的名称或图像，以及) 幻灯片母版和版式 ID 的数据源。 在用户可以插入使用默认幻灯片母版和母版默认版式的幻灯片的情况下，以及在用户可以选择现有幻灯片并创建具有相同幻灯片母版和版式的新幻灯片 (但内容) 不同的情况下，也可以使用 API。 有关详细信息，请参阅 [选择要使用的幻灯片母版和版式](#select-which-slide-master-and-layout-to-use) 。

## <a name="add-a-slide-with-slidecollectionadd"></a>使用 SlideCollection.add 添加幻灯片

使用 [SlideCollection.add 方法添加](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) 幻灯片。 下面是一个简单的示例，其中添加了使用演示文稿的默认幻灯片母版和第一个母版版式的幻灯片。 该方法始终将新幻灯片添加到演示文稿的末尾。 示例如下。

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>选择要使用的幻灯片母版和版式

使用 [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) 参数控制用于新幻灯片的幻灯片母版以及主控形状中的哪个布局。 示例如下。 关于此代码，请注意以下几点：

- 可以包括或同时包含对象的 `AddSlideOptions` 属性。
- 如果使用这两个属性，则指定的布局必须属于指定的主控形状或引发错误。
- `masterId`如果该属性不存在 (或其值为空字符串) ，则使用默认幻灯片母版，`layoutId`并且必须是该幻灯片母版的布局。
- 默认幻灯片母版是演示文稿中最后一张幻灯片使用的幻灯片母版。  (在演示文稿中当前没有幻灯片的异常情况下，默认幻灯片母版是演示文稿中的第一个幻灯片母版。) 
- `layoutId`如果属性不存在 (或其值为空字符串) ，则使用主`masterId`控形状的第一个布局。
- 这两个属性都是三种可能形式之一的字符串：***nnnnnnnnnn*#**， **#* mmmmmmmmm***，或 **_nnnnnnnnnnnn_#* mmmmmmmmm***，其中 *nnnnnnnnnnnn 是* 主控形状或布局的 ID (通常为 10 位数) 和 *mmmmmmmmm* 是主控形状或布局的创建 ID (通常为 6 - 10 位数字) 。 一些示例包括 `2147483690#2908289500`， `2147483690#`和 `#2908289500`.

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

用户无法发现幻灯片母版或版式的 ID 或创建 ID。 因此，实际上只能在知道编码时的 ID 或加载项可以在运行时发现这些 ID 时使用 `AddSlideOptions` 该参数。 由于不能期望用户记住 ID，因此你还需要一种方法来使用户能够选择幻灯片（可能按名称或图像），然后将每个标题或图像与幻灯片的 ID 相关联。

因此， `AddSlideOptions` 该参数主要用于外接程序旨在处理一组特定的幻灯片母版和 ID 已知的布局。 在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (（如幻灯片母版和版式名称或图像) 与相应的 ID 或创建 ID 相关联）。

#### <a name="have-the-user-choose-a-matching-slide"></a>让用户选择匹配的幻灯片

如果外接程序可用于新幻灯片应使用 *现有* 幻灯片所使用的幻灯片母版和版式的相同组合，则加载项可以 (1) 提示用户选择幻灯片， (2) 读取幻灯片母版和版式的 ID。 以下步骤演示如何读取 ID 并添加具有匹配母版和布局的幻灯片。

1. 创建一个函数以获取所选幻灯片的索引。 示例如下。 关于此代码，请注意以下几点：

    - 它使用 Common JavaScript API 的 [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) 方法。
    - 调 `getSelectedDataAsync` 用嵌入到 Promise 返回函数中。 有关为什么以及如何执行此操作的详细信息，请参阅 [承诺返回函数中的包装通用 API](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。
    - `getSelectedDataAsync` 返回一个数组，因为可以选择多个幻灯片。 在此方案中，用户只选择了一张，因此代码将获取第一张 (第 0 张) 幻灯片，这是唯一选择的幻灯片。
    - 幻 `index` 灯片的值是用户在缩略图窗格的幻灯片旁边看到的基于 1 的值。

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

2. 在添加幻灯片的主函数的 [PowerPoint.run () ](/javascript/api/powerpoint#PowerPoint_run_batch_) 中调用新函数。 示例如下。

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

通过获取对幻 [灯片](/javascript/api/powerpoint/powerpoint.slide) 对象的引用来删除幻灯片，该对象代表幻灯片并调用该 `Slide.delete` 方法。 下面是删除第 4 张幻灯片的示例。

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
