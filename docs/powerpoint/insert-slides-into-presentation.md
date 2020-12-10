---
title: 在 PowerPoint 演示文稿中插入和删除幻灯片
description: 了解如何将一个演示文稿中的幻灯片插入到另一个演示文稿中，以及如何删除幻灯片。
ms.date: 12/04/2020
localization_priority: Normal
ms.openlocfilehash: ceb78054a95ac4b26bd71f79a086a00e3dce5278
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/09/2020
ms.locfileid: "49613699"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a>在 PowerPoint 演示文稿中插入和删除幻灯片 (预览) 

PowerPoint 加载项可以使用 PowerPoint 的应用程序特定的 JavaScript 库，将幻灯片从一个演示文稿插入到当前演示文稿中。 您可以控制插入的幻灯片是保留源演示文稿的格式还是目标演示文稿的格式。 您还可以从演示文稿中删除幻灯片。

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

幻灯片插入 Api 主要在演示文稿模板方案中使用：有少量的已知演示文稿充当可由加载项插入的幻灯片池。 在这种情况下，您或客户必须创建并维护一个与选择条件关联的数据源 (如幻灯片标题或与幻灯片 Id) 的图像。 在用户可以插入任意演示文稿中的幻灯片时，也可以使用 Api，但在这种情况下，用户有效地限制为从源演示文稿插入 *所有* 幻灯片。 有关详细信息，请参阅 [选择要插入的幻灯片](#selecting-which-slides-to-insert) 。

将一个演示文稿中的幻灯片插入到另一个演示文稿中有两个步骤。

1. 将源演示文稿文件 ( .pptx) 转换为 base64 格式的字符串。
1. 使用 `insertSlidesFromBase64` 方法可将 base64 文件中的一个或多个幻灯片插入到当前演示文稿中。

## <a name="convert-the-source-presentation-to-base64"></a>将源演示文稿转换为 base64

有多种方法可将文件转换为 base64。 您使用哪种编程语言和库，以及是否在外接程序的服务器端或客户端进行转换取决于您的方案。 最常见的情况是，通过使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 对象在客户端上使用 JavaScript 进行转换。 下面的示例演示了这种做法。

1. 首先获取对源 PowerPoint 文件的引用。 在此示例中，我们将使用 `<input>` 类型的控件 `file` 来提示用户选择文件。 将以下标记添加到加载项页面。

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    此标记将以下屏幕截图中的 UI 添加到页面：

    ![显示 HTML 文件类型输入控件的屏幕截图，其前面有一个说明语句 "选择要从中插入幻灯片的 PowerPoint 演示文稿"。 该控件由标记为 "选择文件" 的按钮，后跟 "未选择文件" 句子。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > 有许多其他方法可以获取 PowerPoint 文件。 例如，如果文件存储在 OneDrive 或 SharePoint 上，则可以使用 Microsoft Graph 下载它。 有关详细信息，请参阅使用 [Microsoft graph 中的文件](/graph/api/resources/onedrive) 和 [使用 Microsoft graph 访问文件](/learn/modules/msgraph-access-file-data/)。

2. 将以下代码添加到加载项的 JavaScript 中，以将函数分配给输入控件的 `change` 事件。  (您 `storeFileAsBase64` 在下一步中创建函数。 ) 

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. 添加以下代码。 请注意有关此代码的以下内容：

    - 该 `reader.readAsDataURL` 方法将文件转换为 base64 并将其存储在 `reader.result` 属性中。 方法完成后，它将触发 `onload` 事件处理程序。
    - `onload`事件处理程序从编码的文件中去除元数据，并将编码后的字符串存储在全局变量中。
    - Base64 编码的字符串在全局范围内存储，因为它将被在后续步骤中创建的其他函数读取。

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## <a name="insert-slides-with-insertslidesfrombase64"></a>插入带 insertSlidesFromBase64 的幻灯片

您的外接程序使用 [insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) 方法将另一个 PowerPoint 演示文稿中的幻灯片插入到当前演示文稿中。 下面是一个简单的示例，其中源演示文稿中的所有幻灯片都插入到当前演示文稿的开头，并且插入的幻灯片保留源文件的格式。 请注意，它 `chosenFileBase64` 是一个包含 base64 编码版本的 PowerPoint 演示文稿文件的全局变量。

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

通过将 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 对象作为第二个参数传递，可以控制插入结果的某些方面，包括插入幻灯片的位置以及它们是否获取源或目标的格式 `insertSlidesFromBase64` 。 示例如下。 关于此代码，请注意以下几点：

- 属性有两个可能的值 `formatting` ： "UseDestinationTheme" 和 "KeepSourceFormatting"。 （可选）您可以使用 `InsertSlideFormatting` enum， (例如， `PowerPoint.InsertSlideFormatting.useDestinationTheme`) 。
- 函数将在由属性指定的幻灯片之后立即在源演示文稿中插入幻灯片 `targetSlideId` 。 此属性的值是包含以下三种格式之一的字符串： ***nnn * #**、* *#* mmmmmmmmm * * * 或 **_nnn_ #* mmmmmmmmm * * *，其中 *nnn* 是幻灯片的 id (通常为3个数字) 并且 *mmmmmmmmm* 是幻灯片的创建 id (通常) 9 个数字。 例如、 `267#763315295` `267#` 和 `#763315295` 。

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

当然，您通常不会在编码时知道目标幻灯片的 ID 或创建 ID。 更常见的情况是，外接程序会要求用户选择目标幻灯片。 以下步骤显示了如何获取当前选定幻灯片的 ***nnn * #** ID 并将其用作目标幻灯片。

1. 使用通用 JavaScript Api 的 [Office.context.docUment getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法创建一个函数，该函数可获取当前选定幻灯片的 ID。 示例如下。 请注意，调用 `getSelectedDataAsync` 被嵌入到承诺返回的函数中。 有关为什么以及如何执行此操作的详细信息，请参阅 [在承诺返回函数中换行 Common-APIs](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. 在 PowerPoint 中调用新函数 [。运行 main 函数的 ( # B1 ](/javascript/api/powerpoint#PowerPoint_run_batch_) ，并将其返回的 ID (传递它返回的 ID。) 作为参数的属性值与 "#" 符号连接 `targetSlideId` `InsertSlideOptions` 。 示例如下。

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### <a name="selecting-which-slides-to-insert"></a>选择要插入的幻灯片

您还可以使用 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 参数来控制源演示文稿中插入的幻灯片。 为此，可通过将源演示文稿的幻灯片 Id 的数组分配给属性来执行此操作 `sourceSlideIds` 。 下面是插入四张幻灯片的示例。 请注意，数组中的每个字符串必须遵循用于该属性的一个或另一个模式 `targetSlideId` 。

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> 幻灯片将按照其在源演示文稿中出现的相对顺序进行插入，而不考虑它们在数组中的显示顺序。

用户无法在源演示文稿中发现幻灯片的 ID 或创建 ID，这是一种切实可行的方法。 因此，仅 `sourceSlideIds` 当您知道编码时的源 id 或加载项可以在运行时从某些数据源检索这些 id 时，才能真正使用属性。 由于无法预期用户能够记住幻灯片 Id，因此还需要一种方法来使用户能够选择幻灯片（如标题或图像），然后将每个标题或图像与幻灯片的 ID 关联起来。

因此，该 `sourceSlideIds` 属性主要用于演示文稿模板方案：外接程序设计为使用一组特定的演示文稿，用作可插入的幻灯片池。 在这种情况下，您或客户必须创建并维护一个与选择条件关联的数据源 (如标题或图像) 与已通过一组可能的源演示文稿构造的幻灯片 Id 或幻灯片创建 Id。

## <a name="delete-slides"></a>删除幻灯片

通过获取对表示幻灯片的 [slide](/javascript/api/powerpoint/powerpoint.slide) 对象的引用并调用方法，可以删除幻灯片 `Slide.delete` 。 下面是一个示例，其中第四张幻灯片被删除。

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
