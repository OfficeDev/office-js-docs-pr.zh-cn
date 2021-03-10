---
title: 在 PowerPoint 演示文稿中插入幻灯片
description: 了解如何将幻灯片从一个演示文稿插入另一个演示文稿。
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 810a398c336c6715cac138840ed8524cff6c0dac
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613911"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>在 PowerPoint 演示文稿中插入幻灯片

PowerPoint 加载项可以使用 PowerPoint 应用程序特定的 JavaScript 库将演示文稿中的幻灯片插入到当前演示文稿中。 您可以控制插入的幻灯片是否保留源演示文稿的格式或目标演示文稿的格式。

幻灯片插入 API 主要用于演示文稿模板方案：有少量已知演示文稿充当加载项可以插入的幻灯片池。 在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (例如幻灯片标题或图像) 幻灯片的 ID。 在用户可以从任何任意演示文稿插入幻灯片的情况下，也可使用 API，但在这种情况下，用户实际上只能插入源演示文稿的所有幻灯片。  有关详细信息 [，请参阅](#selecting-which-slides-to-insert) 选择要插入的幻灯片。

将幻灯片从一个演示文稿插入另一个演示文稿有两个步骤。

1. 将源演示文稿文件 (.pptx) 转换为 base64 格式的字符串。
1. 使用此方法将 base64 文件中一个或多个幻灯片 `insertSlidesFromBase64` 插入当前演示文稿。

## <a name="convert-the-source-presentation-to-base64"></a>将源演示文稿转换为 base64

有许多方法可以将文件转换为 base64。 使用哪种编程语言和库，以及是在加载项的服务器端还是客户端进行转换取决于你的方案。 通常，你将使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 对象在客户端上的 JavaScript 中执行转换。 以下示例演示此做法。

1. 首先获取对源 PowerPoint 文件的引用。 本示例中，我们将使用 `<input>` 类型控件 `file` 提示用户选择文件。 将以下标记添加到外接程序页面。

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    此标记将以下屏幕截图中的 UI 添加到页面：

    ![Screenshot showing an HTML file type input control preceded by an instructional sentence reading "Select a PowerPoint presentation from which to insert slides". 该控件包含一个标记为"选择文件"的按钮，后跟一句"未选择文件"。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > 有许多其他方法可以获取 PowerPoint 文件。 例如，如果文件存储在 OneDrive 或 SharePoint 上，可以使用 Microsoft Graph 下载它。 有关详细信息，请参阅使用 [Microsoft Graph 中的文件和使用](/graph/api/resources/onedrive) [Microsoft Graph 访问文件](/learn/modules/msgraph-access-file-data/)。

2. 将以下代码添加到加载项的 JavaScript，以将函数分配给输入控件 `change` 的事件。  (下一 `storeFileAsBase64` 步创建函数。) 

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. 添加以下代码。 关于此代码，请注意以下事项：

    - 该方法 `reader.readAsDataURL` 将文件转换为 base64，并将其存储在 `reader.result` 属性中。 方法完成后，它将触发 `onload` 事件处理程序。
    - 事件处理程序将剪裁编码文件的元数据，将编码 `onload` 字符串存储在全局变量中。
    - base64 编码的字符串全局存储，因为它由你在稍后步骤创建的另一个函数读取。

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

## <a name="insert-slides-with-insertslidesfrombase64"></a>插入包含 insertSlidesFromBase64 的幻灯片

加载项使用 [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) 方法将另一个 PowerPoint 演示文稿中的幻灯片插入到当前演示文稿中。 下面是一个简单的示例，其中源演示文稿的所有幻灯片都插入到当前演示文稿的开头，并且插入的幻灯片保留源文件的格式。 请注意， `chosenFileBase64` 这是一个包含 PowerPoint 演示文稿文件的 base64 编码版本的全局变量。

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

可以通过将[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)对象作为第二个参数传递给来控制插入结果的某些方面，包括幻灯片的插入位置以及幻灯片是获取源格式还是目标格式。 `insertSlidesFromBase64` 示例如下。 关于此代码，请注意以下几点：

- 该属性有两个可能 `formatting` 的值："UseDestinationTheme"和"KeepSourceFormatting"。 （可选）您可以使用枚举， (`InsertSlideFormatting` 例如 `PowerPoint.InsertSlideFormatting.useDestinationTheme` ，) 。
- 该函数将紧接在属性指定的幻灯片之后插入源演示文稿中的 `targetSlideId` 幻灯片。 此属性的值是三种可能形式之一的字符串：***nnn*#**、* *#* mmmmmmmmm***或 **_nnn_ #* mmmmmmmmm***，其中 *nnn* 是幻灯片的 ID (通常为 3 个数字) 而 *mmmmmmmmm 是* 幻灯片的创建 ID (通常为 9 位数字) 。 一些示例包括 `267#763315295` `267#` ， 和 `#763315295` 。

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

当然，在编码时通常不知道目标幻灯片的 ID 或创建 ID。 通常，加载项会要求用户选择目标幻灯片。 以下步骤显示如何获取当前选定幻灯片的 ***nnn*#** ID，并使用它作为目标幻灯片。

1. 使用通用 JavaScript API 的 [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法创建一个函数，获取当前选定幻灯片的 ID。 示例如下。 请注意，对的 `getSelectedDataAsync` 调用嵌入 Promise 返回函数中。 有关这样做的原因和如何操作，请参阅Common-APIs [返回函数中的 Wrap 对象](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。

 
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

1. 在主函数的[PowerPoint.run () ](/javascript/api/powerpoint#PowerPoint_run_batch_)中调用新函数，并传递它返回的 ID (，该 ID 与"#"符号) 连接作为参数的属性值。 `targetSlideId` `InsertSlideOptions` 示例如下。

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

您还可以使用 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 参数来控制插入源演示文稿中的哪些幻灯片。 为此，请为属性分配源演示文稿幻灯片的一 `sourceSlideIds` 个数组。 下面是插入四张幻灯片的示例。 请注意，数组中的每个字符串必须遵循用于该属性的一种或另一 `targetSlideId` 种模式。

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
> 幻灯片的插入顺序与它们在源演示文稿中的显示顺序相同，而不管它们在数组中的显示顺序如何。

用户无法实际发现源演示文稿中幻灯片的 ID 或创建 ID。 因此，实际上，只有当在编码时知道源标识，或者您的外接程序可以在运行时从某些数据源检索源标识时，才能真正 `sourceSlideIds` 使用该属性。 由于无法要求用户记住幻灯片 ID，因此还需要一种方法使用户可以选择幻灯片（可能按标题或图像选择）然后将每个标题或图像与幻灯片 ID 关联。

因此，该属性主要用于演示文稿模板方案：外接程序旨在处理一组特定的演示文稿，这些演示文稿充当可插入的幻灯片 `sourceSlideIds` 池。 在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (（如标题或图像) ）与从可能源演示文稿集构造的幻灯片 ID 或幻灯片创建 ID 关联。
