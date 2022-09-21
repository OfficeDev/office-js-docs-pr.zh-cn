---
title: 在 PowerPoint 演示文稿中插入幻灯片
description: 了解如何将幻灯片从一个演示文稿插入另一个演示文稿。
ms.date: 03/07/2021
ms.localizationpriority: medium
ms.openlocfilehash: a31933de4272634394dc6c36aafa973c41265471
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857569"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>在 PowerPoint 演示文稿中插入幻灯片

PowerPoint 外接程序可以使用 PowerPoint 的应用程序特定 JavaScript 库将一个演示文稿中的幻灯片插入到当前演示文稿中。 可以控制插入的幻灯片是保留源演示文稿的格式还是目标演示文稿的格式。

幻灯片插入 API 主要用于演示文稿模板方案：有少量的已知演示文稿用作可由外接程序插入的幻灯片池。 在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (相关联，例如幻灯片标题或图像) 幻灯片 ID。 也可以在用户可以从任何任意演示文稿插入幻灯片的情况下使用 API，但在这种情况下，用户实际上仅限于插入源演示文稿中 *的所有* 幻灯片。 有关详细信息，请参阅 [选择要插入的幻灯片](#selecting-which-slides-to-insert) 。

将幻灯片从一个演示文稿插入另一个演示文稿有两个步骤。

1. 将源演示文稿文件 (.pptx) 转换为 base64 格式的字符串。
1. 使用该 `insertSlidesFromBase64` 方法将 base64 文件中的一张或多张幻灯片插入当前演示文稿中。

## <a name="convert-the-source-presentation-to-base64"></a>将源演示文稿转换为 base64

可通过多种方式将文件转换为 base64。 使用哪种编程语言和库，以及是要在外接程序的服务器端转换还是在客户端进行转换，都由你的方案决定。 大多数情况下，你将使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 对象在客户端的 JavaScript 中执行转换。 以下示例演示此做法。

1. 首先获取对源 PowerPoint 文件的引用。 在此示例中，我们将使用 `<input>` 类型的 `file` 控件提示用户选择文件。 将以下标记添加到加载项页面。

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    此标记将以下屏幕截图中的 UI 添加到页面。

    ![显示 HTML 文件类型输入控件的屏幕截图，前面是一个说明性句子，上面写着“选择要从中插入幻灯片的 PowerPoint 演示文稿”。 控件包含一个标记为“选择文件”的按钮，后跟句子“未选择文件”。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > 还有许多其他方法可以获取 PowerPoint 文件。 例如，如果文件存储在 OneDrive 或 SharePoint 上，则可以使用 Microsoft Graph 下载它。 有关详细信息，请参阅 [使用 Microsoft Graph](/graph/api/resources/onedrive) 和 [Microsoft Graph 访问文件](/training/modules/msgraph-access-file-data/)中的文件。

2. 将以下代码添加到加载项的 JavaScript，以便将函数分配给输入控 `change` 件的事件。  (下一步创建函 `storeFileAsBase64` 数。) 

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. 添加以下代码。 对于此代码，请注意以下事项。

    - 该 `reader.readAsDataURL` 方法将文件转换为 base64，并将其存储在属性中 `reader.result` 。 方法完成后，将触发 `onload` 事件处理程序。
    - 事件 `onload` 处理程序从编码的文件中剪裁元数据，并将编码的字符串存储在全局变量中。
    - base64 编码的字符串全局存储，因为它将由在后续步骤中创建的另一个函数读取。

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

## <a name="insert-slides-with-insertslidesfrombase64"></a>使用 insertSlidesFromBase64 插入幻灯片

外接程序使用 [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)) 方法将另一个 PowerPoint 演示文稿中的幻灯片插入当前演示文稿中。 下面是一个简单的示例，其中源演示文稿中的所有幻灯片都插入到当前演示文稿的开头，插入的幻灯片保留源文件的格式。 请注意， `chosenFileBase64` 这是一个全局变量，包含 PowerPoint 演示文稿文件的 base64 编码版本。

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

你可以通过将 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 对象作为第二个参数传递给 `insertSlidesFromBase64`插入结果的某些方面，包括插入幻灯片的位置以及它们是否获取源或目标格式。 示例如下。 关于此代码，请注意以下几点：

- 属性有两个可能的 `formatting` 值：“UseDestinationTheme”和“KeepSourceFormatting”。 （可选）可以使用 `InsertSlideFormatting` 枚举， (例如， `PowerPoint.InsertSlideFormatting.useDestinationTheme`) 。
- 该函数将在属性指定的幻灯片之后立即插入源演示文稿中的 `targetSlideId` 幻灯片。 此属性的值是三种可能形式之一的字符串：***nnn*#**、**#* mmmmmmmmm***或 *nnn mmmmmmm***，其中 *_nnn_#* 是幻灯片的 ID， (通常为 3 位数字) ，*而 mmmmmmmmm* 是幻灯片的创建 ID，通常 (9 位数字) 。 一些示例包括 `267#763315295`， `267#`和 `#763315295`.

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

当然，通常不会在编码时知道目标幻灯片的 ID 或创建 ID。 更常见的是，外接程序会要求用户选择目标幻灯片。 以下步骤演示如何获取当前所选幻灯片的 ***nnn*#** ID，并将其用作目标幻灯片。

1. 创建一个函数，该函数通过使用 Common JavaScript API 的 [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) 方法获取当前所选幻灯片的 ID。 示例如下。 请注意，调 `getSelectedDataAsync` 用嵌入到 Promise-returning 函数中。 有关为何以及如何执行此操作的详细信息，请参阅 [承诺返回函数中的 Wrap Common-APIs](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。

 
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

1. 在主函数的 [PowerPoint.run () ](/javascript/api/powerpoint#PowerPoint_run_batch_) 中调用新函数，并传递它返回的 ID (与“#”符号) 连接为参数属性的`targetSlideId``InsertSlideOptions`值。 示例如下。

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

还可以使用 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 参数来控制从源演示文稿插入的幻灯片。 为此，请将源演示文稿的幻灯片 ID 数组分配给属性 `sourceSlideIds` 。 下面是插入四张幻灯片的示例。 请注意，数组中的每个字符串必须遵循用于属性 `targetSlideId` 的一个或另一个模式。

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
> 幻灯片将按在源演示文稿中显示的相对顺序插入，而不考虑它们在数组中的显示顺序。

用户无法在源演示文稿中发现幻灯片的 ID 或创建 ID。 因此，实际上只能在知道编码时的源 ID 或加载项可以在运行时从某些数据源检索它们时使用 `sourceSlideIds` 该属性。 由于不能期望用户记住幻灯片 ID，因此还需要一种方法来使用户能够选择幻灯片（可能按标题或图像选择幻灯片，然后将每个标题或图像与幻灯片的 ID 相关联）。

因此，该 `sourceSlideIds` 属性主要用于演示文稿模板方案：外接程序旨在处理一组特定的演示文稿，这些演示文稿充当可插入的幻灯片池。 在这种情况下，你或客户必须创建和维护一个数据源，该数据源将选择条件 (相关联，例如标题或图像) 幻灯片 ID 或从可能源演示文稿集构造的幻灯片创建 ID。
