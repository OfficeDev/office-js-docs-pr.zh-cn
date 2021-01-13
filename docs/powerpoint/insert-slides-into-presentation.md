---
title: 在 PowerPoint 演示文稿中插入和删除幻灯片
description: 了解如何将一个演示文稿中的幻灯片插入另一个演示文稿，以及如何删除幻灯片。
ms.date: 01/08/2021
localization_priority: Normal
ms.openlocfilehash: a9a4b2efd1e970d9c45885f9a17046bec4de7e72
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839717"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation"></a><span data-ttu-id="a906e-103">在 PowerPoint 演示文稿中插入和删除幻灯片</span><span class="sxs-lookup"><span data-stu-id="a906e-103">Insert and delete slides in a PowerPoint presentation</span></span>

<span data-ttu-id="a906e-104">PowerPoint 加载项可以使用 PowerPoint 特定于应用程序的 JavaScript 库将一个演示文稿中的幻灯片插入到当前演示文稿中。</span><span class="sxs-lookup"><span data-stu-id="a906e-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="a906e-105">您可以控制插入的幻灯片是否保留源演示文稿的格式或目标演示文稿的格式。</span><span class="sxs-lookup"><span data-stu-id="a906e-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="a906e-106">还可以从演示文稿中删除幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a906e-106">You can also delete slides from the presentation.</span></span>

<span data-ttu-id="a906e-107">幻灯片插入 API 主要用于演示文稿模板方案：少量已知演示文稿充当加载项可以插入的幻灯片池。</span><span class="sxs-lookup"><span data-stu-id="a906e-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="a906e-108">在这种情况下，您或客户必须创建和维护一个数据源，该数据源将选择条件 (如幻灯片标题或图像) 与幻灯片 ID 相关联。</span><span class="sxs-lookup"><span data-stu-id="a906e-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="a906e-109">这些 API 还可用于以下方案：用户可以插入任意演示文稿中的幻灯片，但在这种情况下，用户实际上只能插入源演示文稿的所有幻灯片。 </span><span class="sxs-lookup"><span data-stu-id="a906e-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="a906e-110">有关详细信息 [，请参阅"选择要](#selecting-which-slides-to-insert) 插入的幻灯片"。</span><span class="sxs-lookup"><span data-stu-id="a906e-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="a906e-111">将幻灯片从一个演示文稿插入另一个演示文稿有两个步骤。</span><span class="sxs-lookup"><span data-stu-id="a906e-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="a906e-112">将源演示文稿文件 (.pptx) 转换为 base64 格式的字符串。</span><span class="sxs-lookup"><span data-stu-id="a906e-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="a906e-113">使用此方法 `insertSlidesFromBase64` 将 base64 文件中一个或多个幻灯片插入当前演示文稿。</span><span class="sxs-lookup"><span data-stu-id="a906e-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="a906e-114">将源演示文稿转换为 base64</span><span class="sxs-lookup"><span data-stu-id="a906e-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="a906e-115">有许多方法可以将文件转换为 base64。</span><span class="sxs-lookup"><span data-stu-id="a906e-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="a906e-116">使用哪种编程语言和库，以及是在加载项的服务器端还是客户端进行转换由你的方案决定。</span><span class="sxs-lookup"><span data-stu-id="a906e-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="a906e-117">通常，你将使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 对象在客户端上的 JavaScript 中执行转换。</span><span class="sxs-lookup"><span data-stu-id="a906e-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="a906e-118">以下示例演示此做法。</span><span class="sxs-lookup"><span data-stu-id="a906e-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="a906e-119">首先获取对源 PowerPoint 文件的引用。</span><span class="sxs-lookup"><span data-stu-id="a906e-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="a906e-120">本示例中，我们将使用 `<input>` 类型控件 `file` 提示用户选择文件。</span><span class="sxs-lookup"><span data-stu-id="a906e-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="a906e-121">将以下标记添加到外接程序页面。</span><span class="sxs-lookup"><span data-stu-id="a906e-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="a906e-122">此标记将以下屏幕截图中的 UI 添加到页面：</span><span class="sxs-lookup"><span data-stu-id="a906e-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![Screenshot showing an HTML file type input control preceded by an instructional sentence reading "Select a PowerPoint presentation from which to insert slides".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="a906e-125">有许多其他方法可以获取 PowerPoint 文件。</span><span class="sxs-lookup"><span data-stu-id="a906e-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="a906e-126">例如，如果文件存储在 OneDrive 或 SharePoint 上，可以使用 Microsoft Graph 下载它。</span><span class="sxs-lookup"><span data-stu-id="a906e-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="a906e-127">有关详细信息，请参阅使用 [Microsoft Graph 中的](/graph/api/resources/onedrive) 文件和 [使用 Microsoft Graph 的 Access 文件](/learn/modules/msgraph-access-file-data/)。</span><span class="sxs-lookup"><span data-stu-id="a906e-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="a906e-128">将以下代码添加到加载项的 JavaScript，以将函数分配给输入控件 `change` 的事件。</span><span class="sxs-lookup"><span data-stu-id="a906e-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="a906e-129"> (您将在 `storeFileAsBase64` 下一步.) </span><span class="sxs-lookup"><span data-stu-id="a906e-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="a906e-130">添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="a906e-130">Add the following code.</span></span> <span data-ttu-id="a906e-131">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="a906e-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="a906e-132">该方法 `reader.readAsDataURL` 将文件转换为 base64，并将其存储在 `reader.result` 属性中。</span><span class="sxs-lookup"><span data-stu-id="a906e-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="a906e-133">该方法完成后，将触发 `onload` 事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="a906e-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="a906e-134">事件处理程序从编码文件剪裁元数据，将编码字符串 `onload` 存储在全局变量中。</span><span class="sxs-lookup"><span data-stu-id="a906e-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="a906e-135">base64 编码的字符串全局存储，因为它由你在稍后步骤中创建的另一个函数读取。</span><span class="sxs-lookup"><span data-stu-id="a906e-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="a906e-136">使用 insertSlidesFromBase64 插入幻灯片</span><span class="sxs-lookup"><span data-stu-id="a906e-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="a906e-137">外接程序使用 [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) 方法将另一个 PowerPoint 演示文稿中的幻灯片插入到当前演示文稿中。</span><span class="sxs-lookup"><span data-stu-id="a906e-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="a906e-138">下面是一个简单的示例，其中源演示文稿的所有幻灯片都插入到当前演示文稿的开头，插入的幻灯片保留源文件的格式。</span><span class="sxs-lookup"><span data-stu-id="a906e-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="a906e-139">请注意， `chosenFileBase64` 这是一个全局变量，包含 PowerPoint 演示文稿文件的 base64 编码版本。</span><span class="sxs-lookup"><span data-stu-id="a906e-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="a906e-140">可以通过将 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 对象作为第二个参数传递给控制插入结果的某些方面，包括幻灯片的插入位置以及幻灯片是获取源格式还是目标格式 `insertSlidesFromBase64` 。</span><span class="sxs-lookup"><span data-stu-id="a906e-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="a906e-141">示例如下。</span><span class="sxs-lookup"><span data-stu-id="a906e-141">The following is an example.</span></span> <span data-ttu-id="a906e-142">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="a906e-142">About this code, note:</span></span>

- <span data-ttu-id="a906e-143">该属性有两个可能 `formatting` 的值："UseDestinationTheme"和"KeepSourceFormatting"。</span><span class="sxs-lookup"><span data-stu-id="a906e-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="a906e-144">（可选）您可以使用枚举 `InsertSlideFormatting` ， (例如 `PowerPoint.InsertSlideFormatting.useDestinationTheme`) 。</span><span class="sxs-lookup"><span data-stu-id="a906e-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="a906e-145">该函数将紧接在属性指定的幻灯片之后插入源演示文稿中的 `targetSlideId` 幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a906e-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="a906e-146">此属性的值是三种可能形式之一的字符串：\***nnn\*#**、\* *#* mmm\*\*\*或 \**_nnn_ #* mmm\*\*\*，其中 *nnn* 是幻灯片的 ID (通常为 3 个数字) 而 *mmmmmmmmmmm* 是幻灯片的创建 ID (通常为 9 个数字) 。</span><span class="sxs-lookup"><span data-stu-id="a906e-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="a906e-147">例如 ， `267#763315295` `267#` 和 `#763315295` 。</span><span class="sxs-lookup"><span data-stu-id="a906e-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="a906e-148">当然，在编码时通常不知道目标幻灯片的 ID 或创建 ID。</span><span class="sxs-lookup"><span data-stu-id="a906e-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="a906e-149">通常，加载项会要求用户选择目标幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a906e-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="a906e-150">以下步骤显示如何获取当前选定幻灯片的 \***nnn\*#** ID，并使用它作为目标幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a906e-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="a906e-151">使用通用 JavaScript API 的 [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法创建获取当前选定幻灯片 ID 的函数。</span><span class="sxs-lookup"><span data-stu-id="a906e-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="a906e-152">示例如下。</span><span class="sxs-lookup"><span data-stu-id="a906e-152">The following is an example.</span></span> <span data-ttu-id="a906e-153">请注意，对的 `getSelectedDataAsync` 调用嵌入到 Promise 返回函数中。</span><span class="sxs-lookup"><span data-stu-id="a906e-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="a906e-154">有关这样做的原因和如何操作，请参阅承诺返回函数Common-APIs [自动换行](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。</span><span class="sxs-lookup"><span data-stu-id="a906e-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="a906e-155">在[主函数的 PowerPoint.run () ](/javascript/api/powerpoint#PowerPoint_run_batch_)内调用新函数，并传递它返回的 ID (与"#"符号) 连接作为参数的属性值。 `targetSlideId` `InsertSlideOptions`</span><span class="sxs-lookup"><span data-stu-id="a906e-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="a906e-156">示例如下。</span><span class="sxs-lookup"><span data-stu-id="a906e-156">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="a906e-157">选择要插入的幻灯片</span><span class="sxs-lookup"><span data-stu-id="a906e-157">Selecting which slides to insert</span></span>

<span data-ttu-id="a906e-158">您还可以使用 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 参数控制插入源演示文稿中的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="a906e-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="a906e-159">为此，请为属性分配源演示文稿幻灯片的一个 `sourceSlideIds` 数组。</span><span class="sxs-lookup"><span data-stu-id="a906e-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="a906e-160">下面是插入四张幻灯片的示例。</span><span class="sxs-lookup"><span data-stu-id="a906e-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="a906e-161">请注意，数组中的每个字符串必须遵循用于该属性的一种或另一 `targetSlideId` 种模式。</span><span class="sxs-lookup"><span data-stu-id="a906e-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="a906e-162">幻灯片的插入顺序与它们在源演示文稿中的显示顺序相同，而不管它们在数组中的显示顺序如何。</span><span class="sxs-lookup"><span data-stu-id="a906e-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="a906e-163">用户无法实际发现源演示文稿中幻灯片的 ID 或创建 ID。</span><span class="sxs-lookup"><span data-stu-id="a906e-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="a906e-164">因此，实际上，只有当在编码时知道源标识或加载项可以在运行时从某些数据源中检索源标识时，才能真正使用 `sourceSlideIds` 该属性。</span><span class="sxs-lookup"><span data-stu-id="a906e-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="a906e-165">由于无法要求用户记住幻灯片 ID，因此还需要一种方法使用户能够选择幻灯片（可能按标题或图像选择）然后将每个标题或图像与幻灯片 ID 关联。</span><span class="sxs-lookup"><span data-stu-id="a906e-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="a906e-166">因此，该属性主要用于演示文稿模板方案：外接程序旨在处理一组特定的演示文稿，这些演示文稿充当可插入的幻灯片 `sourceSlideIds` 池。</span><span class="sxs-lookup"><span data-stu-id="a906e-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="a906e-167">在此类方案中，您或客户必须创建和维护一个数据源，该数据源将选择条件 (（如标题或图像) ）与幻灯片 ID 或幻灯片创建 ID（从可能源演示文稿集构造）关联在一起。</span><span class="sxs-lookup"><span data-stu-id="a906e-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="a906e-168">删除幻灯片</span><span class="sxs-lookup"><span data-stu-id="a906e-168">Delete slides</span></span>

<span data-ttu-id="a906e-169">可以通过获取对代表幻灯片的 [Slide](/javascript/api/powerpoint/powerpoint.slide) 对象的引用来删除幻灯片并调用 `Slide.delete` 该方法。</span><span class="sxs-lookup"><span data-stu-id="a906e-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="a906e-170">下面是删除第 4 张幻灯片的示例。</span><span class="sxs-lookup"><span data-stu-id="a906e-170">The following is an example in which the 4th slide is deleted.</span></span>

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
