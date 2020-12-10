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
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a><span data-ttu-id="b3643-103">在 PowerPoint 演示文稿中插入和删除幻灯片 (预览) </span><span class="sxs-lookup"><span data-stu-id="b3643-103">Insert and delete slides in a PowerPoint presentation (preview)</span></span>

<span data-ttu-id="b3643-104">PowerPoint 加载项可以使用 PowerPoint 的应用程序特定的 JavaScript 库，将幻灯片从一个演示文稿插入到当前演示文稿中。</span><span class="sxs-lookup"><span data-stu-id="b3643-104">A PowerPoint add-in can insert slides from one presentation into the current presentation by using PowerPoint's application-specific JavaScript library.</span></span> <span data-ttu-id="b3643-105">您可以控制插入的幻灯片是保留源演示文稿的格式还是目标演示文稿的格式。</span><span class="sxs-lookup"><span data-stu-id="b3643-105">You can control whether the inserted slides keep the formatting of the source presentation or the formatting of the target presentation.</span></span> <span data-ttu-id="b3643-106">您还可以从演示文稿中删除幻灯片。</span><span class="sxs-lookup"><span data-stu-id="b3643-106">You can also delete slides from the presentation.</span></span>

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

<span data-ttu-id="b3643-107">幻灯片插入 Api 主要在演示文稿模板方案中使用：有少量的已知演示文稿充当可由加载项插入的幻灯片池。</span><span class="sxs-lookup"><span data-stu-id="b3643-107">The slide insertion APIs are primarily used in presentation template scenarios: There are a small number of known presentations which serve as pools of slides that can be inserted by the add-in.</span></span> <span data-ttu-id="b3643-108">在这种情况下，您或客户必须创建并维护一个与选择条件关联的数据源 (如幻灯片标题或与幻灯片 Id) 的图像。</span><span class="sxs-lookup"><span data-stu-id="b3643-108">In such a scenario, either you or the customer must create and maintain a data source that correlates the selection criterion (such as slide titles or images) with slide IDs.</span></span> <span data-ttu-id="b3643-109">在用户可以插入任意演示文稿中的幻灯片时，也可以使用 Api，但在这种情况下，用户有效地限制为从源演示文稿插入 *所有* 幻灯片。</span><span class="sxs-lookup"><span data-stu-id="b3643-109">The APIs can also be used in scenarios where the user can insert slides from any arbitrary presentation, but in that scenario the user is effectively limited to inserting *all* the slides from the source presentation.</span></span> <span data-ttu-id="b3643-110">有关详细信息，请参阅 [选择要插入的幻灯片](#selecting-which-slides-to-insert) 。</span><span class="sxs-lookup"><span data-stu-id="b3643-110">See [Selecting which slides to insert](#selecting-which-slides-to-insert) for more information about this.</span></span>

<span data-ttu-id="b3643-111">将一个演示文稿中的幻灯片插入到另一个演示文稿中有两个步骤。</span><span class="sxs-lookup"><span data-stu-id="b3643-111">There are two steps to inserting slides from one presentation into another.</span></span>

1. <span data-ttu-id="b3643-112">将源演示文稿文件 ( .pptx) 转换为 base64 格式的字符串。</span><span class="sxs-lookup"><span data-stu-id="b3643-112">Convert the source presentation file (.pptx) into a base64-formatted string.</span></span>
1. <span data-ttu-id="b3643-113">使用 `insertSlidesFromBase64` 方法可将 base64 文件中的一个或多个幻灯片插入到当前演示文稿中。</span><span class="sxs-lookup"><span data-stu-id="b3643-113">Use the `insertSlidesFromBase64` method to insert one or more slides from the base64 file into the current presentation.</span></span>

## <a name="convert-the-source-presentation-to-base64"></a><span data-ttu-id="b3643-114">将源演示文稿转换为 base64</span><span class="sxs-lookup"><span data-stu-id="b3643-114">Convert the source presentation to base64</span></span>

<span data-ttu-id="b3643-115">有多种方法可将文件转换为 base64。</span><span class="sxs-lookup"><span data-stu-id="b3643-115">There are many ways to convert a file to base64.</span></span> <span data-ttu-id="b3643-116">您使用哪种编程语言和库，以及是否在外接程序的服务器端或客户端进行转换取决于您的方案。</span><span class="sxs-lookup"><span data-stu-id="b3643-116">Which programming language and library you use, and whether to convert on the server-side of your add-in or the client-side is determined by your scenario.</span></span> <span data-ttu-id="b3643-117">最常见的情况是，通过使用 [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) 对象在客户端上使用 JavaScript 进行转换。</span><span class="sxs-lookup"><span data-stu-id="b3643-117">Most commonly, you'll do the conversion in JavaScript on the client-side by using a [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) object.</span></span> <span data-ttu-id="b3643-118">下面的示例演示了这种做法。</span><span class="sxs-lookup"><span data-stu-id="b3643-118">The following example shows this practice.</span></span>

1. <span data-ttu-id="b3643-119">首先获取对源 PowerPoint 文件的引用。</span><span class="sxs-lookup"><span data-stu-id="b3643-119">Begin by getting a reference to the source PowerPoint file.</span></span> <span data-ttu-id="b3643-120">在此示例中，我们将使用 `<input>` 类型的控件 `file` 来提示用户选择文件。</span><span class="sxs-lookup"><span data-stu-id="b3643-120">In this example, we will use an `<input>` control of type `file` to prompt the user to choose a file.</span></span> <span data-ttu-id="b3643-121">将以下标记添加到加载项页面。</span><span class="sxs-lookup"><span data-stu-id="b3643-121">Add the following markup to the add-in page.</span></span>

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    <span data-ttu-id="b3643-122">此标记将以下屏幕截图中的 UI 添加到页面：</span><span class="sxs-lookup"><span data-stu-id="b3643-122">This markup adds the UI in the following screenshot to the page:</span></span>

    ![显示 HTML 文件类型输入控件的屏幕截图，其前面有一个说明语句 "选择要从中插入幻灯片的 PowerPoint 演示文稿"。](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > <span data-ttu-id="b3643-125">有许多其他方法可以获取 PowerPoint 文件。</span><span class="sxs-lookup"><span data-stu-id="b3643-125">There are many other ways to get a PowerPoint file.</span></span> <span data-ttu-id="b3643-126">例如，如果文件存储在 OneDrive 或 SharePoint 上，则可以使用 Microsoft Graph 下载它。</span><span class="sxs-lookup"><span data-stu-id="b3643-126">For example, if the file is stored on OneDrive or SharePoint, you can use Microsoft Graph to download it.</span></span> <span data-ttu-id="b3643-127">有关详细信息，请参阅使用 [Microsoft graph 中的文件](/graph/api/resources/onedrive) 和 [使用 Microsoft graph 访问文件](/learn/modules/msgraph-access-file-data/)。</span><span class="sxs-lookup"><span data-stu-id="b3643-127">For more information, see [Working with files in Microsoft Graph](/graph/api/resources/onedrive) and [Access Files with Microsoft Graph](/learn/modules/msgraph-access-file-data/).</span></span>

2. <span data-ttu-id="b3643-128">将以下代码添加到加载项的 JavaScript 中，以将函数分配给输入控件的 `change` 事件。</span><span class="sxs-lookup"><span data-stu-id="b3643-128">Add the following code to the add-in's JavaScript to assign a function to the input control's `change` event.</span></span> <span data-ttu-id="b3643-129"> (您 `storeFileAsBase64` 在下一步中创建函数。 ) </span><span class="sxs-lookup"><span data-stu-id="b3643-129">(You create the `storeFileAsBase64` function in the next step.)</span></span>

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. <span data-ttu-id="b3643-130">添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="b3643-130">Add the following code.</span></span> <span data-ttu-id="b3643-131">请注意有关此代码的以下内容：</span><span class="sxs-lookup"><span data-stu-id="b3643-131">Note the following about this code,:</span></span>

    - <span data-ttu-id="b3643-132">该 `reader.readAsDataURL` 方法将文件转换为 base64 并将其存储在 `reader.result` 属性中。</span><span class="sxs-lookup"><span data-stu-id="b3643-132">The `reader.readAsDataURL` method converts the file to base64 and stores it in the `reader.result` property.</span></span> <span data-ttu-id="b3643-133">方法完成后，它将触发 `onload` 事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="b3643-133">When the method completes, it triggers the `onload` event handler.</span></span>
    - <span data-ttu-id="b3643-134">`onload`事件处理程序从编码的文件中去除元数据，并将编码后的字符串存储在全局变量中。</span><span class="sxs-lookup"><span data-stu-id="b3643-134">The `onload` event handler trims metadata off of the encoded file and stores the encoded string in a global variable.</span></span>
    - <span data-ttu-id="b3643-135">Base64 编码的字符串在全局范围内存储，因为它将被在后续步骤中创建的其他函数读取。</span><span class="sxs-lookup"><span data-stu-id="b3643-135">The base64-encoded string is stored globally because it will be read by another function that you create in a later step.</span></span>

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

## <a name="insert-slides-with-insertslidesfrombase64"></a><span data-ttu-id="b3643-136">插入带 insertSlidesFromBase64 的幻灯片</span><span class="sxs-lookup"><span data-stu-id="b3643-136">Insert slides with insertSlidesFromBase64</span></span>

<span data-ttu-id="b3643-137">您的外接程序使用 [insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) 方法将另一个 PowerPoint 演示文稿中的幻灯片插入到当前演示文稿中。</span><span class="sxs-lookup"><span data-stu-id="b3643-137">Your add-in inserts slides from another PowerPoint presentation into the current presentation with the [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) method.</span></span> <span data-ttu-id="b3643-138">下面是一个简单的示例，其中源演示文稿中的所有幻灯片都插入到当前演示文稿的开头，并且插入的幻灯片保留源文件的格式。</span><span class="sxs-lookup"><span data-stu-id="b3643-138">The following is a simple example in which all of the slides from the source presentation are inserted at the beginning of the current presentation and the inserted slides keep the formatting of the source file.</span></span> <span data-ttu-id="b3643-139">请注意，它 `chosenFileBase64` 是一个包含 base64 编码版本的 PowerPoint 演示文稿文件的全局变量。</span><span class="sxs-lookup"><span data-stu-id="b3643-139">Note that `chosenFileBase64` is a global variable that holds a base64-encoded version of a PowerPoint presentation file.</span></span>

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

<span data-ttu-id="b3643-140">通过将 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 对象作为第二个参数传递，可以控制插入结果的某些方面，包括插入幻灯片的位置以及它们是否获取源或目标的格式 `insertSlidesFromBase64` 。</span><span class="sxs-lookup"><span data-stu-id="b3643-140">You can control some aspects of the insertion result, including where the slides are inserted and whether they get the source or target formatting , by passing an [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) object as a second parameter to `insertSlidesFromBase64`.</span></span> <span data-ttu-id="b3643-141">示例如下。</span><span class="sxs-lookup"><span data-stu-id="b3643-141">The following is an example.</span></span> <span data-ttu-id="b3643-142">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="b3643-142">About this code, note:</span></span>

- <span data-ttu-id="b3643-143">属性有两个可能的值 `formatting` ： "UseDestinationTheme" 和 "KeepSourceFormatting"。</span><span class="sxs-lookup"><span data-stu-id="b3643-143">There are two possible values for the `formatting` property: "UseDestinationTheme" and "KeepSourceFormatting".</span></span> <span data-ttu-id="b3643-144">（可选）您可以使用 `InsertSlideFormatting` enum， (例如， `PowerPoint.InsertSlideFormatting.useDestinationTheme`) 。</span><span class="sxs-lookup"><span data-stu-id="b3643-144">Optionally, you can use the `InsertSlideFormatting` enum, (e.g., `PowerPoint.InsertSlideFormatting.useDestinationTheme`).</span></span>
- <span data-ttu-id="b3643-145">函数将在由属性指定的幻灯片之后立即在源演示文稿中插入幻灯片 `targetSlideId` 。</span><span class="sxs-lookup"><span data-stu-id="b3643-145">The function will insert the slides from the source presentation immediately after the slide specified by the `targetSlideId` property.</span></span> <span data-ttu-id="b3643-146">此属性的值是包含以下三种格式之一的字符串： \***nnn \* #**、\* *#* mmmmmmmmm \* \* \* 或 \**_nnn_ #* mmmmmmmmm \* \* \*，其中 *nnn* 是幻灯片的 id (通常为3个数字) 并且 *mmmmmmmmm* 是幻灯片的创建 id (通常) 9 个数字。</span><span class="sxs-lookup"><span data-stu-id="b3643-146">The value of this property is a string of one of three possible forms: \***nnn\*#**, \**#* mmmmmmmmm\*\*\*, or \**_nnn_#* mmmmmmmmm\*\*\*, where *nnn* is the slide's ID (typically 3 digits) and *mmmmmmmmm* is the slide's creation ID (typically 9 digits).</span></span> <span data-ttu-id="b3643-147">例如、 `267#763315295` `267#` 和 `#763315295` 。</span><span class="sxs-lookup"><span data-stu-id="b3643-147">Some examples are `267#763315295`, `267#`, and `#763315295`.</span></span>

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

<span data-ttu-id="b3643-148">当然，您通常不会在编码时知道目标幻灯片的 ID 或创建 ID。</span><span class="sxs-lookup"><span data-stu-id="b3643-148">Of course, you typically won't know at coding time the ID or creation ID of the target slide.</span></span> <span data-ttu-id="b3643-149">更常见的情况是，外接程序会要求用户选择目标幻灯片。</span><span class="sxs-lookup"><span data-stu-id="b3643-149">More commonly, an add-in will ask users to select the target slide.</span></span> <span data-ttu-id="b3643-150">以下步骤显示了如何获取当前选定幻灯片的 \***nnn \* #** ID 并将其用作目标幻灯片。</span><span class="sxs-lookup"><span data-stu-id="b3643-150">The following steps show how to get the \***nnn\*#** ID of the currently selected slide and use it as the target slide.</span></span>

1. <span data-ttu-id="b3643-151">使用通用 JavaScript Api 的 [Office.context.docUment getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) 方法创建一个函数，该函数可获取当前选定幻灯片的 ID。</span><span class="sxs-lookup"><span data-stu-id="b3643-151">Create a function that gets the ID of the currently selected slide by using the [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) method of the Common JavaScript APIs.</span></span> <span data-ttu-id="b3643-152">示例如下。</span><span class="sxs-lookup"><span data-stu-id="b3643-152">The following is an example.</span></span> <span data-ttu-id="b3643-153">请注意，调用 `getSelectedDataAsync` 被嵌入到承诺返回的函数中。</span><span class="sxs-lookup"><span data-stu-id="b3643-153">Note that the call to `getSelectedDataAsync` is embedded in a Promise-returning function.</span></span> <span data-ttu-id="b3643-154">有关为什么以及如何执行此操作的详细信息，请参阅 [在承诺返回函数中换行 Common-APIs](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)。</span><span class="sxs-lookup"><span data-stu-id="b3643-154">For more information about why and how to do this, see [Wrap Common-APIs in promise-returning functions](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).</span></span>

 
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

1. <span data-ttu-id="b3643-155">在 PowerPoint 中调用新函数 [。运行 main 函数的 ( # B1 ](/javascript/api/powerpoint#PowerPoint_run_batch_) ，并将其返回的 ID (传递它返回的 ID。) 作为参数的属性值与 "#" 符号连接 `targetSlideId` `InsertSlideOptions` 。</span><span class="sxs-lookup"><span data-stu-id="b3643-155">Call your new function inside the [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) of the main function and pass the ID that it returns (concatenated with the "#" symbol) as the value of the `targetSlideId` property of the `InsertSlideOptions` parameter.</span></span> <span data-ttu-id="b3643-156">示例如下。</span><span class="sxs-lookup"><span data-stu-id="b3643-156">The following is an example.</span></span>

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

### <a name="selecting-which-slides-to-insert"></a><span data-ttu-id="b3643-157">选择要插入的幻灯片</span><span class="sxs-lookup"><span data-stu-id="b3643-157">Selecting which slides to insert</span></span>

<span data-ttu-id="b3643-158">您还可以使用 [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) 参数来控制源演示文稿中插入的幻灯片。</span><span class="sxs-lookup"><span data-stu-id="b3643-158">You can also use the [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) parameter to control which slides from the source presentation are inserted.</span></span> <span data-ttu-id="b3643-159">为此，可通过将源演示文稿的幻灯片 Id 的数组分配给属性来执行此操作 `sourceSlideIds` 。</span><span class="sxs-lookup"><span data-stu-id="b3643-159">You do this by assigning an array of the source presentation's slide IDs to the `sourceSlideIds` property.</span></span> <span data-ttu-id="b3643-160">下面是插入四张幻灯片的示例。</span><span class="sxs-lookup"><span data-stu-id="b3643-160">The following is an example that inserts four slides.</span></span> <span data-ttu-id="b3643-161">请注意，数组中的每个字符串必须遵循用于该属性的一个或另一个模式 `targetSlideId` 。</span><span class="sxs-lookup"><span data-stu-id="b3643-161">Note that each string in the array must follow one or another of the patterns used for the `targetSlideId` property.</span></span>

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
> <span data-ttu-id="b3643-162">幻灯片将按照其在源演示文稿中出现的相对顺序进行插入，而不考虑它们在数组中的显示顺序。</span><span class="sxs-lookup"><span data-stu-id="b3643-162">The slides will be inserted in the same relative order in which they appear in the source presentation, regardless of the order in which they appear in the array.</span></span>

<span data-ttu-id="b3643-163">用户无法在源演示文稿中发现幻灯片的 ID 或创建 ID，这是一种切实可行的方法。</span><span class="sxs-lookup"><span data-stu-id="b3643-163">There is no practical way that users can discover the ID or creation ID of a slide in the source presentation.</span></span> <span data-ttu-id="b3643-164">因此，仅 `sourceSlideIds` 当您知道编码时的源 id 或加载项可以在运行时从某些数据源检索这些 id 时，才能真正使用属性。</span><span class="sxs-lookup"><span data-stu-id="b3643-164">For this reason, you can really only use the `sourceSlideIds` property when either you know the source IDs at coding time or your add-in can retrieve them at runtime from some data source.</span></span> <span data-ttu-id="b3643-165">由于无法预期用户能够记住幻灯片 Id，因此还需要一种方法来使用户能够选择幻灯片（如标题或图像），然后将每个标题或图像与幻灯片的 ID 关联起来。</span><span class="sxs-lookup"><span data-stu-id="b3643-165">Because users cannot be expected to memorize slide IDs, you also need a way to enable the user to select slides, perhaps by title or by an image, and then correlate each title or image with the slide's ID.</span></span>

<span data-ttu-id="b3643-166">因此，该 `sourceSlideIds` 属性主要用于演示文稿模板方案：外接程序设计为使用一组特定的演示文稿，用作可插入的幻灯片池。</span><span class="sxs-lookup"><span data-stu-id="b3643-166">Accordingly, the `sourceSlideIds` property is primarily used in presentation template scenarios: The add-in is designed to work with a specific set of presentations that serve as pools of slides that can be inserted.</span></span> <span data-ttu-id="b3643-167">在这种情况下，您或客户必须创建并维护一个与选择条件关联的数据源 (如标题或图像) 与已通过一组可能的源演示文稿构造的幻灯片 Id 或幻灯片创建 Id。</span><span class="sxs-lookup"><span data-stu-id="b3643-167">In such a scenario, either you or the customer must create and maintain a data source that correlates a selection criterion (such as titles or images) with slide IDs or slide creation IDs that has been constructed from the set of possible source presentations.</span></span>

## <a name="delete-slides"></a><span data-ttu-id="b3643-168">删除幻灯片</span><span class="sxs-lookup"><span data-stu-id="b3643-168">Delete slides</span></span>

<span data-ttu-id="b3643-169">通过获取对表示幻灯片的 [slide](/javascript/api/powerpoint/powerpoint.slide) 对象的引用并调用方法，可以删除幻灯片 `Slide.delete` 。</span><span class="sxs-lookup"><span data-stu-id="b3643-169">You can delete a slide by getting a reference to the [Slide](/javascript/api/powerpoint/powerpoint.slide) object that represents the slide and call the `Slide.delete` method.</span></span> <span data-ttu-id="b3643-170">下面是一个示例，其中第四张幻灯片被删除。</span><span class="sxs-lookup"><span data-stu-id="b3643-170">The following is an example in which the 4th slide is deleted.</span></span>

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
