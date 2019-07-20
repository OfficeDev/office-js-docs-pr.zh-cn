---
title: Word 加载项教程
description: 本教程将介绍如何生成 Word 加载项，用于插入（和替换）文本区域、段落、图像、HTML、表格和内容控件。 此外，还将介绍如何设置文本格式，以及如何插入（和替换）内容控件中的内容。
ms.date: 07/17/2019
ms.prod: word
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: ff57f9b46fbdc50b39890598c78f8d0f194e8a89
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804651"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a><span data-ttu-id="85ee0-104">教程：创建 Word 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="85ee0-104">Tutorial: Create a Word task pane add-in</span></span>

<span data-ttu-id="85ee0-105">在本教程中，将创建 Word 任务窗格加载项，该加载项将：</span><span class="sxs-lookup"><span data-stu-id="85ee0-105">In this tutorial, you'll create a Word task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="85ee0-106">插入文本区域</span><span class="sxs-lookup"><span data-stu-id="85ee0-106">Inserts a range of text</span></span>
> * <span data-ttu-id="85ee0-107">设置文本格式</span><span class="sxs-lookup"><span data-stu-id="85ee0-107">Formats text</span></span>
> * <span data-ttu-id="85ee0-108">替换文本并在各个位置插入文本</span><span class="sxs-lookup"><span data-stu-id="85ee0-108">Replaces text and inserts text in various locations</span></span>
> * <span data-ttu-id="85ee0-109">插入图像、HTML 和表格</span><span class="sxs-lookup"><span data-stu-id="85ee0-109">Inserts images, HTML, and tables</span></span>
> * <span data-ttu-id="85ee0-110">创建和更新内容控件</span><span class="sxs-lookup"><span data-stu-id="85ee0-110">Creates and updates content controls</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="85ee0-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="85ee0-111">Prerequisites</span></span>

<span data-ttu-id="85ee0-112">若要学习本教程，需要安装以下各项。</span><span class="sxs-lookup"><span data-stu-id="85ee0-112">To use this tutorial, you need to have the following installed.</span></span>

- <span data-ttu-id="85ee0-p102">Word 2016 版本 1711（生成号 8730.1000 即点即用）或更高版本。 可能必须成为 Office 预览体验成员，才能获取此版本。 有关详细信息，请参阅[成为 Office 预览体验成员](https://products.office.com/office-insider?tab=tab-1)。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p102">Word 2016, version 1711 (Build 8730.1000 Click-to-Run) or later. You might need to be an Office Insider to get this version. For more information, see [Be an Office Insider](https://products.office.com/office-insider?tab=tab-1).</span></span>

- [<span data-ttu-id="85ee0-116">Node</span><span class="sxs-lookup"><span data-stu-id="85ee0-116">Node</span></span>](https://nodejs.org/en/) 

- <span data-ttu-id="85ee0-117">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="85ee0-117">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

## <a name="create-your-add-in-project"></a><span data-ttu-id="85ee0-118">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="85ee0-118">Create your add-in project</span></span>

<span data-ttu-id="85ee0-119">完成以下步骤以创建将用作本教程基础的 Word 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="85ee0-119">Complete the following steps to create the Word add-in project that you'll use as the basis for this tutorial.</span></span>

1. <span data-ttu-id="85ee0-120">克隆 GitHub 存储库 [Word 加载项教程](https://github.com/OfficeDev/Word-Add-in-Tutorial)。</span><span class="sxs-lookup"><span data-stu-id="85ee0-120">Clone the GitHub repository [Word add-in tutorial](https://github.com/OfficeDev/Word-Add-in-Tutorial).</span></span>

2. <span data-ttu-id="85ee0-121">打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="85ee0-121">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

3. <span data-ttu-id="85ee0-122">运行命令 `npm install`，以安装 package.json 文件中列出的工具和库。</span><span class="sxs-lookup"><span data-stu-id="85ee0-122">Run the command `npm install` to install the tools and libraries listed in the package.json file.</span></span> 

4. <span data-ttu-id="85ee0-123">执行[安装自签名证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)以信任开发计算机操作系统的证书中的步骤。</span><span class="sxs-lookup"><span data-stu-id="85ee0-123">Carry out the steps in [Installing the self-signed certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) to trust the certificate for your development computer's operating system.</span></span>

## <a name="insert-a-range-of-text"></a><span data-ttu-id="85ee0-124">插入文本区域</span><span class="sxs-lookup"><span data-stu-id="85ee0-124">Insert a range of text</span></span>

<span data-ttu-id="85ee0-125">本教程的这一步是，先以编程方式测试加载项是否支持用户的当前版本 Word，再在文档中插入段落。</span><span class="sxs-lookup"><span data-stu-id="85ee0-125">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Word, and then insert a paragraph into the document.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="85ee0-126">编码加载项</span><span class="sxs-lookup"><span data-stu-id="85ee0-126">Code the add-in</span></span>

1. <span data-ttu-id="85ee0-127">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="85ee0-127">Open the project in your code editor.</span></span>

2. <span data-ttu-id="85ee0-128">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-128">Open the file index.html.</span></span>

3. <span data-ttu-id="85ee0-129">将 `TODO1` 替换为以下标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-129">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. <span data-ttu-id="85ee0-130">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-130">Open the app.js file.</span></span>

5. <span data-ttu-id="85ee0-p103">将 `TODO1` 替换为下面的代码。 此代码用于确定用户的 Word 版本是否支持包含本教程所有阶段使用的全部 API 的 Word.js 版本。 在生产加载项中，若要隐藏或禁用调用不受支持的 API 的 UI，请使用条件块的主体。 这样一来，用户仍可以使用 Word 版本支持的加载项部分。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p103">Replace the `TODO1` with the following code. This code determines whether the user's version of Word supports a version of Word.js that includes all the APIs that are used in all the stages of this tutorial. In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs. This will enable the user to still use the parts of the add-in that are supported by their version of Word.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. <span data-ttu-id="85ee0-135">将 `TODO2` 替换为下面的代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-135">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. <span data-ttu-id="85ee0-136">将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="85ee0-136">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="85ee0-137">注意：</span><span class="sxs-lookup"><span data-stu-id="85ee0-137">Note:</span></span>

   - <span data-ttu-id="85ee0-p105">Word.js 业务逻辑会添加到传递给 `Word.run` 的函数中。 此逻辑不会立即执行， 而是添加到挂起命令队列中。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p105">Your Word.js business logic will be added to the function that is passed to `Word.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

   - <span data-ttu-id="85ee0-141">`context.sync` 方法将所有已排入队列的命令都发送到 Word 以供执行。</span><span class="sxs-lookup"><span data-stu-id="85ee0-141">The `context.sync` method sends all queued commands to Word for execution.</span></span>

   - <span data-ttu-id="85ee0-p106">`Word.run` 后跟 `catch` 块。 这是应始终遵循的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p106">The `Word.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO4: Queue commands to insert a paragraph into the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

8. <span data-ttu-id="85ee0-144">将 `TODO4` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="85ee0-144">Replace `TODO4` with the following code.</span></span> <span data-ttu-id="85ee0-145">请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-145">Note:</span></span>

   - <span data-ttu-id="85ee0-146">`insertParagraph` 方法的第一个参数是新段落的文本。</span><span class="sxs-lookup"><span data-stu-id="85ee0-146">The first parameter to the `insertParagraph` method is the text for the new paragraph.</span></span>

   - <span data-ttu-id="85ee0-p108">第二个参数是应在正文中的什么位置插入段落。 如果父对象为正文，其他段落插入选项包括“End”和“Replace”。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p108">The second parameter is the location within the body where the paragraph will be inserted. Other options for insert paragraph, when the parent object is the body, are "End" and "Replace".</span></span>

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office on the web.",
                            "Start");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85ee0-149">测试加载项</span><span class="sxs-lookup"><span data-stu-id="85ee0-149">Test the add-in</span></span>

1. <span data-ttu-id="85ee0-150">打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="85ee0-150">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

2. <span data-ttu-id="85ee0-151">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="85ee0-151">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="85ee0-152">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="85ee0-152">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="85ee0-153">通过以下方法之一旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="85ee0-153">Sideload the add-in by using one of the following methods:</span></span>

    - <span data-ttu-id="85ee0-154">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="85ee0-154">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>

    - <span data-ttu-id="85ee0-155">Web 浏览器: 在[web 上的 office 中旁加载 Office 外接程序](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="85ee0-155">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>

    - <span data-ttu-id="85ee0-156">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="85ee0-156">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

5. <span data-ttu-id="85ee0-157">在 Word 的“开始”\*\*\*\* 菜单中，选择“显示任务窗格”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="85ee0-157">On the **Home** menu of Word, select **Show Taskpane**.</span></span>

6. <span data-ttu-id="85ee0-158">在任务窗格中，选择“插入段落”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="85ee0-158">In the task pane, choose **Insert Paragraph**.</span></span>

7. <span data-ttu-id="85ee0-159">在段落中进行一些更改。</span><span class="sxs-lookup"><span data-stu-id="85ee0-159">Make a change in the paragraph.</span></span>

8. <span data-ttu-id="85ee0-160">再次选择“插入段落”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="85ee0-160">Choose **Insert Paragraph** again.</span></span> <span data-ttu-id="85ee0-161">观察新段落是否位于上一段落之上，因为 `insertParagraph` 方法要在文档正文的“开头”插入内容。</span><span class="sxs-lookup"><span data-stu-id="85ee0-161">Note that the new paragraph is above the previous one because the `insertParagraph` method is inserting at the start of the document's body.</span></span>

    ![Word 教程 - 插入段落](../images/word-tutorial-insert-paragraph.png)

## <a name="format-text"></a><span data-ttu-id="85ee0-163">设置文本格式</span><span class="sxs-lookup"><span data-stu-id="85ee0-163">Format text</span></span>

<span data-ttu-id="85ee0-164">在本教程的此步骤中，你将向文本应用嵌入样式、向文本应用自定义样式并更改文本字体。</span><span class="sxs-lookup"><span data-stu-id="85ee0-164">In this step of the tutorial, you'll apply a built-in style to text, apply a custom style to text, and change the font of text.</span></span>

### <a name="apply-a-built-in-style-to-text"></a><span data-ttu-id="85ee0-165">向文本应用嵌入样式</span><span class="sxs-lookup"><span data-stu-id="85ee0-165">Apply a built-in style to text</span></span>

1. <span data-ttu-id="85ee0-166">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="85ee0-166">Open the project in your code editor.</span></span> 

2. <span data-ttu-id="85ee0-167">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-167">Open the file index.html.</span></span>

3. <span data-ttu-id="85ee0-168">在包含 `insert-paragraph` 按钮的 `div` 正下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-168">Just below the `div` that contains the `insert-paragraph` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. <span data-ttu-id="85ee0-169">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-169">Open the app.js file.</span></span>

5. <span data-ttu-id="85ee0-170">在向 `insert-paragraph` 按钮分配单击处理程序的代码行正下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-170">Just below the line that assigns a click handler to the `insert-paragraph` button, add the following code:</span></span>

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. <span data-ttu-id="85ee0-171">在 `insertParagraph` 函数正下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-171">Just below the `insertParagraph` function, add the following function:</span></span>

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="85ee0-p110">将 `TODO1` 替换为下面的代码。 请注意，此代码向段落应用样式，但也可以向文本区域应用样式。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p110">Replace `TODO1` with the following code. Note that the code applies a style to a paragraph, but styles can also be applied to ranges of text.</span></span>

    ```js
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a><span data-ttu-id="85ee0-174">向文本应用自定义样式</span><span class="sxs-lookup"><span data-stu-id="85ee0-174">Apply a custom style to text</span></span>

1. <span data-ttu-id="85ee0-175">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-175">Open the file index.html.</span></span>

2. <span data-ttu-id="85ee0-176">在包含 `apply-style` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-176">Below the `div` that contains the `apply-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. <span data-ttu-id="85ee0-177">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-177">Open the app.js file.</span></span>

4. <span data-ttu-id="85ee0-178">在向 `apply-style` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-178">Below the line that assigns a click handler to the `apply-style` button, add the following code:</span></span>

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. <span data-ttu-id="85ee0-179">在 `applyStyle` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-179">Below the `applyStyle` function, add the following function:</span></span>

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="85ee0-p111">将 `TODO1` 替换为下面的代码。 请注意，此代码应用的自定义样式尚不存在。 将在**测试加载项**步骤中创建 [MyCustomStyle](#test-the-add-in) 样式。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p111">Replace `TODO1` with the following code. Note that the code applies a custom style that does not exist yet. You'll create a style with the name **MyCustomStyle** in the [Test the add-in](#test-the-add-in) step.</span></span>

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

### <a name="change-the-font-of-text"></a><span data-ttu-id="85ee0-183">更改文本字体</span><span class="sxs-lookup"><span data-stu-id="85ee0-183">Change the font of text</span></span>

1. <span data-ttu-id="85ee0-184">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-184">Open the file index.html.</span></span>

2. <span data-ttu-id="85ee0-185">在包含 `apply-custom-style` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-185">Below the `div` that contains the `apply-custom-style` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. <span data-ttu-id="85ee0-186">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-186">Open the app.js file.</span></span>

4. <span data-ttu-id="85ee0-187">在向 `apply-custom-style` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-187">Below the line that assigns a click handler to the `apply-custom-style` button, add the following code:</span></span>

    ```js
    $('#change-font').click(changeFont);
    ```

5. <span data-ttu-id="85ee0-188">在 `applyCustomStyle` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-188">Below the `applyCustomStyle` function, add the following function:</span></span>

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. <span data-ttu-id="85ee0-p112">将 `TODO1` 替换为下面的代码。 请注意，此代码使用链接到 `ParagraphCollection.getFirst` 方法的 `Paragraph.getNext` 方法，获取对第二个段落的引用。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p112">Replace `TODO1` with the following code. Note that the code gets a reference to the second paragraph by using the `ParagraphCollection.getFirst` method chained to the `Paragraph.getNext` method.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

### <a name="test-the-add-in"></a><span data-ttu-id="85ee0-191">测试加载项</span><span class="sxs-lookup"><span data-stu-id="85ee0-191">Test the add-in</span></span>

1. <span data-ttu-id="85ee0-192">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl+C 两次，停止正在运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="85ee0-192">In the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server.</span></span> <span data-ttu-id="85ee0-193">否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="85ee0-193">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="85ee0-p114">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样才能看到提示并输入生成命令。 生成后，重启服务器。 接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p114">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, you restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="85ee0-198">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="85ee0-198">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="85ee0-199">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="85ee0-199">Run the command `npm start` to start a web server running on localhost.</span></span>   

4. <span data-ttu-id="85ee0-200">通过关闭任务窗格来重新加载它，再选择“开始”\*\*\*\* 菜单上的“显示任务窗格”\*\*\*\*，以重新打开加载项。</span><span class="sxs-lookup"><span data-stu-id="85ee0-200">Reload the task pane by closing it, and then on the **Home** menu select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="85ee0-201">请确保文档中至少有三个段落。</span><span class="sxs-lookup"><span data-stu-id="85ee0-201">Be sure there are at least three paragraphs in the document.</span></span> <span data-ttu-id="85ee0-202">可以选择“插入段落”\*\*\*\* 三次。</span><span class="sxs-lookup"><span data-stu-id="85ee0-202">You can choose **Insert Paragraph** three times.</span></span> <span data-ttu-id="85ee0-203">*仔细检查文档末尾是否没有空白段落。若有，请予以删除。*</span><span class="sxs-lookup"><span data-stu-id="85ee0-203">*Check carefully that there's no blank paragraph at the end of the document. If there is, delete it.*</span></span>

6. <span data-ttu-id="85ee0-p116">在 Word 中，创建自定义样式“MyCustomStyle”。 其中可以包含所需的任何格式。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p116">In Word, create a custom style named "MyCustomStyle". It can have any formatting that you want.</span></span>

7. <span data-ttu-id="85ee0-p117">选择“应用样式”\*\*\*\* 按钮。 第一个段落将采用嵌入样式“明显参考”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p117">Choose the **Apply Style** button. The first paragraph will be styled with the built-in style **Intense Reference**.</span></span>

8. <span data-ttu-id="85ee0-p118">选择“应用自定义样式”\*\*\*\* 按钮。 最后一个段落将采用自定义样式。 （如果好像什么都没有发生，很可能是因为最后一个段落是空白段落。 如果是这样，请向其中添加某文本。）</span><span class="sxs-lookup"><span data-stu-id="85ee0-p118">Choose the **Apply Custom Style** button. The last paragraph will be styled with your custom style. (If nothing seems to happen, the last paragraph might be blank. If so, add some text to it.)</span></span>

9. <span data-ttu-id="85ee0-p119">选择“更改字体”\*\*\*\* 按钮。 第二个段落的字体更改为 18 磅的粗体 Courier New。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p119">Choose the **Change Font** button. The font of the second paragraph changes to 18 pt., bold, Courier New.</span></span>

    ![Word 教程 - 应用样式和字体](../images/word-tutorial-apply-styles-and-font.png)

## <a name="replace-text-and-insert-text"></a><span data-ttu-id="85ee0-215">替换文本和插入文本</span><span class="sxs-lookup"><span data-stu-id="85ee0-215">Replace text and insert text</span></span>

<span data-ttu-id="85ee0-216">本教程的这一步是，在选定文本区域内外添加文本，并替换选定区域的文本。</span><span class="sxs-lookup"><span data-stu-id="85ee0-216">In this step of the tutorial, you'll add text inside and outside of selected ranges of text, and replace the text of a selected range.</span></span>

### <a name="add-text-inside-a-range"></a><span data-ttu-id="85ee0-217">在区域内添加文本</span><span class="sxs-lookup"><span data-stu-id="85ee0-217">Add text inside a range</span></span>

1. <span data-ttu-id="85ee0-218">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="85ee0-218">Open the project in your code editor.</span></span>

2. <span data-ttu-id="85ee0-219">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-219">Open the file index.html.</span></span>

3. <span data-ttu-id="85ee0-220">在包含 `change-font` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-220">Below the `div` that contains the `change-font` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>
    </div>
    ```

4. <span data-ttu-id="85ee0-221">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-221">Open the app.js file.</span></span>

5. <span data-ttu-id="85ee0-222">在向 `change-font` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-222">Below the line that assigns a click handler to the `change-font` button, add the following code:</span></span>

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. <span data-ttu-id="85ee0-223">在 `changeFont` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-223">Below the `changeFont` function, add the following function:</span></span>

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. <span data-ttu-id="85ee0-p120">将 `TODO1` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-p120">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="85ee0-p121">此方法用于在“即点即用”文本区域末尾插入缩写 ["(C2R)"]。 它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p121">The method is intended to insert the abbreviation ["(C2R)"] into the end of the Range whose text is "Click-to-Run". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="85ee0-228">`Range.insertText` 方法的第一个参数是要插入到 `Range` 对象的字符串。</span><span class="sxs-lookup"><span data-stu-id="85ee0-228">The first parameter of the `Range.insertText` method is the string to insert into the `Range` object.</span></span>

   - <span data-ttu-id="85ee0-p122">第二个参数指定了应在区域中的什么位置插入其他文本。 除了“End”外，其他可用选项包括“Start”、“Before”、“After”和“Replace”。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p122">The second parameter specifies where in the range the additional text should be inserted. Besides "End", the other possible options are "Start", "Before", "After", and "Replace".</span></span> 

   - <span data-ttu-id="85ee0-p123">“End”和“After”的区别在于，“End”在现有区域末尾插入新文本，而“After”则是新建包含字符串的区域，并在现有区域后面插入新区域。 同样，“Start”是在现有区域的开头位置插入文本，而“Before”插入的是新区域。 “Replace”将现有区域文本替换为第一个参数中的字符串。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p123">The difference between "End" and "After" is that "End" inserts the new text inside the end of the existing range, but "After" creates a new range with the string and inserts the new range after the existing range. Similarly, "Start" inserts text inside the beginning of the existing range and "Before" inserts a new range. "Replace" replaces the text of the existing range with the string in the first parameter.</span></span>

   - <span data-ttu-id="85ee0-p124">在本教程之前阶段步骤中，正文对象的 insert\* 方法没有“Before”和“After”选项。 这是因为不能将内容置于文档正文外。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p124">You saw in an earlier stage of the tutorial that the insert\* methods of the body object do not have the "Before" and "After" options. This is because you can't put content outside of the document's body.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

8. <span data-ttu-id="85ee0-p125">在下一部分前，将跳过 `TODO2`。 将 `TODO3` 替换为下面的代码。 此代码类似于在本教程第一阶段中创建的代码，区别在于现在是要在文档末尾（而不是开头）插入新段落。 这一新段落将说明，新文本现属于原始区域。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p125">We'll skip over `TODO2` until the next section. Replace `TODO3` with the following code. This code is similar to the code you created in the first stage of the tutorial, except that now you are inserting a new paragraph at the end of the document instead of at the start. This new paragraph will demonstrate that the new text is now part of the original range.</span></span>

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="85ee0-240">添加代码以将文档属性提取到任务窗格的脚本对象</span><span class="sxs-lookup"><span data-stu-id="85ee0-240">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="85ee0-p126">在本系列教程前面的所有函数中，都是将命令排入队列，以对 Office 文档执行*写入*操作。 每个函数结束时都会调用 `context.sync()` 方法，从而将排入队列的命令发送到文档，以供执行。 不过，在上一步中添加的代码调用的是 `originalRange.text` 属性，这与之前编写的函数明显不同，因为 `originalRange` 对象只是任务窗格脚本中的代理对象。 由于它并不了解文档中区域的实际文本，因此它的 `text` 属性无法有实值。 有必要先从文档中提取区域的文本值，再用它设置 `originalRange.text` 的值。 只有这样才能调用 `originalRange.text`，而又不会导致异常抛出。 此提取过程分为三步：</span><span class="sxs-lookup"><span data-stu-id="85ee0-p126">In all the previous functions in this series of tutorials, you queued commands to *write* to the Office document. Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed. But the code you added in the last step calls the `originalRange.text` property, and this is a significant difference from the earlier functions you wrote, because the `originalRange` object is only a proxy object that exists in your task pane's script. It doesn't know what the actual text of the range in the document is, so its `text` property can't have a real value. It is necessary to first fetch the text value of the range from the document and use it to set the value of `originalRange.text`. Only then can `originalRange.text` be called without causing an exception to be thrown. This fetching process has three steps:</span></span>

   1. <span data-ttu-id="85ee0-248">将命令排入队列，以加载（即提取）代码需要读取的属性。</span><span class="sxs-lookup"><span data-stu-id="85ee0-248">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="85ee0-249">调用上下文对象的 `sync`方法，从而向文档发送已排入队列的命令以供执行，并返回请求获取的信息。</span><span class="sxs-lookup"><span data-stu-id="85ee0-249">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="85ee0-250">由于 `sync` 是异步方法，因此请先确保它已完成，然后代码才能调用已提取的属性。</span><span class="sxs-lookup"><span data-stu-id="85ee0-250">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="85ee0-251">只要代码需要从 Office 文档*读取*信息，就必须完成这些步骤。</span><span class="sxs-lookup"><span data-stu-id="85ee0-251">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="85ee0-252">将 `TODO2` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="85ee0-252">Replace `TODO2` with the following code.</span></span>
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.

            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

2. <span data-ttu-id="85ee0-p127">由于不能在同一取消分支代码路径中有两个 `return` 语句，因此请删除 `Word.run` 末尾的最后一行代码 `return context.sync();`。本教程稍后将添加最后一个新 `context.sync` 语句。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p127">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Word.run`. You'll add a new final `context.sync` later in this tutorial.</span></span>

3. <span data-ttu-id="85ee0-255">剪切并粘贴 `doc.body.insertParagraph` 代码行，以替代 `TODO4`。</span><span class="sxs-lookup"><span data-stu-id="85ee0-255">Cut the `doc.body.insertParagraph` line and paste in place of `TODO4`.</span></span>

4. <span data-ttu-id="85ee0-p128">将 `TODO5` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-p128">Replace `TODO5` with the following code. Note:</span></span>

   - <span data-ttu-id="85ee0-258">将 `sync` 方法传递到 `then` 函数可确保它不会在 `insertParagraph` 逻辑已排入队列前运行。</span><span class="sxs-lookup"><span data-stu-id="85ee0-258">Passing the `sync` method to a `then` function ensures that it does not run until the `insertParagraph` logic has been queued.</span></span>

   - <span data-ttu-id="85ee0-259">由于 `then` 方法调用传递给它的任何函数，并且也不想调用 `sync` 两次，因此请从 context.sync 末尾省略掉“()”。</span><span class="sxs-lookup"><span data-stu-id="85ee0-259">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of context.sync.</span></span>

    ```js
    .then(context.sync);
    ```

<span data-ttu-id="85ee0-260">完成后，整个函数应如下所示：</span><span class="sxs-lookup"><span data-stu-id="85ee0-260">When you're done, the entire function should look like the following:</span></span>

```js
function insertTextIntoRange() {
    Word.run(function (context) {

        var doc = context.document;
        var originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");
                }
            )
            .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
```

### <a name="add-text-between-ranges"></a><span data-ttu-id="85ee0-261">在区域间添加文本</span><span class="sxs-lookup"><span data-stu-id="85ee0-261">Add text between ranges</span></span>

1. <span data-ttu-id="85ee0-262">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-262">Open the file index.html.</span></span>

2. <span data-ttu-id="85ee0-263">在包含 `insert-text-into-range` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-263">Below the `div` that contains the `insert-text-into-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>
    </div>
    ```

3. <span data-ttu-id="85ee0-264">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-264">Open the app.js file.</span></span>

4. <span data-ttu-id="85ee0-265">在向 `insert-text-into-range` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-265">Below the line that assigns a click handler to the `insert-text-into-range` button, add the following code:</span></span>

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. <span data-ttu-id="85ee0-266">在 `insertTextIntoRange` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-266">Below the `insertTextIntoRange` function, add the following function:</span></span>

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a new range before the
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="85ee0-p129">将 `TODO1` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-p129">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="85ee0-p130">此方法用于在文本为“Office 365”的区域前添加文本为“Office 2019”的区域。 它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p130">The method is intended to add a range whose text is "Office 2019, " before the range with text "Office 365". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="85ee0-271">`Range.insertText` 方法的第一个参数是要添加的字符串。</span><span class="sxs-lookup"><span data-stu-id="85ee0-271">The first parameter of the `Range.insertText` method is the string to add.</span></span>

   - <span data-ttu-id="85ee0-p131">第二个参数指定了应在区域中的什么位置插入其他文本。 若要详细了解位置选项，请参阅前面介绍的 `insertTextIntoRange` 函数。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p131">The second parameter specifies where in the range the additional text should be inserted. For more details about the location options, see the previous discussion of the `insertTextIntoRange` function.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ```

7. <span data-ttu-id="85ee0-274">将 `TODO2` 替换为下面的代码。</span><span class="sxs-lookup"><span data-stu-id="85ee0-274">Replace `TODO2` with the following code.</span></span>

     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.

                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has
            //        been queued.
    ```

8. <span data-ttu-id="85ee0-p132">将 `TODO3` 替换为下面的代码。 这一新段落将说明，新文本***不***属于原始选定区域。 原始区域中的文本仍与用户选择它时一样。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p132">Replace `TODO3` with the following code. This new paragraph will demonstrate the fact that the new text is ***not*** part of the original selected range. The original range still has only the text it had when it was selected.</span></span>

    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ```

9. <span data-ttu-id="85ee0-278">将 `TODO4` 替换为下面的代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-278">Replace `TODO4` with the following code:</span></span>

    ```js
    .then(context.sync);
    ```

### <a name="replace-the-text-of-a-range"></a><span data-ttu-id="85ee0-279">替换区域文本</span><span class="sxs-lookup"><span data-stu-id="85ee0-279">Replace the text of a range</span></span>

1. <span data-ttu-id="85ee0-280">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-280">Open the file index.html.</span></span>

2. <span data-ttu-id="85ee0-281">在包含 `insert-text-outside-range` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-281">Below the `div` that contains the `insert-text-outside-range` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>
    </div>
    ```

3. <span data-ttu-id="85ee0-282">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-282">Open the app.js file.</span></span>

4. <span data-ttu-id="85ee0-283">在向 `insert-text-outside-range` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-283">Below the line that assigns a click handler to the `insert-text-outside-range` button, add the following code:</span></span>

    ```js
    $('#replace-text').click(replaceText);
    ```

5. <span data-ttu-id="85ee0-284">在 `insertTextBeforeRange` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-284">Below the `insertTextBeforeRange` function, add the following function:</span></span>

    ```js
    function replaceText() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="85ee0-p133">将 `TODO1` 替换为下面的代码。 请注意，此方法用于将字符串“几个”替换为字符串“许多”。 它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p133">Replace `TODO1` with the following code. Note that the method is intended to replace the string "several" with the string "many". It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85ee0-288">测试加载项</span><span class="sxs-lookup"><span data-stu-id="85ee0-288">Test the add-in</span></span>

1. <span data-ttu-id="85ee0-p134">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p134">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="85ee0-p135">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样才能看到提示并输入生成命令。 生成后，重启服务器。 接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p135">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="85ee0-295">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="85ee0-295">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="85ee0-296">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="85ee0-296">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="85ee0-297">通过关闭任务窗格来重新加载它，再选择“开始”\*\*\*\* 菜单上的“显示任务窗格”\*\*\*\*，以重新打开外接程序。</span><span class="sxs-lookup"><span data-stu-id="85ee0-297">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="85ee0-298">在任务窗格中，选择“插入段落”\*\*\*\*，以确保文档开头有一个段落。</span><span class="sxs-lookup"><span data-stu-id="85ee0-298">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph at the start of the document.</span></span>

6. <span data-ttu-id="85ee0-p136">选择某文本。 选择短语“即点即用”最合适。 *请注意，不要在选定区域的前后添加空格。*</span><span class="sxs-lookup"><span data-stu-id="85ee0-p136">Select some text. Selecting the phrase "Click-to-Run" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

7. <span data-ttu-id="85ee0-p137">选择“插入缩写”\*\*\*\* 按钮。 观察“(C2R)”是否已添加。 此外，还请观察，文档底部是否添加了包含整个扩展文本的新段落，因为新字符串已添加到现有区域中。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p137">Choose the **Insert Abbreviation** button. Note that " (C2R)" is added. Note also that at the bottom of the document a new paragraph is added with the entire expanded text because the new string was added to the existing range.</span></span>

8. <span data-ttu-id="85ee0-p138">选择某文本。 选择短语“Office 365”最合适。 *请注意，不要在选定区域的前后添加空格。*</span><span class="sxs-lookup"><span data-stu-id="85ee0-p138">Select some text. Selecting the phrase "Office 365" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

9. <span data-ttu-id="85ee0-p139">选择“添加版本信息”\*\*\*\* 按钮。 观察是否已在“Office 2016”和“Office 365”之间插入“Office 2019”。 此外，还请观察，文档底部是否添加了仅包含最初选定文本的新段落，因为新字符串已变成新区域，而不是添加到原始区域中。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p139">Choose the **Add Version Info** button. Note that "Office 2019, " is inserted between "Office 2016" and "Office 365". Note also that at the bottom of the document a new paragraph is added but it contains only the originally selected text because the new string became a new range rather than being added to the original range.</span></span>

10. <span data-ttu-id="85ee0-p140">选择某文本。 选择字词“几个”最合适。 *请注意，不要在选定区域的前后添加空格。*</span><span class="sxs-lookup"><span data-stu-id="85ee0-p140">Select some text. Selecting the word "several" will make the most sense. *Be careful not to include the preceding or following space in the selection.*</span></span>

11. <span data-ttu-id="85ee0-p141">选择“更改数量术语”\*\*\*\* 按钮。 观察选定文本是否替换为“多个”。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p141">Choose the **Change Quantity Term** button. Note that "many" replaces the selected text.</span></span>

    ![Word 教程 - 添加和替换文本](../images/word-tutorial-text-replace.png)

## <a name="insert-images-html-and-tables"></a><span data-ttu-id="85ee0-317">插入图像、HTML 和表格</span><span class="sxs-lookup"><span data-stu-id="85ee0-317">Insert images, HTML, and tables</span></span>

<span data-ttu-id="85ee0-318">本教程的这一步是，了解如何在文档中插入图像、HTML 和表格。</span><span class="sxs-lookup"><span data-stu-id="85ee0-318">In this step of the tutorial, you'll learn how to insert images, HTML, and tables into the document.</span></span>

### <a name="insert-an-image"></a><span data-ttu-id="85ee0-319">插入图像</span><span class="sxs-lookup"><span data-stu-id="85ee0-319">Insert an image</span></span>

1. <span data-ttu-id="85ee0-320">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="85ee0-320">Open the project in your code editor.</span></span>

2. <span data-ttu-id="85ee0-321">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-321">Open the file index.html.</span></span>

3. <span data-ttu-id="85ee0-322">在包含 `replace-text` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-322">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. <span data-ttu-id="85ee0-323">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-323">Open the app.js file.</span></span>

5. <span data-ttu-id="85ee0-p142">在文件顶部附近的 use-strict 代码行正下方，添加下面的代码行。 此代码行导入另一个文件中的变量。 此变量是用于编码图像的 Base64 字符串。 若要查看已编码字符串，请打开项目根目录中的 base64Image.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p142">Near the top of the file, just below the use-strict line, add the following line. This line imports a variable from another file. The variable is a base 64 string that encodes an image. To see the encoded string, open the base64Image.js file in the root of the project.</span></span>

    ```js
    import { base64Image } from "./base64Image";
    ```

6. <span data-ttu-id="85ee0-328">在向 `replace-text` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-328">Below the line that assigns a click handler to the `replace-text` button, add the following code:</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

7. <span data-ttu-id="85ee0-329">在 `replaceText` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-329">Below the `replaceText` function, add the following function:</span></span>

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

8. <span data-ttu-id="85ee0-p143">将 `TODO1` 替换为下面的代码。 请注意，此代码行在文档末尾插入 Base64 编码图像。 （`Paragraph` 对象还包含 `insertInlinePictureFromBase64` 方法和其他 `insert*` 方法。 有关示例，请参阅下面的 insertHTML 部分。）</span><span class="sxs-lookup"><span data-stu-id="85ee0-p143">Replace `TODO1` with the following code. Note that this line inserts the base 64 encoded image at the end of the document. (The `Paragraph` object also has an `insertInlinePictureFromBase64` method and other `insert*` methods. See the following insertHTML section for an example.)</span></span>

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a><span data-ttu-id="85ee0-334">插入 HTML</span><span class="sxs-lookup"><span data-stu-id="85ee0-334">Insert HTML</span></span>

1. <span data-ttu-id="85ee0-335">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-335">Open the file index.html.</span></span>

2. <span data-ttu-id="85ee0-336">在包含 `insert-image` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-336">Below the `div` that contains the `insert-image` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. <span data-ttu-id="85ee0-337">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-337">Open the app.js file.</span></span>

4. <span data-ttu-id="85ee0-338">在向 `insert-image` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-338">Below the line that assigns a click handler to the `insert-image` button, add the following code:</span></span>

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. <span data-ttu-id="85ee0-339">在 `insertImage` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-339">Below the `insertImage` function, add the following function:</span></span>

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="85ee0-p144">将 `TODO1` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-p144">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="85ee0-342">第一行代码在文档末尾添加空白段落。</span><span class="sxs-lookup"><span data-stu-id="85ee0-342">The first line adds a blank paragraph to the end of the document.</span></span> 

   - <span data-ttu-id="85ee0-p145">第二行代码在段落末尾插入 HTML 字符串；具体而言是两个段落，一个设置使用 Verdana 字体格式，另一个采用 Word 文档的默认样式。 （如前面的 `insertImage` 方法一样，`context.document.body` 对象还包含 `insert*` 方法。）</span><span class="sxs-lookup"><span data-stu-id="85ee0-p145">The second line inserts a string of HTML at the end of the paragraph; specifically two paragraphs, one formatted with Verdana font, the other with the default styling of the Word document. (As you saw in the `insertImage` method earlier, the `context.document.body` object also has the `insert*` methods.)</span></span>

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a><span data-ttu-id="85ee0-345">插入表格</span><span class="sxs-lookup"><span data-stu-id="85ee0-345">Insert a table</span></span>

1. <span data-ttu-id="85ee0-346">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-346">Open the file index.html.</span></span>

2. <span data-ttu-id="85ee0-347">在包含 `insert-html` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-347">Below the `div` that contains the `insert-html` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. <span data-ttu-id="85ee0-348">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-348">Open the app.js file.</span></span>

4. <span data-ttu-id="85ee0-349">在向 `insert-html` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-349">Below the line that assigns a click handler to the `insert-html` button, add the following code:</span></span>

    ```js
    $('#insert-table').click(insertTable);
    ```

5. <span data-ttu-id="85ee0-350">在 `insertHTML` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-350">Below the `insertHTML` function, add the following function:</span></span>

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="85ee0-p146">将 `TODO1` 替换为下面的代码。 请注意，此代码行先使用 `ParagraphCollection.getFirst` 方法获取对第一个段落的引用，再使用 `Paragraph.getNext` 方法获取对第二个段落的引用。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p146">Replace `TODO1` with the following code. Note that this line uses the `ParagraphCollection.getFirst` method to get a reference ot the first paragraph and then uses the `Paragraph.getNext` method to get a reference to the second paragraph.</span></span>

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. <span data-ttu-id="85ee0-353">将 `TODO2` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="85ee0-353">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="85ee0-354">请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-354">Note:</span></span>

   - <span data-ttu-id="85ee0-355">`insertTable` 方法的前两个参数指定行数和列数。</span><span class="sxs-lookup"><span data-stu-id="85ee0-355">The first two parameters of the `insertTable` method specify the number of rows and columns.</span></span>

   - <span data-ttu-id="85ee0-356">第三个参数指定要在哪里插入表格（在此示例中，是在段落后面插入）。</span><span class="sxs-lookup"><span data-stu-id="85ee0-356">The third parameter specifies where to insert the table, in this case after the paragraph.</span></span>

   - <span data-ttu-id="85ee0-357">第四个参数是用于设置表格单元格值的二维数组。</span><span class="sxs-lookup"><span data-stu-id="85ee0-357">The fourth parameter is a two-dimensional array that sets the values of the table cells.</span></span>

   - <span data-ttu-id="85ee0-358">虽然表格采用普通的默认样式，但 `insertTable` 方法返回的 `Table` 对象包含多个成员，其中部分成员用于设置表格样式。</span><span class="sxs-lookup"><span data-stu-id="85ee0-358">The table will have plain default styling, but the `insertTable` method returns a `Table` object with many members, some of which are used to style the table.</span></span>

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85ee0-359">测试加载项</span><span class="sxs-lookup"><span data-stu-id="85ee0-359">Test the add-in</span></span>

1. <span data-ttu-id="85ee0-p148">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl+C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p148">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="85ee0-p149">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样才能看到提示并输入生成命令。 生成后，重启服务器。 接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p149">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="85ee0-366">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="85ee0-366">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="85ee0-367">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="85ee0-367">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="85ee0-368">通过关闭任务窗格来重新加载它，再选择“开始”\*\*\*\* 菜单上的“显示任务窗格”\*\*\*\*，以重新打开外接程序。</span><span class="sxs-lookup"><span data-stu-id="85ee0-368">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="85ee0-369">在任务窗格中，至少选择“插入段落”\*\*\*\* 三次，以确保文档中有多个段落。</span><span class="sxs-lookup"><span data-stu-id="85ee0-369">In the task pane, choose **Insert Paragraph** at least three times to ensure that there are a few paragraphs in the document.</span></span>

6. <span data-ttu-id="85ee0-370">选择“插入图像”\*\*\*\* 按钮，观察图像是否插入在文档末尾。</span><span class="sxs-lookup"><span data-stu-id="85ee0-370">Choose the **Insert Image** button and note that an image is inserted at the end of the document.</span></span>

7. <span data-ttu-id="85ee0-371">选择“插入 HTML”\*\*\*\* 按钮，观察是否在文档末尾插入了两个段落，第一个段落使用 Verdana 字体。</span><span class="sxs-lookup"><span data-stu-id="85ee0-371">Choose the **Insert HTML** button and note that two paragraphs are inserted at the end of the document, and that the first one has Verdana font.</span></span>

8. <span data-ttu-id="85ee0-372">选择“插入表格”\*\*\*\* 按钮，观察是否在第二个段落后面插入了表格。</span><span class="sxs-lookup"><span data-stu-id="85ee0-372">Choose the **Insert Table** button and note that a table is inserted after the second paragraph.</span></span>

    ![Word 教程 - 插入图像、HTML 和表格](../images/word-tutorial-insert-image-html-table.png)

## <a name="create-and-update-content-controls"></a><span data-ttu-id="85ee0-374">创建和更新内容控件</span><span class="sxs-lookup"><span data-stu-id="85ee0-374">Create and update content controls</span></span>

<span data-ttu-id="85ee0-375">本教程的这一步是，了解如何在文档中创建格式文本内容控件，以及如何插入和替换控件的内容。</span><span class="sxs-lookup"><span data-stu-id="85ee0-375">In this step of the tutorial, you'll learn how to create Rich Text content controls in the document, and then how to insert and replace content in the controls.</span></span>

> [!NOTE]
> <span data-ttu-id="85ee0-376">虽然可通过 UI 添加到 Word 文档的内容控件有好几种，但目前 Word.js 仅支持格式文本内容控件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-376">There are several types of content controls that can be added to a Word document through the UI; but currently only Rich Text content controls are supported by Word.js.</span></span>
>
> <span data-ttu-id="85ee0-p150">开始执行本教程的这一步之前，建议通过 Word UI 创建和控制格式文本内容控件，以便熟悉此类控件及其属性。 有关详细信息，请参阅[在 Word 中创建用户填写或打印的表单](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b)。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p150">Before you start this step of the tutorial, we recommend that you create and manipulate Rich Text content controls through the Word UI, so you can be familiar with the controls and their properties. For details, see [Create forms that users complete or print in Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).</span></span>

### <a name="create-a-content-control"></a><span data-ttu-id="85ee0-379">创建内容控件</span><span class="sxs-lookup"><span data-stu-id="85ee0-379">Create a content control</span></span>

1. <span data-ttu-id="85ee0-380">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="85ee0-380">Open the project in your code editor.</span></span>

2. <span data-ttu-id="85ee0-381">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-381">Open the file index.html.</span></span>

3. <span data-ttu-id="85ee0-382">在包含 `replace-text` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-382">Below the `div` that contains the `replace-text` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-content-control">Create Content Control</button>
    </div>
    ```

4. <span data-ttu-id="85ee0-383">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-383">Open the app.js file.</span></span>

5. <span data-ttu-id="85ee0-384">在向 `insert-table` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-384">Below the line that assigns a click handler to the `insert-table` button, add the following code:</span></span>

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. <span data-ttu-id="85ee0-385">在 `insertTable` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-385">Below the `insertTable` function, add the following function:</span></span>

    ```js
    function createContentControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

7. <span data-ttu-id="85ee0-p151">将 `TODO1` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-p151">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="85ee0-p152">此代码用于在内容控件中包装短语“Office 365”。 它做了一个简化假设，即存在字符串，且用户已选择它。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p152">This code is intended to wrap the phrase "Office 365" in a content control. It makes a simplifying assumption that the string is present and the user has selected it.</span></span>

   - <span data-ttu-id="85ee0-390">`ContentControl.title` 属性指定内容控件的可见标题。</span><span class="sxs-lookup"><span data-stu-id="85ee0-390">The `ContentControl.title` property specifies the visible title of the content control.</span></span>

   - <span data-ttu-id="85ee0-391">`ContentControl.tag` 属性指定标记，可用于通过 `ContentControlCollection.getByTag` 方法获取对内容控件的引用，将用于稍后出现的函数。</span><span class="sxs-lookup"><span data-stu-id="85ee0-391">The `ContentControl.tag` property specifies an tag that can be used to get a reference to a content control using the `ContentControlCollection.getByTag` method, which you'll use in a later function.</span></span>

   - <span data-ttu-id="85ee0-p153">`ContentControl.appearance` 属性指定控件的外观。 使用值“Tags”表示，控件包装在开始标记和结束标记中，且开始标记包含内容控件标题。 其他可取值包括“BoundingBox”和“None”。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p153">The `ContentControl.appearance` property specifies the visual look of the control. Using the value "Tags" means that the control will be wrapped in opening and closing tags, and the opening tag will have the content control's title. Other possible values are "BoundingBox" and "None".</span></span>

   - <span data-ttu-id="85ee0-395">`ContentControl.color` 属性指定标记颜色或边界框的边框。</span><span class="sxs-lookup"><span data-stu-id="85ee0-395">The `ContentControl.color` property specifies the color of the tags or the border of the bounding box.</span></span>

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a><span data-ttu-id="85ee0-396">替换内容控件的内容</span><span class="sxs-lookup"><span data-stu-id="85ee0-396">Replace the content of the content control</span></span>

1. <span data-ttu-id="85ee0-397">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="85ee0-397">Open the file index.html.</span></span>

2. <span data-ttu-id="85ee0-398">在包含 `create-content-control` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="85ee0-398">Below the `div` that contains the `create-content-control` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>
    </div>
    ```

3. <span data-ttu-id="85ee0-399">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="85ee0-399">Open the app.js file.</span></span>

4. <span data-ttu-id="85ee0-400">在向 `create-content-control` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="85ee0-400">Below the line that assigns a click handler to the `create-content-control` button, add the following code:</span></span>

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. <span data-ttu-id="85ee0-401">在 `createContentControl` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="85ee0-401">Below the `createContentControl` function, add the following function:</span></span>

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ```

6. <span data-ttu-id="85ee0-p154">将 `TODO1` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="85ee0-p154">Replace `TODO1` with the following code. Note:</span></span>

    - <span data-ttu-id="85ee0-404">`ContentControlCollection.getByTag` 方法将返回指定标记的所有内容控件的 `ContentControlCollection`。</span><span class="sxs-lookup"><span data-stu-id="85ee0-404">The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag.</span></span> <span data-ttu-id="85ee0-405">我们使用 `getFirst` 来获取对所需控件的引用。</span><span class="sxs-lookup"><span data-stu-id="85ee0-405">We use `getFirst` to get a reference to the desired control.</span></span>

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="85ee0-406">测试外接程序</span><span class="sxs-lookup"><span data-stu-id="85ee0-406">Test the add-in</span></span>

1. <span data-ttu-id="85ee0-p156">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl+C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p156">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl+C twice to stop the running web server. Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="85ee0-p157">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样才能看到提示并输入生成命令。 生成后，重启服务器。 接下来的几步操作就是在执行此过程。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p157">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect. In order to do this, you need to kill the server process so that the prompt appears and you can enter the build command. After the build, restart the server. The next few steps carry out this process.</span></span>

2. <span data-ttu-id="85ee0-413">运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="85ee0-413">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by all the hosts where Office Add-ins can run.</span></span>

3. <span data-ttu-id="85ee0-414">运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="85ee0-414">Run the command `npm start` to start a web server running on localhost.</span></span>

4. <span data-ttu-id="85ee0-415">通过关闭任务窗格来重新加载它，再选择“开始”\*\*\*\* 菜单上的“显示任务窗格”\*\*\*\*，以重新打开外接程序。</span><span class="sxs-lookup"><span data-stu-id="85ee0-415">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>

5. <span data-ttu-id="85ee0-416">在任务窗格中，选择“插入段落”\*\*\*\*，以确保文档顶部有包含“Office 365”的段落。</span><span class="sxs-lookup"><span data-stu-id="85ee0-416">In the task pane, choose **Insert Paragraph** to ensure that there is a paragraph with "Office 365" at the top of the document.</span></span>

6. <span data-ttu-id="85ee0-p158">选择刚刚添加的段落中的短语“Office 365”，再选择“创建内容控件”\*\*\*\* 按钮。 观察此短语是否包装在标签为“服务名称”的标记中。</span><span class="sxs-lookup"><span data-stu-id="85ee0-p158">Select the phrase "Office 365" in the paragraph you just added, and then choose the **Create Content Control** button. Note that the phrase is wrapped in tags labelled "Service Name".</span></span>

7. <span data-ttu-id="85ee0-419">选择“重命名服务”\*\*\*\* 按钮，并观察内容控件的文本是否变成“Fabrikam Online Productivity Suite”。</span><span class="sxs-lookup"><span data-stu-id="85ee0-419">Choose the **Rename Service** button and note that the text of the content control changes to "Fabrikam Online Productivity Suite".</span></span>

    ![Word 教程 - 创建内容控件并更改其文本](../images/word-tutorial-content-control.png)

## <a name="next-steps"></a><span data-ttu-id="85ee0-421">后续步骤</span><span class="sxs-lookup"><span data-stu-id="85ee0-421">Next steps</span></span>

<span data-ttu-id="85ee0-422">在本教程中，你已创建 Word 任务窗格加载项，用于在 Word 文档中插入和替换文本、图像和其他内容。</span><span class="sxs-lookup"><span data-stu-id="85ee0-422">In this tutorial, you've created a Word task pane add-in that inserts and replaces text, images, and other content in a Word document.</span></span> <span data-ttu-id="85ee0-423">若要了解有关构建 Word 加载项的详细信息，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="85ee0-423">To learn more about building Word add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="85ee0-424">Word 加载项概述</span><span class="sxs-lookup"><span data-stu-id="85ee0-424">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
