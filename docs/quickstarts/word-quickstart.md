---
title: 生成首个 Word 任务窗格加载项
description: 了解如何使用 Office JS API 生成简单的 Word 任务窗格加载项。
ms.date: 07/17/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 5b65d20a10b98dc3a4ba1e95c4ef52ff91647e97
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308041"
---
# <a name="build-your-first-word-task-pane-add-in"></a><span data-ttu-id="68ff3-103">生成首个 Word 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="68ff3-103">Build your first Word task pane add-in</span></span>

<span data-ttu-id="68ff3-104">_适用于：Windows 版 Word 2016 或更高版本、iPad 版 Word 和 Mac 版 Word_</span><span class="sxs-lookup"><span data-stu-id="68ff3-104">_Applies to: Word 2016 or later on Windows, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="68ff3-105">本文将逐步介绍如何生成 Word 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="68ff3-105">In this article, you'll walk through the process of building a Word task pane add-in.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="68ff3-106">创建加载项</span><span class="sxs-lookup"><span data-stu-id="68ff3-106">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generatortabyeomangenerator"></a>[<span data-ttu-id="68ff3-107">Yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="68ff3-107">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="68ff3-108">先决条件</span><span class="sxs-lookup"><span data-stu-id="68ff3-108">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="68ff3-109">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="68ff3-109">Create the add-in project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="68ff3-110">使用 Yeoman 生成器创建 Word 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="68ff3-110">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="68ff3-111">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="68ff3-111">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="68ff3-112">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="68ff3-112">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="68ff3-113">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="68ff3-113">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="68ff3-114">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="68ff3-114">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="68ff3-115">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="68ff3-115">**Which Office client application would you like to support?**</span></span> `Word`

<span data-ttu-id="68ff3-116">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="68ff3-116">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

### <a name="explore-the-project"></a><span data-ttu-id="68ff3-117">浏览项目</span><span class="sxs-lookup"><span data-stu-id="68ff3-117">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="68ff3-118">试用</span><span class="sxs-lookup"><span data-stu-id="68ff3-118">Try it out</span></span>

1. <span data-ttu-id="68ff3-119">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="68ff3-119">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. <span data-ttu-id="68ff3-120">完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="68ff3-120">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="68ff3-121">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="68ff3-121">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="68ff3-122">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="68ff3-122">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="68ff3-123">如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="68ff3-123">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="68ff3-124">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="68ff3-124">When you run this command, the local web server will start.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="68ff3-125">若要在 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="68ff3-125">To test your add-in in Word, run the following command in the root directory of your project.</span></span> <span data-ttu-id="68ff3-126">这将启动本地的 Web 服务器（如果尚未运行的话），并使用加载的加载项打开 Word。</span><span class="sxs-lookup"><span data-stu-id="68ff3-126">This starts the local web server (if it's not already running) and opens Word with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="68ff3-127">若要在浏览器版 Word 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="68ff3-127">To test your add-in in Word on a browser, run the following command in the root directory of your project.</span></span> <span data-ttu-id="68ff3-128">如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。</span><span class="sxs-lookup"><span data-stu-id="68ff3-128">When you run this command, the local web server will start.</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="68ff3-129">若要使用加载项，请在 Word 网页版中打开新的文档，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="68ff3-129">To use your add-in, open a new document in Word Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

3. <span data-ttu-id="68ff3-130">在 Word 中，打开新的文档，依次选择“**主页**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="68ff3-130">In Word, open a new document, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![突出显示了“显示任务窗格”按钮的 Word 应用程序屏幕截图](../images/word-quickstart-addin-2b.png)

4. <span data-ttu-id="68ff3-132">在任务窗格底部，选择“**运行**”链接，以将文本“Hello World”以蓝色字体添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="68ff3-132">At the bottom of the task pane, choose the **Run** link to add the text "Hello World" to the document in blue font.</span></span>

    ![加载了任务窗格加载项的 Word 应用程序的屏幕截图](../images/word-quickstart-addin-1c.png)

# <a name="visual-studiotabvisualstudio"></a>[<span data-ttu-id="68ff3-134">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="68ff3-134">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="68ff3-135">先决条件</span><span class="sxs-lookup"><span data-stu-id="68ff3-135">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="68ff3-136">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="68ff3-136">Create the add-in project</span></span>

1. <span data-ttu-id="68ff3-137">在 Visual Studio 菜单栏中，依次选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="68ff3-137">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="68ff3-138">在“Visual C#”\*\*\*\* 或“Visual Basic”\*\*\*\* 下的项目类型列表中，展开“Office/SharePoint”\*\*\*\*，选择“加载项”\*\*\*\*，再选择“Word Web 加载项”\*\*\*\* 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="68ff3-138">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="68ff3-139">命名此项目，再选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="68ff3-139">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="68ff3-p106">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="68ff3-p106">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="68ff3-142">探索 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="68ff3-142">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="68ff3-143">更新代码</span><span class="sxs-lookup"><span data-stu-id="68ff3-143">Update the code</span></span>

1. <span data-ttu-id="68ff3-p107">**Home.html** 指定在加载项的任务窗格中呈现的 HTML。 在 **Home.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="68ff3-p107">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion"/>
    </body>
    ```

2. <span data-ttu-id="68ff3-p108">打开 Web 应用项目根目录中的文件“Home.js”\*\*\*\*。 此文件指定加载项脚本。 将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="68ff3-p108">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or later.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or later.');
                }
            });
        });

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="68ff3-p109">打开 Web 应用项目根目录中的文件“Home.css”\*\*\*\*。 此文件指定加载项自定义样式。 将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="68ff3-p109">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto;
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="68ff3-152">更新清单</span><span class="sxs-lookup"><span data-stu-id="68ff3-152">Update the manifest</span></span>

1. <span data-ttu-id="68ff3-153">打开加载项项目中的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="68ff3-153">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="68ff3-154">此文件定义的是加载项设置和功能。</span><span class="sxs-lookup"><span data-stu-id="68ff3-154">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="68ff3-p111">`ProviderName` 元素具有占位符值。 将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="68ff3-p111">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="68ff3-157">`DisplayName` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="68ff3-157">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="68ff3-158">将它替换为“My Office Add-in”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="68ff3-158">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="68ff3-159">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="68ff3-159">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="68ff3-160">将它替换为“A task pane add-in for Word”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="68ff3-160">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="68ff3-161">保存文件。</span><span class="sxs-lookup"><span data-stu-id="68ff3-161">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="68ff3-162">试用</span><span class="sxs-lookup"><span data-stu-id="68ff3-162">Try it out</span></span>

1. <span data-ttu-id="68ff3-p114">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 Word，以测试新建的 Word 加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="68ff3-p114">Using Visual Studio, test the newly created Word add-in by pressing **F5** or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="68ff3-165">在 Word 中，依次选择“开始”\*\*\*\* 选项卡和功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="68ff3-165">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="68ff3-166">（如果使用的是 Office 的一次性购买版本，而不是 Office 365 版本，那么自定义按钮不受支持。</span><span class="sxs-lookup"><span data-stu-id="68ff3-166">(If you are using the one-time purchase version of Office, instead of the Office 365 version, then custom buttons are not supported.</span></span> <span data-ttu-id="68ff3-167">相反，任务窗格将立即打开。）</span><span class="sxs-lookup"><span data-stu-id="68ff3-167">Instead, the task pane will open immediately.)</span></span>

    ![突出显示了“显示任务窗格”按钮的 Word 应用屏幕截图](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="68ff3-169">选择任务窗格中的任意按钮，将样本文字添加到文档。</span><span class="sxs-lookup"><span data-stu-id="68ff3-169">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![加载了样本加载项的 Word 应用的屏幕截图](../images/word-quickstart-addin-1b.png)

---

## <a name="next-steps"></a><span data-ttu-id="68ff3-171">后续步骤</span><span class="sxs-lookup"><span data-stu-id="68ff3-171">Next steps</span></span>

<span data-ttu-id="68ff3-172">恭喜！已成功创建 Word 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="68ff3-172">Congratulations, you've successfully created a Word task pane add-in!</span></span> <span data-ttu-id="68ff3-173">接下来，请详细了解 Word 加载项功能，并跟着 Word 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="68ff3-173">Next, learn more about the capabilities of a Word add-in and build a more complex add-in by following along with the Word add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="68ff3-174">Word 加载项教程</span><span class="sxs-lookup"><span data-stu-id="68ff3-174">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="68ff3-175">另请参阅</span><span class="sxs-lookup"><span data-stu-id="68ff3-175">See also</span></span>

* [<span data-ttu-id="68ff3-176">Word 加载项概述</span><span class="sxs-lookup"><span data-stu-id="68ff3-176">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="68ff3-177">Word 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="68ff3-177">Word add-in code samples</span></span>](https://developer.microsoft.com/zh-CN/office/gallery/?filterBy=Samples,Word)
* [<span data-ttu-id="68ff3-178">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="68ff3-178">Word JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
