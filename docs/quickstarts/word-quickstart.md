---
title: 生成首个 Word 任务窗格加载项
description: ''
ms.date: 05/08/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: f0fda0c7dcdebdc1fd1b6daf4e35c1794a56e950
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952262"
---
# <a name="build-your-first-word-task-pane-add-in"></a><span data-ttu-id="d1e11-102">生成首个 Word 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="d1e11-102">Build your first PowerPoint task pane add-in</span></span>

<span data-ttu-id="d1e11-103">_适用于：Windows 版 Word 2016 或更高版本、Word for iPad、Word for Mac_</span><span class="sxs-lookup"><span data-stu-id="d1e11-103">_Applies to: Word 2016 or later for Windows, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="d1e11-104">本文将逐步介绍如何生成 Word 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="d1e11-104">In this article, you'll walk through the process of building a PowerPoint task pane add-in.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="d1e11-105">创建加载项</span><span class="sxs-lookup"><span data-stu-id="d1e11-105">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="d1e11-106">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="d1e11-106">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="d1e11-107">先决条件</span><span class="sxs-lookup"><span data-stu-id="d1e11-107">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="d1e11-108">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="d1e11-108">Create the add-in project</span></span>

1. <span data-ttu-id="d1e11-109">在 Visual Studio 菜单栏中，依次选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="d1e11-109">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="d1e11-110">在“Visual C#”\*\*\*\* 或“Visual Basic”\*\*\*\* 下的项目类型列表中，展开“Office/SharePoint”\*\*\*\*，选择“加载项”\*\*\*\*，再选择“Word Web 加载项”\*\*\*\* 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="d1e11-110">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="d1e11-111">命名此项目，再选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="d1e11-111">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="d1e11-p101">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="d1e11-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="d1e11-114">探索 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="d1e11-114">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="d1e11-115">更新代码</span><span class="sxs-lookup"><span data-stu-id="d1e11-115">Update the code</span></span>

1. <span data-ttu-id="d1e11-p102">**Home.html** 指定在加载项的任务窗格中呈现的 HTML。 在 **Home.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="d1e11-p102">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="d1e11-p103">打开 Web 应用项目根目录中的文件“Home.js”\*\*\*\*。 此文件指定加载项脚本。 将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="d1e11-p103">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
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

3. <span data-ttu-id="d1e11-p104">打开 Web 应用项目根目录中的文件“Home.css”\*\*\*\*。 此文件指定加载项自定义样式。 将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="d1e11-p104">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="d1e11-124">更新清单</span><span class="sxs-lookup"><span data-stu-id="d1e11-124">Update the manifest</span></span>

1. <span data-ttu-id="d1e11-125">打开加载项项目中的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="d1e11-125">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="d1e11-126">此文件定义的是加载项设置和功能。</span><span class="sxs-lookup"><span data-stu-id="d1e11-126">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="d1e11-p106">`ProviderName` 元素具有占位符值。 将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="d1e11-p106">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="d1e11-129">`DisplayName` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="d1e11-129">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="d1e11-130">将它替换为“My Office Add-in”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="d1e11-130">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="d1e11-131">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="d1e11-131">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="d1e11-132">将它替换为“A task pane add-in for Word”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="d1e11-132">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="d1e11-133">保存文件。</span><span class="sxs-lookup"><span data-stu-id="d1e11-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="d1e11-134">试用</span><span class="sxs-lookup"><span data-stu-id="d1e11-134">Try it out</span></span>

1. <span data-ttu-id="d1e11-p109">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 Word，以测试新建的 Word 加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="d1e11-p109">Using Visual Studio, test the newly created Word add-in by pressing **F5** or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="d1e11-137">在 Word 中，依次选择“开始”\*\*\*\* 选项卡和功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="d1e11-137">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="d1e11-138">（如果使用的是 Office 的一次性购买版本，而不是 Office 365 版本，那么自定义按钮不受支持。</span><span class="sxs-lookup"><span data-stu-id="d1e11-138">(If you are using the one-time purchase version of Office, instead of the Office 365 version, then custom buttons are not supported.</span></span> <span data-ttu-id="d1e11-139">相反，任务窗格将立即打开。）</span><span class="sxs-lookup"><span data-stu-id="d1e11-139">Instead, the task pane will open immediately.)</span></span>

    ![突出显示了“显示任务窗格”按钮的 Word 应用屏幕截图](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="d1e11-141">选择任务窗格中的任意按钮，将样本文字添加到文档。</span><span class="sxs-lookup"><span data-stu-id="d1e11-141">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![加载了样本加载项的 Word 应用的屏幕截图](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="d1e11-143">任意编辑器</span><span class="sxs-lookup"><span data-stu-id="d1e11-143">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="d1e11-144">先决条件</span><span class="sxs-lookup"><span data-stu-id="d1e11-144">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="d1e11-145">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="d1e11-145">Create the add-in project</span></span>

1. <span data-ttu-id="d1e11-146">使用 Yeoman 生成器创建 Word 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="d1e11-146">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="d1e11-147">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="d1e11-147">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="d1e11-148">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="d1e11-148">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
    - <span data-ttu-id="d1e11-149">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="d1e11-149">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="d1e11-150">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="d1e11-150">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="d1e11-151">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="d1e11-151">**Which Office client application would you like to support?**</span></span> `Word`

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-word.png)

    <span data-ttu-id="d1e11-153">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="d1e11-153">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="d1e11-154">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="d1e11-154">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a><span data-ttu-id="d1e11-155">浏览项目</span><span class="sxs-lookup"><span data-stu-id="d1e11-155">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="d1e11-156">试用</span><span class="sxs-lookup"><span data-stu-id="d1e11-156">Try it out</span></span>

1. <span data-ttu-id="d1e11-157">启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="d1e11-157">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="d1e11-158">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="d1e11-158">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="d1e11-159">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="d1e11-159">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> 

    - <span data-ttu-id="d1e11-160">若要在 Word 中测试加载项，请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="d1e11-160">To test your add-in in PowerPoint, run the following command.</span></span> <span data-ttu-id="d1e11-161">运行此命令时，本地 Web 服务器将启动，Word 将打开且加载项已载入。</span><span class="sxs-lookup"><span data-stu-id="d1e11-161">When you run this command, the local web server will start and PowerPoint will open with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="d1e11-162">若要在 Word Online 中测试加载项，请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="d1e11-162">To test your add-in in PowerPoint Online, run the following command.</span></span> <span data-ttu-id="d1e11-163">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="d1e11-163">When you run this command, the local web server will start.</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="d1e11-164">若要使用加载项，请在 Word Online 中打开新的文档，然后按照[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)中的说明旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="d1e11-164">To use your add-in, open a new document in PowerPoint Online and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online).</span></span>

2. <span data-ttu-id="d1e11-165">在 Word 中，打开新的文档，依次选择“**主页**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="d1e11-165">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![突出显示了“显示任务窗格”按钮的 Word 应用程序屏幕截图](../images/word-quickstart-addin-2b.png)

3. <span data-ttu-id="d1e11-167">在任务窗格底部，选择“**运行**”链接，以将文本“Hello World”以蓝色字体添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="d1e11-167">At the bottom of the task pane, choose the **Run** link to insert the text "Hello World" into the current slide.</span></span>

    ![加载了任务窗格加载项的 Word 应用程序的屏幕截图](../images/word-quickstart-addin-1c.png)

---

## <a name="next-steps"></a><span data-ttu-id="d1e11-169">后续步骤</span><span class="sxs-lookup"><span data-stu-id="d1e11-169">Next steps</span></span>

<span data-ttu-id="d1e11-170">恭喜！已成功创建 Word 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="d1e11-170">Congratulations, you've successfully created a PowerPoint task pane add-in!</span></span> <span data-ttu-id="d1e11-171">接下来，请详细了解 Word 加载项功能，并跟着 Word 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="d1e11-171">Next, learn more about the capabilities of a Word add-in and build a more complex add-in by following along with the Word add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="d1e11-172">Word 加载项教程</span><span class="sxs-lookup"><span data-stu-id="d1e11-172">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="d1e11-173">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d1e11-173">See also</span></span>

* [<span data-ttu-id="d1e11-174">Word 加载项概述</span><span class="sxs-lookup"><span data-stu-id="d1e11-174">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* <span data-ttu-id="d1e11-175">
  [Word 加载项代码示例](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)</span><span class="sxs-lookup"><span data-stu-id="d1e11-175">[Word add-in code samples](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)</span></span>
* [<span data-ttu-id="d1e11-176">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="d1e11-176">Word JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)
