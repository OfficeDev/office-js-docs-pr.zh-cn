# <a name="build-your-first-word-add-in"></a><span data-ttu-id="0099d-101">构建您的第一个 Word 外接程序</span><span class="sxs-lookup"><span data-stu-id="0099d-101">Build your first Word add-in</span></span>

<span data-ttu-id="0099d-102">_适用于：Word 2016、Word for iPad、Word for Mac_</span><span class="sxs-lookup"><span data-stu-id="0099d-102">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="0099d-103">本文将逐步介绍如何使用 jQuery 和 Word JavaScript API 生成 Word 加载项。</span><span class="sxs-lookup"><span data-stu-id="0099d-103">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="0099d-104">创建加载项</span><span class="sxs-lookup"><span data-stu-id="0099d-104">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="0099d-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="0099d-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="0099d-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="0099d-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="0099d-107">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="0099d-107">Create the add-in project</span></span>

1. <span data-ttu-id="0099d-108">在 Visual Studio 菜单栏中，依次选择“文件”**** > “新建”**** > “项目”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="0099d-109">在“Visual C#”**** 或“Visual Basic”**** 下的项目类型列表中，展开“Office/SharePoint”****，选择“加载项”****，再选择“Word Web 加载项”**** 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="0099d-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="0099d-110">命名此项目，再选择“确定”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="0099d-p101">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”**** 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="0099d-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="0099d-113">探索 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="0099d-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="0099d-114">更新代码</span><span class="sxs-lookup"><span data-stu-id="0099d-114">Update the code</span></span>

1. <span data-ttu-id="0099d-115">**Home.html** 指定在加载项的任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="0099d-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="0099d-116">在 **Home.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="0099d-117">打开 Web 应用项目根目录中的文件“Home.js”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="0099d-118">此文件指定加载项脚本。</span><span class="sxs-lookup"><span data-stu-id="0099d-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="0099d-119">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-119">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

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

3. <span data-ttu-id="0099d-120">打开 Web 应用项目根目录中的文件“Home.css”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="0099d-121">此文件指定加载项自定义样式。</span><span class="sxs-lookup"><span data-stu-id="0099d-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="0099d-122">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-122">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="0099d-123">更新清单</span><span class="sxs-lookup"><span data-stu-id="0099d-123">Update the manifest</span></span>

1. <span data-ttu-id="0099d-124">打开加载项项目中的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-124">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="0099d-125">此文件定义的是加载项设置和功能。</span><span class="sxs-lookup"><span data-stu-id="0099d-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="0099d-126">元素具有占位符值。`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="0099d-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="0099d-127">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="0099d-127">Replace it with your name.</span></span>

3. <span data-ttu-id="0099d-128">元素的 `DefaultValue` 属性有占位符。`DisplayName`</span><span class="sxs-lookup"><span data-stu-id="0099d-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="0099d-129">将它替换为“My Office Add-in”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="0099d-130">元素的 `DefaultValue` 属性有占位符。`Description`</span><span class="sxs-lookup"><span data-stu-id="0099d-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="0099d-131">将它替换为“A task pane add-in for Word”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-131">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="0099d-132">保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="0099d-133">试用</span><span class="sxs-lookup"><span data-stu-id="0099d-133">Try it out</span></span>

1. <span data-ttu-id="0099d-p109">使用 Visual Studio 的同时，按 F5 或选择“开始”**** 按钮启动 Word，以测试新建的 Word 加载项，功能区中显示有“显示任务窗格”**** 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="0099d-p109">Using Visual Studio, test the newly created Word add-in by pressing F5 or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="0099d-136">在 Word 中，依次选择“主页”**** 选项卡和功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0099d-136">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![突出显示了“显示任务窗格”按钮的 Word 应用屏幕截图](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="0099d-138">选择任务窗格中的任意按钮，将样本文字添加到文档。</span><span class="sxs-lookup"><span data-stu-id="0099d-138">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![加载了样本加载项的 Word 应用的屏幕截图](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="0099d-140">任意编辑器</span><span class="sxs-lookup"><span data-stu-id="0099d-140">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="0099d-141">先决条件</span><span class="sxs-lookup"><span data-stu-id="0099d-141">Prerequisites</span></span>

- [<span data-ttu-id="0099d-142">Node.js</span><span class="sxs-lookup"><span data-stu-id="0099d-142">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="0099d-143">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="0099d-143">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="0099d-144">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="0099d-144">Create the add-in project</span></span>

1. <span data-ttu-id="0099d-145">在本地驱动器上创建文件夹，并将它命名为“`my-word-addin`”。</span><span class="sxs-lookup"><span data-stu-id="0099d-145">Create a folder on your local drive and name it `my-word-addin`.</span></span> <span data-ttu-id="0099d-146">将在其中创建外接程序文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-146">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="0099d-147">转到新文件夹。</span><span class="sxs-lookup"><span data-stu-id="0099d-147">Navigate to your new folder.</span></span>

    ```bash
    cd my-word-addin
    ```

3. <span data-ttu-id="0099d-148">使用 Yeoman 生成器创建 Word 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="0099d-148">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="0099d-149">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="0099d-149">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="0099d-150">**选择一个项目类型：** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="0099d-150">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="0099d-151">**选择一个脚本类型：** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="0099d-151">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="0099d-152">**要如何命名加载项?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="0099d-152">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="0099d-153">**要支持哪一个 Office 客户端应用?:** `Word`</span><span class="sxs-lookup"><span data-stu-id="0099d-153">**Which Office client application would you like to support?:** `Word`</span></span>

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-word-jquery.png)
    
    <span data-ttu-id="0099d-155">完成向导后，生成器将创建项目并安装 Node 支持组件。</span><span class="sxs-lookup"><span data-stu-id="0099d-155">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

### <a name="update-the-code"></a><span data-ttu-id="0099d-156">更新代码</span><span class="sxs-lookup"><span data-stu-id="0099d-156">Update the code</span></span>

1. <span data-ttu-id="0099d-157">在代码编辑器中，打开项目根目录中的“index.html”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-157">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="0099d-158">此文件包含在加载项任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="0099d-158">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="0099d-159">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-159">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="0099d-160">此加载项将显示三个按钮，在用户选择其中任一按钮后，样本文字就会添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="0099d-160">This add-in will display three buttons and when any of the buttons are chosen, boilerplate text will be added to the document.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>Boilerplate text app</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="app.js" type="text/javascript"></script>
            <link href="app.css" rel="stylesheet" type="text/css" />
        </head>
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
    </html>
    ```

2. <span data-ttu-id="0099d-161">打开文件“app.js”****，以指定加载项脚本。</span><span class="sxs-lookup"><span data-stu-id="0099d-161">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="0099d-162">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-162">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="0099d-163">此脚本包含初始化代码以及用于更改 Word 文档的代码（具体方法是通过选择某个按钮将文本插入文档）。</span><span class="sxs-lookup"><span data-stu-id="0099d-163">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

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

3. <span data-ttu-id="0099d-164">打开项目根目录中的文件“app.css”****，以指定加载项自定义样式。</span><span class="sxs-lookup"><span data-stu-id="0099d-164">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="0099d-165">将整个内容替换为以下内容，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-165">Replace the entire contents with the following and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="0099d-166">更新清单</span><span class="sxs-lookup"><span data-stu-id="0099d-166">Update the manifest</span></span>

1. <span data-ttu-id="0099d-167">打开文件“my-office-add-in-manifest.xml”****，以定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="0099d-167">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="0099d-168">元素具有占位符值。`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="0099d-168">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="0099d-169">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="0099d-169">Replace it with your name.</span></span>

3. <span data-ttu-id="0099d-170">元素的 `DefaultValue` 属性有占位符。`Description`</span><span class="sxs-lookup"><span data-stu-id="0099d-170">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="0099d-171">将它替换为“A task pane add-in for Word”****。</span><span class="sxs-lookup"><span data-stu-id="0099d-171">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="0099d-172">保存文件。</span><span class="sxs-lookup"><span data-stu-id="0099d-172">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="0099d-173">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="0099d-173">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="0099d-174">试用</span><span class="sxs-lookup"><span data-stu-id="0099d-174">Try it out</span></span>

1. <span data-ttu-id="0099d-175">请按照运行加载项所用平台对应的说明操作，以在 Word 中旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="0099d-175">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="0099d-176">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="0099d-176">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="0099d-177">Word Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="0099d-177">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="0099d-178">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="0099d-178">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="0099d-179">在 Word 中，依次选择“主页”**** 选项卡和功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0099d-179">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![突出显示了“显示任务窗格”按钮的 Word 应用屏幕截图](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="0099d-181">选择任务窗格中的任意按钮，将样本文字添加到文档。</span><span class="sxs-lookup"><span data-stu-id="0099d-181">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![加载了样本加载项的 Word 应用的屏幕截图](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="0099d-183">后续步骤</span><span class="sxs-lookup"><span data-stu-id="0099d-183">Next steps</span></span>

<span data-ttu-id="0099d-184">恭喜！已使用 jQuery 成功创建 Word 加载项！</span><span class="sxs-lookup"><span data-stu-id="0099d-184">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="0099d-185">接下来，请详细了解 Word 加载项的功能，并跟着 Word 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="0099d-185">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0099d-186">Word 外接程序教程</span><span class="sxs-lookup"><span data-stu-id="0099d-186">Word add-in tutorial</span></span>](../tutorials/word-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="0099d-187">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0099d-187">See also</span></span>

* [<span data-ttu-id="0099d-188">Word 加载项概述</span><span class="sxs-lookup"><span data-stu-id="0099d-188">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* [<span data-ttu-id="0099d-189">Word 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="0099d-189">Word add-in code samples</span></span>](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [<span data-ttu-id="0099d-190">Word JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="0099d-190">Word JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)
