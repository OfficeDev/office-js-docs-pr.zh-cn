---
ms.openlocfilehash: 5d5640a2c7c04817ab5ad87fc1b01e6866917830
ms.sourcegitcommit: 984c425e2ad58577af8f494079923cab165ad36c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/18/2019
ms.locfileid: "28726966"
---
# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="3a3d7-101">生成首个 OneNote 加载项</span><span class="sxs-lookup"><span data-stu-id="3a3d7-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="3a3d7-102">本文将逐步介绍如何使用 jQuery 和 Office JavaScript API 生成 OneNote 加载项。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3a3d7-103">先决条件</span><span class="sxs-lookup"><span data-stu-id="3a3d7-103">Prerequisites</span></span>

- [<span data-ttu-id="3a3d7-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="3a3d7-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="3a3d7-105">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="3a3d7-106">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="3a3d7-106">Create the add-in project</span></span>

1. <span data-ttu-id="3a3d7-107">使用 Yeoman 生成器创建 OneNote 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-107">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="3a3d7-108">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="3a3d7-108">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="3a3d7-109">**选择项目类型:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="3a3d7-109">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="3a3d7-110">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="3a3d7-110">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="3a3d7-111">**要如何命名加载项?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="3a3d7-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="3a3d7-112">**要支持哪一个 Office 客户端应用?:** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="3a3d7-112">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="3a3d7-114">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="3a3d7-115">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-115">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="3a3d7-116">更新代码</span><span class="sxs-lookup"><span data-stu-id="3a3d7-116">Update the code</span></span>

1. <span data-ttu-id="3a3d7-117">在代码编辑器中，打开项目根目录中的“index.html”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-117">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="3a3d7-118">此文件包含在加载项任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-118">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="3a3d7-119">将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-119">Replace the `<body>` element with the following markup and save the file.</span></span> 

    ```html
    <body class="ms-font-m ms-welcome">
        <header class="ms-welcome__header ms-bgColor-themeDark ms-u-fadeIn500">
            <h2 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">OneNote Add-in</h1>
        </header>
        <main id="app-body" class="ms-welcome__main">
            <br />
            <p class="ms-font-m">Enter HTML content here:</p>
            <div class="ms-TextField ms-TextField--placeholder">
                <textarea id="textBox" rows="8" cols="30"></textarea>
            </div>
            <button id="addOutline" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Add outline</span>
            </button>
        </main>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="3a3d7-120">打开文件“**src/index.js**”，以指定加载项的脚本。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-120">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="3a3d7-121">将整个内容替换为下列代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-121">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.onReady(() => {
        // Office is ready
        $(document).ready(() => {
            // The document is ready
            $('#addOutline').click(addOutlineToPage);
        });
    });
    
    async function addOutlineToPage() {
        try {
            await OneNote.run(async context => {
                var html = "<p>" + $("#textBox").val() + "</p>";

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.
                page.load("title");

                // Add text to the page by using the specified HTML.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log("Added outline to page " + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error);
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    ```

4. <span data-ttu-id="3a3d7-122">打开文件 **“app.css**”，以指定加载项自定义样式。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-122">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="3a3d7-123">将整个内容替换为以下内容，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-123">Replace the entire contents with the following and save the file.</span></span>

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 500px;
        min-width: 320px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 5px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 5px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 30px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="3a3d7-124">更新清单</span><span class="sxs-lookup"><span data-stu-id="3a3d7-124">Update the manifest</span></span>

1. <span data-ttu-id="3a3d7-125">打开文件“**manifest.xml**”以定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-125">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="3a3d7-126">`ProviderName` 元素具有占位符值。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="3a3d7-127">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-127">Replace it with your name.</span></span>

3. <span data-ttu-id="3a3d7-128">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-128">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="3a3d7-129">将它替换为“A task pane add-in for OneNote”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-129">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="3a3d7-130">保存文件。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-130">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="3a3d7-131">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="3a3d7-131">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="3a3d7-132">试用</span><span class="sxs-lookup"><span data-stu-id="3a3d7-132">Try it out</span></span>

1. <span data-ttu-id="3a3d7-133">在 [OneNote Online](https://www.onenote.com/notebooks) 中，打开一个笔记本。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-133">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="3a3d7-134">依次选择“插入”>“Office 加载项”\*\*\*\*，打开“Office 加载项”对话框。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-134">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="3a3d7-135">如果使用使用者帐户登录，请依次选择“我的加载项”\*\*\*\* 选项卡和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-135">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="3a3d7-136">如果使用工作或学校帐户登录，请依次选择“我的组织”\*\*\*\* 选项卡和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-136">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="3a3d7-137">下图展示了使用者笔记本的“**我的加载项**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-137">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="3a3d7-138">在“**上传加载项**”对话框中，转到项目文件夹中的 manifest.xml，然后选择“**上传**”。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-138">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="3a3d7-139">在“**开始**”选项卡上，选择位于功能区的“**显示任务窗格**”按钮。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-139">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="3a3d7-140">该加载项窗格在 OneNote 页旁的 iFrame 中打开。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-140">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="3a3d7-141">在文本区域中输入以下 HTML 内容，然后选择“**添加大纲**”。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-141">Enter the following HTML content in the text area, and then choose **Add outline**.</span></span>  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    <span data-ttu-id="3a3d7-142">指定的大纲将添加到页面中。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-142">The outline that you specified is added to the page.</span></span>

    ![通过此演练生成的 OneNote 加载项](../images/onenote-first-add-in-3.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="3a3d7-144">疑难解答和提示</span><span class="sxs-lookup"><span data-stu-id="3a3d7-144">Troubleshooting and tips</span></span>

- <span data-ttu-id="3a3d7-p108">您可以使用浏览器的开发者工具调试外接程序。当您在 Internet Explorer 或 Chrome 中使用 Gulp Web 服务器并进行调试时，您可以本地保存您的更改，然后仅刷新外接程序的 iFrame。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-p108">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="3a3d7-p109">检查 OneNote 对象时，目前可用的属性显示实际值。需要加载的属性显示“未定义”\*\*。展开 `_proto_` 节点，查看已在对象上定义但尚未加载的属性。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-p109">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![调试器中尚未加载的 OneNote 对象](../images/onenote-debug.png)

- <span data-ttu-id="3a3d7-p110">如果您的外接程序使用任何 HTTP 资源，则需要启用浏览器中的混合内容。生产外接程序应当仅使用安全 HTTPS 资源。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-p110">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="3a3d7-153">可以从任何位置打开任务窗格外接程序，但只能在常规页面内容（即不在标题、图像、IFrame 等中）内插入内容外接程序。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-153">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="3a3d7-154">后续步骤</span><span class="sxs-lookup"><span data-stu-id="3a3d7-154">Next steps</span></span>

<span data-ttu-id="3a3d7-155">恭喜！已成功创建 OneNote 加载项！</span><span class="sxs-lookup"><span data-stu-id="3a3d7-155">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="3a3d7-156">接下来，请详细了解与生成 OneNote 加载项有关的核心概念。</span><span class="sxs-lookup"><span data-stu-id="3a3d7-156">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3a3d7-157">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="3a3d7-157">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="3a3d7-158">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3a3d7-158">See also</span></span>

- [<span data-ttu-id="3a3d7-159">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="3a3d7-159">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="3a3d7-160">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="3a3d7-160">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="3a3d7-161">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="3a3d7-161">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="3a3d7-162">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="3a3d7-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
