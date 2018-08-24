# <a name="build-your-first-project-add-in"></a><span data-ttu-id="49ac8-101">生成首个 Project 加载项</span><span class="sxs-lookup"><span data-stu-id="49ac8-101">Build your first Project add-in</span></span>

<span data-ttu-id="49ac8-102">本文将逐步介绍如何使用 jQuery 和 Office JavaScript API 生成 Project 加载项。</span><span class="sxs-lookup"><span data-stu-id="49ac8-102">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="49ac8-103">先决条件</span><span class="sxs-lookup"><span data-stu-id="49ac8-103">Prerequisites</span></span>

- [<span data-ttu-id="49ac8-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="49ac8-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="49ac8-105">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="49ac8-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="49ac8-106">创建加载项</span><span class="sxs-lookup"><span data-stu-id="49ac8-106">Create the add-in</span></span>

1. <span data-ttu-id="49ac8-107">在本地驱动器上创建一个文件夹，并命名为 `my-project-addin`。</span><span class="sxs-lookup"><span data-stu-id="49ac8-107">Create a folder on your local drive and name it `my-project-addin`.</span></span> <span data-ttu-id="49ac8-108">将在其中创建外接程序文件。</span><span class="sxs-lookup"><span data-stu-id="49ac8-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="49ac8-109">转到新文件夹。</span><span class="sxs-lookup"><span data-stu-id="49ac8-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-project-addin
    ```

3. <span data-ttu-id="49ac8-110">使用 Yeoman 生成器创建 Project 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="49ac8-110">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="49ac8-111">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="49ac8-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="49ac8-112">**选择一个项目类型：** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="49ac8-112">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="49ac8-113">**选择一个脚本类型：** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="49ac8-113">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="49ac8-114">**要如何命名加载项?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="49ac8-114">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="49ac8-115">**要支持哪一个 Office 客户端应用?:** `Project`</span><span class="sxs-lookup"><span data-stu-id="49ac8-115">**Which Office client application would you like to support?:** `Project`</span></span>

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-project-jquery.png)
    
    <span data-ttu-id="49ac8-117">完成向导后，生成器将创建项目并安装 Node 支持组件。</span><span class="sxs-lookup"><span data-stu-id="49ac8-117">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="49ac8-118">更新代码</span><span class="sxs-lookup"><span data-stu-id="49ac8-118">Update the code</span></span>

1. <span data-ttu-id="49ac8-119">在代码编辑器中，打开项目根目录中的“index.html”****。</span><span class="sxs-lookup"><span data-stu-id="49ac8-119">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="49ac8-120">此文件包含在加载项任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="49ac8-120">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="49ac8-121">将 `<body>` 元素内的 `<header>` 元素替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="49ac8-121">Replace the `<header>` element inside the `<body>` element with the following markup.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

3. <span data-ttu-id="49ac8-122">将 `<body>` 元素内的 `<main>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="49ac8-122">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
            <h3>Try it out</h3>
            <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
            <br/><br/>
            <button class="ms-Button" id="get-task">Get Task data</button>
            <br/>
            <h4>Results:</h4>
            <textarea id="result" rows="6" cols="25"></textarea>
        </div>
    </div>
    ```

4. <span data-ttu-id="49ac8-p104">打开文件 **index/ src.js** 以指定加载项的脚本。使用以下代码替换全部内容并保存文件。</span><span class="sxs-lookup"><span data-stu-id="49ac8-p104">Open the file **app.js** to specify the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. <span data-ttu-id="49ac8-125">|||UNTRANSLATED_CONTENT_START|||Open the file **app.css** in the root of the project to specify the custom styles for the add-in.|||UNTRANSLATED_CONTENT_END|||</span><span class="sxs-lookup"><span data-stu-id="49ac8-125">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="49ac8-126">将整个内容替换为以下内容，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="49ac8-126">Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="49ac8-127">更新清单</span><span class="sxs-lookup"><span data-stu-id="49ac8-127">Update the manifest</span></span>

1. <span data-ttu-id="49ac8-128">打开文件“my-office-add-in-manifest.xml”****，以定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="49ac8-128">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="49ac8-129">元素具有占位符值。`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="49ac8-129">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="49ac8-130">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="49ac8-130">Replace it with your name.</span></span>

3. <span data-ttu-id="49ac8-131">元素的 `DefaultValue` 属性有占位符。`Description`</span><span class="sxs-lookup"><span data-stu-id="49ac8-131">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="49ac8-132">将它替换为“A task pane add-in for Project”****。</span><span class="sxs-lookup"><span data-stu-id="49ac8-132">Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="49ac8-133">保存文件。</span><span class="sxs-lookup"><span data-stu-id="49ac8-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="49ac8-134">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="49ac8-134">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="49ac8-135">试用</span><span class="sxs-lookup"><span data-stu-id="49ac8-135">Try it out</span></span>

1. <span data-ttu-id="49ac8-136">在 Project 中，创建至少有一个任务的简单项目。</span><span class="sxs-lookup"><span data-stu-id="49ac8-136">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="49ac8-137">请按照运行加载项所用平台对应的说明操作，以在 Project 中旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="49ac8-137">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="49ac8-138">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="49ac8-138">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="49ac8-139">Project Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="49ac8-139">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="49ac8-140">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="49ac8-140">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="49ac8-141">在 Project 中，选择一个任务。</span><span class="sxs-lookup"><span data-stu-id="49ac8-141">In Project, select a task.</span></span>

    ![Project 中已选择一个任务的项目计划的屏幕截图](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="49ac8-143">在任务窗格中，选择“获取任务 GUID”**** 按钮，将任务 GUID 写入到“结果”**** 文本框。</span><span class="sxs-lookup"><span data-stu-id="49ac8-143">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![Project 中已选择一个任务的项目计划，且任务 GUID 写入到任务窗格中文本框的屏幕截图](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="49ac8-145">在任务窗格中，选择“获取任务数据”**** 按钮，将选定任务的多个属性写入到“结果”**** 文本框。</span><span class="sxs-lookup"><span data-stu-id="49ac8-145">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![Project 中已选择一个任务的项目计划，且多个任务属性写入到任务窗格中文本框的屏幕截图](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="49ac8-147">后续步骤</span><span class="sxs-lookup"><span data-stu-id="49ac8-147">Next steps</span></span>

<span data-ttu-id="49ac8-148">恭喜！已成功创建 Project 加载项！</span><span class="sxs-lookup"><span data-stu-id="49ac8-148">Congratulations, you've successfully created a Project add-in!</span></span> <span data-ttu-id="49ac8-149">接下来，请详细了解 Project 加载项功能，并探索常见方案。</span><span class="sxs-lookup"><span data-stu-id="49ac8-149">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="49ac8-150">Project 加载项</span><span class="sxs-lookup"><span data-stu-id="49ac8-150">Project add-ins</span></span>](../project/project-add-ins.md)
