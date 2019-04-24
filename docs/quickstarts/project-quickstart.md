---
title: 生成首个 Project 加载项
description: ''
ms.date: 01/17/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 4d0dfa98d36d6da56fe2b9687922371eea29062a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450766"
---
# <a name="build-your-first-project-add-in"></a><span data-ttu-id="83a3f-102">生成首个 Project 加载项</span><span class="sxs-lookup"><span data-stu-id="83a3f-102">Build your first Project add-in</span></span>

<span data-ttu-id="83a3f-103">本文将逐步介绍如何使用 jQuery 和 Office JavaScript API 生成 Project 加载项。</span><span class="sxs-lookup"><span data-stu-id="83a3f-103">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="83a3f-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="83a3f-104">Prerequisites</span></span>

- [<span data-ttu-id="83a3f-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="83a3f-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="83a3f-106">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="83a3f-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="83a3f-107">创建加载项</span><span class="sxs-lookup"><span data-stu-id="83a3f-107">Create the add-in</span></span>

1. <span data-ttu-id="83a3f-108">使用 Yeoman 生成器创建 Project 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="83a3f-108">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="83a3f-109">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="83a3f-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="83a3f-110">**选择项目类型:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="83a3f-110">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="83a3f-111">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="83a3f-111">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="83a3f-112">**要如何命名加载项?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="83a3f-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="83a3f-113">**要支持哪一个 Office 客户端应用?:** `Project`</span><span class="sxs-lookup"><span data-stu-id="83a3f-113">**Which Office client application would you like to support?:** `Project`</span></span>

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-project-jquery.png)
    
    <span data-ttu-id="83a3f-115">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="83a3f-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="83a3f-116">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="83a3f-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="83a3f-117">更新代码</span><span class="sxs-lookup"><span data-stu-id="83a3f-117">Update the code</span></span>

1. <span data-ttu-id="83a3f-p102">在代码编辑器中，打开项目根目录中的“index.html”\*\*\*\*。 此文件包含在加载项任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="83a3f-p102">In your code editor, open **index.html** in the root of the project. This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="83a3f-120">用以下标记替换 `<body>` 元素。</span><span class="sxs-lookup"><span data-stu-id="83a3f-120">Replace the `<body>` element with the following markup.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
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
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="83a3f-121">打开文件 **src/index.js**，指定加载项的脚本。</span><span class="sxs-lookup"><span data-stu-id="83a3f-121">Open the file **src/index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="83a3f-122">将整个内容替换为下列代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="83a3f-122">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        var taskGuid;

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        });

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

4. <span data-ttu-id="83a3f-p104">打开项目根目录中的文件“app.css”\*\*\*\*，以指定加载项自定义样式。 将整个内容替换为以下内容，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="83a3f-p104">Open the file **app.css** in the root of the project to specify the custom styles for the add-in. Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="83a3f-125">更新清单</span><span class="sxs-lookup"><span data-stu-id="83a3f-125">Update the manifest</span></span>

1. <span data-ttu-id="83a3f-126">打开文件“**manifest.xml**”以定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="83a3f-126">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="83a3f-p105">`ProviderName` 元素具有占位符值。 将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="83a3f-p105">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="83a3f-129">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="83a3f-129">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="83a3f-130">将它替换为“A task pane add-in for Project”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="83a3f-130">Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="83a3f-131">保存文件。</span><span class="sxs-lookup"><span data-stu-id="83a3f-131">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="83a3f-132">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="83a3f-132">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="83a3f-133">试用</span><span class="sxs-lookup"><span data-stu-id="83a3f-133">Try it out</span></span>

1. <span data-ttu-id="83a3f-134">在 Project 中，创建至少有一个任务的简单项目。</span><span class="sxs-lookup"><span data-stu-id="83a3f-134">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="83a3f-135">请按照运行加载项所用平台对应的说明操作，以在 Project 中旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="83a3f-135">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="83a3f-136">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="83a3f-136">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="83a3f-137">Project Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="83a3f-137">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="83a3f-138">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="83a3f-138">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="83a3f-139">在 Project 中，选择一个任务。</span><span class="sxs-lookup"><span data-stu-id="83a3f-139">In Project, select a task.</span></span>

    ![Project 中已选择一个任务的项目计划的屏幕截图](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="83a3f-141">在任务窗格中，选择“获取任务 GUID”\*\*\*\* 按钮，将任务 GUID 写入到“结果”\*\*\*\* 文本框。</span><span class="sxs-lookup"><span data-stu-id="83a3f-141">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![Project 中已选择一个任务的项目计划，且任务 GUID 写入到任务窗格中文本框的屏幕截图](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="83a3f-143">在任务窗格中，选择“获取任务数据”\*\*\*\* 按钮，将选定任务的多个属性写入到“结果”\*\*\*\* 文本框。</span><span class="sxs-lookup"><span data-stu-id="83a3f-143">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![Project 中已选择一个任务的项目计划，且多个任务属性写入到任务窗格中文本框的屏幕截图](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="83a3f-145">后续步骤</span><span class="sxs-lookup"><span data-stu-id="83a3f-145">Next steps</span></span>

<span data-ttu-id="83a3f-p107">恭喜！已成功创建 Project 加载项！ 接下来，请详细了解 Project 加载项功能，并探索常见方案。</span><span class="sxs-lookup"><span data-stu-id="83a3f-p107">Congratulations, you've successfully created a Project add-in! Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="83a3f-148">Project 加载项</span><span class="sxs-lookup"><span data-stu-id="83a3f-148">Project add-ins</span></span>](../project/project-add-ins.md)

