---
title: 生成首个 Project 任务窗格加载项
description: ''
ms.date: 09/06/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 0a7684f3d1bd4f404ba42a798908bb9d2ba2f8d2
ms.sourcegitcommit: ce7e7087a4550b9c090dc565fee5eac08a2985a2
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/06/2019
ms.locfileid: "36782280"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="17b54-102">生成首个 Project 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="17b54-102">Build your first Project task pane add-in</span></span>

<span data-ttu-id="17b54-103">本文将逐步介绍如何生成 Project 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="17b54-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="17b54-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="17b54-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="17b54-105">Windows 版 Project 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="17b54-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="17b54-106">创建加载项</span><span class="sxs-lookup"><span data-stu-id="17b54-106">Create the add-in</span></span>

<span data-ttu-id="17b54-107">使用 Yeoman 生成器创建 Project 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="17b54-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="17b54-108">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="17b54-108">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="17b54-109">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="17b54-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="17b54-110">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="17b54-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="17b54-111">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="17b54-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="17b54-112">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="17b54-112">**Which Office client application would you like to support?**</span></span> `Project`

![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-project.png)

<span data-ttu-id="17b54-114">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="17b54-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="17b54-115">浏览项目</span><span class="sxs-lookup"><span data-stu-id="17b54-115">Explore the project</span></span>

<span data-ttu-id="17b54-116">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="17b54-116">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="17b54-117">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="17b54-117">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="17b54-118">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="17b54-118">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="17b54-119">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="17b54-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="17b54-120">**./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Office 托管应用程序之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="17b54-120">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="17b54-121">更新代码</span><span class="sxs-lookup"><span data-stu-id="17b54-121">Update the code</span></span>

<span data-ttu-id="17b54-122">在代码编辑器中，打开文件 **./src/taskpane/taskpane.js** 并在 **run** 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="17b54-122">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="17b54-123">此代码使用 Office JavaScript API 设置所选任务的 `Name` 字段和 `Notes` 字段。</span><span class="sxs-lookup"><span data-stu-id="17b54-123">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a><span data-ttu-id="17b54-124">试用</span><span class="sxs-lookup"><span data-stu-id="17b54-124">Try it out</span></span>

1. <span data-ttu-id="17b54-125">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="17b54-125">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="17b54-126">启动本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="17b54-126">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="17b54-127">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="17b54-127">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="17b54-128">如果系统在运行以下命令后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="17b54-128">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="17b54-129">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="17b54-129">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="17b54-130">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="17b54-130">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

3. <span data-ttu-id="17b54-131">在 Project 中，创建一个简单的项目计划。</span><span class="sxs-lookup"><span data-stu-id="17b54-131">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="17b54-132">按照[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中的说明，在 Project 中加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="17b54-132">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="17b54-133">在项目中选择单个任务。</span><span class="sxs-lookup"><span data-stu-id="17b54-133">Select a single task within the project.</span></span>

6. <span data-ttu-id="17b54-134">在任务窗格的底部，选择“**运行**”链接以重命名所选任务并向所选任务添加备注。</span><span class="sxs-lookup"><span data-stu-id="17b54-134">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![加载了任务窗格加载项的 Project 应用程序的屏幕截图](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="17b54-136">后续步骤</span><span class="sxs-lookup"><span data-stu-id="17b54-136">Next steps</span></span>

<span data-ttu-id="17b54-137">恭喜！已成功创建 Project 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="17b54-137">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="17b54-138">接下来，请详细了解 Project 加载项功能，并探索常见方案。</span><span class="sxs-lookup"><span data-stu-id="17b54-138">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="17b54-139">Project 加载项</span><span class="sxs-lookup"><span data-stu-id="17b54-139">Project add-ins</span></span>](../project/project-add-ins.md)

