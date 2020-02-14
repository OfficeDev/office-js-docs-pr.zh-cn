---
title: 生成首个 Project 任务窗格加载项
description: 了解如何使用 Office JS API 生成简单的 Project 任务窗格加载项。
ms.date: 01/16/2020
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 821cdc9f32b0fbc2b48e2a92259f340e65a03f64
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950619"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="f98ee-103">生成首个 Project 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="f98ee-103">Build your first Project task pane add-in</span></span>

<span data-ttu-id="f98ee-104">本文将逐步介绍如何生成 Project 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="f98ee-104">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f98ee-105">先决条件</span><span class="sxs-lookup"><span data-stu-id="f98ee-105">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="f98ee-106">Windows 版 Project 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="f98ee-106">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="f98ee-107">创建加载项</span><span class="sxs-lookup"><span data-stu-id="f98ee-107">Create the add-in</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="f98ee-108">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="f98ee-108">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="f98ee-109">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="f98ee-109">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="f98ee-110">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="f98ee-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="f98ee-111">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="f98ee-111">**Which Office client application would you like to support?**</span></span> `Project`

![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-project.png)

<span data-ttu-id="f98ee-113">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="f98ee-113">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="f98ee-114">浏览项目</span><span class="sxs-lookup"><span data-stu-id="f98ee-114">Explore the project</span></span>

<span data-ttu-id="f98ee-115">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="f98ee-115">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="f98ee-116">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="f98ee-116">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="f98ee-117">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="f98ee-117">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="f98ee-118">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="f98ee-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="f98ee-119">**./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Office 托管应用程序之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="f98ee-119">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="f98ee-120">更新代码</span><span class="sxs-lookup"><span data-stu-id="f98ee-120">Update the code</span></span>

<span data-ttu-id="f98ee-121">在代码编辑器中，打开文件 **./src/taskpane/taskpane.js** 并在 **run** 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="f98ee-121">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="f98ee-122">此代码使用 Office JavaScript API 设置所选任务的 `Name` 字段和 `Notes` 字段。</span><span class="sxs-lookup"><span data-stu-id="f98ee-122">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="f98ee-123">试用</span><span class="sxs-lookup"><span data-stu-id="f98ee-123">Try it out</span></span>

1. <span data-ttu-id="f98ee-124">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="f98ee-124">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="f98ee-125">启动本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="f98ee-125">Start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f98ee-126">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="f98ee-126">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f98ee-127">如果系统在运行以下命令后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="f98ee-127">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="f98ee-128">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="f98ee-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="f98ee-129">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="f98ee-129">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

3. <span data-ttu-id="f98ee-130">在 Project 中，创建一个简单的项目计划。</span><span class="sxs-lookup"><span data-stu-id="f98ee-130">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="f98ee-131">按照[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中的说明，在 Project 中加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="f98ee-131">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="f98ee-132">在项目中选择单个任务。</span><span class="sxs-lookup"><span data-stu-id="f98ee-132">Select a single task within the project.</span></span>

6. <span data-ttu-id="f98ee-133">在任务窗格的底部，选择“**运行**”链接以重命名所选任务并向所选任务添加备注。</span><span class="sxs-lookup"><span data-stu-id="f98ee-133">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![加载了任务窗格加载项的 Project 应用程序的屏幕截图](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="f98ee-135">后续步骤</span><span class="sxs-lookup"><span data-stu-id="f98ee-135">Next steps</span></span>

<span data-ttu-id="f98ee-136">恭喜！已成功创建 Project 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="f98ee-136">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="f98ee-137">接下来，请详细了解 Project 加载项功能，并探索常见方案。</span><span class="sxs-lookup"><span data-stu-id="f98ee-137">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f98ee-138">Project 加载项</span><span class="sxs-lookup"><span data-stu-id="f98ee-138">Project add-ins</span></span>](../project/project-add-ins.md)

## <a name="see-also"></a><span data-ttu-id="f98ee-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="f98ee-139">See also</span></span>

- [<span data-ttu-id="f98ee-140">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="f98ee-140">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="f98ee-141">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="f98ee-141">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="f98ee-142">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="f98ee-142">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
