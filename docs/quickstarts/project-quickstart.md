---
title: 生成首个 Project 任务窗格加载项
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: d61f8d83b88dbe69ff0ba9cd4b0afef77a4f03d6
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952245"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="29e98-102">生成首个 Project 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="29e98-102">Build your first PowerPoint task pane add-in</span></span>

<span data-ttu-id="29e98-103">本文将逐步介绍如何生成 Project 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="29e98-103">In this article, you'll walk through the process of building a PowerPoint task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="29e98-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="29e98-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="29e98-105">Windows 版 Project 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="29e98-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="29e98-106">创建加载项</span><span class="sxs-lookup"><span data-stu-id="29e98-106">Create the add-in</span></span>

1. <span data-ttu-id="29e98-107">使用 Yeoman 生成器创建 Project 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="29e98-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="29e98-108">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="29e98-108">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="29e98-109">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="29e98-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
    - <span data-ttu-id="29e98-110">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="29e98-110">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="29e98-111">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="29e98-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="29e98-112">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="29e98-112">**Which Office client application would you like to support?**</span></span> `Project`

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-project.png)
    
    <span data-ttu-id="29e98-114">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="29e98-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="29e98-115">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="29e98-115">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## <a name="explore-the-project"></a><span data-ttu-id="29e98-116">浏览项目</span><span class="sxs-lookup"><span data-stu-id="29e98-116">Explore the project</span></span>

<span data-ttu-id="29e98-117">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="29e98-117">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="29e98-118">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="29e98-118">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="29e98-119">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="29e98-119">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="29e98-120">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="29e98-120">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="29e98-121">**./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Office 托管应用程序之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="29e98-121">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="29e98-122">更新代码</span><span class="sxs-lookup"><span data-stu-id="29e98-122">Update the code</span></span>

<span data-ttu-id="29e98-123">在代码编辑器中，打开文件 **./src/taskpane/taskpane.js** 并在 **run** 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="29e98-123">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="29e98-124">此代码使用 Office JavaScript API 设置所选任务的 `Name` 字段和 `Notes` 字段。</span><span class="sxs-lookup"><span data-stu-id="29e98-124">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="29e98-125">试用</span><span class="sxs-lookup"><span data-stu-id="29e98-125">Try it out</span></span>

1. <span data-ttu-id="29e98-126">通过运行以下命令启用本地 Web 服务器：</span><span class="sxs-lookup"><span data-stu-id="29e98-126">Start the local web server by running the following command:</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="29e98-127">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="29e98-127">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="29e98-128">如果系统在运行 `npm start` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="29e98-128">If you are prompted to install a certificate after you run `npm start`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> 

2. <span data-ttu-id="29e98-129">在 Project 中，创建一个简单的项目计划。</span><span class="sxs-lookup"><span data-stu-id="29e98-129">In Project, create a simple project plan.</span></span>

3. <span data-ttu-id="29e98-130">按照[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)中的说明，在 Project 中加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="29e98-130">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

4. <span data-ttu-id="29e98-131">在项目中选择单个任务。</span><span class="sxs-lookup"><span data-stu-id="29e98-131">Select a single task within the project.</span></span>

5. <span data-ttu-id="29e98-132">在任务窗格的底部，选择“**运行**”链接以重命名所选任务并向所选任务添加备注。</span><span class="sxs-lookup"><span data-stu-id="29e98-132">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![加载了任务窗格加载项的 Project 应用程序的屏幕截图](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="29e98-134">后续步骤</span><span class="sxs-lookup"><span data-stu-id="29e98-134">Next steps</span></span>

<span data-ttu-id="29e98-135">恭喜！已成功创建 Project 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="29e98-135">Congratulations, you've successfully created a PowerPoint task pane add-in!</span></span> <span data-ttu-id="29e98-136">接下来，请详细了解 Project 加载项功能，并探索常见方案。</span><span class="sxs-lookup"><span data-stu-id="29e98-136">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="29e98-137">Project 加载项</span><span class="sxs-lookup"><span data-stu-id="29e98-137">Project add-ins</span></span>](../project/project-add-ins.md)

