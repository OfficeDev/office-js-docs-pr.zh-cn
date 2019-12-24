---
title: 生成首个 OneNote 任务窗格加载项
description: ''
ms.date: 12/24/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 2e8c560aa02de690fa4e6abae25d0625379e26ad
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851563"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="35a53-102">生成首个 OneNote 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="35a53-102">Build your first OneNote task pane add-in</span></span>

<span data-ttu-id="35a53-103">本文将逐步介绍如何生成 OneNote 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="35a53-103">In this article, you'll walk through the process of building a OneNote task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="35a53-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="35a53-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="35a53-105">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="35a53-105">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="35a53-106">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="35a53-106">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="35a53-107">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="35a53-107">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="35a53-108">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="35a53-108">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="35a53-109">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="35a53-109">**Which Office client application would you like to support?**</span></span> `OneNote`

![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-onenote.png)

<span data-ttu-id="35a53-111">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="35a53-111">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="35a53-112">浏览项目</span><span class="sxs-lookup"><span data-stu-id="35a53-112">Explore the project</span></span>

<span data-ttu-id="35a53-113">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="35a53-113">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="35a53-114">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="35a53-114">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="35a53-115">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="35a53-115">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="35a53-116">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="35a53-116">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="35a53-117">**./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Office 托管应用程序之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="35a53-117">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="35a53-118">更新代码</span><span class="sxs-lookup"><span data-stu-id="35a53-118">Update the code</span></span>

<span data-ttu-id="35a53-119">在代码编辑器中，打开文件 **./src/taskpane/taskpane.js** 并在 **run** 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="35a53-119">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="35a53-120">此代码使用 OneNote JavaScript API 设置页面标题并在页面正文添加大纲。</span><span class="sxs-lookup"><span data-stu-id="35a53-120">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

```js
try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a><span data-ttu-id="35a53-121">试用</span><span class="sxs-lookup"><span data-stu-id="35a53-121">Try it out</span></span>

1. <span data-ttu-id="35a53-122">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="35a53-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="35a53-123">启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="35a53-123">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="35a53-124">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="35a53-124">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="35a53-125">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="35a53-125">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="35a53-126">如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="35a53-126">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="35a53-127">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="35a53-127">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="35a53-128">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="35a53-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="35a53-129">如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。</span><span class="sxs-lookup"><span data-stu-id="35a53-129">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="35a53-130">在 [OneNote 网页版](https://www.onenote.com/notebooks)中，打开笔记本并新建页面。</span><span class="sxs-lookup"><span data-stu-id="35a53-130">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="35a53-131">依次选择“插入”>“Office 加载项”\*\*\*\*，打开“Office 加载项”对话框。</span><span class="sxs-lookup"><span data-stu-id="35a53-131">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="35a53-132">如果使用使用者帐户登录，请依次选择“我的加载项”\*\*\*\* 选项卡和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="35a53-132">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="35a53-133">如果使用工作或学校帐户登录，请依次选择“我的组织”\*\*\*\* 选项卡和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="35a53-133">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="35a53-134">下图展示了使用者笔记本的“**我的加载项**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="35a53-134">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="35a53-135">在“**上传加载项**”对话框中，转到项目文件夹中的 manifest.xml，然后选择“**上传**”。</span><span class="sxs-lookup"><span data-stu-id="35a53-135">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

6. <span data-ttu-id="35a53-136">在“**开始**”选项卡上，选择位于功能区的“**显示任务窗格**”按钮。</span><span class="sxs-lookup"><span data-stu-id="35a53-136">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="35a53-137">该加载项窗格在 OneNote 页旁的 iFrame 中打开。</span><span class="sxs-lookup"><span data-stu-id="35a53-137">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="35a53-138">在任务窗格底部，选择“**运行**”链接以设置页面标题并在页面正文中添加大纲。</span><span class="sxs-lookup"><span data-stu-id="35a53-138">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![通过此演练生成的 OneNote 加载项](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="35a53-140">后续步骤</span><span class="sxs-lookup"><span data-stu-id="35a53-140">Next steps</span></span>

<span data-ttu-id="35a53-141">恭喜！已成功创建 OneNote 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="35a53-141">Congratulations, you've successfully created a OneNote task pane add-in!</span></span> <span data-ttu-id="35a53-142">接下来，请详细了解与生成 OneNote 加载项有关的核心概念。</span><span class="sxs-lookup"><span data-stu-id="35a53-142">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="35a53-143">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="35a53-143">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="35a53-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="35a53-144">See also</span></span>

* [<span data-ttu-id="35a53-145">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="35a53-145">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="35a53-146">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="35a53-146">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)
* <span data-ttu-id="35a53-147">[开发 Office 加载项](../develop/develop-overview.md)</span><span class="sxs-lookup"><span data-stu-id="35a53-147">[](../develop/develop-overview.md)Develop Office Add-ins with Angular</span></span>
- [<span data-ttu-id="35a53-148">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="35a53-148">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="35a53-149">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="35a53-149">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="35a53-150">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="35a53-150">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

