---
title: 生成首个 OneNote 任务窗格加载项
description: 了解如何使用 Office JS API 生成简单的 OneNote 任务窗格加载项。
ms.date: 10/14/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: de0729a483057a61be3793e299995aa05d287441
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132289"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="280e0-103">生成首个 OneNote 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="280e0-103">Build your first OneNote task pane add-in</span></span>

<span data-ttu-id="280e0-104">本文将逐步介绍如何生成 OneNote 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="280e0-104">In this article, you'll walk through the process of building a OneNote task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="280e0-105">先决条件</span><span class="sxs-lookup"><span data-stu-id="280e0-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="280e0-106">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="280e0-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="280e0-107">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="280e0-107">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="280e0-108">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="280e0-108">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="280e0-109">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="280e0-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="280e0-110">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="280e0-110">**Which Office client application would you like to support?**</span></span> `OneNote`

![显示命令行界面中 Yeoman 生成器的提示和回答的屏幕截图](../images/yo-office-onenote.png)

<span data-ttu-id="280e0-112">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="280e0-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="280e0-113">浏览项目</span><span class="sxs-lookup"><span data-stu-id="280e0-113">Explore the project</span></span>

<span data-ttu-id="280e0-114">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="280e0-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span>

- <span data-ttu-id="280e0-115">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="280e0-115">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="280e0-116">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="280e0-116">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="280e0-117">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="280e0-117">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="280e0-118">**./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Office 客户端应用程序之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="280e0-118">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="280e0-119">更新代码</span><span class="sxs-lookup"><span data-stu-id="280e0-119">Update the code</span></span>

<span data-ttu-id="280e0-120">在代码编辑器中，打开文件 **./src/taskpane/taskpane.js** 并在 `run` 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="280e0-120">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="280e0-121">此代码使用 OneNote JavaScript API 设置页面标题并在页面正文添加大纲。</span><span class="sxs-lookup"><span data-stu-id="280e0-121">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="280e0-122">试用</span><span class="sxs-lookup"><span data-stu-id="280e0-122">Try it out</span></span>

1. <span data-ttu-id="280e0-123">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="280e0-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="280e0-124">启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="280e0-124">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="280e0-125">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="280e0-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="280e0-126">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="280e0-126">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="280e0-127">你可能还必须以管理员身份运行命令提示符或终端才能进行更改。</span><span class="sxs-lookup"><span data-stu-id="280e0-127">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="280e0-128">如果在 Mac 上测试加载项，请先运行以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="280e0-128">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="280e0-129">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="280e0-129">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="280e0-130">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="280e0-130">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="280e0-131">如果你运行此命令，本地 Web 服务器将启动（如果尚未运行的话）。</span><span class="sxs-lookup"><span data-stu-id="280e0-131">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="280e0-132">在 [OneNote 网页版](https://www.onenote.com/notebooks)中，打开笔记本并新建页面。</span><span class="sxs-lookup"><span data-stu-id="280e0-132">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="280e0-133">依次选择“插入”>“Office 加载项”，打开“Office 加载项”对话框。</span><span class="sxs-lookup"><span data-stu-id="280e0-133">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="280e0-134">如果使用使用者帐户登录，请依次选择“我的加载项”选项卡和“上传我的加载项”。</span><span class="sxs-lookup"><span data-stu-id="280e0-134">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="280e0-135">如果使用工作或教育帐户登录，请依次选择“**我的组织**”选项卡和“**上传我的加载项**”。</span><span class="sxs-lookup"><span data-stu-id="280e0-135">If you're signed in with your work or education account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span>

    <span data-ttu-id="280e0-136">下图展示了使用者笔记本的“**我的加载项**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="280e0-136">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="Screenshot of the Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="280e0-137">在“**上传加载项**”对话框中，转到项目文件夹中的 manifest.xml，然后选择“**上传**”。</span><span class="sxs-lookup"><span data-stu-id="280e0-137">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span>

6. <span data-ttu-id="280e0-138">在“**开始**”选项卡上，选择位于功能区的“**显示任务窗格**”按钮。</span><span class="sxs-lookup"><span data-stu-id="280e0-138">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="280e0-139">该加载项窗格在 OneNote 页旁的 iFrame 中打开。</span><span class="sxs-lookup"><span data-stu-id="280e0-139">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="280e0-140">在任务窗格底部，选择“**运行**”链接以设置页面标题并在页面正文中添加大纲。</span><span class="sxs-lookup"><span data-stu-id="280e0-140">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![显示根据本演练生成的加载项的屏幕截图：在 OneNote 中显示任务窗格功能区按钮和任务窗格](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="280e0-142">后续步骤</span><span class="sxs-lookup"><span data-stu-id="280e0-142">Next steps</span></span>

<span data-ttu-id="280e0-143">恭喜！已成功创建 OneNote 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="280e0-143">Congratulations, you've successfully created a OneNote task pane add-in!</span></span> <span data-ttu-id="280e0-144">接下来，请详细了解与生成 OneNote 加载项有关的核心概念。</span><span class="sxs-lookup"><span data-stu-id="280e0-144">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="280e0-145">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="280e0-145">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="280e0-146">另请参阅</span><span class="sxs-lookup"><span data-stu-id="280e0-146">See also</span></span>

- [<span data-ttu-id="280e0-147">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="280e0-147">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="280e0-148">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="280e0-148">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="280e0-149">OneNote JavaScript API 编程概述</span><span class="sxs-lookup"><span data-stu-id="280e0-149">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="280e0-150">OneNote JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="280e0-150">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="280e0-151">Rubric Grader 示例</span><span class="sxs-lookup"><span data-stu-id="280e0-151">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
