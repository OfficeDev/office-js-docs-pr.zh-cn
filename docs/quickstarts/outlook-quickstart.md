---
title: 生成首个 Outlook 加载项
description: 了解如何使用 Office JS API 生成简单的 Outlook 任务窗格加载项。
ms.date: 06/23/2020
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 6eb510986ebb9cd8a4362f8c11f744e694bef5f7
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094160"
---
# <a name="build-your-first-outlook-add-in"></a><span data-ttu-id="917e9-103">生成首个 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="917e9-103">Build your first Outlook add-in</span></span>

<span data-ttu-id="917e9-104">本文将逐步介绍如何生成显示选定邮件的至少一个属性的 Outlook 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="917e9-104">In this article, you'll walk through the process of building an Outlook task pane add-in that displays at least one property of a selected message.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="917e9-105">创建加载项</span><span class="sxs-lookup"><span data-stu-id="917e9-105">Create the add-in</span></span>

<span data-ttu-id="917e9-106">可以使用[适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)或 Visual Studio 创建 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="917e9-106">You can create an Office Add-in by using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) or Visual Studio.</span></span> <span data-ttu-id="917e9-107">Yeoman 生成器将创建一个可通过 Visual Studio Code 或任何其他编辑器管理的 Node.js 项目，而 Visual Studio 将创建一个 Visual Studio 解决方案。</span><span class="sxs-lookup"><span data-stu-id="917e9-107">The Yeoman generator creates a Node.js project that can be managed with Visual Studio Code or any other editor, whereas Visual Studio creates a Visual Studio solution.</span></span>  <span data-ttu-id="917e9-108">选择适合于想要使用的方法的选项卡，然后按照说明创建加载项并在本地测试。</span><span class="sxs-lookup"><span data-stu-id="917e9-108">Select the tab for the one you'd like to use and then follow the instructions to create your add-in and test it locally.</span></span>

# <a name="yeoman-generator"></a>[<span data-ttu-id="917e9-109">Yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="917e9-109">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="917e9-110">先决条件</span><span class="sxs-lookup"><span data-stu-id="917e9-110">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- <span data-ttu-id="917e9-111">[Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）</span><span class="sxs-lookup"><span data-stu-id="917e9-111">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

- <span data-ttu-id="917e9-112">最新版本的 [Yeoman](https://github.com/yeoman/yo) 和[适用于 Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="917e9-112">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="917e9-113">即便先前已安装了 Yeoman 生成器，我们还是建议你通过 npm 将包更新为最新版本。</span><span class="sxs-lookup"><span data-stu-id="917e9-113">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="917e9-114">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="917e9-114">Create the add-in project</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="917e9-115">**选择项目类型** - `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="917e9-115">**Choose a project type** - `Office Add-in Task Pane project`</span></span>

    - <span data-ttu-id="917e9-116">**选择脚本类型** - `Javascript`</span><span class="sxs-lookup"><span data-stu-id="917e9-116">**Choose a script type** - `Javascript`</span></span>

    - <span data-ttu-id="917e9-117">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="917e9-117">**What do you want to name your add-in?**</span></span> - `My Office Add-in`

    - <span data-ttu-id="917e9-118">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="917e9-118">**Which Office client application would you like to support?**</span></span> - `Outlook`

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-outlook.png)
    
    <span data-ttu-id="917e9-120">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="917e9-120">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. <span data-ttu-id="917e9-121">导航到 Web 应用程序项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="917e9-121">Navigate to the root folder of the web application project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a><span data-ttu-id="917e9-122">浏览项目</span><span class="sxs-lookup"><span data-stu-id="917e9-122">Explore the project</span></span>

<span data-ttu-id="917e9-123">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="917e9-123">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="917e9-124">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="917e9-124">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="917e9-125">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="917e9-125">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="917e9-126">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="917e9-126">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="917e9-127">**./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Outlook 之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="917e9-127">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and Outlook.</span></span>

### <a name="update-the-code"></a><span data-ttu-id="917e9-128">更新代码</span><span class="sxs-lookup"><span data-stu-id="917e9-128">Update the code</span></span>

1. <span data-ttu-id="917e9-129">在代码编辑器中，打开文件 **./src/taskpane/taskpane.html** 并将整个 `<main>` 元素（位于 `<body>` 元素中）替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="917e9-129">In your code editor, open the file **./src/taskpane/taskpane.html** and replace the entire `<main>` element (within the `<body>` element) with the following markup.</span></span> <span data-ttu-id="917e9-130">此新标记将添加标签，其中 **./src/taskpane/taskpane.js** 中的脚本将写入数据。</span><span class="sxs-lookup"><span data-stu-id="917e9-130">This new markup adds a label where the script in **./src/taskpane/taskpane.js** will write data.</span></span>

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. <span data-ttu-id="917e9-131">在代码编辑器中，打开文件 **./src/taskpane/taskpane.js** 并在 `run` 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="917e9-131">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="917e9-132">此代码使用 Office JavaScript API 获取当前邮件的引用并将其 `subject` 属性值写入任务窗格。</span><span class="sxs-lookup"><span data-stu-id="917e9-132">This code uses the Office JavaScript API to get a reference to the current message and write its `subject` property value to the task pane.</span></span>

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a><span data-ttu-id="917e9-133">试用</span><span class="sxs-lookup"><span data-stu-id="917e9-133">Try it out</span></span>

> [!NOTE]
> <span data-ttu-id="917e9-134">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="917e9-134">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="917e9-135">如果系统在运行以下命令后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="917e9-135">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

1. <span data-ttu-id="917e9-136">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="917e9-136">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="917e9-137">运行此命令时，本地 Web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="917e9-137">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="917e9-138">按照[旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="917e9-138">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="917e9-139">在 Outlook 中，选择或打开邮件。</span><span class="sxs-lookup"><span data-stu-id="917e9-139">In Outlook, select or open a message.</span></span>

1. <span data-ttu-id="917e9-140">选择“**主页**”选项卡（或“**邮件**”选项卡，如果在新窗口中打开了邮件），然后选择功能区的“**显示任务窗格**”按钮以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="917e9-140">Choose the **Home** tab (or the **Message** tab if you opened the message in a new window), and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Outlook 中邮件窗口的屏幕截图，突出显示了加载项按钮](../images/quick-start-button-1.png)

1. <span data-ttu-id="917e9-142">滚动至任务窗格的底部并选择“**运行**”链接，将邮件主题写入任务窗格。</span><span class="sxs-lookup"><span data-stu-id="917e9-142">Scroll to the bottom of the task pane and choose the **Run** link to write the message subject to the task pane.</span></span>

    ![加载项任务窗格截屏，高亮显示运行链接](../images/quick-start-task-pane-2.png)

    ![加载项任务窗格的屏幕截图，其中显示邮件主题](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a><span data-ttu-id="917e9-145">后续步骤</span><span class="sxs-lookup"><span data-stu-id="917e9-145">Next steps</span></span>

<span data-ttu-id="917e9-146">祝贺！已成功创建首个 Outlook 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="917e9-146">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="917e9-147">接下来，将继续学习 [Outlook 加载项教程](../tutorials/outlook-tutorial.md)，详细了解 Outlook 加载项的功能，以及如何生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="917e9-147">Next, learn more about the capabilities of an Outlook add-in and build a more complex add-in by following along with the [Outlook add-in tutorial](../tutorials/outlook-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="917e9-148">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="917e9-148">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="917e9-149">先决条件</span><span class="sxs-lookup"><span data-stu-id="917e9-149">Prerequisites</span></span>

- <span data-ttu-id="917e9-150">安装了 **Office/SharePoint 开发**工作负载的 [Visual Studio 2019](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="917e9-150">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="917e9-151">如果之前已安装 Visual Studio 2019，请[使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发**工作负载。</span><span class="sxs-lookup"><span data-stu-id="917e9-151">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span>

- <span data-ttu-id="917e9-152">Office 365</span><span class="sxs-lookup"><span data-stu-id="917e9-152">Office 365</span></span>

    > [!NOTE]
    > <span data-ttu-id="917e9-153">如果没有 Microsoft 365 订阅，可以通过注册 [Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个免费订阅。</span><span class="sxs-lookup"><span data-stu-id="917e9-153">If you do not have a Microsoft 365 subscription, you can get a free one by signing up for the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="917e9-154">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="917e9-154">Create the add-in project</span></span>

1. <span data-ttu-id="917e9-155">在 Visual Studio 菜单栏中，依次选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="917e9-155">On the Visual Studio menu bar, choose **File** > **New** > **Project**.</span></span>

1. <span data-ttu-id="917e9-156">在“Visual C#”\*\*\*\* 或“Visual Basic”\*\*\*\* 下的项目类型列表中，展开“Office/SharePoint”\*\*\*\*，选择“加载项”\*\*\*\*，然后选择“Outlook Web 加载项”\*\*\*\* 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="917e9-156">In the list of project types under **Visual C#** or **Visual Basic**, expand **Office/SharePoint**, choose **Add-ins**, and then choose **Outlook Web Add-in** as the project type.</span></span>

1. <span data-ttu-id="917e9-157">命名此项目，再选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="917e9-157">Name the project, and then choose **OK**.</span></span>

1. <span data-ttu-id="917e9-158">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。</span><span class="sxs-lookup"><span data-stu-id="917e9-158">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="917e9-159">**MessageRead.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="917e9-159">The **MessageRead.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="917e9-160">浏览 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="917e9-160">Explore the Visual Studio solution</span></span>

<span data-ttu-id="917e9-161">在用户完成向导后，Visual Studio 会创建一个包含两个项目的解决方案。</span><span class="sxs-lookup"><span data-stu-id="917e9-161">When you've completed the wizard, Visual Studio creates a solution that contains two projects.</span></span>

|<span data-ttu-id="917e9-162">**项目**</span><span class="sxs-lookup"><span data-stu-id="917e9-162">**Project**</span></span>|<span data-ttu-id="917e9-163">**说明**</span><span class="sxs-lookup"><span data-stu-id="917e9-163">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="917e9-164">加载项项目</span><span class="sxs-lookup"><span data-stu-id="917e9-164">Add-in project</span></span>|<span data-ttu-id="917e9-165">仅包含 XML 清单文件，内含描述加载项的所有设置。</span><span class="sxs-lookup"><span data-stu-id="917e9-165">Contains only an XML manifest file, which contains all the settings that describe your add-in.</span></span> <span data-ttu-id="917e9-166">这些设置有助于 Office 主机确定应在何时激活加载项，以及应在哪里显示加载项。</span><span class="sxs-lookup"><span data-stu-id="917e9-166">These settings help the Office host determine when your add-in should be activated and where the add-in should appear.</span></span> <span data-ttu-id="917e9-167">Visual Studio 生成了此文件的内容，以便于用户能够立即运行项目并使用外接程序。</span><span class="sxs-lookup"><span data-stu-id="917e9-167">Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately.</span></span> <span data-ttu-id="917e9-168">可以通过修改 XML 文件随时更改这些设置。</span><span class="sxs-lookup"><span data-stu-id="917e9-168">You can change these settings any time by modifying the XML file.</span></span>|
|<span data-ttu-id="917e9-169">Web 应用项目</span><span class="sxs-lookup"><span data-stu-id="917e9-169">Web application project</span></span>|<span data-ttu-id="917e9-170">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages.</span><span class="sxs-lookup"><span data-stu-id="917e9-170">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages.</span></span> <span data-ttu-id="917e9-171">While you develop your add-in, Visual Studio hosts the web application on your local IIS server.</span><span class="sxs-lookup"><span data-stu-id="917e9-171">While you develop your add-in, Visual Studio hosts the web application on your local IIS server.</span></span> <span data-ttu-id="917e9-172">When you're ready to publish the add-in, you'll need to deploy this web application project to a web server.</span><span class="sxs-lookup"><span data-stu-id="917e9-172">When you're ready to publish the add-in, you'll need to deploy this web application project to a web server.</span></span>|

### <a name="update-the-code"></a><span data-ttu-id="917e9-173">更新代码</span><span class="sxs-lookup"><span data-stu-id="917e9-173">Update the code</span></span>

1. <span data-ttu-id="917e9-174">**MessageRead.html** 指定将在加载项的任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="917e9-174">**MessageRead.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="917e9-175">在 **MessageRead.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="917e9-175">In **MessageRead.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```HTML
    <body class="ms-font-m ms-welcome">
        <div class="ms-Fabric content-main">
            <h1 class="ms-font-xxl">Message properties</h1>
            <table class="ms-Table ms-Table--selectable">
                <thead>
                    <tr>
                        <th>Property</th>
                        <th>Value</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td><strong>Id</strong></td>
                        <td class="prop-val"><code><label id="item-id"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Subject</strong></td>
                        <td class="prop-val"><code><label id="item-subject"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>Message Id</strong></td>
                        <td class="prop-val"><code><label id="item-internetMessageId"></label></code></td>
                    </tr>
                    <tr>
                        <td><strong>From</strong></td>
                        <td class="prop-val"><code><label id="item-from"></label></code></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </body>
    ```

1. <span data-ttu-id="917e9-176">打开 Web 应用项目的根文件夹中的文件“MessageRead.js”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="917e9-176">Open the file **MessageRead.js** in the root of the web application project.</span></span> <span data-ttu-id="917e9-177">此文件指定的是加载项脚本。</span><span class="sxs-lookup"><span data-stu-id="917e9-177">This file specifies the script for the add-in.</span></span> <span data-ttu-id="917e9-178">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="917e9-178">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.onReady(function () {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                loadItemProps(Office.context.mailbox.item);
            });
        });

        function loadItemProps(item) {
            // Write message property values to the task pane
            $('#item-id').text(item.itemId);
            $('#item-subject').text(item.subject);
            $('#item-internetMessageId').text(item.internetMessageId);
            $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        }
    })();
    ```

1. <span data-ttu-id="917e9-179">打开 Web 应用项目的根文件夹中的文件“MessageRead.css”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="917e9-179">Open the file **MessageRead.css** in the root of the web application project.</span></span> <span data-ttu-id="917e9-180">此文件指定的是加载项自定义样式。</span><span class="sxs-lookup"><span data-stu-id="917e9-180">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="917e9-181">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="917e9-181">Replace the entire contents with the following code and save the file.</span></span>

    ```CSS
    html,
    body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    td.prop-val {
        word-break: break-all;
    }

    .content-main {
        margin: 10px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="917e9-182">更新清单</span><span class="sxs-lookup"><span data-stu-id="917e9-182">Update the manifest</span></span>

1. <span data-ttu-id="917e9-183">Open the XML manifest file in the Add-in project.</span><span class="sxs-lookup"><span data-stu-id="917e9-183">Open the XML manifest file in the Add-in project.</span></span> <span data-ttu-id="917e9-184">This file defines the add-in's settings and capabilities.</span><span class="sxs-lookup"><span data-stu-id="917e9-184">This file defines the add-in's settings and capabilities.</span></span>

1. <span data-ttu-id="917e9-185">The `ProviderName` element has a placeholder value.</span><span class="sxs-lookup"><span data-stu-id="917e9-185">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="917e9-186">Replace it with your name.</span><span class="sxs-lookup"><span data-stu-id="917e9-186">Replace it with your name.</span></span>

1. <span data-ttu-id="917e9-187">`DisplayName` 元素的 `DefaultValue` 属性具有占位符。</span><span class="sxs-lookup"><span data-stu-id="917e9-187">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="917e9-188">将其替换为 `My Office Add-in`。</span><span class="sxs-lookup"><span data-stu-id="917e9-188">Replace it with `My Office Add-in`.</span></span>

1. <span data-ttu-id="917e9-189">`Description` 元素的 `DefaultValue` 属性具有占位符。</span><span class="sxs-lookup"><span data-stu-id="917e9-189">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="917e9-190">将其替换为 `My First Outlook add-in`。</span><span class="sxs-lookup"><span data-stu-id="917e9-190">Replace it with `My First Outlook add-in`.</span></span>

1. <span data-ttu-id="917e9-191">保存文件。</span><span class="sxs-lookup"><span data-stu-id="917e9-191">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="My First Outlook add-in"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="917e9-192">试用</span><span class="sxs-lookup"><span data-stu-id="917e9-192">Try it out</span></span>

1. <span data-ttu-id="917e9-193">在 Visual Studio 中，按 F5 或选择“开始”\*\*\*\* 按钮测试新建的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="917e9-193">Using Visual Studio, test the newly created Outlook add-in by pressing F5 or choosing the **Start** button.</span></span> <span data-ttu-id="917e9-194">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="917e9-194">The add-in will be hosted locally on IIS.</span></span>

1. <span data-ttu-id="917e9-195">在“连接到 Exchange 电子邮件帐户”\*\*\*\* 对话框中，输入你的 [Microsoft 帐户](https://account.microsoft.com/account)的电子邮件地址和密码，然后选择“连接”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="917e9-195">In the **Connect to Exchange email account** dialog box, enter the email address and password for your [Microsoft account](https://account.microsoft.com/account) and then choose **Connect**.</span></span> <span data-ttu-id="917e9-196">如果 Outlook.com 登录页是在浏览器中打开，请使用先前输入的相同凭据登录电子邮件帐户。</span><span class="sxs-lookup"><span data-stu-id="917e9-196">When the Outlook.com login page opens in a browser, sign in to your email account with the same credentials as you entered previously.</span></span>

    > [!NOTE]
    > <span data-ttu-id="917e9-197">如果“**连接到 Exchange 电子邮件帐户**”对话框重复提示你登录，则可能已对你 Microsoft 365 租户上的帐户禁用基本身份验证。</span><span class="sxs-lookup"><span data-stu-id="917e9-197">If the **Connect to Exchange email account** dialog box repeatedly prompts you to sign in, Basic Auth may be disabled for accounts on your Microsoft 365 tenant.</span></span> <span data-ttu-id="917e9-198">若要测试此加载项，请改用 [Microsoft 帐户](https://account.microsoft.com/account)登录。</span><span class="sxs-lookup"><span data-stu-id="917e9-198">To test this add-in, sign in using a [Microsoft account](https://account.microsoft.com/account) instead.</span></span>

1. <span data-ttu-id="917e9-199">在 Outlook 网页版中，选择或打开邮件。</span><span class="sxs-lookup"><span data-stu-id="917e9-199">In Outlook on the web, select or open a message.</span></span>

1. <span data-ttu-id="917e9-200">在邮件中，查找包含加载项按钮的溢出菜单的省略号。</span><span class="sxs-lookup"><span data-stu-id="917e9-200">Within the message, locate the ellipsis for the overflow menu containing the add-in's button.</span></span>

    ![Outlook 网页版中邮件窗口的屏幕截图，其中突出显示省略号](../images/quick-start-button-owa-1.png)

1. <span data-ttu-id="917e9-202">在 "溢出" 菜单中，找到加载项的按钮。</span><span class="sxs-lookup"><span data-stu-id="917e9-202">Within the overflow menu, locate the add-in's button.</span></span>

    ![Outlook 网页版中邮件窗口的屏幕截图，其中突出显示加载项按钮](../images/quick-start-button-owa-2.png)

1. <span data-ttu-id="917e9-204">单击此按钮，打开加载项的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="917e9-204">Click the button to open the add-in's task pane.</span></span>

    ![Outlook 网页版中加载项任务窗格的屏幕截图，其中显示邮件属性](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > <span data-ttu-id="917e9-206">如果任务窗格未加载，请尝试通过在同一台计算机上的浏览器中打开它来进行验证。</span><span class="sxs-lookup"><span data-stu-id="917e9-206">If the task pane doesn't load, try to verify by opening it in a browser on the same machine.</span></span>

### <a name="next-steps"></a><span data-ttu-id="917e9-207">后续步骤</span><span class="sxs-lookup"><span data-stu-id="917e9-207">Next steps</span></span>

<span data-ttu-id="917e9-208">祝贺！已成功创建首个 Outlook 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="917e9-208">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="917e9-209">接下来，了解有关[使用 Visual Studio 开发 Office 加载项](../develop/develop-add-ins-visual-studio.md)的详细信息。</span><span class="sxs-lookup"><span data-stu-id="917e9-209">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---
