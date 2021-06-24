---
title: 生成首个 Outlook 加载项
description: 了解如何使用 Office JS API 生成简单的 Outlook 任务窗格加载项。
ms.date: 06/10/2021
ms.prod: outlook
localization_priority: Priority
ms.openlocfilehash: 75a2b4d5a3ec995971b6cca1e47c77572d274555
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076600"
---
# <a name="build-your-first-outlook-add-in"></a><span data-ttu-id="142f9-103">生成首个 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="142f9-103">Build your first Outlook add-in</span></span>

<span data-ttu-id="142f9-104">本文将逐步介绍如何生成显示选定邮件的至少一个属性的 Outlook 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="142f9-104">In this article, you'll walk through the process of building an Outlook task pane add-in that displays at least one property of a selected message.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="142f9-105">创建加载项</span><span class="sxs-lookup"><span data-stu-id="142f9-105">Create the add-in</span></span>

<span data-ttu-id="142f9-106">可以使用[适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)或 Visual Studio 创建 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="142f9-106">You can create an Office Add-in by using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) or Visual Studio.</span></span> <span data-ttu-id="142f9-107">Yeoman 生成器将创建一个可通过 Visual Studio Code 或任何其他编辑器管理的 Node.js 项目，而 Visual Studio 将创建一个 Visual Studio 解决方案。</span><span class="sxs-lookup"><span data-stu-id="142f9-107">The Yeoman generator creates a Node.js project that can be managed with Visual Studio Code or any other editor, whereas Visual Studio creates a Visual Studio solution.</span></span> <span data-ttu-id="142f9-108">选择适合于想要使用的方法的选项卡，然后按照说明创建加载项并在本地测试。</span><span class="sxs-lookup"><span data-stu-id="142f9-108">Select the tab for the one you'd like to use and then follow the instructions to create your add-in and test it locally.</span></span>

# <a name="yeoman-generator"></a>[<span data-ttu-id="142f9-109">Yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="142f9-109">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="142f9-110">先决条件</span><span class="sxs-lookup"><span data-stu-id="142f9-110">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]

- <span data-ttu-id="142f9-111">[Node.js](https://nodejs.org/)（最新的 [LTS](https://nodejs.org/about/releases) 版本）</span><span class="sxs-lookup"><span data-stu-id="142f9-111">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

- <span data-ttu-id="142f9-112">最新版本的 [Yeoman](https://github.com/yeoman/yo) 和[适用于 Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="142f9-112">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="142f9-113">即便先前已安装了 Yeoman 生成器，我们还是建议你通过 npm 将包更新为最新版本。</span><span class="sxs-lookup"><span data-stu-id="142f9-113">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="142f9-114">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="142f9-114">Create the add-in project</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="142f9-115">**选择项目类型** - `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="142f9-115">**Choose a project type** - `Office Add-in Task Pane project`</span></span>

    - <span data-ttu-id="142f9-116">**选择脚本类型** - `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="142f9-116">**Choose a script type** - `JavaScript`</span></span>

    - <span data-ttu-id="142f9-117">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="142f9-117">**What do you want to name your add-in?**</span></span> - `My Office Add-in`

    - <span data-ttu-id="142f9-118">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="142f9-118">**Which Office client application would you like to support?**</span></span> - `Outlook`

    ![显示命令行界面中 Yeoman 生成器的提示和回答的屏幕截图。](../images/yo-office-outlook-1.png)

    <span data-ttu-id="142f9-120">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="142f9-120">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. <span data-ttu-id="142f9-121">导航到 Web 应用程序项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="142f9-121">Navigate to the root folder of the web application project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="explore-the-project"></a><span data-ttu-id="142f9-122">浏览项目</span><span class="sxs-lookup"><span data-stu-id="142f9-122">Explore the project</span></span>

<span data-ttu-id="142f9-123">使用 Yeoman 生成器创建的加载项项目包含适合于基础任务窗格加载项的示例代码。</span><span class="sxs-lookup"><span data-stu-id="142f9-123">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span>

- <span data-ttu-id="142f9-124">项目根目录中的 **./manifest.xml** 文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="142f9-124">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="142f9-125">**./src/taskpane/taskpane.html** 文件包含组成任务窗格的 HTML。</span><span class="sxs-lookup"><span data-stu-id="142f9-125">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="142f9-126">**./src/taskpane/taskpane.css** 文件包含应用于任务窗格中的内容的 CSS。</span><span class="sxs-lookup"><span data-stu-id="142f9-126">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="142f9-127">**./src/taskpane/taskpane.js** 文件包含用于加快任务窗格与 Outlook 之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="142f9-127">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and Outlook.</span></span>

### <a name="update-the-code"></a><span data-ttu-id="142f9-128">更新代码</span><span class="sxs-lookup"><span data-stu-id="142f9-128">Update the code</span></span>

1. <span data-ttu-id="142f9-129">在代码编辑器中，打开文件 **./src/taskpane/taskpane.html** 并将整个 `<main>` 元素（位于 `<body>` 元素中）替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="142f9-129">In your code editor, open the file **./src/taskpane/taskpane.html** and replace the entire `<main>` element (within the `<body>` element) with the following markup.</span></span> <span data-ttu-id="142f9-130">此新标记将添加标签，其中 **./src/taskpane/taskpane.js** 中的脚本将写入数据。</span><span class="sxs-lookup"><span data-stu-id="142f9-130">This new markup adds a label where the script in **./src/taskpane/taskpane.js** will write data.</span></span>

    ```html
    <main id="app-body" class="ms-welcome__main" style="display: none;">
        <h2 class="ms-font-xl"> Discover what Office Add-ins can do for you today! </h2>
        <p><label id="item-subject"></label></p>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. <span data-ttu-id="142f9-131">在代码编辑器中，打开文件 **./src/taskpane/taskpane.js** 并在 `run` 函数中添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="142f9-131">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="142f9-132">此代码使用 Office JavaScript API 获取当前邮件的引用并将其 `subject` 属性值写入任务窗格。</span><span class="sxs-lookup"><span data-stu-id="142f9-132">This code uses the Office JavaScript API to get a reference to the current message and write its `subject` property value to the task pane.</span></span>

    ```js
    // Get a reference to the current message
    var item = Office.context.mailbox.item;

    // Write message property value to the task pane
    document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
    ```

### <a name="try-it-out"></a><span data-ttu-id="142f9-133">试用</span><span class="sxs-lookup"><span data-stu-id="142f9-133">Try it out</span></span>

> [!NOTE]
> <span data-ttu-id="142f9-134">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="142f9-134">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="142f9-135">如果系统在运行以下命令后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="142f9-135">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="142f9-136">你可能还必须以管理员身份运行命令提示符或终端才能进行更改。</span><span class="sxs-lookup"><span data-stu-id="142f9-136">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

1. <span data-ttu-id="142f9-137">在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="142f9-137">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="142f9-138">如果运行此命令，本地 Web 服务器将启动（如果尚未运行），并将旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="142f9-138">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="142f9-139">按照[旁加载 Outlook 加载项以供测试](../outlook/sideload-outlook-add-ins-for-testing.md)中的说明操作，旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="142f9-139">Follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="142f9-140">在 Outlook 中，在[阅读窗格](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0)中查看邮件，或在其自己的窗口中打开邮件。</span><span class="sxs-lookup"><span data-stu-id="142f9-140">In Outlook, view a message in the [Reading Pane](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0), or open the message in its own window.</span></span>

1. <span data-ttu-id="142f9-141">选择“**主页**”选项卡（或“**邮件**”选项卡，如果在新窗口中打开了邮件），然后选择功能区的“**显示任务窗格**”按钮以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="142f9-141">Choose the **Home** tab (or the **Message** tab if you opened the message in a new window), and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![显示 Outlook 中突出显示“加载项”功能区按钮的消息窗口屏幕截图。](../images/quick-start-button-1.png)

    > [!NOTE]
    > <span data-ttu-id="142f9-143">如果在任务窗格中收到错误“我们无法从本地主机打开此加载项”，请按照[疑难解答文章中](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost)中所述步骤进行操作。</span><span class="sxs-lookup"><span data-stu-id="142f9-143">If you receive the error "We can't open this add-in from localhost" in the task pane, follow the steps outlined in the [troubleshooting article](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).</span></span>

1. <span data-ttu-id="142f9-144">滚动至任务窗格的底部并选择“**运行**”链接，将邮件主题写入任务窗格。</span><span class="sxs-lookup"><span data-stu-id="142f9-144">Scroll to the bottom of the task pane and choose the **Run** link to write the message subject to the task pane.</span></span>

    ![显示突出显示“运行”链接的加载项任务窗格屏幕截图。](../images/quick-start-task-pane-2.png)

    ![加载项任务窗格的屏幕截图，其中显示邮件主题。](../images/quick-start-task-pane-3.png)

### <a name="next-steps"></a><span data-ttu-id="142f9-147">后续步骤</span><span class="sxs-lookup"><span data-stu-id="142f9-147">Next steps</span></span>

<span data-ttu-id="142f9-148">祝贺！已成功创建首个 Outlook 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="142f9-148">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="142f9-149">接下来，将继续学习 [Outlook 加载项教程](../tutorials/outlook-tutorial.md)，详细了解 Outlook 加载项的功能，以及如何生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="142f9-149">Next, learn more about the capabilities of an Outlook add-in and build a more complex add-in by following along with the [Outlook add-in tutorial](../tutorials/outlook-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="142f9-150">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="142f9-150">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="142f9-151">先决条件</span><span class="sxs-lookup"><span data-stu-id="142f9-151">Prerequisites</span></span>

- <span data-ttu-id="142f9-152">安装了 **Office/SharePoint 开发** 工作负载的 [Visual Studio 2019](https://www.visualstudio.com/vs/)</span><span class="sxs-lookup"><span data-stu-id="142f9-152">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="142f9-153">如果之前已安装 Visual Studio 2019，请 [使用 Visual Studio 安装程序](/visualstudio/install/modify-visual-studio)，以确保安装 **Office/SharePoint 开发** 工作负载。</span><span class="sxs-lookup"><span data-stu-id="142f9-153">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span>

- <span data-ttu-id="142f9-154">Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="142f9-154">Microsoft 365</span></span>

    > [!NOTE]
    > <span data-ttu-id="142f9-155">如果没有 Microsoft 365 订阅，可以通过注册 [Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个免费订阅。</span><span class="sxs-lookup"><span data-stu-id="142f9-155">If you do not have a Microsoft 365 subscription, you can get a free one by signing up for the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="142f9-156">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="142f9-156">Create the add-in project</span></span>

1. <span data-ttu-id="142f9-157">在 Visual Studio 菜单栏中，依次选择“文件” > “新建” > “项目”。</span><span class="sxs-lookup"><span data-stu-id="142f9-157">On the Visual Studio menu bar, choose **File** > **New** > **Project**.</span></span>

1. <span data-ttu-id="142f9-158">在“Visual C#”或“Visual Basic”下的项目类型列表中，展开“Office/SharePoint”，选择“加载项”，然后选择“Outlook Web 加载项”作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="142f9-158">In the list of project types under **Visual C#** or **Visual Basic**, expand **Office/SharePoint**, choose **Add-ins**, and then choose **Outlook Web Add-in** as the project type.</span></span>

1. <span data-ttu-id="142f9-159">命名此项目，再选择“确定”。</span><span class="sxs-lookup"><span data-stu-id="142f9-159">Name the project, and then choose **OK**.</span></span>

1. <span data-ttu-id="142f9-160">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”中。</span><span class="sxs-lookup"><span data-stu-id="142f9-160">Visual Studio creates a solution and its two projects appear in **Solution Explorer**.</span></span> <span data-ttu-id="142f9-161">**MessageRead.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="142f9-161">The **MessageRead.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="142f9-162">浏览 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="142f9-162">Explore the Visual Studio solution</span></span>

<span data-ttu-id="142f9-163">在用户完成向导后，Visual Studio 会创建一个包含两个项目的解决方案。</span><span class="sxs-lookup"><span data-stu-id="142f9-163">When you've completed the wizard, Visual Studio creates a solution that contains two projects.</span></span>

|<span data-ttu-id="142f9-164">**项目**</span><span class="sxs-lookup"><span data-stu-id="142f9-164">**Project**</span></span>|<span data-ttu-id="142f9-165">**说明**</span><span class="sxs-lookup"><span data-stu-id="142f9-165">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="142f9-166">加载项项目</span><span class="sxs-lookup"><span data-stu-id="142f9-166">Add-in project</span></span>|<span data-ttu-id="142f9-p108">仅包含 XML 清单文件，该文件包含描述加载项的全部设置。这些设置有助于 Office 应用程序确定应何时激活加载项，以及应在何处显示加载项。Visual Studio 会为你生成此文件的内容，因此你可以直接运行相应项目并使用加载项。可以通过修改该 XML 文件随时更改这些设置。</span><span class="sxs-lookup"><span data-stu-id="142f9-p108">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office application determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You can change these settings any time by modifying the XML file.</span></span>|
|<span data-ttu-id="142f9-171">Web 应用项目</span><span class="sxs-lookup"><span data-stu-id="142f9-171">Web application project</span></span>|<span data-ttu-id="142f9-p109">包含加载项的内容页，包括开发 Office 感知 HTML 和 JavaScript 页面所需的全部文件和文件引用。开发加载项时，Visual Studio 在本地 IIS 服务器上托管 Web 应用。准备好发布加载项后，需要将此 Web 应用项目部署到 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="142f9-p109">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish the add-in, you'll need to deploy this web application project to a web server.</span></span>|

### <a name="update-the-code"></a><span data-ttu-id="142f9-175">更新代码</span><span class="sxs-lookup"><span data-stu-id="142f9-175">Update the code</span></span>

1. <span data-ttu-id="142f9-176">**MessageRead.html** 指定将在加载项的任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="142f9-176">**MessageRead.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="142f9-177">在 **MessageRead.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="142f9-177">In **MessageRead.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

1. <span data-ttu-id="142f9-178">打开 Web 应用项目的根文件夹中的文件“MessageRead.js”。</span><span class="sxs-lookup"><span data-stu-id="142f9-178">Open the file **MessageRead.js** in the root of the web application project.</span></span> <span data-ttu-id="142f9-179">此文件指定的是加载项脚本。</span><span class="sxs-lookup"><span data-stu-id="142f9-179">This file specifies the script for the add-in.</span></span> <span data-ttu-id="142f9-180">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="142f9-180">Replace the entire contents with the following code and save the file.</span></span>

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

1. <span data-ttu-id="142f9-181">打开 Web 应用项目的根文件夹中的文件“MessageRead.css”。</span><span class="sxs-lookup"><span data-stu-id="142f9-181">Open the file **MessageRead.css** in the root of the web application project.</span></span> <span data-ttu-id="142f9-182">此文件指定的是加载项自定义样式。</span><span class="sxs-lookup"><span data-stu-id="142f9-182">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="142f9-183">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="142f9-183">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="142f9-184">更新清单</span><span class="sxs-lookup"><span data-stu-id="142f9-184">Update the manifest</span></span>

1. <span data-ttu-id="142f9-p113">打开加载项项目中的 XML 清单文件。 此文件定义的是加载项设置和功能。</span><span class="sxs-lookup"><span data-stu-id="142f9-p113">Open the XML manifest file in the Add-in project. This file defines the add-in's settings and capabilities.</span></span>

1. <span data-ttu-id="142f9-p114">`ProviderName` 元素具有占位符值。 将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="142f9-p114">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

1. <span data-ttu-id="142f9-189">`DisplayName` 元素的 `DefaultValue` 属性具有占位符。</span><span class="sxs-lookup"><span data-stu-id="142f9-189">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="142f9-190">将其替换为 `My Office Add-in`。</span><span class="sxs-lookup"><span data-stu-id="142f9-190">Replace it with `My Office Add-in`.</span></span>

1. <span data-ttu-id="142f9-191">`Description` 元素的 `DefaultValue` 属性具有占位符。</span><span class="sxs-lookup"><span data-stu-id="142f9-191">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="142f9-192">将其替换为 `My First Outlook add-in`。</span><span class="sxs-lookup"><span data-stu-id="142f9-192">Replace it with `My First Outlook add-in`.</span></span>

1. <span data-ttu-id="142f9-193">保存文件。</span><span class="sxs-lookup"><span data-stu-id="142f9-193">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="My First Outlook add-in"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="142f9-194">试用</span><span class="sxs-lookup"><span data-stu-id="142f9-194">Try it out</span></span>

1. <span data-ttu-id="142f9-195">在 Visual Studio 中，按 F5 或选择“开始”按钮测试新建的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="142f9-195">Using Visual Studio, test the newly created Outlook add-in by pressing F5 or choosing the **Start** button.</span></span> <span data-ttu-id="142f9-196">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="142f9-196">The add-in will be hosted locally on IIS.</span></span>

1. <span data-ttu-id="142f9-197">在“连接到 Exchange 电子邮件帐户”对话框中，输入你的 [Microsoft 帐户](https://account.microsoft.com/account)的电子邮件地址和密码，然后选择“连接”。</span><span class="sxs-lookup"><span data-stu-id="142f9-197">In the **Connect to Exchange email account** dialog box, enter the email address and password for your [Microsoft account](https://account.microsoft.com/account) and then choose **Connect**.</span></span> <span data-ttu-id="142f9-198">如果 Outlook.com 登录页是在浏览器中打开，请使用先前输入的相同凭据登录电子邮件帐户。</span><span class="sxs-lookup"><span data-stu-id="142f9-198">When the Outlook.com login page opens in a browser, sign in to your email account with the same credentials as you entered previously.</span></span>

    > [!NOTE]
    > <span data-ttu-id="142f9-199">如果“**连接到 Exchange 电子邮件帐户**”对话框重复提示你登录或收到未获授权错误，则 Microsoft 365 租户上的帐户可能禁用基本身份验证。</span><span class="sxs-lookup"><span data-stu-id="142f9-199">If the **Connect to Exchange email account** dialog box repeatedly prompts you to sign in or you receive an error that you are unauthorized, Basic Auth may be disabled for accounts on your Microsoft 365 tenant.</span></span> <span data-ttu-id="142f9-200">要测试此加载项，请在 Web 加载项项目属性对话框中将 **使用多重身份验证** 属性设置为 True 后再次尝试登录，或改为使用 [Microsoft 帐户](https://account.microsoft.com/account) 登录。</span><span class="sxs-lookup"><span data-stu-id="142f9-200">To test this add-in, try signing in again after setting the **Use multi-factor auth** property to True in the Web Add-in project properties dialog, or sign in using a [Microsoft account](https://account.microsoft.com/account) instead.</span></span>

1. <span data-ttu-id="142f9-201">在 Outlook 网页版中，选择或打开邮件。</span><span class="sxs-lookup"><span data-stu-id="142f9-201">In Outlook on the web, select or open a message.</span></span>

1. <span data-ttu-id="142f9-202">在邮件中，查找包含加载项按钮的溢出菜单的省略号。</span><span class="sxs-lookup"><span data-stu-id="142f9-202">Within the message, locate the ellipsis for the overflow menu containing the add-in's button.</span></span>

    ![Outlook 网页版中邮件窗口的屏幕截图，其中突出显示省略号。](../images/quick-start-button-owa-1.png)

1. <span data-ttu-id="142f9-204">在"溢出"菜单中，找到加载项的按钮。</span><span class="sxs-lookup"><span data-stu-id="142f9-204">Within the overflow menu, locate the add-in's button.</span></span>

    ![Outlook 网页版中邮件窗口的屏幕截图，其中突出显示加载项按钮。](../images/quick-start-button-owa-2.png)

1. <span data-ttu-id="142f9-206">单击此按钮，打开加载项的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="142f9-206">Click the button to open the add-in's task pane.</span></span>

    ![Outlook 网页版中加载项任务窗格的屏幕截图，其中显示邮件属性。](../images/quick-start-task-pane-owa-1.png)

    > [!NOTE]
    > <span data-ttu-id="142f9-208">如果任务窗格未加载，请尝试通过在同一台计算机上的浏览器中打开它来进行验证。</span><span class="sxs-lookup"><span data-stu-id="142f9-208">If the task pane doesn't load, try to verify by opening it in a browser on the same machine.</span></span>

### <a name="next-steps"></a><span data-ttu-id="142f9-209">后续步骤</span><span class="sxs-lookup"><span data-stu-id="142f9-209">Next steps</span></span>

<span data-ttu-id="142f9-210">祝贺！已成功创建首个 Outlook 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="142f9-210">Congratulations, you've successfully created your first Outlook task pane add-in!</span></span> <span data-ttu-id="142f9-211">接下来，了解有关[使用 Visual Studio 开发 Office 加载项](../develop/develop-add-ins-visual-studio.md)的详细信息。</span><span class="sxs-lookup"><span data-stu-id="142f9-211">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---
