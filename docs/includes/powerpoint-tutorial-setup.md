<span data-ttu-id="96387-101">在本教程中，请先设置开发项目。</span><span class="sxs-lookup"><span data-stu-id="96387-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="96387-102">此为 PowerPoint 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="96387-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="96387-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="96387-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="96387-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="96387-104">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="setup"></a><span data-ttu-id="96387-105">设置</span><span class="sxs-lookup"><span data-stu-id="96387-105">Setup</span></span>

<span data-ttu-id="96387-106">在本教程中，将使用 Visual Studio 创建加载项。</span><span class="sxs-lookup"><span data-stu-id="96387-106">In this tutorial, you'll create an add-in using Visual Studio.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="96387-107">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="96387-107">Create the add-in project</span></span>

1. <span data-ttu-id="96387-108">在 Visual Studio 菜单栏中，依次选择“文件”**** > “新建”**** > “项目”****。</span><span class="sxs-lookup"><span data-stu-id="96387-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="96387-109">在“Visual C#”**** 或“Visual Basic”**** 下的项目类型列表中，展开“Office/SharePoint”****，选择“加载项”****，再选择“PowerPoint Web 加载项”**** 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="96387-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="96387-110">将项目命名为“HelloWorld”****，再选择“确定”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="96387-110">Name the project **HelloWorld**, and then choose the **OK** button.</span></span>

4. <span data-ttu-id="96387-111">在“创建 Office 加载项”**** 对话框窗口中，选择“将新功能添加到 PowerPoint”****，再选择“完成”**** 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="96387-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="96387-p102">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”**** 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="96387-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![PowerPoint 教程 - 显示 HelloWorld 解决方案中 2 个项目的 Visual Studio 解决方案资源管理器窗口](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="96387-115">探索 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="96387-115">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="96387-116">更新代码</span><span class="sxs-lookup"><span data-stu-id="96387-116">Update code</span></span> 

<span data-ttu-id="96387-117">请按照下面的步骤编辑加载项代码，以创建在本教程后续步骤中实现加载项功能的框架。</span><span class="sxs-lookup"><span data-stu-id="96387-117">Edit the add-in code as follows, to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="96387-118">**Home.html** 指定在加载项任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="96387-118">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="96387-119">在 **Home.html** 文件中，查找包含 `id="content-main"` 的 **div**，并将找到的整个 **div** 替换为以下标记，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="96387-119">In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. <span data-ttu-id="96387-120">打开 Web 应用程序项目根目录中的文件 **Home.js**。</span><span class="sxs-lookup"><span data-stu-id="96387-120">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="96387-121">此文件指定加载项脚本。</span><span class="sxs-lookup"><span data-stu-id="96387-121">This file specifies the script for the add-in.</span></span> <span data-ttu-id="96387-122">将整个内容替换为下列代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="96387-122">Replace the entire contents with the following code and save the file.</span></span>

    ```javascript
    (function () {
        "use strict";

        var messageBanner;

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        };

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```
