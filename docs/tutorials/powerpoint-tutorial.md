---
title: PowerPoint 加载项教程
description: 在本教程中，将生成 PowerPoint 加载项，用于插入图像、插入文本、获取幻灯片元数据，以及在幻灯片之间导航。
ms.date: 12/31/2018
ms.prod: powerpoint
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 9f5e30929c0881c0216b7ca77fbfa4b989fabc6e
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742413"
---
# <a name="tutorial-create-a-powerpoint-task-pane-add-in"></a><span data-ttu-id="dda27-103">教程：创建 PowerPoint 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="dda27-103">Tutorial: Create a PowerPoint task pane add-in</span></span>

<span data-ttu-id="dda27-104">在本教程中，将使用 Visual Studio 创建 PowerPoint 任务窗格加载项：</span><span class="sxs-lookup"><span data-stu-id="dda27-104">In this tutorial, you'll use Visual Studio to create an PowerPoint task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="dda27-105">向幻灯片添加一天中的[必应](https://www.bing.com)照片</span><span class="sxs-lookup"><span data-stu-id="dda27-105">Adds the [Bing](https://www.bing.com) photo of the day to a slide</span></span>
> * <span data-ttu-id="dda27-106">向幻灯片添加文本</span><span class="sxs-lookup"><span data-stu-id="dda27-106">Adds text to a slide</span></span>
> * <span data-ttu-id="dda27-107">获取幻灯片元数据</span><span class="sxs-lookup"><span data-stu-id="dda27-107">Gets slide metadata</span></span>
> * <span data-ttu-id="dda27-108">在幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="dda27-108">Navigates between slides</span></span>

## <a name="prerequisites"></a><span data-ttu-id="dda27-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="dda27-109">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="dda27-110">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="dda27-110">Create your add-in project</span></span>

<span data-ttu-id="dda27-111">完成以下步骤以使用 Visual Studio 创建 PowerPoint 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="dda27-111">Complete the following steps to create a PowerPoint add-in project using Visual Studio.</span></span>

1. <span data-ttu-id="dda27-112">在 Visual Studio 菜单栏中，依次选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dda27-112">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="dda27-113">在“Visual C#”\*\*\*\* 或“Visual Basic”\*\*\*\* 下的项目类型列表中，展开“Office/SharePoint”\*\*\*\*，选择“加载项”\*\*\*\*，再选择“PowerPoint Web 加载项”\*\*\*\* 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="dda27-113">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="dda27-114">将项目命名为“HelloWorld”\*\*\*\*，再选择“确定”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-114">Name the project **HelloWorld**, and then choose the **OK** button.</span></span>

4. <span data-ttu-id="dda27-115">在“创建 Office 加载项”\*\*\*\* 对话框窗口中，选择“将新功能添加到 PowerPoint”\*\*\*\*，再选择“完成”\*\*\*\* 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="dda27-115">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="dda27-p101">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="dda27-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![PowerPoint 教程 - 显示 HelloWorld 解决方案中 2 个项目的 Visual Studio 解决方案资源管理器窗口](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="dda27-119">探索 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="dda27-119">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="dda27-120">更新代码</span><span class="sxs-lookup"><span data-stu-id="dda27-120">Update code</span></span> 

<span data-ttu-id="dda27-121">请按照下面的步骤编辑加载项代码，以创建在本教程后续步骤中实现加载项功能的框架。</span><span class="sxs-lookup"><span data-stu-id="dda27-121">Edit the add-in code as follows to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="dda27-122">**Home.html** 指定在加载项的任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="dda27-122">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="dda27-123">在 **Home.html** 文件中，查找包含 `id="content-main"` 的 **div**，并将找到的整个 **div** 替换为以下标记，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="dda27-123">In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

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

2. <span data-ttu-id="dda27-124">打开 Web 应用程序项目根目录中的文件 **Home.js**。</span><span class="sxs-lookup"><span data-stu-id="dda27-124">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="dda27-125">此文件指定加载项脚本。</span><span class="sxs-lookup"><span data-stu-id="dda27-125">This file specifies the script for the add-in.</span></span> <span data-ttu-id="dda27-126">将整个内容替换为下列代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="dda27-126">Replace the entire contents with the following code and save the file.</span></span>

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

## <a name="insert-an-image"></a><span data-ttu-id="dda27-127">插入图像</span><span class="sxs-lookup"><span data-stu-id="dda27-127">Insert an image</span></span>

<span data-ttu-id="dda27-128">完成以下步骤以添加用于检索一天中的[必应](https://www.bing.com)照片的代码，并将该图像插入幻灯片中。</span><span class="sxs-lookup"><span data-stu-id="dda27-128">Complete the following steps to add code that retrieves the [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.</span></span>

1. <span data-ttu-id="dda27-129">使用解决方案资源管理器，将 **Controllers** 新文件夹添加到 **HelloWorldWeb** 项目。</span><span class="sxs-lookup"><span data-stu-id="dda27-129">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![PowerPoint 教程 - 突出显示 HelloWorldWeb 目中 Controllers 文件夹的 Visual Studio 解决方案资源管理器窗口](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="dda27-131">右键单击“Controllers”\*\*\*\* 文件夹，并依次选择“添加”>“新基架项...”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="dda27-131">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="dda27-132">在“添加基架”\*\*\*\* 对话框窗口中，依次选择“Web API 2 控制器 - 空”\*\*\*\* 和“添加”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-132">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="dda27-133">在“添加控制器”\*\*\*\* 对话框窗口中，输入“PhotoController”\*\*\*\* 作为控制器名称，再选择“添加”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-133">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button.</span></span> <span data-ttu-id="dda27-134">此时，Visual Studio 创建并打开 **PhotoController.cs** 文件。</span><span class="sxs-lookup"><span data-stu-id="dda27-134">Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="dda27-135">将 **PhotoController.cs** 文件的全部内容替换为下列代码，以调用必应服务来检索 Base64 编码字符串形式的一天中照片。</span><span class="sxs-lookup"><span data-stu-id="dda27-135">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string.</span></span> <span data-ttu-id="dda27-136">使用 Office JavaScript API 将图像插入文档时，必须将图像数据指定为 Base64 编码字符串。</span><span class="sxs-lookup"><span data-stu-id="dda27-136">When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. <span data-ttu-id="dda27-137">在 **Home.html** 文件中，将 `TODO1` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="dda27-137">In the **Home.html** file, replace `TODO1` with the following markup.</span></span> <span data-ttu-id="dda27-138">此标记定义在加载项任务窗格内显示的“插入图像”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-138">This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="ms-Button ms-Button--primary" id="insert-image">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Insert Image</span>
        <span class="ms-Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="dda27-139">在 **Home.js** 文件中，将 `TODO1` 替换为下列代码，以分配“插入图像”\*\*\*\* 按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="dda27-139">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```javascript
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="dda27-140">在 **Home.js** 文件中，将 `TODO2` 替换为下列代码，以定义 **insertImage** 函数。</span><span class="sxs-lookup"><span data-stu-id="dda27-140">In the **Home.js** file, replace `TODO2` with the following code to define the **insertImage** function.</span></span> <span data-ttu-id="dda27-141">此函数从必应 Web 服务提取图像，再调用 `insertImageFromBase64String` 函数将相应图像插入文档。</span><span class="sxs-lookup"><span data-stu-id="dda27-141">This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

    ```javascript
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. <span data-ttu-id="dda27-142">在 **Home.js** 文件中，将 `TODO3` 替换为下列代码，以定义 `insertImageFromBase64String` 函数。</span><span class="sxs-lookup"><span data-stu-id="dda27-142">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function.</span></span> <span data-ttu-id="dda27-143">此函数使用 Office JavaScript API 将图像插入文档。</span><span class="sxs-lookup"><span data-stu-id="dda27-143">This function uses the Office JavaScript API to insert the image into the document.</span></span> <span data-ttu-id="dda27-144">注意：</span><span class="sxs-lookup"><span data-stu-id="dda27-144">Note:</span></span> 

    - <span data-ttu-id="dda27-145">`coercionType` 选项被指定为 `setSelectedDataAsyc` 请求的第二个参数，指明了要插入的数据的类型。</span><span class="sxs-lookup"><span data-stu-id="dda27-145">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsyc` request indicates the type of data being inserted.</span></span> 

    - <span data-ttu-id="dda27-146">`asyncResult` 对象封装 `setSelectedDataAsync` 请求的结果，包括状态和错误消息（如果请求失败的话）。</span><span class="sxs-lookup"><span data-stu-id="dda27-146">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

    ```javascript
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="dda27-147">测试加载项</span><span class="sxs-lookup"><span data-stu-id="dda27-147">Test the add-in</span></span>

1. <span data-ttu-id="dda27-p109">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 PowerPoint，以测试新建的 PowerPoint 加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="dda27-p109">Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="dda27-151">在 PowerPoint 中，选择功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="dda27-151">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="dda27-153">在任务窗格中，选择“插入图像”\*\*\*\* 按钮，将一天中的必应照片添加到当前幻灯片。</span><span class="sxs-lookup"><span data-stu-id="dda27-153">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![突出显示“插入图像”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="dda27-155">在 Visual Studio 中，按 **Shift + F5** 或选择“停止”\*\*\*\* 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="dda27-155">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="dda27-156">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="dda27-156">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="customize-user-interface-ui-elements"></a><span data-ttu-id="dda27-158">自定义用户界面 (UI) 元素</span><span class="sxs-lookup"><span data-stu-id="dda27-158">Customize User Interface (UI) elements</span></span>

<span data-ttu-id="dda27-159">完成以下步骤以添加用于自定义任务窗格 UI 的标记。</span><span class="sxs-lookup"><span data-stu-id="dda27-159">Complete the following steps to add markup that customizes the task pane UI.</span></span>

1. <span data-ttu-id="dda27-160">在 **Home.html** 文件中，将 `TODO2` 替换为以下标记，以将页眉部分和标题添加到任务窗格。</span><span class="sxs-lookup"><span data-stu-id="dda27-160">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane.</span></span> <span data-ttu-id="dda27-161">注意：</span><span class="sxs-lookup"><span data-stu-id="dda27-161">Note:</span></span>

    - <span data-ttu-id="dda27-162">以 `ms-` 开头的样式由 [Office UI Fabric](../design/office-ui-fabric.md) 进行定义，后者是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。</span><span class="sxs-lookup"><span data-stu-id="dda27-162">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365.</span></span> <span data-ttu-id="dda27-163">**Home.html** 文件包含对 Fabric 样式表的引用。</span><span class="sxs-lookup"><span data-stu-id="dda27-163">The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="dda27-164">在 **Home.html** 文件中，查找包含 `class="footer"` 的 **div**，并删除找到的整个 **div**，以从任务窗格中删除页脚部分。</span><span class="sxs-lookup"><span data-stu-id="dda27-164">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="dda27-165">测试加载项</span><span class="sxs-lookup"><span data-stu-id="dda27-165">Test the add-in</span></span>

1. <span data-ttu-id="dda27-166">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 PowerPoint，以测试 PowerPoint 加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-166">Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="dda27-167">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="dda27-167">The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="dda27-169">在 PowerPoint 中，选择功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="dda27-169">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="dda27-171">请注意，任务窗格现在包含页眉部分和标题，并且不再包含页脚部分。</span><span class="sxs-lookup"><span data-stu-id="dda27-171">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![突出显示“插入图像”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="dda27-173">在 Visual Studio 中，按 **Shift + F5** 或选择“停止”\*\*\*\* 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="dda27-173">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="dda27-174">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="dda27-174">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="insert-text"></a><span data-ttu-id="dda27-176">插入文本</span><span class="sxs-lookup"><span data-stu-id="dda27-176">Insert text</span></span>

<span data-ttu-id="dda27-177">完成以下步骤以添加用于将文本插入到标题幻灯片的代码，该幻灯片包含一天中的[必应](https://www.bing.com)照片。</span><span class="sxs-lookup"><span data-stu-id="dda27-177">Complete the following steps to add code that inserts text into the title slide which contains the [Bing](https://www.bing.com) photo of the day.</span></span>

1. <span data-ttu-id="dda27-178">在 **Home.html** 文件中，将 `TODO3` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="dda27-178">In the **Home.html** file, replace `TODO3` with the following markup.</span></span> <span data-ttu-id="dda27-179">此标记定义在加载项任务窗格内显示的“插入文本”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-179">This markup defines the **Insert Text** button that will appear within the add-in's task pane.</span></span>

    ```html
        <br /><br />
        <button class="ms-Button ms-Button--primary" id="insert-text">
            <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="ms-Button-label">Insert Text</span>
            <span class="ms-Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. <span data-ttu-id="dda27-180">在 **Home.js** 文件中，将 `TODO4` 替换为下列代码，以分配“插入文本”\*\*\*\* 按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="dda27-180">In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.</span></span>

    ```javascript
    $('#insert-text').click(insertText);
    ```

3. <span data-ttu-id="dda27-181">在 **Home.js** 文件中，将 `TODO5` 替换为下列代码，以定义 **insertText** 函数。</span><span class="sxs-lookup"><span data-stu-id="dda27-181">In the **Home.js** file, replace `TODO5` with the following code to define the **insertText** function.</span></span> <span data-ttu-id="dda27-182">此函数将文本插入当前幻灯片。</span><span class="sxs-lookup"><span data-stu-id="dda27-182">This function inserts text into the current slide.</span></span>

    ```javascript
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="dda27-183">测试加载项</span><span class="sxs-lookup"><span data-stu-id="dda27-183">Test the add-in</span></span>

1. <span data-ttu-id="dda27-184">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-184">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="dda27-185">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="dda27-185">The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="dda27-187">在 PowerPoint 中，选择功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="dda27-187">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="dda27-189">在任务窗格中，选择“插入图像”\*\*\*\* 按钮，将一天中的必应照片添加到当前幻灯片，再为包含标题文本框的幻灯片选择一种设计。</span><span class="sxs-lookup"><span data-stu-id="dda27-189">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.</span></span>

    ![突出显示“插入图像”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. <span data-ttu-id="dda27-191">将光标置于标题幻灯片上的文本框中，再选择任务窗格中的“插入文本”\*\*\*\* 按钮，向幻灯片添加文本。</span><span class="sxs-lookup"><span data-stu-id="dda27-191">Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.</span></span>

    ![突出显示“插入文本”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-text.png)


5. <span data-ttu-id="dda27-193">在 Visual Studio 中，按 **Shift + F5** 或选择“停止”\*\*\*\* 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="dda27-193">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="dda27-194">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="dda27-194">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="get-slide-metadata"></a><span data-ttu-id="dda27-196">获取幻灯片元数据</span><span class="sxs-lookup"><span data-stu-id="dda27-196">Get slide metadata</span></span>

<span data-ttu-id="dda27-197">完成以下步骤以添加用于检索所选幻灯片的元数据的代码。</span><span class="sxs-lookup"><span data-stu-id="dda27-197">Complete the following steps to add code that retrieves metadata for the selected slide.</span></span>

1. <span data-ttu-id="dda27-198">在 **Home.html** 文件中，将 `TODO4` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="dda27-198">In the **Home.html** file, replace `TODO4` with the following markup.</span></span> <span data-ttu-id="dda27-199">此标记定义在加载项任务窗格内显示的“获取幻灯片元数据”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-199">This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="get-slide-metadata">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Get Slide Metadata</span>
        <span class="ms-Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="dda27-200">在 **Home.js** 文件中，将 `TODO6` 替换为下列代码，以分配“获取幻灯片元数据”\*\*\*\* 按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="dda27-200">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```javascript
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="dda27-201">在 **Home.js** 文件中，将 `TODO7` 替换为下列代码，以定义 **getSlideMetadata** 函数。</span><span class="sxs-lookup"><span data-stu-id="dda27-201">In the **Home.js** file, replace `TODO7` with the following code to define the **getSlideMetadata** function.</span></span> <span data-ttu-id="dda27-202">此函数检索选定一张或多张幻灯片的元数据，并将它写入加载项任务窗格内的弹出对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="dda27-202">This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

    ```javascript
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="dda27-203">测试加载项</span><span class="sxs-lookup"><span data-stu-id="dda27-203">Test the add-in</span></span>

1. <span data-ttu-id="dda27-204">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-204">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="dda27-205">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="dda27-205">The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="dda27-207">在 PowerPoint 中，选择功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="dda27-207">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="dda27-209">在任务窗格中，选择“获取幻灯片元数据”\*\*\*\* 按钮，以获取选定幻灯片的元数据。</span><span class="sxs-lookup"><span data-stu-id="dda27-209">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide.</span></span> <span data-ttu-id="dda27-210">此时，幻灯片元数据写入到任务窗格底部的弹出对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="dda27-210">The slide metadata is written to the popup dialog window at the bottom of the task pane.</span></span> <span data-ttu-id="dda27-211">在此示例中，JSON 元数据中的 `slides` 数组包含一个对象，用于指定选定幻灯片的 `id`、`title` 和 `index`。</span><span class="sxs-lookup"><span data-stu-id="dda27-211">In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide.</span></span> <span data-ttu-id="dda27-212">如果在检索幻灯片元数据时选择了多张幻灯片，JSON 元数据中的 `slides` 数组会对每张选定幻灯片都包含一个对象。</span><span class="sxs-lookup"><span data-stu-id="dda27-212">If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![突出显示“获取幻灯片元数据”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="dda27-214">在 Visual Studio 中，按 **Shift + F5** 或选择“停止”\*\*\*\* 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="dda27-214">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="dda27-215">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="dda27-215">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="navigate-between-slides"></a><span data-ttu-id="dda27-217">在幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="dda27-217">Navigate between slides</span></span>

<span data-ttu-id="dda27-218">完成以下步骤以添加用于在文档幻灯片之间导航的代码。</span><span class="sxs-lookup"><span data-stu-id="dda27-218">Complete the following steps to add code that navigates between the slides of a document.</span></span>

1. <span data-ttu-id="dda27-219">在 **Home.html** 文件中，将 `TODO5` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="dda27-219">In the **Home.html** file, replace `TODO5` with the following markup.</span></span> <span data-ttu-id="dda27-220">此标记定义在加载项任务窗格内显示的四个导航按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-220">This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-first-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to First Slide</span>
        <span class="ms-Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-next-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Next Slide</span>
        <span class="ms-Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-previous-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Previous Slide</span>
        <span class="ms-Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-last-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Last Slide</span>
        <span class="ms-Button-description">Go to the last slide.</span>
    </button>
    ```

2. <span data-ttu-id="dda27-221">在 **Home.js** 文件中，将 `TODO8` 替换为下列代码，以分配四个导航按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="dda27-221">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```javascript
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="dda27-222">在 **Home.js** 文件中，将 `TODO9` 替换为下列代码，以定义导航函数。</span><span class="sxs-lookup"><span data-stu-id="dda27-222">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="dda27-223">以下各函数均使用 `goToByIdAsync` 函数，以根据幻灯片在文档中的位置（第一张、最后一张、上一张和下一张）选择幻灯片。</span><span class="sxs-lookup"><span data-stu-id="dda27-223">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, and next).</span></span>

    ```javascript
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="dda27-224">测试加载项</span><span class="sxs-lookup"><span data-stu-id="dda27-224">Test the add-in</span></span>

1. <span data-ttu-id="dda27-225">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-225">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="dda27-226">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="dda27-226">The add-in will be hosted locally on IIS.</span></span>

    ![突出显示“开始”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="dda27-228">在 PowerPoint 中，选择功能区中的“显示任务窗格”\*\*\*\* 按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="dda27-228">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![“开始”功能区中突出显示“显示任务窗格”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)


3. <span data-ttu-id="dda27-230">使用“开始”\*\*\*\* 选项卡功能区中的“新建幻灯片”\*\*\*\* 按钮，将两张新幻灯片添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="dda27-230">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span> 

4. <span data-ttu-id="dda27-231">在任务窗格中，选择“前往第一张幻灯片”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-231">In the task pane, choose the **Go to First Slide** button.</span></span> <span data-ttu-id="dda27-232">此时，选择并显示文档中的第一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="dda27-232">The first slide in the document is selected and displayed.</span></span>

    ![突出显示“前往第一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="dda27-234">在任务窗格中，选择“前往下一张幻灯片”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-234">In the task pane, choose the **Go to Next Slide** button.</span></span> <span data-ttu-id="dda27-235">此时，选择并显示文档中的下一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="dda27-235">The next slide in the document is selected and displayed.</span></span>

    ![突出显示“前往下一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="dda27-237">在任务窗格中，选择“前往上一张幻灯片”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-237">In the task pane, choose the **Go to Previous Slide** button.</span></span> <span data-ttu-id="dda27-238">此时，选择并显示文档中的上一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="dda27-238">The previous slide in the document is selected and displayed.</span></span>

    ![突出显示“前往上一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="dda27-240">在任务窗格中，选择“前往最后一张幻灯片”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="dda27-240">In the task pane, choose the **Go to Last Slide** button.</span></span> <span data-ttu-id="dda27-241">此时，选择并显示文档中的最后一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="dda27-241">The last slide in the document is selected and displayed.</span></span>

    ![突出显示“前往最后一张幻灯片”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="dda27-243">在 Visual Studio 中，按 **Shift + F5** 或选择“停止”\*\*\*\* 按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="dda27-243">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="dda27-244">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="dda27-244">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![突出显示“停止”按钮的 Visual Studio 屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="next-steps"></a><span data-ttu-id="dda27-246">后续步骤</span><span class="sxs-lookup"><span data-stu-id="dda27-246">Next steps</span></span>

<span data-ttu-id="dda27-247">在本教程中，你已创建 PowerPoint 加载项，用于插入图像、插入文本、获取幻灯片元数据，以及在幻灯片之间导航。</span><span class="sxs-lookup"><span data-stu-id="dda27-247">In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.</span></span> <span data-ttu-id="dda27-248">若要了解有关构建 PowerPoint 加载项的详细信息，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="dda27-248">To learn more about building PowerPoint add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="dda27-249">PowerPoint 加载项概述</span><span class="sxs-lookup"><span data-stu-id="dda27-249">PowerPoint add-ins overview</span></span>](../powerpoint/powerpoint-add-ins.md)
