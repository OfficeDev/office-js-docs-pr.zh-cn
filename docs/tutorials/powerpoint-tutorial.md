---
title: PowerPoint 加载项教程
description: 在本教程中，将生成 PowerPoint 加载项，用于插入图像、插入文本、获取幻灯片元数据，以及在幻灯片之间导航。
ms.date: 10/14/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: 72b7abb8f67ad634025abd80b5bc9bb987ff6868
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132373"
---
# <a name="tutorial-create-a-powerpoint-task-pane-add-in"></a><span data-ttu-id="aed40-103">教程：创建 PowerPoint 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="aed40-103">Tutorial: Create a PowerPoint task pane add-in</span></span>

<span data-ttu-id="aed40-104">在本教程中，将使用 Visual Studio 创建 PowerPoint 任务窗格加载项：</span><span class="sxs-lookup"><span data-stu-id="aed40-104">In this tutorial, you'll use Visual Studio to create an PowerPoint task pane add-in that:</span></span>

> [!div class="checklist"]
>
> - <span data-ttu-id="aed40-105">向幻灯片添加一天中的[必应](https://www.bing.com)照片</span><span class="sxs-lookup"><span data-stu-id="aed40-105">Adds the [Bing](https://www.bing.com) photo of the day to a slide</span></span>
> - <span data-ttu-id="aed40-106">向幻灯片添加文本</span><span class="sxs-lookup"><span data-stu-id="aed40-106">Adds text to a slide</span></span>
> - <span data-ttu-id="aed40-107">获取幻灯片元数据</span><span class="sxs-lookup"><span data-stu-id="aed40-107">Gets slide metadata</span></span>
> - <span data-ttu-id="aed40-108">在幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="aed40-108">Navigates between slides</span></span>

## <a name="prerequisites"></a><span data-ttu-id="aed40-109">先决条件</span><span class="sxs-lookup"><span data-stu-id="aed40-109">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="aed40-110">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="aed40-110">Create your add-in project</span></span>

<span data-ttu-id="aed40-111">完成以下步骤以使用 Visual Studio 创建 PowerPoint 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="aed40-111">Complete the following steps to create a PowerPoint add-in project using Visual Studio.</span></span>

1. <span data-ttu-id="aed40-112">选择“**创建新项目**”。</span><span class="sxs-lookup"><span data-stu-id="aed40-112">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="aed40-113">使用搜索框，输入“**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="aed40-113">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="aed40-114">选择“**PowerPoint 外接程序**”，然后选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="aed40-114">Choose **PowerPoint Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="aed40-115">将项目命名为“`HelloWorld`”，然后选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="aed40-115">Name the project `HelloWorld`, and select **Create**.</span></span>

4. <span data-ttu-id="aed40-116">在“创建 Office 加载项”对话框窗口中，选择“将新功能添加到 PowerPoint”，再选择“完成”以创建项目。</span><span class="sxs-lookup"><span data-stu-id="aed40-116">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="aed40-p102">此时，Visual Studio 创建解决方案，且它的两个项目显示在“**解决方案资源管理器**”中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="aed40-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![Visual Studio 解决方案资源管理器窗口的屏幕截图，显示 HelloWorld 解决方案中的 2 个项目：HelloWorld 和 HelloWorldWeb。](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="aed40-120">浏览 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="aed40-120">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="aed40-121">更新代码</span><span class="sxs-lookup"><span data-stu-id="aed40-121">Update code</span></span>

<span data-ttu-id="aed40-122">请按照下面的步骤编辑加载项代码，以创建在本教程后续步骤中实现加载项功能的框架。</span><span class="sxs-lookup"><span data-stu-id="aed40-122">Edit the add-in code as follows to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="aed40-p103">**Home.html** 指定在加载项任务窗格中呈现的 HTML。 在 **Home.html** 文件中，查找包含  的 `id="content-main"`，并将找到的整个 **div** 替换为以下标记，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="aed40-p103">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

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

2. <span data-ttu-id="aed40-p104">打开 Web 应用项目根目录中的文件 **Home.js**。 此文件指定加载项脚本。 将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="aed40-p104">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    (function () {
        "use strict";

        var messageBanner;

        Office.onReady(function () {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        });

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

## <a name="insert-an-image"></a><span data-ttu-id="aed40-128">插入图像</span><span class="sxs-lookup"><span data-stu-id="aed40-128">Insert an image</span></span>

<span data-ttu-id="aed40-129">完成以下步骤以添加用于检索一天中的[必应](https://www.bing.com)照片的代码，并将该图像插入幻灯片中。</span><span class="sxs-lookup"><span data-stu-id="aed40-129">Complete the following steps to add code that retrieves the [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.</span></span>

1. <span data-ttu-id="aed40-130">使用解决方案资源管理器，将 **Controllers** 新文件夹添加到 **HelloWorldWeb** 项目。</span><span class="sxs-lookup"><span data-stu-id="aed40-130">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![Visual Studio 解决方案资源管理器窗口的屏幕截图，显示在 HelloWorldWeb 项目中突出显示的“Controllers”文件夹](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="aed40-132">右键单击“**Controllers**”文件夹，并依次选择“添加”>“**新基架项...**”。</span><span class="sxs-lookup"><span data-stu-id="aed40-132">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="aed40-133">在“添加基架”对话框窗口中，依次选择“Web API 2 控制器 - 空”和“添加”按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-133">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="aed40-134">在“添加控制器”对话框窗口中，输入“PhotoController”作为控制器名称，再选择“添加”按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-134">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button.</span></span> <span data-ttu-id="aed40-135">此时，Visual Studio 创建并打开 **PhotoController.cs** 文件。</span><span class="sxs-lookup"><span data-stu-id="aed40-135">Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="aed40-p106">将 **PhotoController.cs** 文件的全部内容替换为下列代码，以调用必应服务来检索 Base64 编码字符串形式的一天中照片。 使用 Office JavaScript API 将图像插入文档时，必须将图像数据指定为 Base64 编码字符串。</span><span class="sxs-lookup"><span data-stu-id="aed40-p106">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string. When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

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

6. <span data-ttu-id="aed40-138">在 **Home.html** 文件中，将 `TODO1` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="aed40-138">In the **Home.html** file, replace `TODO1` with the following markup.</span></span> <span data-ttu-id="aed40-139">此标记定义在加载项任务窗格内显示的“插入图像”按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-139">This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="aed40-140">在 **Home.js** 文件中，将 `TODO1` 替换为下列代码，以分配“插入图像”按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="aed40-140">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="aed40-p108">在 **Home.js** 文件中，将 `TODO2` 替换为下列代码，以定义 `insertImage` 函数。 此函数从必应 Web 服务提取图像，再调用 `insertImageFromBase64String` 函数将相应图像插入文档。</span><span class="sxs-lookup"><span data-stu-id="aed40-p108">In the **Home.js** file, replace `TODO2` with the following code to define the `insertImage` function. This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

    ```js
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

9. <span data-ttu-id="aed40-p109">在 **Home.js** 文件中，将 `TODO3` 替换为下列代码，以定义 `insertImageFromBase64String` 函数。 此函数使用 Office JavaScript API 将图像插入文档。 注意：</span><span class="sxs-lookup"><span data-stu-id="aed40-p109">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function. This function uses the Office JavaScript API to insert the image into the document. Note:</span></span>

    - <span data-ttu-id="aed40-146">`coercionType` 选项被指定为 `setSelectedDataAsync` 请求的第二个参数，指明了要插入的数据的类型。</span><span class="sxs-lookup"><span data-stu-id="aed40-146">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsync` request indicates the type of data being inserted.</span></span>

    - <span data-ttu-id="aed40-147">`asyncResult` 对象封装 `setSelectedDataAsync` 请求的结果，包括状态和错误消息（如果请求失败的话）。</span><span class="sxs-lookup"><span data-stu-id="aed40-147">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

    ```js
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

### <a name="test-the-add-in"></a><span data-ttu-id="aed40-148">测试加载项</span><span class="sxs-lookup"><span data-stu-id="aed40-148">Test the add-in</span></span>

1. <span data-ttu-id="aed40-p110">使用 Visual Studio 的同时，按 **F5** 或选择“**开始**”按钮启动 PowerPoint，以测试新建的 PowerPoint 加载项，功能区中显示有“**显示任务窗格**”加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="aed40-p110">Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![显示在 Visual Studio 中突出显示的“开始”按钮的屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="aed40-152">在 PowerPoint 中，选择功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="aed40-152">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![屏幕截图显示 PowerPoint 中主功能区上突出显示的“显示任务窗格”按钮](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="aed40-154">在任务窗格中，选择“**插入图像**”按钮，将一天中的必应照片添加到当前幻灯片。</span><span class="sxs-lookup"><span data-stu-id="aed40-154">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![突出显示“插入图像”按钮的 PowerPoint 加载项屏幕截图](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="aed40-156">在 Visual Studio 中，按 **Shift + F5** 或选择“**停止**”按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="aed40-156">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="aed40-157">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="aed40-157">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![显示在 Visual Studio 中突出显示的“停止”按钮的屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="customize-user-interface-ui-elements"></a><span data-ttu-id="aed40-159">自定义用户界面 (UI) 元素</span><span class="sxs-lookup"><span data-stu-id="aed40-159">Customize User Interface (UI) elements</span></span>

<span data-ttu-id="aed40-160">完成以下步骤以添加用于自定义任务窗格 UI 的标记。</span><span class="sxs-lookup"><span data-stu-id="aed40-160">Complete the following steps to add markup that customizes the task pane UI.</span></span>

1. <span data-ttu-id="aed40-p112">在 **Home.html** 文件中，将 `TODO2` 替换为以下标记，以将页眉部分和标题添加到任务窗格。 注意：</span><span class="sxs-lookup"><span data-stu-id="aed40-p112">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane. Note:</span></span>

    - <span data-ttu-id="aed40-p113">以 `ms-` 开头的样式由 [Office UI Fabric](../design/office-ui-fabric.md) 进行定义，后者是用于生成 Office 和 Office 365 用户体验的 JavaScript 前端框架。 **Home.html** 文件包含对 Fabric 样式表的引用。</span><span class="sxs-lookup"><span data-stu-id="aed40-p113">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365. The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="aed40-165">在 **Home.html** 文件中，查找包含 `class="footer"` 的 **div**，并删除找到的整个 **div**，以从任务窗格中删除页脚部分。</span><span class="sxs-lookup"><span data-stu-id="aed40-165">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="aed40-166">测试加载项</span><span class="sxs-lookup"><span data-stu-id="aed40-166">Test the add-in</span></span>

1. <span data-ttu-id="aed40-167">使用 Visual Studio 的同时，按 **F5** 或选择“开始”按钮启动 PowerPoint，以测试 PowerPoint 加载项，功能区中显示有“显示任务窗格”加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-167">Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="aed40-168">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="aed40-168">The add-in will be hosted locally on IIS.</span></span>

    ![显示在 Visual Studio 中突出显示的“开始”按钮的屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="aed40-170">在 PowerPoint 中，选择功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="aed40-170">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![屏幕截图显示 PowerPoint 主功能区上突出显示的“显示任务窗格”按钮](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="aed40-172">请注意，任务窗格现在包含页眉部分和标题，并且不再包含页脚部分。</span><span class="sxs-lookup"><span data-stu-id="aed40-172">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![带有“插入图像”按钮的 PowerPoint 加载项的屏幕截图](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="aed40-174">在 Visual Studio 中，按 **Shift + F5** 或选择“**停止**”按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="aed40-174">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="aed40-175">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="aed40-175">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![显示在 Visual Studio 中突出显示的“停止”按钮的屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="insert-text"></a><span data-ttu-id="aed40-177">插入文本</span><span class="sxs-lookup"><span data-stu-id="aed40-177">Insert text</span></span>

<span data-ttu-id="aed40-178">完成以下步骤以添加用于将文本插入到标题幻灯片的代码，该幻灯片包含一天中的[必应](https://www.bing.com)照片。</span><span class="sxs-lookup"><span data-stu-id="aed40-178">Complete the following steps to add code that inserts text into the title slide which contains the [Bing](https://www.bing.com) photo of the day.</span></span>

1. <span data-ttu-id="aed40-179">在 **Home.html** 文件中，将 `TODO3` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="aed40-179">In the **Home.html** file, replace `TODO3` with the following markup.</span></span> <span data-ttu-id="aed40-180">此标记定义在加载项任务窗格内显示的“插入文本”按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-180">This markup defines the **Insert Text** button that will appear within the add-in's task pane.</span></span>

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. <span data-ttu-id="aed40-181">在 **Home.js** 文件中，将 `TODO4` 替换为下列代码，以分配“插入文本”按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="aed40-181">In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.</span></span>

    ```js
    $('#insert-text').click(insertText);
    ```

3. <span data-ttu-id="aed40-p117">在 **Home.js** 文件中，将 `TODO5` 替换为下列代码，以定义 `insertText` 函数。 此函数将文本插入当前幻灯片。</span><span class="sxs-lookup"><span data-stu-id="aed40-p117">In the **Home.js** file, replace `TODO5` with the following code to define the `insertText` function. This function inserts text into the current slide.</span></span>

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="aed40-184">测试加载项</span><span class="sxs-lookup"><span data-stu-id="aed40-184">Test the add-in</span></span>

1. <span data-ttu-id="aed40-185">使用 Visual Studio 的同时，按 **F5** 或选择“开始”按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-185">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="aed40-186">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="aed40-186">The add-in will be hosted locally on IIS.</span></span>

    ![在 Visual Studio 中突出显示的“开始”按钮的屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="aed40-188">在 PowerPoint 中，选择功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="aed40-188">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![突出显示 PowerPoint 中主功能区上的“显示任务窗格”按钮的屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="aed40-190">在任务窗格中，选择“**插入图像**”按钮，将一天中的必应照片添加到当前幻灯片，再为包含标题文本框的幻灯片选择一种设计。</span><span class="sxs-lookup"><span data-stu-id="aed40-190">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.</span></span>

    ![突出显示当前幻灯片，并在加载项中突出显示“插入图像”按钮的 PowerPoint 屏幕截图](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. <span data-ttu-id="aed40-192">将光标置于标题幻灯片上的文本框中，再选择任务窗格中的“**插入文本**”按钮，向幻灯片添加文本。</span><span class="sxs-lookup"><span data-stu-id="aed40-192">Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.</span></span>

    ![在加载项中突出显示“插入文本”按钮的 PowerPoint 屏幕截图](../images/powerpoint-tutorial-insert-text.png)

5. <span data-ttu-id="aed40-194">在 Visual Studio 中，按 **Shift + F5** 或选择“**停止**”按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="aed40-194">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="aed40-195">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="aed40-195">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![在 Visual Studio 中突出显示的“停止”按钮的屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="get-slide-metadata"></a><span data-ttu-id="aed40-197">获取幻灯片元数据</span><span class="sxs-lookup"><span data-stu-id="aed40-197">Get slide metadata</span></span>

<span data-ttu-id="aed40-198">完成以下步骤以添加用于检索所选幻灯片的元数据的代码。</span><span class="sxs-lookup"><span data-stu-id="aed40-198">Complete the following steps to add code that retrieves metadata for the selected slide.</span></span>

1. <span data-ttu-id="aed40-199">在 **Home.html** 文件中，将 `TODO4` 替换为以下标记。</span><span class="sxs-lookup"><span data-stu-id="aed40-199">In the **Home.html** file, replace `TODO4` with the following markup.</span></span> <span data-ttu-id="aed40-200">此标记定义在加载项任务窗格内显示的“获取幻灯片元数据”按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-200">This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="aed40-201">在 **Home.js** 文件中，将 `TODO6` 替换为下列代码，以分配“获取幻灯片元数据”按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="aed40-201">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="aed40-p121">在 **Home.js** 文件中，将 `TODO7` 替换为下列代码，以定义 `getSlideMetadata` 函数。 此函数检索选定一张或多张幻灯片的元数据，并将它写入加载项任务窗格内的弹出对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="aed40-p121">In the **Home.js** file, replace `TODO7` with the following code to define the `getSlideMetadata` function. This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

    ```js
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

### <a name="test-the-add-in"></a><span data-ttu-id="aed40-204">测试加载项</span><span class="sxs-lookup"><span data-stu-id="aed40-204">Test the add-in</span></span>

1. <span data-ttu-id="aed40-205">使用 Visual Studio 的同时，按 **F5** 或选择“开始”按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-205">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="aed40-206">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="aed40-206">The add-in will be hosted locally on IIS.</span></span>

    ![在 Visual Studio 中突出显示“开始”按钮的屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="aed40-208">在 PowerPoint 中，选择功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="aed40-208">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![突出显示 PowerPoint 主功能区上的“显示任务窗格”按钮的屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="aed40-210">在任务窗格中，选择“**获取幻灯片元数据**”按钮，以获取选定幻灯片的元数据。</span><span class="sxs-lookup"><span data-stu-id="aed40-210">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide.</span></span> <span data-ttu-id="aed40-211">此时，幻灯片元数据写入到任务窗格底部的弹出对话框窗口。</span><span class="sxs-lookup"><span data-stu-id="aed40-211">The slide metadata is written to the popup dialog window at the bottom of the task pane.</span></span> <span data-ttu-id="aed40-212">在此示例中，JSON 元数据中的 `slides` 数组包含一个对象，用于指定选定幻灯片的 `id`、`title` 和 `index`。</span><span class="sxs-lookup"><span data-stu-id="aed40-212">In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide.</span></span> <span data-ttu-id="aed40-213">如果在检索幻灯片元数据时选择了多张幻灯片，JSON 元数据中的 `slides` 数组会对每张选定幻灯片都包含一个对象。</span><span class="sxs-lookup"><span data-stu-id="aed40-213">If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![加载项中突出显示“获取幻灯片元数据”按钮的 PowerPoint 屏幕截图](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="aed40-215">在 Visual Studio 中，按 **Shift + F5** 或选择“**停止**”按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="aed40-215">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="aed40-216">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="aed40-216">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![在 Visual Studio 中突出显示“停止”按钮的屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="navigate-between-slides"></a><span data-ttu-id="aed40-218">在幻灯片之间导航</span><span class="sxs-lookup"><span data-stu-id="aed40-218">Navigate between slides</span></span>

<span data-ttu-id="aed40-219">完成以下步骤以添加用于在文档幻灯片之间导航的代码。</span><span class="sxs-lookup"><span data-stu-id="aed40-219">Complete the following steps to add code that navigates between the slides of a document.</span></span>

1. <span data-ttu-id="aed40-p125">在 **Home.html** 文件中，将 `TODO5` 替换为以下标记。 此标记定义在加载项任务窗格内显示的四个导航按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-p125">In the **Home.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="go-to-first-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to First Slide</span>
        <span class="Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-next-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Next Slide</span>
        <span class="Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-previous-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Previous Slide</span>
        <span class="Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-last-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Last Slide</span>
        <span class="Button-description">Go to the last slide.</span>
    </button>
    ```

2. <span data-ttu-id="aed40-222">在 **Home.js** 文件中，将 `TODO8` 替换为下列代码，以分配四个导航按钮的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="aed40-222">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="aed40-223">在 **Home.js** 文件中，将 `TODO9` 替换为下列代码，以定义导航函数。</span><span class="sxs-lookup"><span data-stu-id="aed40-223">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="aed40-224">以下各函数均使用 `goToByIdAsync` 函数，以根据幻灯片在文档中的位置（第一张、最后一张、上一张和下一张）选择幻灯片。</span><span class="sxs-lookup"><span data-stu-id="aed40-224">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, and next).</span></span>

    ```js
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

### <a name="test-the-add-in"></a><span data-ttu-id="aed40-225">测试加载项</span><span class="sxs-lookup"><span data-stu-id="aed40-225">Test the add-in</span></span>

1. <span data-ttu-id="aed40-226">使用 Visual Studio 的同时，按 **F5** 或选择“开始”按钮启动 PowerPoint，以测试加载项，功能区中显示有“显示任务窗格”加载项按钮。</span><span class="sxs-lookup"><span data-stu-id="aed40-226">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="aed40-227">加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="aed40-227">The add-in will be hosted locally on IIS.</span></span>

    ![显示 Visual Studio 工具栏上突出显示“开始”按钮的屏幕截图](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="aed40-229">在 PowerPoint 中，选择功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="aed40-229">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![显示 PowerPoint 中主功能区上突出显示“显示任务窗格”按钮的屏幕截图](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="aed40-231">使用“**开始**”选项卡功能区中的“**新建幻灯片**”按钮，将两张新幻灯片添加到文档中。</span><span class="sxs-lookup"><span data-stu-id="aed40-231">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span>

4. <span data-ttu-id="aed40-p128">在任务窗格中，选择 **“前往第一张幻灯片”** 按钮。 此时，选择并显示文档中的第一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="aed40-p128">In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.</span></span>

    ![在加载项中突出显示“转到第一张幻灯片”按钮的 PowerPoint 屏幕截图](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="aed40-p129">在任务窗格中，选择 **“前往下一张幻灯片”** 按钮。 此时，选择并显示文档中的下一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="aed40-p129">In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.</span></span>

    ![加载项中突出显示“转到下一张幻灯片”按钮的 PowerPoint 屏幕截图](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="aed40-p130">在任务窗格中，选择 **“前往上一张幻灯片”** 按钮。 此时，选择并显示文档中的上一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="aed40-p130">In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.</span></span>

    ![在加载项中突出显示“转到上一张幻灯片”按钮的 PowerPoint 屏幕截图](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="aed40-p131">在任务窗格中，选择 **“前往最后一张幻灯片”** 按钮。 此时，选择并显示文档中的最后一张幻灯片。</span><span class="sxs-lookup"><span data-stu-id="aed40-p131">In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.</span></span>

    ![加载项中突出显示“转到最后一张幻灯片”按钮的 PowerPoint 屏幕截图](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="aed40-244">在 Visual Studio 中，按 **Shift + F5** 或选择“**停止**”按钮，以停止加载项。</span><span class="sxs-lookup"><span data-stu-id="aed40-244">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="aed40-245">PowerPoint 在加载项停止时自动关闭。</span><span class="sxs-lookup"><span data-stu-id="aed40-245">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![显示 Visual Studio 工具栏上突出显示“停止”按钮的屏幕截图](../images/powerpoint-tutorial-stop.png)

## <a name="next-steps"></a><span data-ttu-id="aed40-247">后续步骤</span><span class="sxs-lookup"><span data-stu-id="aed40-247">Next steps</span></span>

<span data-ttu-id="aed40-248">在本教程中，你已创建 PowerPoint 加载项，用于插入图像、插入文本、获取幻灯片元数据，以及在幻灯片之间导航。</span><span class="sxs-lookup"><span data-stu-id="aed40-248">In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.</span></span> <span data-ttu-id="aed40-249">若要了解有关构建 PowerPoint 加载项的详细信息，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="aed40-249">To learn more about building PowerPoint add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="aed40-250">PowerPoint 加载项概述</span><span class="sxs-lookup"><span data-stu-id="aed40-250">PowerPoint add-ins overview</span></span>](../powerpoint/powerpoint-add-ins.md)

## <a name="see-also"></a><span data-ttu-id="aed40-251">另请参阅</span><span class="sxs-lookup"><span data-stu-id="aed40-251">See also</span></span>

- [<span data-ttu-id="aed40-252">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="aed40-252">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="aed40-253">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="aed40-253">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
