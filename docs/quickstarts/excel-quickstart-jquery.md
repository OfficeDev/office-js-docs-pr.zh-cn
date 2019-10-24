---
title: 生成首个 Excel 任务窗格加载项
description: 了解如何使用 Office JS API 生成简单的 Excel 任务窗格加载项。
ms.date: 10/17/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 3ade0eb77f525ebd593a475736ab81742d915b94
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626969"
---
# <a name="build-an-excel-task-pane-add-in"></a><span data-ttu-id="1607d-103">生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="1607d-103">Build an Excel task pane add-in</span></span>

<span data-ttu-id="1607d-104">本文将逐步介绍如何生成 Excel 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="1607d-104">In this article, you'll walk through the process of building an Excel task pane add-in.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="1607d-105">创建加载项</span><span class="sxs-lookup"><span data-stu-id="1607d-105">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generatortabyeomangenerator"></a>[<span data-ttu-id="1607d-106">Yeoman 生成器</span><span class="sxs-lookup"><span data-stu-id="1607d-106">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="1607d-107">先决条件</span><span class="sxs-lookup"><span data-stu-id="1607d-107">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="1607d-108">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="1607d-108">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="1607d-109">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="1607d-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="1607d-110">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="1607d-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="1607d-111">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="1607d-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="1607d-112">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="1607d-112">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman 生成器](../images/yo-office-excel.png)

<span data-ttu-id="1607d-114">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="1607d-114">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a><span data-ttu-id="1607d-115">浏览项目</span><span class="sxs-lookup"><span data-stu-id="1607d-115">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="1607d-116">试用</span><span class="sxs-lookup"><span data-stu-id="1607d-116">Try it out</span></span>

1. <span data-ttu-id="1607d-117">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="1607d-117">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="1607d-118">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="1607d-118">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="1607d-120">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="1607d-120">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="1607d-121">在任务窗格的底部，选择“**运行**”链接，价格选定范围的颜色设为黄色。</span><span class="sxs-lookup"><span data-stu-id="1607d-121">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-3c.png)

# <a name="visual-studiotabvisualstudio"></a>[<span data-ttu-id="1607d-123">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="1607d-123">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="1607d-124">先决条件</span><span class="sxs-lookup"><span data-stu-id="1607d-124">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="1607d-125">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="1607d-125">Create the add-in project</span></span>


1. <span data-ttu-id="1607d-126">在 Visual Studio 中，选择“**新建项目**”。</span><span class="sxs-lookup"><span data-stu-id="1607d-126">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="1607d-127">使用搜索框，输入“**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="1607d-127">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="1607d-128">选择“**Excel Web 加载项**”，然后选择“**下一步**”。</span><span class="sxs-lookup"><span data-stu-id="1607d-128">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="1607d-129">对项目命名，然后选择“**创建**”。</span><span class="sxs-lookup"><span data-stu-id="1607d-129">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="1607d-130">在“创建 Office 加载项”\*\*\*\* 对话框窗口中，选择“将新功能添加到 Excel”\*\*\*\*，再选择“完成”\*\*\*\* 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="1607d-130">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="1607d-p102">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="1607d-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="1607d-133">探索 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="1607d-133">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="1607d-134">更新代码</span><span class="sxs-lookup"><span data-stu-id="1607d-134">Update the code</span></span>

1. <span data-ttu-id="1607d-p103">**Home.html** 指定在加载项的任务窗格中呈现的 HTML。 在 **Home.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="1607d-p103">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

2. <span data-ttu-id="1607d-p104">打开 Web 应用项目根目录中的文件“Home.js”\*\*\*\*。 此文件指定加载项脚本。 将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="1607d-p104">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="1607d-p105">打开 Web 应用项目根目录中的文件“Home.css”\*\*\*\*。 此文件指定加载项自定义样式。 将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="1607d-p105">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="1607d-143">更新清单</span><span class="sxs-lookup"><span data-stu-id="1607d-143">Update the manifest</span></span>

1. <span data-ttu-id="1607d-144">打开加载项项目中的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="1607d-144">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="1607d-145">此文件定义的是加载项设置和功能。</span><span class="sxs-lookup"><span data-stu-id="1607d-145">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="1607d-p107">`ProviderName` 元素具有占位符值。 将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="1607d-p107">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="1607d-148">`DisplayName` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="1607d-148">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="1607d-149">将它替换为“My Office Add-in”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1607d-149">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="1607d-150">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="1607d-150">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="1607d-151">将它替换为“A task pane add-in for Excel”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1607d-151">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="1607d-152">保存文件。</span><span class="sxs-lookup"><span data-stu-id="1607d-152">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="1607d-153">试用</span><span class="sxs-lookup"><span data-stu-id="1607d-153">Try it out</span></span>

1. <span data-ttu-id="1607d-p110">使用 Visual Studio 的同时，按 **F5** 或选择“开始”\*\*\*\* 按钮启动 Excel，以测试新建的 Excel 加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="1607d-p110">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="1607d-156">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="1607d-156">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="1607d-158">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="1607d-158">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="1607d-159">在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="1607d-159">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="1607d-161">后续步骤</span><span class="sxs-lookup"><span data-stu-id="1607d-161">Next steps</span></span>

<span data-ttu-id="1607d-162">恭喜，你已成功创建 Excel 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="1607d-162">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="1607d-163">接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="1607d-163">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="1607d-164">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="1607d-164">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="1607d-165">另请参阅</span><span class="sxs-lookup"><span data-stu-id="1607d-165">See also</span></span>

* [<span data-ttu-id="1607d-166">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="1607d-166">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="1607d-167">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="1607d-167">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="1607d-168">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="1607d-168">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="1607d-169">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="1607d-169">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
