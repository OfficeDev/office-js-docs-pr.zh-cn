# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="bb353-101">使用 jQuery 生成 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="bb353-101">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="bb353-102">本文将逐步介绍如何使用 jQuery 和 Excel JavaScript API 生成 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="bb353-102">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="bb353-103">创建加载项</span><span class="sxs-lookup"><span data-stu-id="bb353-103">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="bb353-104">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="bb353-104">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="bb353-105">先决条件</span><span class="sxs-lookup"><span data-stu-id="bb353-105">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="bb353-106">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="bb353-106">Create the add-in project</span></span>

1. <span data-ttu-id="bb353-107">在 Visual Studio 菜单栏中，依次选择“文件”\*\*\*\* > “新建”\*\*\*\* > “项目”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-107">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="bb353-108">在“Visual C#”\*\*\*\* 或“Visual Basic”\*\*\*\* 下的项目类型列表中，展开“Office/SharePoint”\*\*\*\*，选择“加载项”\*\*\*\*，再选择“Excel Web 加载项”\*\*\*\* 作为项目类型。</span><span class="sxs-lookup"><span data-stu-id="bb353-108">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="bb353-109">命名此项目，再选择“确定”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-109">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="bb353-110">在“创建 Office 加载项”\*\*\*\* 对话框窗口中，选择“将新功能添加到 Excel”\*\*\*\*，再选择“完成”\*\*\*\* 以创建项目。</span><span class="sxs-lookup"><span data-stu-id="bb353-110">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="bb353-p101">此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”\*\*\*\* 中。**Home.html** 文件在 Visual Studio 中打开。</span><span class="sxs-lookup"><span data-stu-id="bb353-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="bb353-113">探索 Visual Studio 解决方案</span><span class="sxs-lookup"><span data-stu-id="bb353-113">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="bb353-114">更新代码</span><span class="sxs-lookup"><span data-stu-id="bb353-114">Update the code</span></span>

1. <span data-ttu-id="bb353-115">**Home.html** 指定在加载项的任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="bb353-115">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="bb353-116">在 **Home.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-116">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
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

2. <span data-ttu-id="bb353-117">打开 Web 应用项目根目录中的文件“Home.js”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-117">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="bb353-118">此文件指定加载项脚本。</span><span class="sxs-lookup"><span data-stu-id="bb353-118">This file specifies the script for the add-in.</span></span> <span data-ttu-id="bb353-119">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-119">Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

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

3. <span data-ttu-id="bb353-120">打开 Web 应用项目根目录中的文件“Home.css”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-120">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="bb353-121">此文件指定加载项自定义样式。</span><span class="sxs-lookup"><span data-stu-id="bb353-121">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="bb353-122">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-122">Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="bb353-123">更新清单</span><span class="sxs-lookup"><span data-stu-id="bb353-123">Update the manifest</span></span>

1. <span data-ttu-id="bb353-124">打开加载项项目中的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-124">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="bb353-125">此文件定义的是加载项设置和功能。</span><span class="sxs-lookup"><span data-stu-id="bb353-125">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="bb353-126">`ProviderName` 元素具有占位符值。</span><span class="sxs-lookup"><span data-stu-id="bb353-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="bb353-127">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="bb353-127">Replace it with your name.</span></span>

3. <span data-ttu-id="bb353-128">`DisplayName` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="bb353-128">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="bb353-129">将它替换为“My Office Add-in”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-129">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="bb353-130">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="bb353-130">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="bb353-131">将它替换为“A task pane add-in for Excel”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-131">Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="bb353-132">保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-132">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="bb353-133">试用</span><span class="sxs-lookup"><span data-stu-id="bb353-133">Try it out</span></span>

1. <span data-ttu-id="bb353-p109">使用 Visual Studio 的同时，按 F5 或选择“开始”\*\*\*\* 按钮启动 Excel，以测试新建的 Excel 加载项，功能区中显示有“显示任务窗格”\*\*\*\* 加载项按钮。加载项本地托管在 IIS 上。</span><span class="sxs-lookup"><span data-stu-id="bb353-p109">Using Visual Studio, test the newly created Excel add-in by pressing F5 or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="bb353-136">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="bb353-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="bb353-138">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="bb353-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="bb353-139">在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="bb353-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="bb353-141">任意编辑器</span><span class="sxs-lookup"><span data-stu-id="bb353-141">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="bb353-142">先决条件</span><span class="sxs-lookup"><span data-stu-id="bb353-142">Prerequisites</span></span>

- [<span data-ttu-id="bb353-143">Node.js</span><span class="sxs-lookup"><span data-stu-id="bb353-143">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="bb353-144">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="bb353-144">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="bb353-145">创建 Web 应用</span><span class="sxs-lookup"><span data-stu-id="bb353-145">Create the web app</span></span>

1. <span data-ttu-id="bb353-146">使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="bb353-146">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="bb353-147">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="bb353-147">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="bb353-148">**选择项目类型:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="bb353-148">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="bb353-149">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="bb353-149">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="bb353-150">**要如何命名加载项?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="bb353-150">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="bb353-151">**要支持哪一个 Office 客户端应用？：**`Excel`</span><span class="sxs-lookup"><span data-stu-id="bb353-151">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman 生成器](../images/yo-office-jquery.png)
    
    <span data-ttu-id="bb353-153">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="bb353-153">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="bb353-154">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="bb353-154">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="bb353-155">更新代码</span><span class="sxs-lookup"><span data-stu-id="bb353-155">Update the code</span></span> 

1. <span data-ttu-id="bb353-156">在代码编辑器中，打开项目根目录中的“index.html”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-156">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="bb353-157">此文件指定在加载项任务窗格中呈现的 HTML。</span><span class="sxs-lookup"><span data-stu-id="bb353-157">This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 
 
2. <span data-ttu-id="bb353-158">在“**index.html**”中，将 `body` 标记替换为以下标记，再保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-158">Within **index.html**, replace the `body` tag with the following markup and save the file.</span></span>
 
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
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>    
    ```

3. <span data-ttu-id="bb353-159">打开文件“**src/index.js**”，以指定加载项的脚本。</span><span class="sxs-lookup"><span data-stu-id="bb353-159">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="bb353-160">将整个内容替换为下列代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-160">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

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

4. <span data-ttu-id="bb353-161">打开文件“app.css”\*\*\*\*，以指定加载项自定义样式。</span><span class="sxs-lookup"><span data-stu-id="bb353-161">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="bb353-162">将整个内容替换为以下代码，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-162">Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="bb353-163">更新清单</span><span class="sxs-lookup"><span data-stu-id="bb353-163">Update the manifest</span></span>

1. <span data-ttu-id="bb353-164">打开文件“**manifest.xml**”以定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="bb353-164">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="bb353-165">`ProviderName` 元素具有占位符值。</span><span class="sxs-lookup"><span data-stu-id="bb353-165">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="bb353-166">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="bb353-166">Replace it with your name.</span></span>

3. <span data-ttu-id="bb353-167">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="bb353-167">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="bb353-168">将它替换为“A task pane add-in for Excel”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="bb353-168">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="bb353-169">保存文件。</span><span class="sxs-lookup"><span data-stu-id="bb353-169">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="bb353-170">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="bb353-170">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="bb353-171">试用</span><span class="sxs-lookup"><span data-stu-id="bb353-171">Try it out</span></span>

1. <span data-ttu-id="bb353-172">请按照运行加载项所用平台对应的说明操作，以在 Excel 中旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="bb353-172">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="bb353-173">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="bb353-173">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="bb353-174">Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="bb353-174">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="bb353-175">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="bb353-175">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="bb353-176">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="bb353-176">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="bb353-178">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="bb353-178">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="bb353-179">在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="bb353-179">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="bb353-181">后续步骤</span><span class="sxs-lookup"><span data-stu-id="bb353-181">Next steps</span></span>

<span data-ttu-id="bb353-p116">恭喜！已使用 jQuery 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="bb353-p116">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="bb353-184">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="bb353-184">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="bb353-185">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bb353-185">See also</span></span>

* [<span data-ttu-id="bb353-186">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="bb353-186">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="bb353-187">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="bb353-187">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="bb353-188">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="bb353-188">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="bb353-189">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="bb353-189">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
