在本教程中，请先设置开发项目。 

> [!NOTE]
> 此为 PowerPoint 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [PowerPoint 加载项教程](../tutorials/powerpoint-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="prerequisites"></a>先决条件

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="setup"></a>设置

在本教程中，将使用 Visual Studio 创建加载项。

### <a name="create-the-add-in-project"></a>创建加载项项目

1. 在 Visual Studio 菜单栏中，依次选择“文件”**** > “新建”**** > “项目”****。
    
2. 在“Visual C#”****或“Visual Basic”****下的项目类型列表中，展开“Office/SharePoint”****，选择“加载项”****，再选择“PowerPoint Web 加载项”****作为项目类型。 

3. 将项目命名为“HelloWorld”****，再选择“确定”****按钮。

4. 在“创建 Office 加载项”****对话框窗口中，选择“将新功能添加到 PowerPoint”****，再选择“完成”****以创建项目。

5. 此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”****中。**Home.html** 文件在 Visual Studio 中打开。

     ![PowerPoint 教程 - 显示 HelloWorld 解决方案中 2 个项目的 Visual Studio 解决方案资源管理器窗口](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a>探索 Visual Studio 解决方案

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a>更新代码 

请按照下面的步骤编辑加载项代码，以创建在本教程后续步骤中实现加载项功能的框架。

1. **Home.html** 指定在加载项任务窗格中呈现的 HTML。 在 **Home.html** 文件中，查找包含 `id="content-main"` 的 **div**，并将找到的整个 **div** 替换为以下标记，再保存此文件。

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

2. 打开 Web 应用程序项目根目录中的文件 **Home.js**。 此文件指定加载项脚本。 将整个内容替换为下列代码，并保存文件。

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
