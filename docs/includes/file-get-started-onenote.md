# <a name="build-your-first-onenote-add-in"></a>生成首个 OneNote 加载项

本文将逐步介绍如何使用 jQuery 和 Office JavaScript API 生成 OneNote 加载项。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org)

- 全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a>创建加载项项目

1. 在本地驱动器上创建文件夹，并将它命名为“`my-onenote-addin`”。 将在其中创建外接程序文件。

2. 转到新文件夹。

    ```bash
    cd my-onenote-addin
    ```

3. 使用 Yeoman 生成器创建 OneNote 加载项项目。 运行下面的命令，再回答如下所示的提示问题：

    ```bash
    yo office
    ```

    - **选择一个项目类型：** `Jquery`
    - **选择一个脚本类型：** `Javascript`
    - **要将你的外接程序命名为什么?:** `My Office Add-in`
    - **要支持哪一个 Office 客户端应用程序?:** `Onenote`

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-onenote-jquery.png)
    
    完成向导后，生成器将创建项目并安装 Node 支持组件。


## <a name="update-the-code"></a>更新代码

1. 在代码编辑器中，打开项目根目录中的“index.html”****。 此文件包含在加载项任务窗格中呈现的 HTML。

2. 将 `<body>` 元素内的 `<main>` 元素替换为以下标记，并保存文件。 这会使用 [Office UI Fabric 组件](http://dev.office.com/fabric/components)添加文本区域和按钮。

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. 打开文件“app.js”****，以指定加载项脚本。 将整个内容替换为以下代码，并保存文件。

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                app.initialize();

                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a>更新清单

1. 打开文件“one-note-add-in-manifest.xml”****，以定义加载项的设置和功能。

2. 元素具有占位符值。`ProviderName` 将其替换为你的姓名。

3. 元素的 `DefaultValue` 属性有占位符。`Description` 将它替换为“A task pane add-in for OneNote”****。

4. 保存文件。

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a>启动开发人员服务器

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a>试用

1. 在 [OneNote Online](https://www.onenote.com/notebooks) 中，打开一个笔记本。

2. 依次选择“插入”>“Office 加载项”****，打开“Office 加载项”对话框。

    - 如果使用使用者帐户登录，请依次选择“我的加载项”**** 选项卡和“上传我的加载项”****。

    - 如果使用工作或学校帐户登录，请依次选择“我的组织”**** 选项卡和“上传我的加载项”****。 

    下图展示了使用者笔记本的“我的加载项”**** 选项卡。

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. |||UNTRANSLATED_CONTENT_START|||In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.|||UNTRANSLATED_CONTENT_END||| 

4. 在**主页**选项卡，选择功能区中的**显示任务窗格**按钮。 该外接程序在 OneNote 页面旁的 iFrame 中打开。

5. 在文本区域中输入一些文本，然后选择**添加大纲**。 您输入的文本将添加至页面。 

    ![通过此步骤生成的 OneNote 外接程序](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a>疑难解答和提示

- 您可以使用浏览器的开发者工具调试外接程序。当您在 Internet Explorer 或 Chrome 中使用 Gulp Web 服务器并进行调试时，您可以本地保存您的更改，然后仅刷新外接程序的 iFrame。

- 检查 OneNote 对象时，目前可用的属性显示实际值。需要加载的属性显示“未定义”**。展开 `_proto_` 节点，查看已在对象上定义但尚未加载的属性。

   ![调试器中尚未加载的 OneNote 对象](../images/onenote-debug.png)

- 如果您的外接程序使用任何 HTTP 资源，则需要启用浏览器中的混合内容。生产外接程序应当仅使用安全 HTTPS 资源。

- 可以从任何位置打开任务窗格外接程序，但只能在常规页面内容（即不在标题、图像、IFrame 等中）内插入内容外接程序。 

## <a name="next-steps"></a>后续步骤

恭喜！已成功创建 OneNote 加载项！ 接下来，请详细了解与生成 OneNote 加载项有关的核心概念。

> [!div class="nextstepaction"]
> [OneNote JavaScript API 编程概述](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>另请参阅

- [OneNote JavaScript API 编程概述](../onenote/onenote-add-ins-programming-overview.md)
- [OneNote JavaScript API 参考](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 加载项平台概述](../overview/office-add-ins.md)
