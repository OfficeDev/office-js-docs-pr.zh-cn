# <a name="build-your-first-project-add-in"></a>生成首个 Project 加载项

本文将逐步介绍如何使用 jQuery 和 Office JavaScript API 生成 Project 加载项。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org)

- 全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a>创建加载项

1. 使用 Yeoman 生成器创建 Project 加载项项目。 运行下面的命令，再回答如下所示的提示问题：

    ```bash
    yo office
    ```

    - **选择项目类型:** `Office Add-in project using Jquery framework`
    - **选择脚本类型:** `Javascript`
    - **要如何命名加载项?:** `My Office Add-in`
    - **要支持哪一个 Office 客户端应用?:** `Project`

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-project-jquery.png)
    
    完成此向导后，生成器会创建项目，并安装支持的 Node 组件。
    
2. 导航到项目的根文件夹。

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>更新代码

1. 在代码编辑器中，打开项目根目录中的“index.html”****。 此文件包含在加载项任务窗格中呈现的 HTML。

2. 用以下标记替换 `<body>` 元素。

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
                <h3>Try it out</h3>
                <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
                <br/><br/>
                <button class="ms-Button" id="get-task">Get Task data</button>
                <br/>
                <h4>Results:</h4>
                <textarea id="result" rows="6" cols="25"></textarea>
            </div>
        </div>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. 打开文件 **src/index.js**，指定加载项的脚本。 将整个内容替换为下列代码，并保存文件。

    ```js
    'use strict';

    (function () {

        var taskGuid;

        // The initialize function must be run each time a new page is loaded
        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        };

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. 打开项目根目录中的文件“app.css”****，以指定加载项自定义样式。 将整个内容替换为以下内容，并保存文件。

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

## <a name="update-the-manifest"></a>更新清单

1. 打开文件“**manifest.xml**”以定义加载项的设置和功能。

2. `ProviderName` 元素具有占位符值。 将其替换为你的姓名。

3. `Description` 元素的 `DefaultValue` 属性有占位符。 将它替换为“A task pane add-in for Project”****。

4. 保存文件。

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a>启动开发人员服务器

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>试用

1. 在 Project 中，创建至少有一个任务的简单项目。

2. 请按照运行加载项所用平台对应的说明操作，以在 Project 中旁加载加载项。

    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Project Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

3. 在 Project 中，选择一个任务。

    ![Project 中已选择一个任务的项目计划的屏幕截图](../images/project_quickstart_addin_1.png)

4. 在任务窗格中，选择“获取任务 GUID”**** 按钮，将任务 GUID 写入到“结果”**** 文本框。

    ![Project 中已选择一个任务的项目计划，且任务 GUID 写入到任务窗格中文本框的屏幕截图](../images/project_quickstart_addin_2.png)

5. 在任务窗格中，选择“获取任务数据”**** 按钮，将选定任务的多个属性写入到“结果”**** 文本框。

    ![Project 中已选择一个任务的项目计划，且多个任务属性写入到任务窗格中文本框的屏幕截图](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a>后续步骤

恭喜！已成功创建 Project 加载项！ 接下来，请详细了解 Project 加载项功能，并探索常见方案。

> [!div class="nextstepaction"]
> [Project 加载项](../project/project-add-ins.md)
