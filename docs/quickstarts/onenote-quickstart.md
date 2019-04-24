---
title: 生成首个 OneNote 加载项
description: ''
ms.date: 03/19/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 378d691d1994a2d22166afc5338007400f7a48af
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450883"
---
# <a name="build-your-first-onenote-add-in"></a>生成首个 OneNote 加载项

本文将逐步介绍如何使用 jQuery 和 Office JavaScript API 生成 OneNote 加载项。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org)

- 全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a>创建加载项项目

1. 使用 Yeoman 生成器创建 OneNote 加载项项目。 运行下面的命令，再回答如下所示的提示问题：

    ```bash
    yo office
    ```

    - **选择项目类型:** `Office Add-in project using Jquery framework`
    - **选择脚本类型:** `Javascript`
    - **要如何命名加载项?:** `My Office Add-in`
    - **要支持哪一个 Office 客户端应用?:** `Onenote`

    ![有关 Yeoman 生成器提示和回答的屏幕截图](../images/yo-office-onenote-jquery.png)
    
    完成此向导后，生成器会创建项目，并安装支持的 Node 组件。
    
2. 导航到项目的根文件夹。

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>更新代码

1. 在代码编辑器中，打开项目根目录中的“index.html”****。 此文件包含在加载项任务窗格中呈现的 HTML。

2. 将 `<body>` 元素替换为以下标记，并保存文件。 

    ```html
    <body class="ms-font-m ms-welcome">
        <header class="ms-welcome__header ms-bgColor-themeDark ms-u-fadeIn500">
            <h2 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">OneNote Add-in</h1>
        </header>
        <main id="app-body" class="ms-welcome__main">
            <br />
            <p class="ms-font-m">Enter HTML content here:</p>
            <div class="ms-TextField ms-TextField--placeholder">
                <textarea id="textBox" rows="8" cols="30"></textarea>
            </div>
            <button id="addOutline" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Add outline</span>
            </button>
        </main>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. 打开文件“**src/index.js**”，以指定加载项的脚本。 将整个内容替换为下列代码，并保存文件。

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.onReady(() => {
        // Office is ready
        $(document).ready(() => {
            // The document is ready
            $('#addOutline').click(addOutlineToPage);
        });
    });
    
    async function addOutlineToPage() {
        try {
            await OneNote.run(async context => {
                var html = "<p>" + $("#textBox").val() + "</p>";

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.
                page.load("title");

                // Add text to the page by using the specified HTML.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log("Added outline to page " + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error);
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    ```

4. 打开文件 **“app.css**”，以指定加载项自定义样式。 将整个内容替换为以下内容，并保存文件。

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 500px;
        min-width: 320px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 5px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 5px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 30px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

## <a name="update-the-manifest"></a>更新清单

1. 打开文件“**manifest.xml**”以定义加载项的设置和功能。

2. `ProviderName` 元素具有占位符值。 将其替换为你的姓名。

3. `Description` 元素的 `DefaultValue` 属性有占位符。 将它替换为“A task pane add-in for OneNote”****。

4. 保存文件。

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
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

    下图展示了使用者笔记本的“**我的加载项**”选项卡。

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. 在“**上传加载项**”对话框中，转到项目文件夹中的 manifest.xml，然后选择“**上传**”。 

4. 在“**开始**”选项卡上，选择位于功能区的“**显示任务窗格**”按钮。 该加载项窗格在 OneNote 页旁的 iFrame 中打开。

5. 在文本区域中输入以下 HTML 内容，然后选择“**添加大纲**”。  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    指定的大纲将添加到页面中。

    ![通过此演练生成的 OneNote 加载项](../images/onenote-first-add-in-3.png)

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
- [OneNote JavaScript API 参考](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [Rubric Grader 示例](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Office 加载项平台概述](../overview/office-add-ins.md)

