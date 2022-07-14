---
title: 生成首个 Excel 任务窗格加载项
description: 了解如何使用 Office JS API 生成简单的 Excel 任务窗格加载项。
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 11af3d28db9d8b22790eb2e95821d5c9f34d194f
ms.sourcegitcommit: d1541634997c26f8ea22057b1d39074b01757c27
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/14/2022
ms.locfileid: "66789120"
---
# <a name="build-an-excel-task-pane-add-in"></a>生成 Excel 任务窗格加载项

本文将逐步介绍如何生成 Excel 任务窗格加载项。

## <a name="create-the-add-in"></a>创建加载项

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generator"></a>[Yeoman 生成器](#tab/yeomangenerator)

## <a name="prerequisites"></a>先决条件

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a>创建加载项项目

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **选择项目类型:** `Office Add-in Task Pane project`
- **选择脚本类型:** `Javascript`
- **要如何命名加载项?** `My Office Add-in`
- **要支持哪一个 Office 客户端应用程序?** `Excel`

![Yeoman Office 加载项生成器命令行界面的屏幕截图。](../images/yo-office-excel.png)

完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a>浏览项目

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a>试用

1. 导航到项目的根文件夹。

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

3. 在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。

    ![显示突出显示“显示任务窗格”按钮的 Excel 主页菜单屏幕截图。](../images/excel-quickstart-addin-3b.png)

4. 选择工作表中的任何一系列单元格。

5. 在任务窗格的底部，选择“**运行**”链接，价格选定范围的颜色设为黄色。

    ![Excel 的屏幕截图，其中“加载项”任务窗格处于打开状态，并且“加载项”任务窗格中突出显示“运行”按钮。](../images/excel-quickstart-addin-3c.png)

### <a name="next-steps"></a>后续步骤

恭喜！已成功创建 Excel 任务窗格加载项！接下来，请详细了解 Excel 加载项功能，并跟着[ Excel ](../tutorials/excel-tutorial.md)加载项教程一起操作，以生成更复杂的加载项。

# <a name="visual-studio"></a>[Visual Studio](#tab/visualstudio)

### <a name="prerequisites"></a>先决条件

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a>创建加载项项目

1. 在 Visual Studio 中，选择“**新建项目**”。

2. 使用搜索框，输入“**加载项**”。 选择“**Excel Web 加载项**”，然后选择“**下一步**”。

3. 命名项目为 **ExcelWebAddIn1** 并选择“**创建**”。

4. 在“创建 Office 加载项”对话框窗口中，选择“将新功能添加到 Excel”，再选择“完成”以创建项目。

5. 此时，Visual Studio 创建解决方案，且它的两个项目显示在“解决方案资源管理器”中。**Home.html** 文件在 Visual Studio 中打开。

### <a name="explore-the-visual-studio-solution"></a>探索 Visual Studio 解决方案

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a>更新代码

1. **Home.html** 指定在加载项的任务窗格中呈现的 HTML。 在 **Home.html** 中，将 `<body>` 元素替换为以下标记，并保存文件。

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

2. 打开 Web 应用项目根目录中的文件“Home.js”****。 此文件指定加载项脚本。 将整个内容替换为以下代码，并保存文件。

    [!include[Information about the use of ES6 JavaScript](../includes/modern-js-note.md)]

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

        async function setColor() {
            await Excel.run(async (context) => {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                await context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. 打开 Web 应用项目根目录中的文件“Home.css”****。 此文件指定加载项自定义样式。 将整个内容替换为以下代码，并保存文件。

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

### <a name="update-the-manifest"></a>更新清单

1. 在 **解决方案资源管理器** 中，转到 **ExcelWebAddIn1** 加载项项目并打开 **ExcelWebAddIn1Manifest** 目录。 此目录包含清单文件，即 **ExcelWebAddIn1.xml**。 此文件定义了加载项设置和功能。 请参阅上一节 [探索Visual Studio 解决方案](#explore-the-visual-studio-solution) ，详细了解由 Visual Studio 解决方案创建的两个项目。

2. `ProviderName` 元素具有占位符值。 将其替换为你的姓名。

3. `DisplayName` 元素的 `DefaultValue` 属性有占位符。 将它替换为“My Office Add-in”。

4. `Description` 元素的 `DefaultValue` 属性有占位符。 将它替换为“A task pane add-in for Excel”。

5. 保存文件。

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a>试用

1. 使用 Visual Studio 时，测试新建的 Excel 加载项，具体方法为按 **F5** 或选择“开始”按钮来启动 Excel，功能区中显示“显示任务窗格”加载项按钮。 加载项本地托管在 IIS 上。 如果系统要求你信任证书，请信任证书，这样此加载项才能连接到它的 Office 应用程序。

2. 在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。

    ![Excel 主页菜单的屏幕截图，突出显示“显示任务窗格”按钮。](../images/excel-quickstart-addin-2a.png)

3. 选择工作表中的任何一系列单元格。

4. 在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。

    ![Excel 屏幕截图，其中加载项任务窗格处于打开状态。](../images/excel-quickstart-addin-2c.png)

[!include[Console tool note](../includes/console-tool-note.md)]

### <a name="next-steps"></a>后续步骤

恭喜，你已成功创建 Excel 任务窗格加载项！ 接下来，了解有关[使用 Visual Studio 开发 Office 加载项](../develop/develop-add-ins-visual-studio.md)的详细信息。

---

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
- [Excel 加载项中的 Word JavaScript 对象模型](../excel/excel-add-ins-core-concepts.md)
- [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Excel JavaScript API 参考](../reference/overview/excel-add-ins-reference-overview.md)
- [使用 Visual Studio Code 发布](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
