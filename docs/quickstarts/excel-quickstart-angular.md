---
title: 使用 Angular 生成 Excel 加载项
description: ''
ms.date: 10/19/2018
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: a7600b5e53fa9c7361f2dd0c117d97937ec954c7
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742371"
---
# <a name="build-an-excel-add-in-using-angular"></a>使用 Angular 生成 Excel 加载项

在本文中，你将完成使用 Angular 和 Excel JavaScript API 生成 Excel 加载项的过程。

## <a name="prerequisites"></a>先决条件

- [Node.js](https://nodejs.org)

- 全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a>创建 Web 应用

1. 使用 Yeoman 生成器创建 Excel 加载项项目。 运行下面的命令，再回答如下所示的提示问题：

    ```bash
    yo office
    ```

    - **选择项目类型:** `Office Add-in project using Angular framework`
    - **选择脚本类型:** `Typescript`
    - **要如何命名加载项?:** `My Office Add-in`
    - **要支持哪一个 Office 客户端应用？：**`Excel`

    ![Yeoman 生成器](../images/yo-office-excel-angular.png)
    
    完成此向导后，生成器会创建项目，并安装支持的 Node 组件。

2. 导航到项目的根文件夹。

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a>更新代码

1. 在代码编辑器中，打开文件“**app.css**”，将以下样式添加到文件的末尾，然后保存文件。

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
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. 打开文件“**src/app/app.component.html**”，将全部内容替换为以下代码，然后保存文件。

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>{{welcomeMessage}}</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <br />
            <div role="button" class="ms-Button" (click)="setColor()">
                <span class="ms-Button-label">Set color</span>
                <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--ChevronRight"></i></span>
            </div>
        </div>
    </div>
    ```

3. 打开文件“**src/app/app.component.ts**”，将全部内容替换为以下代码，然后保存文件。

    ```typescript
    import { Component } from '@angular/core';
    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    const template = require('./app.component.html');

    @Component({
        selector: 'app-home',
        template
    })
    export default class AppComponent {
        welcomeMessage = 'Welcome';

        async setColor() {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

    }
    ```

## <a name="update-the-manifest"></a>更新清单

1. 打开文件“**manifest.xml**”以定义加载项的设置和功能。 

2. `ProviderName` 元素具有占位符值。 将其替换为你的姓名。

3. `Description` 元素的 `DefaultValue` 属性有占位符。 将它替换为“A task pane add-in for Excel”****。

4. 保存文件。

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a>启动开发人员服务器

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a>试用

1. 请按照运行加载项和在 Excel 中旁加载加载项时所用平台对应的说明操作。

    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

   
2. 在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2b.png)

3. 选择工作表中的任何一系列单元格。

4. 在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>后续步骤

恭喜！已使用 Angular 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。

> [!div class="nextstepaction"]
> [Excel 加载项教程](../tutorials/excel-tutorial.md)

## <a name="see-also"></a>另请参阅

* [Excel 加载项教程](../tutorials/excel-tutorial-create-table.md)
* [Excel JavaScript API 基本编程概念](../excel/excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Excel JavaScript API 参考](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

