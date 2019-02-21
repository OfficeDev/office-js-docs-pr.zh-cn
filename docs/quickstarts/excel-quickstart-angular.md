---
title: 使用 Angular 生成 Excel 加载项
description: ''
ms.date: 10/19/2018
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: da3e0ce9c6cb075ff828344956cfe25fb579a8fa
ms.sourcegitcommit: 03773fef3d2a380028ba0804739d2241d4b320e5
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2019
ms.locfileid: "30091251"
---
# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="0dee7-102">使用 Angular 生成 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="0dee7-102">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="0dee7-103">本文将逐步介绍如何使用 Angular 和 Excel JavaScript API 生成 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="0dee7-103">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0dee7-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="0dee7-104">Prerequisites</span></span>

- [<span data-ttu-id="0dee7-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="0dee7-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="0dee7-106">全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="0dee7-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="0dee7-107">创建 Web 应用</span><span class="sxs-lookup"><span data-stu-id="0dee7-107">Create the web app</span></span>

1. <span data-ttu-id="0dee7-108">使用 Yeoman 生成器创建 Excel 加载项项目。</span><span class="sxs-lookup"><span data-stu-id="0dee7-108">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="0dee7-109">运行下面的命令，再回答如下所示的提示问题：</span><span class="sxs-lookup"><span data-stu-id="0dee7-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="0dee7-110">**选择项目类型:** `Office Add-in project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="0dee7-110">**Choose a project type:** `Office Add-in project using Angular framework`</span></span>
    - <span data-ttu-id="0dee7-111">**选择脚本类型:** `Typescript`</span><span class="sxs-lookup"><span data-stu-id="0dee7-111">**Choose a script type:** `Typescript`</span></span>
    - <span data-ttu-id="0dee7-112">**要如何命名加载项?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="0dee7-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="0dee7-113">**要支持哪一个 Office 客户端应用？：**`Excel`</span><span class="sxs-lookup"><span data-stu-id="0dee7-113">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Yeoman 生成器](../images/yo-office-excel-angular.png)
    
    <span data-ttu-id="0dee7-115">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="0dee7-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="0dee7-116">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="0dee7-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="0dee7-117">更新代码</span><span class="sxs-lookup"><span data-stu-id="0dee7-117">Update the code</span></span>

1. <span data-ttu-id="0dee7-118">在代码编辑器中，打开文件“**app.css**”，将以下样式添加到文件的末尾，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="0dee7-118">In your code editor, open the file **app.css**, add the following styles to the end of the file, and save the file.</span></span>

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

2. <span data-ttu-id="0dee7-119">打开文件“**src/app/app.component.html**”，将全部内容替换为以下代码，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="0dee7-119">Open the file **src/app/app.component.html**, replace the entire contents with the following code, and save the file.</span></span>

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

3. <span data-ttu-id="0dee7-120">打开文件“**src/app/app.component.ts**”，将全部内容替换为以下代码，然后保存文件。</span><span class="sxs-lookup"><span data-stu-id="0dee7-120">Open the file **src/app/app.component.ts**, replace the entire contents with the following code, and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="0dee7-121">更新清单</span><span class="sxs-lookup"><span data-stu-id="0dee7-121">Update the manifest</span></span>

1. <span data-ttu-id="0dee7-122">打开文件“**manifest.xml**”以定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="0dee7-122">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="0dee7-123">`ProviderName` 元素具有占位符值。</span><span class="sxs-lookup"><span data-stu-id="0dee7-123">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="0dee7-124">将其替换为你的姓名。</span><span class="sxs-lookup"><span data-stu-id="0dee7-124">Replace it with your name.</span></span>

3. <span data-ttu-id="0dee7-125">`Description` 元素的 `DefaultValue` 属性有占位符。</span><span class="sxs-lookup"><span data-stu-id="0dee7-125">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="0dee7-126">将它替换为“A task pane add-in for Excel”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="0dee7-126">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="0dee7-127">保存文件。</span><span class="sxs-lookup"><span data-stu-id="0dee7-127">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="0dee7-128">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="0dee7-128">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="0dee7-129">试用</span><span class="sxs-lookup"><span data-stu-id="0dee7-129">Try it out</span></span>

1. <span data-ttu-id="0dee7-130">请按照运行加载项和在 Excel 中旁加载加载项时所用平台对应的说明操作。</span><span class="sxs-lookup"><span data-stu-id="0dee7-130">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="0dee7-131">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="0dee7-131">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="0dee7-132">Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="0dee7-132">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="0dee7-133">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="0dee7-133">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="0dee7-134">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="0dee7-134">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="0dee7-136">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="0dee7-136">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="0dee7-137">在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="0dee7-137">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="0dee7-139">后续步骤</span><span class="sxs-lookup"><span data-stu-id="0dee7-139">Next steps</span></span>

<span data-ttu-id="0dee7-p104">恭喜！已使用 Angular 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="0dee7-p104">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0dee7-142">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="0dee7-142">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="0dee7-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0dee7-143">See also</span></span>

* [<span data-ttu-id="0dee7-144">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="0dee7-144">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="0dee7-145">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="0dee7-145">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="0dee7-146">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="0dee7-146">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="0dee7-147">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="0dee7-147">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

