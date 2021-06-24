---
title: 使用 Vue 生成 Excel 任务窗格加载项
description: 了解如何使用 Office JS API 和 Vue 生成简单的 Excel 任务窗格加载项。
ms.date: 06/16/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: ec216e84e9aa4bc7eabec4b20c7a2dd271ca1718
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076614"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="e174d-103">使用 Vue 生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="e174d-103">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="e174d-104">本文将逐步介绍如何使用 Vue 和 Excel JavaScript API 生成 Excel 任务加载项。</span><span class="sxs-lookup"><span data-stu-id="e174d-104">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e174d-105">先决条件</span><span class="sxs-lookup"><span data-stu-id="e174d-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="e174d-106">全局安装 [Vue CLI](https://cli.vuejs.org/)。</span><span class="sxs-lookup"><span data-stu-id="e174d-106">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="e174d-107">生成新 Vue 应用程序</span><span class="sxs-lookup"><span data-stu-id="e174d-107">Generate a new Vue app</span></span>

<span data-ttu-id="e174d-p101">使用 Vue CLI 生成新的 Vue 应用。从终端运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="e174d-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="e174d-110">然后选择“Vue 3”的 `Default` 预设（如果愿意，可以选择使用“Vue 2”）。</span><span class="sxs-lookup"><span data-stu-id="e174d-110">Then select the `Default` preset for "Vue 3" (you may choose to use "Vue 2" if you'd prefer).</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="e174d-111">生成清单文件</span><span class="sxs-lookup"><span data-stu-id="e174d-111">Generate the manifest file</span></span>

<span data-ttu-id="e174d-112">每个加载项都需要定义自己设置和功能的清单文件。</span><span class="sxs-lookup"><span data-stu-id="e174d-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="e174d-113">转到应用程序文件夹。</span><span class="sxs-lookup"><span data-stu-id="e174d-113">Navigate to your app folder.</span></span>

    ```command&nbsp;line
    cd my-add-in
    ```

2. <span data-ttu-id="e174d-114">通过运行以下命令，使用 Yeoman 生成器生成加载项清单文件：</span><span class="sxs-lookup"><span data-stu-id="e174d-114">Use the Yeoman generator to generate the manifest file for your add-in by running the following command:</span></span>

    ```command&nbsp;line
    yo office
    ```

    > [!NOTE]
    > <span data-ttu-id="e174d-115">运行该`yo office`命令时，可能会收到有关 Yeoman 和 Office 加载项 CLI 工具的数据收集策略的提示。</span><span class="sxs-lookup"><span data-stu-id="e174d-115">When you run the `yo office` command, you may receive prompts about the data collection policies of Yeoman and the Office Add-in CLI tools.</span></span> <span data-ttu-id="e174d-116">根据你的需要，使用提供的信息来响应提示。</span><span class="sxs-lookup"><span data-stu-id="e174d-116">Use the information that's provided to respond to the prompts as you see fit.</span></span> <span data-ttu-id="e174d-117">如果在对第二条提示的响应中选择“**退出**”，则在准备好创建加载项项目时，需要再次运行 `yo office` 命令。</span><span class="sxs-lookup"><span data-stu-id="e174d-117">If you choose **Exit** in response to the second prompt, you'll need to run the `yo office` command again when you're ready to create your add-in project.</span></span>

    <span data-ttu-id="e174d-118">出现提示时，请提供以下信息来创建加载项项目：</span><span class="sxs-lookup"><span data-stu-id="e174d-118">When prompted, provide the following information to create your add-in project:</span></span>

    - <span data-ttu-id="e174d-119">**选择项目类型:** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="e174d-119">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
    - <span data-ttu-id="e174d-120">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="e174d-120">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="e174d-121">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="e174d-121">**Which Office client application would you like to support?**</span></span> `Excel`

    ![项目类型设置为“仅清单” 的 Yeoman Office 加载项生成器命令行界面屏幕截图。](../images/yo-office-manifest-only-vue.png)

<span data-ttu-id="e174d-123">完成向导后，会创建一个 `My Office Add-in` 文件夹，其中包含一个 `manifest.xml` 文件。</span><span class="sxs-lookup"><span data-stu-id="e174d-123">After you complete the wizard, it creates a `My Office Add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="e174d-124">你将在本快速入门结束时使用该清单旁加载和测试你的加载项。</span><span class="sxs-lookup"><span data-stu-id="e174d-124">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

> [!TIP]
> <span data-ttu-id="e174d-125">创建加载项项目后，可忽略 Yeoman 生成器提供的 *后续步骤* 指南。</span><span class="sxs-lookup"><span data-stu-id="e174d-125">You can ignore the *next steps* guidance that the Yeoman generator provides after the add-in project's been created.</span></span> <span data-ttu-id="e174d-126">本文中的分步说明提供了完成本教程所需的全部指南。</span><span class="sxs-lookup"><span data-stu-id="e174d-126">The step-by-step instructions within this article provide all of the guidance you'll need to complete this tutorial.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="e174d-127">保护应用</span><span class="sxs-lookup"><span data-stu-id="e174d-127">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

1. <span data-ttu-id="e174d-128">要为应用启用 HTTPS，请使用以下内容在 Vue 项目的根文件夹中创建一个 `vue.config.js` 文件：</span><span class="sxs-lookup"><span data-stu-id="e174d-128">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

    ```js
    var fs = require("fs");
    var path = require("path");
    var homedir = require('os').homedir()
  
    module.exports = {
      devServer: {
        port: 3000,
        https: true,
        key: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.key`)),
        cert: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/localhost.crt`)),
        ca: fs.readFileSync(path.resolve(`${homedir}/.office-addin-dev-certs/ca.crt`))
      }
    }
    ```

2. <span data-ttu-id="e174d-129">在终端中，运行以下命令以安装加载项证书。</span><span class="sxs-lookup"><span data-stu-id="e174d-129">From the terminal, run the following command to install the add-in's certificates.</span></span>

   ```command&nbsp;line
   npx office-addin-dev-certs install
   ```

## <a name="update-the-app"></a><span data-ttu-id="e174d-130">更新应用</span><span class="sxs-lookup"><span data-stu-id="e174d-130">Update the app</span></span>

1. <span data-ttu-id="e174d-131">打开 `public/index.html` 文件，在紧靠 `</head>` 标记的前面添加以下 `<script>` 标记：</span><span class="sxs-lookup"><span data-stu-id="e174d-131">Open the `public/index.html` file and add the following `<script>` tag immediately before the `</head>` tag:</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="e174d-132">打开 `src/main.js`，将内容替换为以下代码：</span><span class="sxs-lookup"><span data-stu-id="e174d-132">Open `src/main.js` and replace the contents with the following code:</span></span>

   ```js
   import { createApp } from 'vue'
   import App from './App.vue'

   window.Office.onReady(() => {
       createApp(App).mount('#app');
   });
   ```

3. <span data-ttu-id="e174d-133">打开 `src/App.vue`，将文件内容替换为以下代码：</span><span class="sxs-lookup"><span data-stu-id="e174d-133">Open `src/App.vue` and replace the file contents with the following code:</span></span>

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div class="content-main">
           <div class="padding">
             <p>
               Choose the button below to set the color of the selected range to
               green.
             </p>
             <br />
             <h3>Try it out</h3>
             <button @click="onSetColor">Set color</button>
           </div>
         </div>
       </div>
     </div>
   </template>

   <script>
     export default {
       name: 'App',
       methods: {
         onSetColor() {
           window.Excel.run(async context => {
             const range = context.workbook.getSelectedRange();
             range.format.fill.color = 'green';
             await context.sync();
           });
         }
       }
     };
   </script>

   <style>
     .content-header {
       background: #2a8dd4;
       color: #fff;
       position: absolute;
       top: 0;
       left: 0;
       width: 100%;
       height: 80px;
       overflow: hidden;
     }

     .content-main {
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
   </style>
   ```

## <a name="start-the-dev-server"></a><span data-ttu-id="e174d-134">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="e174d-134">Start the dev server</span></span>

1. <span data-ttu-id="e174d-135">通过终端运行下面的命令，以启动开发人员服务器。</span><span class="sxs-lookup"><span data-stu-id="e174d-135">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="e174d-136">在 Web 浏览器中，导航到 `https://localhost:3000`（请注意 `https`）。</span><span class="sxs-lookup"><span data-stu-id="e174d-136">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="e174d-137">如果 `https://localhost:3000` 上的页面空白但没有任何证书错误，这表示它正常工作。</span><span class="sxs-lookup"><span data-stu-id="e174d-137">If the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="e174d-138">Office 初始化后装载 Vue 应用，因此它仅显示 Excel 环境中的内容。</span><span class="sxs-lookup"><span data-stu-id="e174d-138">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="e174d-139">试用</span><span class="sxs-lookup"><span data-stu-id="e174d-139">Try it out</span></span>

1. <span data-ttu-id="e174d-140">请按照运行加载项和在 Excel 中旁加载加载项时所用平台对应的说明操作。</span><span class="sxs-lookup"><span data-stu-id="e174d-140">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="e174d-141">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="e174d-141">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="e174d-142">Web 浏览器：[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="e174d-142">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="e174d-143">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="e174d-143">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="e174d-144">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="e174d-144">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Excel 主页菜单的屏幕截图，突出显示“显示任务窗格”按钮。](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="e174d-146">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="e174d-146">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="e174d-147">在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="e174d-147">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Excel 屏幕截图，其中加载项任务窗格处于打开状态。](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="e174d-149">后续步骤</span><span class="sxs-lookup"><span data-stu-id="e174d-149">Next steps</span></span>

<span data-ttu-id="e174d-p106">恭喜！已使用 Vue 成功创建 Excel 任务窗格！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，以生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="e174d-p106">Congratulations, you've successfully created an Excel task pane add-in using Vue! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e174d-152">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="e174d-152">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="e174d-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e174d-153">See also</span></span>

* [<span data-ttu-id="e174d-154">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="e174d-154">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="e174d-155">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="e174d-155">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="e174d-156">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="e174d-156">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="e174d-157">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="e174d-157">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="e174d-158">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="e174d-158">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
