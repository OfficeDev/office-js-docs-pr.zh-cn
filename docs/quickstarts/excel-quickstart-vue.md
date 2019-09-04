---
title: 使用 Vue 生成 Excel 任务窗格加载项
description: ''
ms.date: 09/04/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 9947852a586570345ba9f3dfe09340af6d01ace6
ms.sourcegitcommit: 78998a9f0ebb81c4dd2b77574148b16fe6725cfc
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/03/2019
ms.locfileid: "36715626"
---
# <a name="build-an-excel-task-pane-add-in-using-vue"></a><span data-ttu-id="5418a-102">使用 Vue 生成 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="5418a-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="5418a-103">本文将逐步介绍如何使用 Vue 和 Excel JavaScript API 生成 Excel 任务加载项。</span><span class="sxs-lookup"><span data-stu-id="5418a-103">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5418a-104">先决条件</span><span class="sxs-lookup"><span data-stu-id="5418a-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="5418a-105">全局安装 [Vue CLI](https://cli.vuejs.org/)。</span><span class="sxs-lookup"><span data-stu-id="5418a-105">Install the [Vue CLI](https://cli.vuejs.org/) globally.</span></span>

  ```command&nbsp;line
  npm install -g @vue/cli
  ```

## <a name="generate-a-new-vue-app"></a><span data-ttu-id="5418a-106">生成新 Vue 应用程序</span><span class="sxs-lookup"><span data-stu-id="5418a-106">Generate a new Vue app</span></span>

<span data-ttu-id="5418a-p101">使用 Vue CLI 生成新的 Vue 应用。从终端运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="5418a-p101">Use the Vue CLI to generate a new Vue app. From the terminal, run the following command and then answer the prompts as described below.</span></span>

```command&nbsp;line
vue create my-add-in
```

<span data-ttu-id="5418a-109">然后，选择 `default` 预设项。</span><span class="sxs-lookup"><span data-stu-id="5418a-109">Then select the `default` preset.</span></span> <span data-ttu-id="5418a-110">如果系统提示你使用 Yarn 或 NPM 作为包，可任选其一。</span><span class="sxs-lookup"><span data-stu-id="5418a-110">If you are prompted to use either Yarn or NPM as a package you can choose either one.</span></span>

## <a name="generate-the-manifest-file"></a><span data-ttu-id="5418a-111">生成清单文件</span><span class="sxs-lookup"><span data-stu-id="5418a-111">Generate the manifest file</span></span>

<span data-ttu-id="5418a-112">每个加载项都需要定义自己设置和功能的清单文件。</span><span class="sxs-lookup"><span data-stu-id="5418a-112">Each add-in requires a manifest file to define its settings and capabilities.</span></span>

1. <span data-ttu-id="5418a-113">转到应用程序文件夹。</span><span class="sxs-lookup"><span data-stu-id="5418a-113">Navigate to your app folder.</span></span>

   ```command&nbsp;line
   cd my-add-in
   ```

2. <span data-ttu-id="5418a-p103">使用 Yeoman 生成器生成加载项清单文件。 运行下面的命令，再回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="5418a-p103">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

   ```command&nbsp;line
   yo office
   ```

   ![Yeoman 生成器](../images/yo-office-manifest-only-vue.png)

   - <span data-ttu-id="5418a-117">**选择项目类型:** `Office Add-in project containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="5418a-117">**Choose a project type:** `Office Add-in project containing the manifest only`</span></span>
   - <span data-ttu-id="5418a-118">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="5418a-118">**What do you want to name your add-in?**</span></span> `my-office-add-in`
   - <span data-ttu-id="5418a-119">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="5418a-119">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="5418a-120">完成向导后，会创建一个 `my-office-add-in` 文件夹，其中包含一个 `manifest.xml` 文件。</span><span class="sxs-lookup"><span data-stu-id="5418a-120">After you complete the wizard, it creates a `my-office-add-in` folder, which contains a `manifest.xml` file.</span></span> <span data-ttu-id="5418a-121">你将在本快速入门结束时使用该清单旁加载和测试你的加载项。</span><span class="sxs-lookup"><span data-stu-id="5418a-121">You will use the manifest to sideload and test your add-in at the end of the quick start.</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="5418a-122">保护应用</span><span class="sxs-lookup"><span data-stu-id="5418a-122">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="5418a-123">要为应用启用 HTTPS，请使用以下内容在 Vue 项目的根文件夹中创建一个 `vue.config.js` 文件：</span><span class="sxs-lookup"><span data-stu-id="5418a-123">To enable HTTPS for your app, create a `vue.config.js` file in the root folder of the Vue project with the following contents:</span></span>

```js
module.exports = {
  devServer: {
    port: 3000,
    https: true
  }
};
```

## <a name="update-the-app"></a><span data-ttu-id="5418a-124">更新应用</span><span class="sxs-lookup"><span data-stu-id="5418a-124">Update the app</span></span>

1. <span data-ttu-id="5418a-125">打开 `public/index.html` 文件，在紧靠 `</head>` 标记的前面添加以下 `<script>` 标记：</span><span class="sxs-lookup"><span data-stu-id="5418a-125">Open `public/index.html`, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

   ```html
   <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
   ```

2. <span data-ttu-id="5418a-126">打开 `src/main.js`，将内容替换为以下代码：</span><span class="sxs-lookup"><span data-stu-id="5418a-126">Open the `src/main.js` file and replace its contents with the following code:</span></span>

   ```js
   import Vue from 'vue';
   import App from './App.vue';

   Vue.config.productionTip = false;

   window.Office.initialize = () => {
     new Vue({
       render: h => h(App)
     }).$mount('#app');
   };
   ```

3. <span data-ttu-id="5418a-127">打开 `src/App.vue`，将文件内容替换为以下代码：</span><span class="sxs-lookup"><span data-stu-id="5418a-127">Open the `src/App.vue` file and replace its contents with the following code:</span></span>

   ```html
   <template>
     <div id="app">
       <div class="content">
         <div class="content-header">
           <div class="padding">
             <h1>Welcome</h1>
           </div>
         </div>
         <div id="content-main">
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

## <a name="start-the-dev-server"></a><span data-ttu-id="5418a-128">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="5418a-128">Start the dev server</span></span>

1. <span data-ttu-id="5418a-129">通过终端运行下面的命令，以启动开发人员服务器。</span><span class="sxs-lookup"><span data-stu-id="5418a-129">From the terminal, run the following command to start the dev server.</span></span>

   ```command&nbsp;line
   npm run serve
   ```

2. <span data-ttu-id="5418a-130">在 Web 浏览器中，导航到 `https://localhost:3000`（请注意 `https`）。</span><span class="sxs-lookup"><span data-stu-id="5418a-130">In a web browser, navigate to `https://localhost:3000` (notice the `https`).</span></span> <span data-ttu-id="5418a-131">如果浏览器指明网站证书不受信任，则需要[将计算机配置为信任此证书](https://github.com/OfficeDev/generator-office/blob/fd600bbe00747e64aa5efb9846295a3f66d428aa/src/docs/ssl.md#add-certification-file-through-ie)。</span><span class="sxs-lookup"><span data-stu-id="5418a-131">If your browser indicates that the site's certificate is not trusted, you will need to configure your computer to trust the certificate.</span></span>

3. <span data-ttu-id="5418a-132">如果 `https://localhost:3000` 上的页面空白但没有任何证书错误，这表示它正常工作。</span><span class="sxs-lookup"><span data-stu-id="5418a-132">When the page on `https://localhost:3000` is blank and without any certificate errors, it means that it is working.</span></span> <span data-ttu-id="5418a-133">Office 初始化后装载 Vue 应用，因此它仅显示 Excel 环境中的内容。</span><span class="sxs-lookup"><span data-stu-id="5418a-133">The Vue App is mounted after Office is initialized, so it only shows things inside of an Excel environment.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="5418a-134">试用</span><span class="sxs-lookup"><span data-stu-id="5418a-134">Try it out</span></span>

1. <span data-ttu-id="5418a-135">请按照运行加载项和在 Excel 中旁加载加载项时所用平台对应的说明操作。</span><span class="sxs-lookup"><span data-stu-id="5418a-135">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

   - <span data-ttu-id="5418a-136">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="5418a-136">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
   - <span data-ttu-id="5418a-137">Web 浏览器：[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span><span class="sxs-lookup"><span data-stu-id="5418a-137">Web browser: [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)</span></span>
   - <span data-ttu-id="5418a-138">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="5418a-138">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="5418a-139">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="5418a-139">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

   ![Excel 加载项按钮](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="5418a-141">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="5418a-141">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="5418a-142">在任务窗格中，选择“**设置颜色**”按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="5418a-142">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

   ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="5418a-144">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5418a-144">Next steps</span></span>

<span data-ttu-id="5418a-145">祝贺，你已使用 Vue 成功创建了 Excel 任务窗格加载项！</span><span class="sxs-lookup"><span data-stu-id="5418a-145">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="5418a-146">接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="5418a-146">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5418a-147">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="5418a-147">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="5418a-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5418a-148">See also</span></span>

* [<span data-ttu-id="5418a-149">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="5418a-149">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="5418a-150">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="5418a-150">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="5418a-151">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="5418a-151">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="5418a-152">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="5418a-152">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
