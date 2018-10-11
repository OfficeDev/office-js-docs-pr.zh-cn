# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="61fcd-101">使用 Angular 生成 Excel 加载项</span><span class="sxs-lookup"><span data-stu-id="61fcd-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="61fcd-102">在本文中，你将完成使用 Angular 和 Excel JavaScript API 生成 Excel 加载项的过程。</span><span class="sxs-lookup"><span data-stu-id="61fcd-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="61fcd-103">先决条件</span><span class="sxs-lookup"><span data-stu-id="61fcd-103">Prerequisites</span></span>

- <span data-ttu-id="61fcd-104">检查是否已有 [Angular CLI 必备组件](https://github.com/angular/angular-cli#prerequisites)，并安装缺少的任何必备组件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-104">Check whether you already have the [Angular CLI prerequisites](https://github.com/angular/angular-cli#prerequisites) and install any prerequistes that you are missing.</span></span>

- <span data-ttu-id="61fcd-105">全局安装 [Angular CLI](https://github.com/angular/angular-cli)。</span><span class="sxs-lookup"><span data-stu-id="61fcd-105">Install the [Angular CLI](https://github.com/angular/angular-cli) globally.</span></span> 

    ```bash
    npm install -g @angular/cli
    ```

- <span data-ttu-id="61fcd-106">全局安装最新版本的 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="61fcd-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a><span data-ttu-id="61fcd-107">生成新的 Angular 应用</span><span class="sxs-lookup"><span data-stu-id="61fcd-107">Generate a new Angular app</span></span>

<span data-ttu-id="61fcd-p101">使用 Angular CLI 生成 Angular 应用。在终端运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="61fcd-p101">Use the Angular CLI to generate your Angular app. From the terminal, run the following command:</span></span>

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a><span data-ttu-id="61fcd-110">生成清单文件</span><span class="sxs-lookup"><span data-stu-id="61fcd-110">Generate the manifest file</span></span>

<span data-ttu-id="61fcd-111">加载项清单文件定义加载项的设置和功能。</span><span class="sxs-lookup"><span data-stu-id="61fcd-111">An add-in's manifest file defines its settings and capabilities.</span></span>

1. <span data-ttu-id="61fcd-112">转到应用文件夹。</span><span class="sxs-lookup"><span data-stu-id="61fcd-112">Navigate to your app folder.</span></span>

    ```bash
    cd my-addin
    ```

2. <span data-ttu-id="61fcd-p102">使用 Yeoman 生成器生成加载项的清单文件。运行下面的命令，再回答提示问题，如以下所示。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p102">Use the Yeoman generator to generate the manifest file for your add-in. Run the following command and then answer the prompts as shown below.</span></span>

    ```bash
    yo office 
    ```

    - <span data-ttu-id="61fcd-115">**选择一个项目类型：** `Office Add-in containing the manifest only`</span><span class="sxs-lookup"><span data-stu-id="61fcd-115">**Choose a project type:** `Office Add-in containing the manifest only`</span></span>
    - <span data-ttu-id="61fcd-116">**要将你的外接程序命名为什么?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="61fcd-116">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="61fcd-117">**要支持哪一个 Office 客户端应用程序?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="61fcd-117">**Which Office client application would you like to support?:** `Excel`</span></span>

    <span data-ttu-id="61fcd-118">完成向导后，可以使用清单文件和资源文件来构建项目。</span><span class="sxs-lookup"><span data-stu-id="61fcd-118">After you complete the wizard, a manifest file and resource file are available for you to build your project.</span></span>

    ![Yeoman 生成器](../images/yo-office.png)
    
    > [!NOTE]
    > <span data-ttu-id="61fcd-120">如果系统提示覆盖 \*\* package.json\*\* ，请回答\*\* 否\*\*（不覆盖）。</span><span class="sxs-lookup"><span data-stu-id="61fcd-120">If you're prompted to overwrite **package.json**, answer **No** (do not overwrite).</span></span>

## <a name="secure-the-app"></a><span data-ttu-id="61fcd-121">保护应用</span><span class="sxs-lookup"><span data-stu-id="61fcd-121">Secure the app</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

<span data-ttu-id="61fcd-p103">对于本快速入门，可以使用\*\* Office 加载项的 Yeoman 生成器**提供的证书。由于已在全局范围内安装了生成器（作为此快速入门的**先决条件\*\*的一部分），因此只需将证书从全局安装位置复制到应用文件夹即可。下面逐步介绍如何完成此过程。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p103">For this quick start, you can use the certificates that the **Yeoman generator for Office Add-ins** provides. You've already installed the generator globally (as part of the **Prerequisites** for this quick start), so you'll just need to copy the certificates from the global install location into your app folder. The following steps describe how to complete this process.</span></span>

1. <span data-ttu-id="61fcd-125">在终端运行以下命令，以确定其中安装了全局 **npm** 库的文件夹：</span><span class="sxs-lookup"><span data-stu-id="61fcd-125">From the terminal, run the following command to identify the folder where global **npm** libraries are installed:</span></span>

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > <span data-ttu-id="61fcd-126">由该命令生成的第一行输出指定在其中安装全局 **npm** 库的文件夹。</span><span class="sxs-lookup"><span data-stu-id="61fcd-126">The first line of output that's generated by this command specifies the folder where global **npm** libraries are installed.</span></span>          
    
2. <span data-ttu-id="61fcd-p104">使用文件资源管理器转到 `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` 文件夹。从此位置将 `certs` 文件夹复制到剪贴板。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p104">Using File Explorer, navigate to the `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` folder. From that location, copy the `certs` folder to your clipboard.</span></span>

3. <span data-ttu-id="61fcd-129">转到在上一部分中的第 1 步创建的 Angular 应用的根文件夹，并将 `certs` 文件夹从剪贴板粘贴到此文件夹中。</span><span class="sxs-lookup"><span data-stu-id="61fcd-129">Navigate to the root folder of the Angular app that you created in step 1 of the previous section, and paste the `certs` folder from your clipboard into that folder.</span></span>

## <a name="update-the-app"></a><span data-ttu-id="61fcd-130">更新应用</span><span class="sxs-lookup"><span data-stu-id="61fcd-130">Update the app</span></span>

1. <span data-ttu-id="61fcd-p105">在代码编辑器中，打开项目根目录中的 **package.json**。将 `start` 脚本修改为指定服务器应使用 SSL 和端口 3000 运行，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p105">In your code editor, open **package.json** in the root of the project. Modify the `start` script to specify that the server should run using SSL and port 3000, and save the file.</span></span>

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. <span data-ttu-id="61fcd-p106">打开项目根目录中的 **.angular cli.json** 。将 **defaults** 对象修改为指定证书文件位置，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p106">Open **.angular-cli.json** in the root of the project. Modify the **defaults** object to specify the location of the certificate files, and save the file.</span></span>

    ```json
    "defaults": {
      "styleExt": "css",
      "component": {},
      "serve": {
        "sslKey": "certs/server.key",
        "sslCert": "certs/server.crt"
      }
    }
    ```

3. <span data-ttu-id="61fcd-135">打开 **src/index.html**，紧靠 `</head>` 标记前面添加以下 `<script>` 标记，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-135">Open **src/index.html**, add the following `<script>` tag immediately before the `</head>` tag, and save the file.</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. <span data-ttu-id="61fcd-136">打开 **src/main.ts** ，将 `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` 替换为以下代码，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-136">Open **src/main.ts**, replace `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` with the following code, and save the file.</span></span> 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. <span data-ttu-id="61fcd-137">打开 **src/polyfills.ts**，在其他所有现有 `import` 语句上方添加以下代码行，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-137">Open **src/polyfills.ts**, add the following line of code above all other existing `import` statements, and save the file.</span></span>

    ```typescript
    import 'core-js/client/shim';
    ```

6. <span data-ttu-id="61fcd-138">在 **src/polyfills.ts** 中，取消注释以下代码行，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-138">In **src/polyfills.ts**, uncomment the following lines, and save the file.</span></span>

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

7. <span data-ttu-id="61fcd-139">打开 **src/app/app.component.html**，将文件内容替换为以下 HTML，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-139">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span> 

    ```html
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
            <button (click)="onSetColor()">Set color</button>
        </div>
    </div>
    ```

8. <span data-ttu-id="61fcd-140">打开 **src/app/app.component.css**，将文件内容替换为以下 CSS 代码，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-140">Open **src/app/app.component.css**, replace file contents with the following CSS code, and save the file.</span></span>

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

9. <span data-ttu-id="61fcd-141">打开 **src/app/app.component.ts**，将文件内容替换为下列代码，再保存此文件。</span><span class="sxs-lookup"><span data-stu-id="61fcd-141">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span> 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onSetColor() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="61fcd-142">启动开发人员服务器</span><span class="sxs-lookup"><span data-stu-id="61fcd-142">Start the dev server</span></span>

1. <span data-ttu-id="61fcd-143">通过终端运行下面的命令，以启动开发人员服务器。</span><span class="sxs-lookup"><span data-stu-id="61fcd-143">From the terminal, run the following command to start the dev server.</span></span>

    ```bash
    npm run start
    ```

2. <span data-ttu-id="61fcd-p107">在 Web 浏览器中，转到 `https://localhost:3000`。如果浏览器指明网站证书不受信任，需要将此证书添加为受信任的证书。有关详细信息，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p107">In a web browser, navigate to `https://localhost:3000`. If your browser indicates that the site's certificate is not trusted, you will need to add the certificate as a trusted certificate. See [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for details.</span></span>

    > [!NOTE]
    > <span data-ttu-id="61fcd-p108">Chrome（Web 浏览器）可能会继续指明网站证书不受信任，即使已完成[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中所述的过程，也是如此。可以忽略 Chrome 中的此警告，并转到 Internet Explorer 或 Microsoft Edge 中的 `https://localhost:3000`，以验证证书是否受信任。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p108">Chrome (web browser) may continue to indicate the the site's certificate is not trusted, even after you have completed the process described in [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). You can disregard this warning in Chrome and can verify that the certificate is trusted by navigating to `https://localhost:3000` in either Internet Explorer or Microsoft Edge.</span></span> 

3. <span data-ttu-id="61fcd-149">如果浏览器在加载加载项页面后没有显示任何证书错误，就可以准备测试加载项了。</span><span class="sxs-lookup"><span data-stu-id="61fcd-149">After your browser loads the add-in page without any certificate errors, you're ready test your add-in.</span></span> 

## <a name="try-it-out"></a><span data-ttu-id="61fcd-150">试用</span><span class="sxs-lookup"><span data-stu-id="61fcd-150">Try it out</span></span>

1. <span data-ttu-id="61fcd-151">请按照运行加载项和在 Excel 中旁加载加载项时所用平台对应的说明操作。</span><span class="sxs-lookup"><span data-stu-id="61fcd-151">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="61fcd-152">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="61fcd-152">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="61fcd-153">Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="61fcd-153">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="61fcd-154">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="61fcd-154">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="61fcd-155">在 Excel 中，依次选择**主页**选项卡和功能区中的**显示任务窗格**按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="61fcd-155">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="61fcd-157">选择工作表中的任何一系列单元格。</span><span class="sxs-lookup"><span data-stu-id="61fcd-157">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="61fcd-158">在任务窗格中，选择**设置颜色**按钮，将选定区域的颜色设置为绿色。</span><span class="sxs-lookup"><span data-stu-id="61fcd-158">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="61fcd-160">后续步骤</span><span class="sxs-lookup"><span data-stu-id="61fcd-160">Next steps</span></span>

<span data-ttu-id="61fcd-p109">恭喜！已使用 Angular 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="61fcd-p109">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="61fcd-163">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="61fcd-163">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="61fcd-164">另请参阅</span><span class="sxs-lookup"><span data-stu-id="61fcd-164">See also</span></span>

* [<span data-ttu-id="61fcd-165">Excel 加载项教程</span><span class="sxs-lookup"><span data-stu-id="61fcd-165">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="61fcd-166">使用 Excel JavaScript API 的基本编程概念</span><span class="sxs-lookup"><span data-stu-id="61fcd-166">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="61fcd-167">Excel 加载项代码示例</span><span class="sxs-lookup"><span data-stu-id="61fcd-167">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="61fcd-168">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="61fcd-168">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
