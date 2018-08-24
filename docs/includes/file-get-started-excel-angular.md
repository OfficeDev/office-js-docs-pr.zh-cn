# <a name="build-an-excel-add-in-using-angular"></a>使用 Angular 生成 Excel 加载项

在本文中，你将完成使用 Angular 和 Excel JavaScript API 生成 Excel 加载项的过程。

## <a name="prerequisites"></a>先决条件

- 检查是否已有 [Angular CLI 必备组件](https://github.com/angular/angular-cli#prerequisites)，并安装缺少的任何必备组件。

- 全局安装 [Angular CLI](https://github.com/angular/angular-cli)。 

    ```bash
    npm install -g @angular/cli
    ```

- 全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a>生成新的 Angular 应用

使用 Angular CLI 生成 Angular 应用。 在终端运行以下命令：

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a>生成清单文件

加载项清单文件定义加载项的设置和功能。

1. 转到应用程序文件夹。

    ```bash
    cd my-addin
    ```

2. 使用 Yeoman 生成器生成加载项清单文件。 运行下面的命令，再回答如下所示的提示问题。

    ```bash
    yo office 
    ```

    - **选择一个项目类型：** `Office Add-in containing the manifest only`
    - **要将你的外接程序命名为什么?:** `My Office Add-in`
    - **要支持哪一个 Office 客户端应用程序?:** `Excel`

    完成向导后，可以使用清单文件和资源文件来构建项目。

    ![Yeoman 生成器](../images/yo-office.png)
    
    > [!NOTE]
    > 如果系统提示覆盖 **package.json**，请回答“否”****（不覆盖）。

## <a name="secure-the-app"></a>保护应用程序

[!include[HTTPS guidance](../includes/https-guidance.md)]

对于本快速入门，可以使用**适用于 Office 外接程序的 Yeoman 生成器**提供的证书。 由于已在全局范围内安装了生成器（作为此快速入门的**先决条件**的一部分），因此只需将证书从全局安装位置复制到应用程序文件夹即可。 下面逐步介绍了如何完成此过程。

1. 在终端运行以下命令，以确定其中安装了全局 **npm** 库的文件夹：

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > 由该命令生成的第一行输出指定在其中安装全局 **npm** 库的文件夹。          
    
2. 使用文件资源管理器转到 `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` 文件夹。 从此位置将 `certs` 文件夹复制到剪贴板。

3. 转到在上一部分中的第 1 步创建的 Angular 应用程序的根文件夹，并将 `certs` 文件夹从剪贴板粘贴到此文件夹中。

## <a name="update-the-app"></a>更新应用程序

1. 在代码编辑器中，打开项目根目录中的 **package.json**。 将 `start` 脚本修改为指定服务器应使用 SSL 和端口 3000 运行，并保存文件。

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. 打开项目根目录中的 **.angular-cli.json**。 将 **defaults** 对象修改为指定证书文件位置，并保存文件。

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

3. 打开 **src/index.html**，紧靠 `</head>` 标记前面添加以下 `<script>` 标记，再保存此文件。

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. 打开“src/main.ts”****，将 `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` 替换为以下代码，再保存此文件。 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. 打开“src/polyfills.ts”****，在其他所有现有 `import` 语句上方添加以下代码行，再保存此文件。

    ```typescript
    import 'core-js/client/shim';
    ```

6. 在“src/polyfills.ts”**** 中，取消注释以下代码行，再保存此文件。

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

7. 打开“src/app/app.component.html”****，将文件内容替换为以下 HTML，再保存此文件。 

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

8. 打开“src/app/app.component.css”****，将文件内容替换为以下 CSS 代码，再保存此文件。

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

9. 打开 **src/app/app.component.ts**，将文件内容替换为下列代码，再保存此文件。 

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

## <a name="start-the-dev-server"></a>启动开发人员服务器

1. 通过终端运行下面的命令，以启动开发人员服务器。

    ```bash
    npm run start
    ```

2. 在 Web 浏览器中，转到 `https://localhost:3000`。如果浏览器指明网站证书不受信任，需要将此证书添加为受信任的证书。有关详细信息，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

    > [!NOTE]
    > Chrome（Web 浏览器）可能会继续指明网站证书不受信任，即使已完成[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)中所述的过程，也是如此。 可以忽略 Chrome 中的此警告，并转到 Internet Explorer 或 Microsoft Edge 中的 `https://localhost:3000`，以验证证书是否受信任。 

3. 如果浏览器在加载加载项页面后没有显示任何证书错误，就可以准备测试加载项了。 

## <a name="try-it-out"></a>试用

1. 请按照运行加载项和在 Excel 中旁加载加载项时所用平台对应的说明操作。

    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

   
2. 在 Excel 中，依次选择“主页”**** 选项卡和功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2a.png)

3. 选择工作表中的任何一系列单元格。

4. 在任务窗格中，选择“设置颜色”**** 按钮，将选定区域的颜色设置为绿色。

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>后续步骤

恭喜！已使用 Angular 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。

> [!div class="nextstepaction"]
> [Excel 加载项教程](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>另请参阅

* [Excel 加载项教程](../tutorials/excel-tutorial-create-table.md)
* [Excel JavaScript API 核心概念](../excel/excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API 参考](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
