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

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>生成清单文件并旁加载加载项

加载项的清单文件定义其设置和功能。

1. 转到应用程序文件夹。

    ```bash
    cd my-addin
    ```

2. 使用 Yeoman 生成器生成加载项的清单文件。 运行下面的命令，再回答提示问题，如以下屏幕截图所示：

    ```bash
    yo office
    ```
    - **是否要为项目新建子文件夹?：**`No`
    - **要将你的外接程序命名为什么?:** `My Office Add-in`
    - **要支持哪一个 Office 客户端应用程序?:** `Excel`
    - **是否要新建加载项?：**`No`

    然后，生成器会询问是否要打开“resource.html”****。在本教程中，无需打开此文件。不过，如果感兴趣，也可以自行打开。选择“是”或“否”完成向导，并允许生成器执行它的工作。

    ![Yeoman 生成器](../images/yo-office.png)
    
    > [!NOTE]
    > 如果系统提示覆盖 **package.json**，请回答“否”****（不覆盖）。

3. 打开清单文件（即应用根目录中名称以“manifest.xml”结尾的文件）。将所有 `https://localhost:3000` 都替换为 `http://localhost:4200`，再保存此文件。

    > [!TIP]
    > 除了将端口号更改为 **4200** 外，还请务必将协议更改为 **http**。

4. 请按照运行加载项所用平台对应的说明操作，以在 Excel 中旁加载加载项。

    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>更新应用

1. 打开“src/index.html”****，紧靠 `</head>` 标记前面添加以下 `<script>` 标记，再保存此文件。

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. 打开“src/main.ts”****，将 `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` 替换为以下代码，再保存此文件。 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

3. 打开“src/polyfills.ts”****，在其他所有现有 `import` 语句上方添加以下代码行，再保存此文件。

    ```typescript
    import 'core-js/client/shim';
    ```

4. 在“src/polyfills.ts”****中，取消注释以下代码行，再保存此文件。

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

5. 打开“src/app/app.component.html”****，将文件内容替换为以下 HTML，再保存此文件。 

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

6. 打开“src/app/app.component.css”****，将文件内容替换为以下 CSS 代码，再保存此文件。

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

7. 打开“src/app/app.component.ts”****，将文件内容替换为以下代码，再保存此文件。 

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

## <a name="try-it-out"></a>试用

1. 通过终端运行以下命令，启动开发服务器。

    ```bash
    npm start
    ```
   
2. 在 Excel 中，依次选择“开始”****选项卡和功能区中的“显示任务窗格”****按钮，打开加载项任务窗格。

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2a.png)

3. 选择工作表中的任何一系列单元格。

4. 在任务窗格中，选择“设置颜色”****按钮，将选定区域的颜色设置为绿色。

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>后续步骤

恭喜！已使用 Angular 成功创建 Excel 加载项！ 接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。

> [!div class="nextstepaction"]
> [Excel 加载项教程](../tutorials/excel-tutorial-create-table.md)

## <a name="see-also"></a>另请参阅

* [Excel 加载项教程](../tutorials/excel-tutorial-create-table.md)
* [Excel JavaScript API 核心概念](../excel/excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API 参考](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)

