# <a name="build-an-excel-add-in-using-jquery"></a>使用 jQuery 生成 Excel 加载项

在本文中，你将完成使用 jQuery 和 Excel JavaScript API 生成 Excel 加载项的过程。

## <a name="prerequisites"></a>先决条件

如果以前未执行过此操作，则需要全局安装 [Yeoman](https://github.com/yeoman/yo) 和 [适用于 Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a>创建 Web 应用

1. 在本地驱动器上创建一个文件夹，并命名为“my-addin”****。 将在其中创建应用程序文件。

2. 转到应用程序文件夹。

    ```bash
    cd my-addin
    ```

3. 使用 Yeoman 生成器生成加载项的清单文件。 运行下面的命令，再回答提示问题，如以下屏幕截图所示：

    ```bash
    yo office
    ```
    ![Yeoman 生成器](../../images/yo-office-jquery.png)


4. 在代码编辑器中，打开项目根目录中的 **index.html**。 该文件指定将在加载项的任务窗格中呈现的 HTML。 
 
5. 将生成的 `header` 标记替换为以下标记。
 
    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

6. 将生成的 `main` 标记替换为以下标记，然后保存文件。

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button class="ms-Button" id="set-color">Set color</button>
        </div>
    </div>
    ```

7. 打开文件 **app.js** 以指定加载项的脚本。 将生成的被立即调用的函数表达式替换为以下代码并保存该文件。

    ```js
    (function () {
        "use strict";

        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

8. 打开文件 **app.css** 以指定加载项的自定义样式。 将内容（版权注释除外）替换为以下内容并保存该文件。

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

## <a name="configure-the-manifest-file-and-sideload-the-add-in"></a>配置清单文件并旁加载加载项

1. 打开文件 **my-office-add-in-manifest.xml** 以定义加载项的设置和功能。 

2. **ProviderName** 标记具有占位符值。 将其更改为 `Microsoft`。

3. **DisplayName** 标记的 **DefaultValue** 具有占位符值。 将其更改为 `A task pane add-in for Excel`。 

4. 保存但不关闭文件。

## <a name="configure-to-use-http"></a>配置为使用 HTTP

Office Web 加载项应使用 HTTPS，而不是 HTTP，即使在开发时也是如此。 但是，为了快速启动并运行加载项，此快速启动过程将使用 HTTP。 若要实现此目的，请执行以下步骤：

1. 在清单文件 **my-office-add-in-manifest.xml** 中，将“https”全部替换为“http”。 然后保存并关闭文件。

2. 打开项目根目录中的 **bsconfig.json** 文件。 将 **https** 属性的值更改为 `false`。 保存文件。


## <a name="try-it-out"></a>试用

1. 遵循将用于运行并在 Excel 中旁加载加载项的平台所适用的说明。

    - Windows：[在 Windows 上旁加载 Office 加载项进行测试](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. 打开项目根目录下的 bash 终端，运行以下命令启动开发服务器。

    ```bash
    npm start
    ```

   > **注意**：浏览器窗口将打开，其中包含加载项。 关闭此窗口。

3. 在 Excel 中，选择“主页”****选项卡，然后选择功能区中的“显示任务窗格”****按钮，以打开加载项任务窗格。

    ![Excel 加载项按钮](../../images/excel_quickstart_addin_2a.png)

4. 选择工作表中的任意单元格区域。

5. 在此任务窗格中，选择“为我设置颜色”****窗格按钮，将所选区域的颜色设置为绿色。

    ![Excel 加载项](../../images/excel_quickstart_addin_2b.png)

## <a name="next-steps"></a>后续步骤

祝贺你，你已使用 jQuery 成功创建了 Excel 加载项！ 接下来，详细了解关于生成 Excel 加载项的[核心概念](excel-add-ins-core-concepts.md)。

## <a name="additional-resources"></a>其他资源

* [Excel JavaScript API 核心概念](excel-add-ins-core-concepts.md)
* [通过脚本实验室探索代码段](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Excel 加载项代码示例](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API 参考](../../reference/excel/excel-add-ins-reference-overview.md)
