# <a name="build-an-excel-add-in-using-react"></a>使用 React 生成 Excel 加载项

本文将逐步介绍如何使用 React 和 Excel JavaScript API 生成 Excel 加载项。

## <a name="environment"></a>环境

- **Office 桌面**：确保你安装了最新版本的 Office。 加载项命令需要内部版本 16.0.6769.0000 或更高版本（推荐 **16.0.6868.0000**）。 学习如何 [安装最新版本的 Office 应用程序](http://aka.ms/latestoffice)。 
 
- **Office Online**：没有额外的设置。 请注意，对工作/学校帐户的 Office Online 命令的支持处于预览状态。

## <a name="prerequisites"></a>先决条件

- 全局安装 [Create React App](https://github.com/facebookincubator/create-react-app)。

    ```bash
    npm install -g create-react-app
    ```

- 全局安装最新版 [Yeoman](https://github.com/yeoman/yo) 和 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)。

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a>生成新的 React 应用

使用 Create React App 生成 React 应用。 在终端运行以下命令：

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>生成清单文件并旁加载加载项

每个加载项都需要用于定义其设置和功能的清单文件。

1. 转到应用程序文件夹。

    ```bash
    cd my-addin
    ```

2. 使用 Yeoman 生成器生成加载项的清单文件。 运行下面的命令，再回答提示问题，如以下屏幕截图所示：

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

3. 请按照运行加载项所用平台对应的说明操作，以在 Excel 中旁加载加载项。

    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>更新应用

1. 打开“public/index.html”****，紧靠 `</head>` 标记前面添加以下 `<script>` 标记，再保存此文件。

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. 打开“src/index.js”****，将 `ReactDOM.render(<App />, document.getElementById('root'));` 替换为以下代码，再保存此文件。 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. 打开“src/App.js”****，将文件内容替换为以下代码，再保存此文件。 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. 打开“src/App.css”****，将文件内容替换为以下 CSS 代码，再保存此文件。 

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

## <a name="try-it-out"></a>试用

1. 通过终端运行下面的命令，以启动开发人员服务器。

    Windows：
    ```bash
    set HTTPS=true&&npm start
    ```

    先决条件
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > 此时，浏览器窗口打开，其中包含加载项。请关闭此窗口。

2. 在 Excel 中，依次选择“主页”**** 选项卡和功能区中的“显示任务窗格”**** 按钮，以打开加载项任务窗格。

    ![Excel 加载项按钮](../images/excel-quickstart-addin-2b.png)

3. 选择工作表中的任何一系列单元格。

4. 在任务窗格中，选择“设置颜色”**** 按钮，将选定区域的颜色设置为绿色。

    ![Excel 加载项](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>后续步骤

恭喜！已使用 React 成功创建 Excel 加载项！接下来，请详细了解 Excel 加载项功能，并跟着 Excel 加载项教程一起操作，生成更复杂的加载项。

> [!div class="nextstepaction"]
> [Excel 加载项教程](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>另请参阅

* [Excel 加载项教程](../tutorials/excel-tutorial-create-table.md)
* [Excel JavaScript API 核心概念](../excel/excel-add-ins-core-concepts.md)
* [Excel 加载项代码示例](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Excel JavaScript API 参考](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
