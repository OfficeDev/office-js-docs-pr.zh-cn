---
title: 教程：Microsoft Excel自定义函数和任务窗格之间共享数据和事件
description: 学习如何在Microsoft Excel中的自定义函数和任务窗格之间共享数据和事件。
ms.date: 11/29/2021
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 69dbb7c2b57d09f3d71397db0b1d56babf7c64a6
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242052"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>教程：Microsoft Excel自定义函数和任务窗格之间共享数据和事件

共享全局数据，并通过共享运行时在 Excel 加载项的任务窗格和自定义函数之间发送事件。 对于大多数自定义函数方案，建议使用共享运行时，除非有特定的理由需要使用非任务窗格 (UI-less) 自定义函数。 本教程假设你已经熟悉使用Yo Office生成器来创建插件项目。 如果尚未完成[Excel 自定义函数教程](excel-tutorial-create-custom-functions.md)，请考虑完成它。

## <a name="create-the-add-in-project"></a>创建加载项项目

使用 [Office 加载项的 Yeoman 生成器](https://github.com/OfficeDev/generator-office) 来创建 Excel 加载项项目。

- 要生成带自定义函数的 Excel 加载项，请运行以下命令。
    
    ```command&nbsp;line
    yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true
    ```

生成器创建项目并安装支持节点组件。

## <a name="configure-the-manifest"></a>配置清单

请按照以下步骤将加载项项目配置为使用共享运行时。

1. 启动 Visual Studio Code 并打开生成的加载项项目。
1. 打开 **manifest.xml** 文件。
1. 替换（或添加）以下 `<Requirements>` 部分 XML，以要求 [共享运行时要求集](../reference/requirement-sets/shared-runtime-requirement-sets.md)。

    ```xml
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    ```

    更新后，清单 XML 应按以下顺序显示。

    ```xml
    <Hosts>
      <Host Name="..."/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. 查找 `<VersionOverrides>` 部分并添加以下 `<Runtimes>` 部分。 生存期需要 **较长**，以便在关闭任务窗格时加载项代码仍可运行。 `resid` 值是 **Taskpane.Url**，它引用 **manifest.xml** 文件底部附近的 `<bt:Urls>` 部分中指定的 **taskpane.html** 文件位置。
    
    ```xml
    <Runtimes>
      <Runtime resid="Taskpane.Url" lifetime="long" />
    </Runtimes>
    ```
    
    > [!IMPORTANT]
    > 必须按照以下 XML 中显示的确切顺序在 `<Host xsi:type="...">` 元素之后输入 `<Runtimes>` 部分。

    ```xml
    <VersionOverrides ...>
      <Hosts>
        <Host xsi:type="...">
          <Runtimes>
            <Runtime resid="Taskpane.Url" lifetime="long" />
          </Runtimes>
        ...
        </Host>
    ```
    
    > [!NOTE]
    > 如果加载项包含清单中的 `Runtimes` 元素（共享运行时所需），并且满足将 Microsoft Edge 与 WebView2（基于 Chromium）一起使用的条件，则它使用该 WebView2 控件。 如果不满足条件，则使用 Internet Explorer 11，而不考虑 Windows 或 Microsoft 365 版本。 有关详细信息，请参阅 [运行时](../reference/manifest/runtimes.md) 和 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。

1. 查找 `<Page>` 元素。然后将源位置从 **Functions.Page.Url** 更改为 **Taskpane.Url**。

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. 查找 `<FunctionFile ...>` 标记并将 `resid` 从 **Commands.Url** 更改为  **Taskpane.Url**。

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. 保存 **manifest.xml** 文件。

## <a name="configure-the-webpackconfigjs-file"></a>配置 webpack.config.js 文件

**webpack.config.js** 将生成多个运行时加载程序。 你需要对其进行修改，以通过 **taskpane.html** 文件仅加载共享 JavaScript 运行时。 

1. 打开 **webpack.config.js** 文件。
1. 转到 `plugins:` 部分。
1. 删除以下 `functions.html` 插件（如果存在）。
    
    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. 删除以下 `commands.html` 插件（如果存在）。

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. 如果删除了 `functions` 或 `commands` 插件，请将其添加为 `chunks`。 如果同时删除了 `functions` 和 `commands` 插件，则以下 JavaScript 将显示更新的条目。
    
    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```
    
1. 保存更改并重新生成项目。

   ```command&nbsp;line
   npm run build
   ```
    
    > [!NOTE]
    > 还可以删除 **functions.html** 和 **commands.html** 文件。 **taskpane.html** 将通过你刚才进行的 webpack 更新将 **functions.js** 和 **commands.js** 代码加载到共享 JavaScript 运行时中。
    
1. 保存更改并运行项目。 确保加载和运行时没有错误。
    
   ```command&nbsp;line
   npm run start
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>共享自定义函数和任务窗格代码之间的状态

由于自定义函数在与任务窗格代码相同的上下文中运行，因此可以直接共享状态，无需使用 **Storage** 对象。 以下说明介绍了如何在自定义函数和任务窗格代码之间共享全局变量。

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>创建用于获取或存储共享状态的自定义函数

1. 在 Visual Studio Code 中，打开文件 **src/functions/functions.js**。
2. 在第 1 行，将以下代码插入到最顶部。 这将初始化一个名为 **sharedState** 的全局变量。

   ```js
   window.sharedState = "empty";
   ```

3. 添加以下代码，创建将值存储到 **sharedState** 变量的自定义函数。

   ```js
   /**
    * Saves a string value to shared state with the task pane
    * @customfunction STOREVALUE
    * @param {string} value String to write to shared state with task pane.
    * @return {string} A success value
    */
   function storeValue(sharedValue) {
     window.sharedState = sharedValue;
     return "value stored";
   }
   ```

4. 添加以下代码，创建获取 **sharedState** 变量的当前值的自定义函数。

   ```js
   /**
    * Gets a string value from shared state with the task pane
    * @customfunction GETVALUE
    * @returns {string} String value of the shared state with task pane.
    */
   function getValue() {
     return window.sharedState;
   }
   ```

5. 保存文件。

### <a name="create-task-pane-controls-to-work-with-global-data"></a>创建任务窗格控件以处理全局数据

1. 打开 **src/taskpane/taskpane.html** 文件。
2. 紧贴在结尾的 `</head>` 元素前，添加以下脚本元素。

   ```html
   <script src="../functions/functions.js"></script>
   ```

3. 关闭 `</main>` 元素后，添加以下 HTML。 该 HTML 创建两个用于获取或存储全局数据的文本框和按钮。

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong> into a cell to retrieve it.
     </li>
     <li>
       To send data to the task pane, in a cell, enter
       <strong>=CONTOSO.STOREVALUE("new value")</strong>
     </li>
     <li>Select <strong>Get</strong> to display the value in the task pane.</li>
   </ol>

   <p>Store new value to shared state</p>
   <div>
     <input type="text" id="storeBox" />
     <button onclick="storeSharedValue()">Store</button>
   </div>

   <p>Get shared state value</p>
   <div>
     <input type="text" id="getBox" />
     <button onclick="getSharedValue()">Get</button>
   </div>
   ```

4. 在结束 `</body>` 元素之前，添加以下脚本。当用户要存储或获取全局数据时，此代码将处理按钮单击事件。

   ```js
   <script>
   function storeSharedValue() {
     let sharedValue = document.getElementById('storeBox').value;
     window.sharedState = sharedValue;
   }

   function getSharedValue() {
     document.getElementById('getBox').value = window.sharedState;
   }
   </script>
   ```

5. 保存文件。
6. 生成项目

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>在自定义函数和任务窗格之间尝试共享数据

- 使用以下命令启动项目。

  ```command line
  npm run start
  ```

Excel 启动后，可使用“任务窗格”按钮来存储或获取共享数据。 在自定义函数的单元格中输入 `=CONTOSO.GETVALUE()`，以检索相同的共享数据。 或使用 `=CONTOSO.STOREVALUE("new value")` 将共享数据更改为新值。

> [!NOTE]
> 如本文所示配置项目，可在自定义函数和任务窗格之间共享上下文。 通过自定义函数可以调用一些Office API。 更多细节请参见[从自定义函数调用Microsoft Excel APIs](../excel/call-excel-apis-from-custom-function.md)。

## <a name="see-also"></a>另请参阅

- [将 Office 加载项配置为使用共享 JavaScript 运行时](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
