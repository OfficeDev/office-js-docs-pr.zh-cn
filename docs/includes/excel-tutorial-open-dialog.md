本教程的最后一步是，在加载项中打开对话框，将消息从对话框进程传递到任务窗格进程，再关闭对话框。 Office 加载项对话框是*非模式*窗口。也就是说，用户可以继续与主机 Office 应用中的文档，以及与任务窗格中的主机页进行交互。

> [!NOTE]
> 此为 Excel 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="create-the-dialog-page"></a>创建对话框页面

1. 在代码编辑器中打开项目。
2. 在项目的根目录（其中包含 index.html）中，创建 popup.html 文件。
3. 将下面的标记添加到 popup.html 中。请注意以下几点：
   - 此页面包含可供用户输入用户名的 `<input>`，并包含将用户名发送到任务窗格中用户名显示页面的按钮。
   - 此标记加载在后续步骤中创建的 popup.js 脚本。
   - 此标记还加载 Office.JS 库和 jQuery，因为 popup.js 将使用它们。

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
        
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css">
    
            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>
    
        </head>
         <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
         <div class="padding">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
         </div>        
        <div class="padding">
            <input id="name-box" type="text"/>
        <div>
        <div class="padding">
            <button id="ok-button" class="ms-Button">OK</button>
        </div>
    </body>
    </html>
    ```

4. 在项目的根目录中，创建 popup.js 文件。
5. 将下面的代码添加到 popup.js 中。请注意以下几点：
   - *所有调用 Office.JS 库中 API 的页面都必须向 `Office.initialize` 属性分配函数。* 如果不需要初始化，函数可以主体是空的，但此属性既不得未定义，也不得分配到空值或非函数值。 有关示例，请参阅项目根目录中的 app.js 文件。 分配代码必须先于任何 Office.JS 调用运行，因此分配代码位于页面加载的脚本文件中，正如本例所示。
   - jQuery `ready` 函数在 `initialize` 方法内调用。应在 `Office.initialize` 函数内加载、初始化或启动其他 JavaScript 库的代码，这几乎就是一条普遍性规则。

    ```js
    (function () {
    "use strict";

        Office.initialize = function() {        
            $(document).ready(function () {  
    
                // TODO1: Assign handler to the OK button.
    
            });
        }

        // TODO2: Create the OK button handler
    
    }());    
    ```

6. 将 `TODO1` 替换为下列代码。 将在下一步中创建 `sendStringToParentPage` 函数。

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. 将 `TODO2` 替换为以下代码。 方法将它的参数传递到父页面（在此示例中，为任务窗格中的页面）。`messageParent` 参数可以是布尔值或字符串，其中包含可串行化为字符串的任何内容（如 XML 或 JSON）。 

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. 保存文件。

   > [!NOTE]
   > popup.html 文件及其加载的 popup.js 文件在完全独立于加载项任务窗格的 Internet Explorer 进程中运行。 如果将 popup.js 转换为与 app.js 文件相同的 bundle.js 文件，加载项必须加载 bundle.js 文件的两个副本，这就违背了绑定目的。 此外，popup.js 文件不包含任何不受 IE 支持的 JavaScript。 出于这两点原因，此加载项根本不会转换 popup.js。 


## <a name="open-the-dialog-from-the-task-pane"></a>从任务窗格打开对话框

1. 打开文件 index.html。
2. 在包含 `freeze-header` 按钮的 `div` 下方，添加下列标记：

    ```html
    <div class="padding">            
        <button class="ms-Button" id="open-dialog">Open Dialog</button>          
    </div>
    ```

3. 对话框会提示用户输入用户名，并将用户名传递到任务窗格。 任务窗格将在标签中显示用户名。 在刚刚添加的 `div` 正下方，添加下列标记：

    ```html
    <div class="padding">            
        <label id="user-name"></label>            
    </div>
    ```

4. 打开 app.js 文件。

5. 在向 `freeze-header` 按钮分配单击处理程序的代码行下方，添加下列代码。 方法是在后续步骤中创建。`openDialog`

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. 在 `freezeHeader` 函数下方，添加下列声明。此变量用于保留父页面执行上下文中的对象，以用作对话框页面执行上下文的中间对象。

    ```js
    let dialog = null;
    ```

7. 在 `dialog` 声明下方，添加下列函数。 关于此代码，请务必注意它*不*包含的内容，即不含 `Excel.run` 调用。 这是因为对话框打开 API 跨所有 Office 主机共享，所以它属于 Office JavaScript 公用 API，而不属于 Excel 专用 API。

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ``` 

8. 将 `TODO1` 替换为以下代码。请注意以下几点：
   - 方法在屏幕中央打开对话框。`displayDialogAsync`
   - 第一个参数是要打开的页面 URL。
   - 第二个参数用于传递选项。`height` 和 `width` 是 Office 应用程序窗口大小百分比。 
   
    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},
        
        // TODO2: Add callback parameter.
    );
    ``` 

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a>处理对话框发送的消息并关闭对话框

1. 继续使用 app.js 文件，将 `TODO2` 替换为下列代码。请注意以下几点：
   - 回调在对话框成功打开后，且当用户在对话框中执行任何操作前立即执行。
   - 对象用作父页面执行上下文和对话框页面执行上下文的中间对象。`result.value`
   - 函数将在后续步骤中创建。`processMessage` 此处理程序将处理通过 `messageParent` 函数调用从对话框页面发送的任何值。

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. 在 `openDialog` 函数下方，添加下列函数。

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## <a name="test-the-add-in"></a>测试加载项

1. 如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。

     > [!NOTE]
     > 虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样就可以通提示符输入生成命令。 生成后，重启服务器。 接下来的几步执行的就是此进程。

1. 运行命令 `npm run build`，将 ES6 源代码转换为 Internet Explorer 支持的旧版 JavaScript（Excel 在后台用来运行 Excel 加载项）。
2. 运行命令 `npm start`，启动在 localhost 上运行的 Web 服务器。
4. 通过关闭任务窗格来重新加载它，再选择“主页”**** 菜单上的“显示任务窗格”****，重新打开加载项。
6. 选择任务窗格中的“打开对话框”**** 按钮。 
7. 对话框打开后，拖动它并重设大小。 请注意，既可以与工作表进行交互，也可以按任务窗格上的其他按钮。 不过，无法从相同的任务窗格页面启动第二个对话框。
8. 在对话框中，输入用户名，再选择“确定”****。 此时，用户名显示在任务窗格上，且对话框关闭。
9. （可选）注释掉 `processMessage` 函数中的代码行 `dialog.close();`。 然后，重复执行此部分的步骤。 这样一来，对话框便会继续处于打开状态，可供用户更改用户名。 按右上角的“X”**** 按钮，可手动关闭对话框。

    ![Excel 教程 - 对话框](../images/excel-tutorial-dialog-open.png)

