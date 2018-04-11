本教程的这一步是，了解如何在文档中创建格式文本内容控件，以及如何插入和替换控件的内容。 

> [!NOTE]
> 此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。

开始执行本教程的这一步之前，建议通过 Word UI 创建和控制格式文本内容控件，以便熟悉此类控件及其属性。 有关详细信息，请参阅[在 Word 中创建用户填写或打印的表单](https://support.office.com/zh-cn/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b)。

> [!NOTE]
> 虽然可通过 UI 添加到 Word 文档的内容控件有好几种，但目前 Word.js 仅支持格式文本内容控件。


## <a name="create-a-content-control"></a>创建内容控件

1. 在代码编辑器中打开项目。 
2. 打开文件 index.html。
3. 在包含 `replace-text` 按钮的 `div` 下方，添加下列标记：

    ```html
    <div class="padding">            
        <button class="ms-Button" id="create-content-control">Create Content Control</button>            
    </div>
    ```

4. 打开 app.js 文件。

5. 在向 `insert-table` 按钮分配单击处理程序的代码行下方，添加下列代码：

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. 在 `insertTable` 函数下方，添加下列函数：

    ```js
    function createContentControl() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to create a content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. 将 `TODO1` 替换为下面的代码。请注意以下几点：
   - 此代码用于在内容控件中包装短语“Office 365”。 它做了一个简化假设，即存在字符串，且用户已选择它。
   - `ContentControl.title` 属性指定内容控件的可见标题。 
   - `ContentControl.tag` 属性指定标记，可用于通过 `ContentControlCollection.getByTag` 方法获取对内容控件的引用，将用于稍后出现的函数。 
   - `ContentControl.appearance` 属性指定控件的外观。 使用值“Tags”表示，控件包装在开始标记和结束标记中，且开始标记包含内容控件标题。 其他可取值包括“BoundingBox”和“None”。
   - `ContentControl.color` 属性指定标记颜色或边界框的边框。

    ```js
    const serviceNameRange = context.document.getSelection();
    const serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ``` 

## <a name="replace-the-content-of-the-content-control"></a>替换内容控件的内容

1. 打开文件 index.html。
3. 在包含 `create-content-control` 按钮的 `div` 下方，添加下列标记：
    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>            
    </div>
    ```

4. 打开 app.js 文件。

5. 在向 `create-content-control` 按钮分配单击处理程序的代码行下方，添加下列代码：

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

6. 在 `createContentControl` 函数下方，添加下列函数：

    ```js    function replaceContentInControl() {      Word.run(function (context) {
            
            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. Replace `TODO1` with the following code. 
    > [!NOTE]
    > The `ContentControlCollection.getByTag` method returns a `ContentControlCollection` of all content controls of the specified tag. We use `getFirst` to get a reference to the desired control.

    ```js
    const serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ``` 

## <a name="test-the-add-in"></a>测试加载项

1. 如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl+C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”****文件夹。
     > [!NOTE]
     > 虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样才能看到提示并输入生成命令。 生成后，重启服务器。 接下来的几步操作就是在执行此过程。
2. 运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。
3. 运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。
4. 通过关闭任务窗格来重新加载它，再选择“开始”****菜单上的“显示任务窗格”****，以重新打开加载项。
5. 在任务窗格中，选择“插入段落”****，以确保文档顶部有包含“Office 365”的段落。
6. 选择刚刚添加的段落中的短语“Office 365”，再选择“创建内容控件”****按钮。 观察此短语是否包装在标签为“服务名称”的标记中。
7. 选择“重命名服务”****按钮，并观察内容控件的文本是否变成“Fabrikam Online Productivity Suite”。

    ![Word 教程 - 创建内容控件并更改其文本](../images/word-tutorial-content-control.png)
