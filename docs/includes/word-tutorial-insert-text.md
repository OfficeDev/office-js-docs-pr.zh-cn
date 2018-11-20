本教程的这一步是，先以编程方式测试加载项是否支持用户的当前版本 Word，再在文档中插入段落。

> [!NOTE]
> 此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="code-the-add-in"></a>编码加载项

1. 在代码编辑器中打开项目。
2. 打开文件 index.html。
3. 将 `TODO1` 替换为以下标记：

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. 打开 app.js 文件。
5. 将 `TODO1` 替换为下面的代码。 此代码用于确定用户的 Word 版本是否支持包含本教程所有阶段使用的全部 API 的 Word.js 版本。 在生产加载项中，若要隐藏或禁用调用不受支持的 API 的 UI，请使用条件块的主体。 这样一来，用户仍可以使用 Word 版本支持的加载项部分。

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }
    ```

6. 将 `TODO2` 替换为下面的代码：

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. 将 `TODO3` 替换为以下代码。 请注意以下几点：
   - Word.js 业务逻辑会添加到传递给 `Word.run` 的函数中。 此逻辑不会立即执行， 而是添加到挂起命令队列中。
   - `context.sync` 方法将所有已排入队列的命令都发送到 Word 以供执行。
   - `Word.run` 后跟 `catch` 块。 这是应始终遵循的最佳做法。 

    ```js
    function insertParagraph() {
        Word.run(function (context) {

            // TODO4: Queue commands to insert a paragraph into the document.

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

8. 将 `TODO4` 替换为下面的代码。请注意以下几点：
   - `insertParagraph` 方法的第一个参数是新段落的文本。
   - 第二个参数是应在正文中的什么位置插入段落。 如果父对象为正文，其他段落插入选项包括“End”和“Replace”。

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");
    ```

## <a name="test-the-add-in"></a>测试加载项

1. 打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。
2. 运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。
3. 运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。
4. 通过以下方法之一旁加载加载项：
    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. 在 Word 的“开始”**** 菜单中，选择“显示任务窗格”****。
6. 在任务窗格中，选择“插入段落”****。
7. 在段落中进行一些更改。
8. 再次选择“插入段落”****。 观察新段落是否位于上一段落之上，因为 `insertParagraph` 方法要在文档正文的“开头”插入内容。

    ![Word 教程 - 插入段落](../images/word-tutorial-insert-paragraph.png)
