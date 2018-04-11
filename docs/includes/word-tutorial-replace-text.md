本教程的这一步是，在选定文本区域内外添加文本，并替换选定区域的文本。 

> [!NOTE]
> 此页面介绍了 Word 加载项教程的步骤之一。如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Word 加载项教程](../tutorials/word-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="add-text-inside-a-range"></a>在区域内添加文本

1. 在代码编辑器中打开项目。 
2. 打开文件 index.html。
3. 在包含 `change-font` 按钮的 `div` 下方，添加下列标记：

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>            
    </div>
    ```

4. 打开 app.js 文件。

5. 在向 `change-font` 按钮分配单击处理程序的代码行下方，添加下列代码：

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. 在 `changeFont` 函数下方，添加下列函数：

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the 
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

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
   - 此方法用于在“即点即用”文本区域末尾插入缩写 ["(C2R)"]。 它做了一个简化假设，即存在字符串，且用户已选择它。
   - `Range.insertText` 方法的第一个参数是要插入到 `Range` 对象的字符串。
   - 第二个参数指定了应在区域中的什么位置插入其他文本。 除了“End”外，其他可用选项包括“Start”、“Before”、“After”和“Replace”。 
   - “End”和“After”的区别在于，“End”在现有区域末尾插入新文本，而“After”则是新建包含字符串的区域，并在现有区域后面插入新区域。 同样，“Start”是在现有区域的开头位置插入文本，而“Before”插入的是新区域。 “Replace”将现有区域文本替换为第一个参数中的字符串。
   - 在本教程之前阶段步骤中，正文对象的 insert* 方法没有“Before”和“After”选项。 这是因为不能将内容置于文档正文外。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ``` 

8. 在下一部分前，将跳过 `TODO2`。 将 `TODO3` 替换为下面的代码。 此代码类似于在本教程第一阶段中创建的代码，区别在于现在是要在文档末尾（而不是开头）插入新段落。 这一新段落将说明，新文本现属于原始区域。
 
    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text,
                             "End");
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>添加代码以将文档属性提取到任务窗格的脚本对象

在本系列教程前面的所有函数中，都是将命令排入队列，以对 Office 文档执行*写入*操作。 每个函数结束时都会调用 `context.sync()` 方法，从而将排入队列的命令发送到文档，以供执行。 不过，在上一步中添加的代码调用的是 `originalRange.text` 属性，这与之前编写的函数明显不同，因为 `originalRange` 对象只是任务窗格脚本中的代理对象。 由于它并不了解文档中区域的实际文本，因此它的 `text` 属性无法有实值。 有必要先从文档中提取区域的文本值，再用它设置 `originalRange.text` 的值。 只有这样才能调用 `originalRange.text`，而又不会导致异常抛出。 此提取过程分为三步：

   1. 将命令排入队列，以加载（即提取）代码需要读取的属性。
   2. 调用上下文对象的 `sync`方法，从而向文档发送已排入队列的命令以供执行，并返回请求获取的信息。
   3. 由于 `sync` 是异步方法，因此请先确保它已完成，然后代码才能调用已提取的属性。

只要代码需要从 Office 文档*读取*信息，就必须完成这些步骤。

1. 将 `TODO2` 替换为下面的代码。
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.
    
            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

2. 由于不能在同一取消分支代码路径中有两个 `return` 语句，因此请删除 `Word.run` 末尾的最后一行代码 `return context.sync();`。本教程稍后将添加最后一个新 `context.sync` 语句。 
3. 剪切并粘贴 `doc.body.insertParagraph` 代码行，以替代 `TODO4`。 
4. 将 `TODO5` 替换为下面的代码。请注意以下几点：
   - 将 `sync` 方法传递到 `then` 函数可确保它不会在 `insertParagraph` 逻辑已排入队列前运行。
   - 由于 `then` 方法调用传递给它的任何函数，并且也不想调用 `sync` 两次，因此请从 context.sync 末尾省略掉“()”。

    ```js
    .then(context.sync);
    ```

完成后，整个函数应如下所示：

  
```js
function insertTextIntoRange() {
    Word.run(function (context) {
        
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {        
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");            
                }
            )
            .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
``` 

## <a name="add-text-between-ranges"></a>在区域间添加文本

1. 打开文件 index.html。
2. 在包含 `insert-text-into-range` 按钮的 `div` 下方，添加下列标记：

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>            
    </div>
    ```

3. 打开 app.js 文件。

4. 在向 `insert-text-into-range` 按钮分配单击处理程序的代码行下方，添加下列代码：

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. 在 `insertTextIntoRange` 函数下方，添加下列函数：

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert a new range before the 
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the 
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. 将 `TODO1` 替换为下面的代码。请注意以下几点：
   - 此方法用于在文本为“Office 365”的区域前添加文本为“Office 2019”的区域。 它做了一个简化假设，即存在字符串，且用户已选择它。
   - `Range.insertText` 方法的第一个参数是要添加的字符串。
   - 第二个参数指定了应在区域中的什么位置插入其他文本。 若要详细了解位置选项，请参阅前面介绍的 `insertTextIntoRange` 函数。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ``` 

7. 将 `TODO2` 替换为下面的代码。 
 
     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.
    
                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

8. 将 `TODO3` 替换为下面的代码。 这一新段落将说明，新文本***不***属于原始选定区域。 原始区域中的文本仍与用户选择它时一样。
 
    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ``` 

9. 将 `TODO4` 替换为下面的代码：

    ```js
    .then(context.sync);
    ```


## <a name="replace-the-text-of-a-range"></a>替换区域文本

1. 打开文件 index.html。
2. 在包含 `insert-text-outside-range` 按钮的 `div` 下方，添加下列标记：

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>            
    </div>
    ```

3. 打开 app.js 文件。

4. 在向 `insert-text-outside-range` 按钮分配单击处理程序的代码行下方，添加下列代码：

    ```js
    $('#replace-text').click(replaceText);
    ```

5. 在 `insertTextBeforeRange` 函数下方，添加下列函数：

    ```js
    function replaceText() {
        Word.run(function (context) {
             
            // TODO1: Queue commands to replace the text.

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

6. 将 `TODO1` 替换为下面的代码。 请注意，此方法用于将字符串“几个”替换为字符串“许多”。 它做了一个简化假设，即存在字符串，且用户已选择它。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace"); 
    ``` 

## <a name="test-the-add-in"></a>测试加载项

1. 如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”****文件夹。

     > [!NOTE]
     > 虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样才能看到提示并输入生成命令。 生成后，重启服务器。 接下来的几步操作就是在执行此过程。

2. 运行命令 `npm run build`，以将 ES6 源代码转换为所有可运行 Office 加载项的主机支持的旧版 JavaScript。
3. 运行命令 `npm start`，以启动在 localhost 上运行的 Web 服务器。
4. 通过关闭任务窗格来重新加载它，再选择“开始”****菜单上的“显示任务窗格”****，以重新打开加载项。
5. 在任务窗格中，选择“插入段落”****，以确保文档开头有一个段落。
6. 选择某文本。 选择短语“即点即用”最合适。 *请注意，不要在选定区域的前后添加空格。*
7. 选择“插入缩写”****按钮。 观察“(C2R)”是否已添加。 此外，还请观察，文档底部是否添加了包含整个扩展文本的新段落，因为新字符串已添加到现有区域中。
8. 选择某文本。 选择短语“Office 365”最合适。 *请注意，不要在选定区域的前后添加空格。*
9. 选择“添加版本信息”****按钮。 观察是否已在“Office 2016”和“Office 365”之间插入“Office 2019”。 此外，还请观察，文档底部是否添加了仅包含最初选定文本的新段落，因为新字符串已变成新区域，而不是添加到原始区域中。
10. 选择某文本。 选择字词“几个”最合适。 *请注意，不要在选定区域的前后添加空格。*
11. 选择“更改数量术语”****按钮。 观察选定文本是否替换为“多个”。

    ![Word 教程 - 添加和替换文本](../images/word-tutorial-text-replace.png)
