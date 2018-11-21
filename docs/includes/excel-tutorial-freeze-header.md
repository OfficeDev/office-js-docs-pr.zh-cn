如果表格很长，导致用户必须滚动才能看到一些行，那么标题行可能会在滚动时不可见。 本教程的这一步是，冻结以前创建的表格的标题行，让它在用户向下滚动工作表时依然可见。 

> [!NOTE]
> 此为 Excel 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="freeze-the-tables-header-row"></a>冻结表的标题行

1. 在代码编辑器中打开项目。
2. 打开文件 index.html。
3. 在包含 `create-chart` 按钮的 `div` 下方，添加下列标记：

    ```html
    <div class="padding">
        <button class="ms-Button" id="freeze-header">Freeze Header</button>
    </div>
    ```

4. 打开 app.js 文件。

5. 在向 `create-chart` 按钮分配单击处理程序的代码行下方，添加下列代码：

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. 在 `createChart` 函数下方，添加下列函数：

    ```js
    function freezeHeader() {
        Excel.run(function (context) {

            // TODO1: Queue commands to keep the header visible when the user scrolls.

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

7. 将 `TODO1` 替换为以下代码。请注意以下几点：
   - `Worksheet.freezePanes` 集合是工作表中的一组窗格，在工作表滚动时就地固定或冻结。
   - `freezeRows` 方法需要使用要就地固定的行数（自顶部算起）作为参数。传递 `1` 可以就地固定第一行。

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

## <a name="test-the-add-in"></a>测试加载项

1. 如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。

     > [!NOTE]
     > 虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样就可以通提示符输入生成命令。 生成后，重启服务器。 接下来的几步执行的就是此进程。

1. 运行命令 `npm run build`，将 ES6 源代码转换为 Internet Explorer 支持的旧版 JavaScript（Excel 在后台用来运行 Excel 加载项）。
2. 运行命令 `npm start`，启动在 localhost 上运行的 Web 服务器。
4. 通过关闭任务窗格来重新加载它，再选择“主页”**** 菜单上的“显示任务窗格”****，重新打开加载项。
6. 如果表格在工作表中，请删除它。
7. 在任务窗格中，选择“**创建表**”。
8. 选择“**冻结标题**”按钮。
9. 向下滚动工作表，直到在上面的行不可见时表格标题在顶部依然可见。

    ![Excel 教程 - 冻结标题](../images/excel-tutorial-freeze-header.png)
