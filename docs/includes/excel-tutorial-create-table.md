本教程的这一步是，以编程方式测试加载项是否支持用户的当前版本 Excel，向工作表中添加表格，使用数据填充表格，并设置格式。

> [!NOTE]
> 此为 Excel 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="code-the-add-in"></a>编码加载项

1. 在代码编辑器中打开项目。 
2. 打开文件 index.html。
3. 将 `TODO1` 替换为以下标记：

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. 打开 app.js 文件。
5. 将 `TODO1` 替换为以下代码。 此代码用于确定用户的 Excel 版本是否支持包含本系列教程将使用的所有 API 的 Excel.js 版本。 在生产加载项中，若要隐藏或禁用调用不受支持的 API 的 UI，请使用条件块的主体。 这样一来，用户仍可以使用 Excel 版本支持的加载项部分。

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    } 
    ```

6. 将 `TODO2` 替换为以下代码：

    ```js
    $('#create-table').click(createTable);
    ```

7. 将 `TODO3` 替换为以下代码。 请注意以下几点：
   - Excel.js 业务逻辑将添加到传递给 `Excel.run` 的函数。 此逻辑不立即执行。 相反，它会被添加到挂起的命令队列中。
   - `context.sync` 方法将所有已排入队列的命令发送到 Excel 以供执行。
   - `Excel.run` 后跟 `catch` 块。 这是应始终遵循的最佳做法。 

    ```js
    function createTable() {
        Excel.run(function (context) {
            
            // TODO4: Queue table creation logic here.

            // TODO5: Queue commands to populate the table with data.

            // TODO6: Queue commands to format the table.

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

8. 将 `TODO4` 替换为以下代码。请注意以下几点：
   - 此代码通过使用工作表的表格集合的 `add` 方法来创建表格，即使是空的，也始终存在。 这是创建 Excel.js 对象的标准方式。 没有类构造函数 API，切勿使用 `new` 运算符创建 Excel 对象。 相反，请添加到父集合对象。 
   - `add` 方法的第一个参数仅是表格最上面一行的范围，而不是表格最终使用的整个范围。 这是因为当加载项填充数据行时（在下一步中），它将新行添加到表中，而不是将值写入现有行的单元格。 这是更为常见的模式，因为在创建表时表的行数通常是未知的。 
   - 表名称必须在整个工作簿中都是唯一的，而不仅仅是在工作表一级。

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ``` 

9. 将 `TODO5` 替换为以下代码。请注意以下几点：
   - 范围的单元格值是通过一组数组进行设置。
   - 表格中的新行是通过调用表格的行集合的 `add` 方法进行创建。 通过在作为第二个参数传递的父数组中添加多个单元格值数组，可以在一次 `add` 调用中添加多个行。

    ```js
    expensesTable.getHeaderRowRange().values = 
        [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
        ["1/1/2017", "The Phone Company", "Communications", "120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
        ["1/11/2017", "Bellows College", "Education", "350.1"],
        ["1/15/2017", "Trey Research", "Other", "135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
    ]);
    ``` 

10. 将 `TODO6` 替换为以下代码。请注意以下几点：
   - 此代码将从零开始编制的索引传递给表格的列集合的 `getItemAt` 方法，以获取对“金额”****列的引用。 

     > [!NOTE]
     > Excel.js 集合对象（如 `TableCollection`、`WorksheetCollection` 和 `TableColumnCollection`）有 `items` 属性，此属性是子对象类型的数组（如 `Table`、`Worksheet` 或 `TableColumn`），但 `*Collection` 对象本身并不是数组。

   - 然后，此代码将“金额”****列的范围格式化为欧元（精确到小数点后两位）。 
   - 最后，它确保了列宽和行高足以容纳最长（或最高）的数据项。 请注意，此代码必须获取要格式化的 `Range` 对象。 `TableColumn` 和 `TableRow` 对象没有格式属性。

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ``` 

## <a name="test-the-add-in"></a>测试加载项

1. 打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”****文件夹。
2. 运行命令 `npm run build`，将 ES6 源代码转换为 Internet Explorer 支持的旧版 JavaScript（Excel 在后台用来运行 Excel 加载项）。
3. 运行命令 `npm start`，启动在 localhost 上运行的 Web 服务器。   
4. 通过以下方法之一旁加载加载项：
    - Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. 在“主页”****菜单上，选择“显示任务窗格”****。
6. 在任务窗格中，选择“创建表格”****。

    ![Excel 教程 - 创建表格](../images/excel-tutorial-create-table.png)
