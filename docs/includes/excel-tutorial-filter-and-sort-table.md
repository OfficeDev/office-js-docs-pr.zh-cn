本教程的这一步是，筛选并排序之前创建的表。

> [!NOTE]
> 此为 Excel 加载项分步教程页面。 如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。

## <a name="filter-the-table"></a>筛选表

1. 在代码编辑器中打开项目。 
2. 打开文件 index.html。
3. 在包含 `create-table` 按钮的 `div` 正下方，添加下列标记：

    ```html
    <div class="padding">            
        <button class="ms-Button" id="filter-table">Filter Table</button>            
    </div>
    ```

4. 打开 app.js 文件。

5. 在向 `create-table` 按钮分配单击处理程序的代码行正下方，添加下列代码：

    ```js
    $('#filter-table').click(filterTable);
    ```

6. 在 `createTable` 函数正下方，添加下列函数：

    ```js
    function filterTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to filter out all expense categories except 
            //        Groceries and Education.

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
   - 代码先将列名称传递给 `getItem` 方法（而不是像 `createTable` 方法一样将列索引传递给 `getItemAt` 方法），获取对需要筛选的列的引用。 由于用户可以移动表格列，因此给定索引处的列可能会在表格创建后更改。 所以，更安全的做法是，使用列名称获取对列的引用。 上一教程安全地使用了 `getItemAt`，因为是在与创建表格完全相同的方法中使用了它，所以用户没有机会移动列。
   - `applyValuesFilter` 方法是对 `Filter` 对象执行的多种筛选方法之一。

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(["Education", "Groceries"]);
    ``` 

## <a name="sort-the-table"></a>排序表格

1. 打开文件 index.html。
2. 在包含 `filter-table` 按钮的 `div` 下方，添加下列标记：

    ```html
    <div class="padding">            
        <button class="ms-Button" id="sort-table">Sort Table</button>            
    </div>
    ```

3. 打开 app.js 文件。

4. 在向 `filter-table` 按钮分配单击处理程序的代码行下方，添加下列代码：

    ```js
    $('#sort-table').click(sortTable);
    ```

5. 在 `filterTable` 函数下方，添加下列函数。

    ```js
    function sortTable() {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to sort the table by Merchant name.

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
   - 此代码创建一组 `SortField` 对象，其中只有一个成员，因为加载项只对“商家”列进行了排序。
   - `SortField` 对象的 `key` 属性是要排序的列的从零开始编制索引。
   - `Table` 的 `sort` 成员是 `TableSort` 对象，并不是方法。 `SortField` 传递到 `TableSort` 对象的 `apply` 方法。

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const sortFields = [
        { 
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ``` 

## <a name="test-the-add-in"></a>测试加载项

1. 如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。 否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”文件夹。

     > [!NOTE]
     > 虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。 为此，需要终止服务器进程，这样就可以通提示符输入生成命令。 生成后，重启服务器。 接下来的几步执行的就是此进程。

1. 运行命令 `npm run build`，将 ES6 源代码转换为 Internet Explorer 支持的旧版 JavaScript（Excel 在后台用来运行 Excel 加载项）。
2. 运行命令 `npm start`，启动在 localhost 上运行的 Web 服务器。
4. 通过关闭任务窗格来重新加载它，再选择“主页”菜单上的“显示任务窗格”，重新打开加载项。
5. 如果出于某种原因在工作表中打不开表格，请在任务窗格中选择“创建表格” 
6. 选择“筛选表格”和“排序表格”（按顺序和倒序中的任一顺序排序皆可）。

    ![Excel 教程 - 筛选和排序表格](../images/excel-tutorial-filter-and-sort-table.png)
