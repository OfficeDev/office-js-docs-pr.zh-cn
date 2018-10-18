<span data-ttu-id="919fd-101">本教程的这一步是，以编程方式测试加载项是否支持用户的当前版本 Excel，向工作表中添加表格，使用数据填充表格，并设置格式。</span><span class="sxs-lookup"><span data-stu-id="919fd-101">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

> [!NOTE]
> <span data-ttu-id="919fd-102">此为 Excel 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="919fd-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="919fd-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="919fd-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="code-the-add-in"></a><span data-ttu-id="919fd-104">编码加载项</span><span class="sxs-lookup"><span data-stu-id="919fd-104">Code the add-in</span></span>

1. <span data-ttu-id="919fd-105">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="919fd-105">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="919fd-106">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="919fd-106">Open the file index.html.</span></span>
3. <span data-ttu-id="919fd-107">将 `TODO1` 替换为以下标记：</span><span class="sxs-lookup"><span data-stu-id="919fd-107">Replace the `TODO1` with the following markup:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button>
    ```

4. <span data-ttu-id="919fd-108">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="919fd-108">Open the app.js file.</span></span>
5. <span data-ttu-id="919fd-109">将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="919fd-109">Replace the `TODO1` with the following code.</span></span> <span data-ttu-id="919fd-110">此代码用于确定用户的 Excel 版本是否支持包含本系列教程将使用的所有 API 的 Excel.js 版本。</span><span class="sxs-lookup"><span data-stu-id="919fd-110">This code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="919fd-111">在生产加载项中，若要隐藏或禁用调用不受支持的 API 的 UI，请使用条件块的主体。</span><span class="sxs-lookup"><span data-stu-id="919fd-111">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="919fd-112">这样一来，用户仍可以使用 Excel 版本支持的加载项部分。</span><span class="sxs-lookup"><span data-stu-id="919fd-112">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    ```js
    if (!Office.context.requirements.isSetSupported('ExcelApi', 1.7)) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    } 
    ```

6. <span data-ttu-id="919fd-113">将 `TODO2` 替换为以下代码：</span><span class="sxs-lookup"><span data-stu-id="919fd-113">Replace the `TODO2` with the following code:</span></span>

    ```js
    $('#create-table').click(createTable);
    ```

7. <span data-ttu-id="919fd-114">将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="919fd-114">Replace the `TODO3` with the following code.</span></span> <span data-ttu-id="919fd-115">请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="919fd-115">Note the following:</span></span>
   - <span data-ttu-id="919fd-116">Excel.js 业务逻辑将添加到传递给 `Excel.run` 的函数。</span><span class="sxs-lookup"><span data-stu-id="919fd-116">Your Excel.js business logic will be added to the function that is passed to `Excel.run`.</span></span> <span data-ttu-id="919fd-117">此逻辑不立即执行。</span><span class="sxs-lookup"><span data-stu-id="919fd-117">This logic does not execute immediately.</span></span> <span data-ttu-id="919fd-118">相反，它会被添加到挂起的命令队列中。</span><span class="sxs-lookup"><span data-stu-id="919fd-118">Instead, it is added to a queue of pending commands.</span></span>
   - <span data-ttu-id="919fd-119">方法将所有已排入队列的命令发送到 Excel 以供执行。`context.sync`</span><span class="sxs-lookup"><span data-stu-id="919fd-119">The `context.sync` method sends all queued commands to Excel for execution.</span></span>
   - <span data-ttu-id="919fd-120">后跟 `catch` 块。`Excel.run`</span><span class="sxs-lookup"><span data-stu-id="919fd-120">The `Excel.run` is followed by a `catch` block.</span></span> <span data-ttu-id="919fd-121">这是应始终遵循的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="919fd-121">This is a best practice that you should always follow.</span></span> 

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

8. <span data-ttu-id="919fd-p106">将 `TODO4` 替换为下面的代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="919fd-p106">Replace `TODO4` with the following code. Note:</span></span>
   - <span data-ttu-id="919fd-124">此代码通过使用工作表的表格集合的 `add` 方法来创建表格，即使是空的，也始终存在。</span><span class="sxs-lookup"><span data-stu-id="919fd-124">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty.</span></span> <span data-ttu-id="919fd-125">这是创建 Excel.js 对象的标准方式。</span><span class="sxs-lookup"><span data-stu-id="919fd-125">This is the standard way that Excel.js objects are created.</span></span> <span data-ttu-id="919fd-126">没有类构造函数 API，切勿使用 `new` 运算符创建 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="919fd-126">There are no class constructor APIs, and you never use a `new` operator to create an Excel object.</span></span> <span data-ttu-id="919fd-127">相反，请添加到父集合对象。</span><span class="sxs-lookup"><span data-stu-id="919fd-127">Instead, you add to a parent collection object.</span></span> 
   - <span data-ttu-id="919fd-128">方法的第一个参数仅是表格最上面一行的范围，而不是表格最终使用的整个范围。`add`</span><span class="sxs-lookup"><span data-stu-id="919fd-128">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use.</span></span> <span data-ttu-id="919fd-129">这是因为当加载项填充数据行时（在下一步中），它将新行添加到表中，而不是将值写入现有行的单元格。</span><span class="sxs-lookup"><span data-stu-id="919fd-129">This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows.</span></span> <span data-ttu-id="919fd-130">这是更为常见的模式，因为在创建表时表的行数通常是未知的。</span><span class="sxs-lookup"><span data-stu-id="919fd-130">This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span> 
   - <span data-ttu-id="919fd-131">表名称必须在整个工作簿中都是唯一的，而不仅仅是在工作表一级。</span><span class="sxs-lookup"><span data-stu-id="919fd-131">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ``` 

9. <span data-ttu-id="919fd-p109">将 `TODO5` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="919fd-p109">Replace `TODO5` with the following code. Note:</span></span>
   - <span data-ttu-id="919fd-134">范围的单元格值是通过一组数组进行设置。</span><span class="sxs-lookup"><span data-stu-id="919fd-134">The cell values of a range are set with an array of arrays.</span></span>
   - <span data-ttu-id="919fd-135">表格中的新行是通过调用表格的行集合的 `add` 方法进行创建。</span><span class="sxs-lookup"><span data-stu-id="919fd-135">New rows are created in a table by calling the `add` method of the table's row collection.</span></span> <span data-ttu-id="919fd-136">通过在作为第二个参数传递的父数组中添加多个单元格值数组，可以在一次 `add` 调用中添加多个行。</span><span class="sxs-lookup"><span data-stu-id="919fd-136">You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

10. <span data-ttu-id="919fd-p111">将 `TODO6` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="919fd-p111">Replace `TODO6` with the following code. Note:</span></span>
   - <span data-ttu-id="919fd-139">此代码将从零开始编制的索引传递给表格的列集合的 `getItemAt` 方法，以获取对“金额”\*\*\*\* 列的引用。</span><span class="sxs-lookup"><span data-stu-id="919fd-139">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span> 

     > [!NOTE]
     > <span data-ttu-id="919fd-140">Excel.js 集合对象（如 `TableCollection`、`WorksheetCollection` 和 `TableColumnCollection`）有 `items` 属性，此属性是子对象类型的数组（如 `Table`、`Worksheet` 或 `TableColumn`），但 `*Collection` 对象本身并不是数组。</span><span class="sxs-lookup"><span data-stu-id="919fd-140">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

   - <span data-ttu-id="919fd-141">然后，此代码将“金额”\*\*\*\* 列的范围格式化为欧元（精确到小数点后两位）。</span><span class="sxs-lookup"><span data-stu-id="919fd-141">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 
   - <span data-ttu-id="919fd-142">最后，它确保了列宽和行高足以容纳最长（或最高）的数据项。</span><span class="sxs-lookup"><span data-stu-id="919fd-142">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item.</span></span> <span data-ttu-id="919fd-143">请注意，此代码必须获取要格式化的 `Range` 对象。</span><span class="sxs-lookup"><span data-stu-id="919fd-143">Notice that the code must get `Range` objects to format.</span></span> <span data-ttu-id="919fd-144">`TableColumn` 和 `TableRow` 对象没有格式属性。</span><span class="sxs-lookup"><span data-stu-id="919fd-144">`TableColumn` and `TableRow` objects do not have format properties.</span></span>

        ```js
        expensesTable.columns.getItemAt(3).getRange().numberFormat = [['€#,##0.00']];
        expensesTable.getRange().format.autofitColumns();
        expensesTable.getRange().format.autofitRows();
        ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="919fd-145">测试加载项</span><span class="sxs-lookup"><span data-stu-id="919fd-145">Test the add-in</span></span>

1. <span data-ttu-id="919fd-146">打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="919fd-146">Open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>
2. <span data-ttu-id="919fd-147">运行命令 `npm run build`，将 ES6 源代码转换为 Internet Explorer 支持的旧版 JavaScript（Excel 在后台用来运行 Excel 加载项）。</span><span class="sxs-lookup"><span data-stu-id="919fd-147">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
3. <span data-ttu-id="919fd-148">运行命令 `npm start`，启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="919fd-148">Run the command `npm start` to start a web server running on localhost.</span></span>   
4. <span data-ttu-id="919fd-149">通过以下方法之一旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="919fd-149">Sideload the add-in by using one of the following methods:</span></span>
    - <span data-ttu-id="919fd-150">Windows：[在 Windows 上旁加载 Office 加载项](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="919fd-150">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="919fd-151">Excel Online：[在 Office Online 中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="919fd-151">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="919fd-152">iPad 和 Mac：[在 iPad 和 Mac 上旁加载 Office 加载项](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="919fd-152">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>
5. <span data-ttu-id="919fd-153">在“主页”\*\*\*\* 菜单上，选择“显示任务窗格”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="919fd-153">On the **Home** menu, choose **Show Taskpane**.</span></span>
6. <span data-ttu-id="919fd-154">在任务窗格中，选择“创建表格”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="919fd-154">In the taskpane, choose **Create Table**.</span></span>

    ![Excel 教程 - 创建表格](../images/excel-tutorial-create-table.png)
