<span data-ttu-id="45f55-101">本教程的这一步是，使用先前创建的表中的数据创建图表，再设置图表格式。</span><span class="sxs-lookup"><span data-stu-id="45f55-101">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

> [!NOTE]
> <span data-ttu-id="45f55-102">此为 Excel 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="45f55-102">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="45f55-103">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="45f55-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="chart-table-data"></a><span data-ttu-id="45f55-104">将表数据绘制成图表</span><span class="sxs-lookup"><span data-stu-id="45f55-104">Chart table data</span></span>

1. <span data-ttu-id="45f55-105">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="45f55-105">Open the project in your code editor.</span></span>
2. <span data-ttu-id="45f55-106">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="45f55-106">Open the file index.html.</span></span>
3. <span data-ttu-id="45f55-107">在包含 `sort-table` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="45f55-107">Below the `div` that contains the `sort-table` button, add the following markup:</span></span>

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-chart">Create Chart</button>
    </div>
    ```

4. <span data-ttu-id="45f55-108">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="45f55-108">Open the app.js file.</span></span>

5. <span data-ttu-id="45f55-109">在向 `sort-chart` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="45f55-109">Below the line that assigns a click handler to the `sort-chart` button, add the following code:</span></span>

    ```js
    $('#create-chart').click(createChart);
    ```

6. <span data-ttu-id="45f55-110">在 `sortTable` 函数下方，添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="45f55-110">Below the `sortTable` function add the following function.</span></span>

    ```js
    function createChart() {
        Excel.run(function (context) {

            // TODO1: Queue commands to get the range of data to be charted.

            // TODO2: Queue command to create the chart and define its type.

            // TODO3: Queue commands to position and format the chart.

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

7. <span data-ttu-id="45f55-p102">将 `TODO1` 替换为下列代码。请注意，为了排除标题行，此代码使用 `Table.getDataBodyRange` 方法（而不是 `getRange` 方法），获取要绘制成图表的数据的范围。</span><span class="sxs-lookup"><span data-stu-id="45f55-p102">Replace `TODO1` with the following code. Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    const dataRange = expensesTable.getDataBodyRange();
    ```

8. <span data-ttu-id="45f55-113">将 `TODO2` 替换为下列代码。</span><span class="sxs-lookup"><span data-stu-id="45f55-113">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="45f55-114">请注意以下参数：</span><span class="sxs-lookup"><span data-stu-id="45f55-114">Note the following parameters:</span></span>
   - <span data-ttu-id="45f55-p104">`add` 方法的第一个参数指定图表类型。有几十种类型。</span><span class="sxs-lookup"><span data-stu-id="45f55-p104">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>
   - <span data-ttu-id="45f55-117">第二个参数指定要在图表中添加的数据的范围。</span><span class="sxs-lookup"><span data-stu-id="45f55-117">The second parameter specifies the range of data to include in the chart.</span></span>
   - <span data-ttu-id="45f55-118">第三个参数确定是按行方向还是按列方向绘制表格中的一系列数据点。</span><span class="sxs-lookup"><span data-stu-id="45f55-118">The third parameter determines whether a series of data points from the table should be charted rowwise or columnwise.</span></span> <span data-ttu-id="45f55-119">选项 `auto` 指示 Excel 确定最佳方法。</span><span class="sxs-lookup"><span data-stu-id="45f55-119">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    let chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

9. <span data-ttu-id="45f55-120">将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="45f55-120">Replace `TODO3` with the following code.</span></span> <span data-ttu-id="45f55-121">此代码的大部分内容非常直观明了。</span><span class="sxs-lookup"><span data-stu-id="45f55-121">Most of this code is self-explanatory.</span></span> <span data-ttu-id="45f55-122">请注意几下几点：</span><span class="sxs-lookup"><span data-stu-id="45f55-122">Note:</span></span>
   - <span data-ttu-id="45f55-123">`setPosition` 方法的参数指定应包含图表的工作表区域的左上角和右下角单元格。</span><span class="sxs-lookup"><span data-stu-id="45f55-123">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart.</span></span> <span data-ttu-id="45f55-124">Excel 可以调整行宽等设置，以便图表能够适应所提供的空间。</span><span class="sxs-lookup"><span data-stu-id="45f55-124">Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   - <span data-ttu-id="45f55-125">“系列”是指表格列中的一组数据点。</span><span class="sxs-lookup"><span data-stu-id="45f55-125">A "series" is a set of data points from a column of the table.</span></span> <span data-ttu-id="45f55-126">因为表格中只有一个非字符串列，所以 Excel 推断此列就是要绘制成图表的唯一一列数据点。</span><span class="sxs-lookup"><span data-stu-id="45f55-126">Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart.</span></span> <span data-ttu-id="45f55-127">它将其他列解释为图表标签。</span><span class="sxs-lookup"><span data-stu-id="45f55-127">It interprets the other columns as chart labels.</span></span> <span data-ttu-id="45f55-128">因此，图表中只有一个系列，它的索引为 0。</span><span class="sxs-lookup"><span data-stu-id="45f55-128">So there will be just one series in the chart and it will have index 0.</span></span> <span data-ttu-id="45f55-129">这是要标记为“金额（欧元）”的系列。</span><span class="sxs-lookup"><span data-stu-id="45f55-129">This is the one to label with "Value in €".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in €';
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="45f55-130">测试加载项</span><span class="sxs-lookup"><span data-stu-id="45f55-130">Test the add-in</span></span>


1. <span data-ttu-id="45f55-131">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="45f55-131">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="45f55-132">否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”\*\*\*\* 文件夹。</span><span class="sxs-lookup"><span data-stu-id="45f55-132">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="45f55-133">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。</span><span class="sxs-lookup"><span data-stu-id="45f55-133">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="45f55-134">为此，需要终止服务器进程，这样就可以通提示符输入生成命令。</span><span class="sxs-lookup"><span data-stu-id="45f55-134">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="45f55-135">生成后，重启服务器。</span><span class="sxs-lookup"><span data-stu-id="45f55-135">After the build, you restart the server.</span></span> <span data-ttu-id="45f55-136">接下来的几步执行的就是此进程。</span><span class="sxs-lookup"><span data-stu-id="45f55-136">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="45f55-137">运行命令 `npm run build`，将 ES6 源代码转换为 Internet Explorer 支持的旧版 JavaScript（Excel 在后台用来运行 Excel 加载项）。</span><span class="sxs-lookup"><span data-stu-id="45f55-137">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="45f55-138">运行命令 `npm start`，启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="45f55-138">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="45f55-139">通过关闭任务窗格来重新加载它，再选择“**开始**”菜单上的“**显示任务窗格**”，以重新打开加载项。</span><span class="sxs-lookup"><span data-stu-id="45f55-139">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
5. <span data-ttu-id="45f55-140">如果出于某种原因在工作表中打不开表格，请在任务窗格中依次选择“**创建表**”、“**筛选表**”和“**排序表**”按钮（按顺序和倒序中的任一顺序排序皆可）。</span><span class="sxs-lookup"><span data-stu-id="45f55-140">If for any reason the table is not in the open worksheet, in the taskpane, choose **Create Table** and then **Filter Table** and **Sort Table** buttons, in either order.</span></span>
6. <span data-ttu-id="45f55-141">选择“**创建图表**”按钮。</span><span class="sxs-lookup"><span data-stu-id="45f55-141">Choose the **Create Chart** button.</span></span> <span data-ttu-id="45f55-142">此时，图表创建完成，其中仅包含筛选出的行中的数据。</span><span class="sxs-lookup"><span data-stu-id="45f55-142">A chart is created and only the data from the rows that have been filtered are included.</span></span> <span data-ttu-id="45f55-143">底部数据点上的标签按图表的排序顺序进行排序，即按商家名称的字母倒序排序。</span><span class="sxs-lookup"><span data-stu-id="45f55-143">The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Excel 教程 - 创建图表](../images/excel-tutorial-create-chart.png)
