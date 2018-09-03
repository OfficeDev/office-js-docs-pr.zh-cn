<span data-ttu-id="a5385-101">如果表格很长，导致用户必须滚动才能看到一些行，那么标题行可能会在滚动时不可见。</span><span class="sxs-lookup"><span data-stu-id="a5385-101">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight.</span></span> <span data-ttu-id="a5385-102">本教程的这一步是，冻结以前创建的表格的标题行，让它在用户向下滚动工作表时依然可见。</span><span class="sxs-lookup"><span data-stu-id="a5385-102">In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span> 

> [!NOTE]
> <span data-ttu-id="a5385-103">此为 Excel 加载项分步教程页面。</span><span class="sxs-lookup"><span data-stu-id="a5385-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="a5385-104">如果是通过搜索引擎结果或其他直接链接到达此页面，请转到 [Excel 加载项教程](../tutorials/excel-tutorial.yml)介绍性页面，从头开始学习本教程。</span><span class="sxs-lookup"><span data-stu-id="a5385-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="freeze-the-tables-header-row"></a><span data-ttu-id="a5385-105">冻结表的标题行</span><span class="sxs-lookup"><span data-stu-id="a5385-105">Freeze the table's header row</span></span>

1. <span data-ttu-id="a5385-106">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="a5385-106">Open the project in your code editor.</span></span> 
2. <span data-ttu-id="a5385-107">打开文件 index.html。</span><span class="sxs-lookup"><span data-stu-id="a5385-107">Open the file index.html.</span></span>
3. <span data-ttu-id="a5385-108">在包含 `create-chart` 按钮的 `div` 下方，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="a5385-108">Below the `div` that contains the `create-chart` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="freeze-header">Freeze Header</button>            
    </div>
    ```

4. <span data-ttu-id="a5385-109">打开 app.js 文件。</span><span class="sxs-lookup"><span data-stu-id="a5385-109">Open the app.js file.</span></span>

5. <span data-ttu-id="a5385-110">在向 `create-chart` 按钮分配单击处理程序的代码行下方，添加下列代码：</span><span class="sxs-lookup"><span data-stu-id="a5385-110">Below the line that assigns a click handler to the `create-chart` button, add the following code:</span></span>

    ```js
    $('#freeze-header').click(freezeHeader);
    ```

6. <span data-ttu-id="a5385-111">在 `createChart` 函数下方，添加下列函数：</span><span class="sxs-lookup"><span data-stu-id="a5385-111">Below the `createChart` function add the following function:</span></span>

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

7. <span data-ttu-id="a5385-p103">将 `TODO1` 替换为以下代码。请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="a5385-p103">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="a5385-114">`Worksheet.freezePanes` 集合是工作表中的一组窗格，在工作表滚动时就地固定或冻结。</span><span class="sxs-lookup"><span data-stu-id="a5385-114">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>
   - <span data-ttu-id="a5385-p104">`freezeRows` 方法需要使用要就地固定的行数（自顶部算起）作为参数。传递 `1` 可以就地固定第一行。</span><span class="sxs-lookup"><span data-stu-id="a5385-p104">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ``` 

## <a name="test-the-add-in"></a><span data-ttu-id="a5385-117">测试加载项</span><span class="sxs-lookup"><span data-stu-id="a5385-117">Test the add-in</span></span>

1. <span data-ttu-id="a5385-118">如果上一阶段教程中的 Git Bash 窗口或已启用 Node.JS 的系统命令提示符仍处于打开状态，请按 Ctrl-C 两次，停止正在运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="a5385-118">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="a5385-119">否则，打开 Git Bash 窗口或已启用 Node.JS 的系统命令提示符，并转到项目的“开始”**** 文件夹。</span><span class="sxs-lookup"><span data-stu-id="a5385-119">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="a5385-120">虽然只要更改任意文件（包括 app.js 文件），浏览器同步服务器就会在任务窗格中重新加载加载项，但它不会重新转换 JavaScript。因此，必须重复执行生成命令，这样对 app.js 做出的更改才会生效。</span><span class="sxs-lookup"><span data-stu-id="a5385-120">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="a5385-121">为此，需要终止服务器进程，这样就可以通提示符输入生成命令。</span><span class="sxs-lookup"><span data-stu-id="a5385-121">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="a5385-122">生成后，重启服务器。</span><span class="sxs-lookup"><span data-stu-id="a5385-122">After the build, you restart the server.</span></span> <span data-ttu-id="a5385-123">接下来的几步执行的就是此进程。</span><span class="sxs-lookup"><span data-stu-id="a5385-123">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="a5385-124">运行命令 `npm run build`，将 ES6 源代码转换为 Internet Explorer 支持的旧版 JavaScript（Excel 在后台用来运行 Excel 加载项）。</span><span class="sxs-lookup"><span data-stu-id="a5385-124">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="a5385-125">运行命令 `npm start`，启动在 localhost 上运行的 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="a5385-125">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="a5385-126">通过关闭任务窗格来重新加载它，再选择“主页”**** 菜单上的“显示任务窗格”****，重新打开加载项。</span><span class="sxs-lookup"><span data-stu-id="a5385-126">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="a5385-127">如果表格在工作表中，请删除它。</span><span class="sxs-lookup"><span data-stu-id="a5385-127">If the table is in the worksheet, delete it.</span></span>
7. <span data-ttu-id="a5385-128">在任务窗格中，选择“创建表格”****。</span><span class="sxs-lookup"><span data-stu-id="a5385-128">In the taskpane, choose **Create Table**.</span></span> 
8. <span data-ttu-id="a5385-129">选择“冻结标题”**** 按钮。</span><span class="sxs-lookup"><span data-stu-id="a5385-129">Choose the **Freeze Header** button.</span></span>
9. <span data-ttu-id="a5385-130">向下滚动工作表，直到在上面的行不可见时表格标题在顶部依然可见。</span><span class="sxs-lookup"><span data-stu-id="a5385-130">Scroll down the worksheet enough to to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Excel 教程 - 冻结标题](../images/excel-tutorial-freeze-header.png)
