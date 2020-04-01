---
title: Excel 加载项教程
description: 在本教程中，你将学习如何构建一个 Excel 外接程序，用于创建、填充、筛选和排序表格、创建图表、冻结表格标题、保护工作表并打开对话框。
ms.date: 03/30/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: bc14027600a108f47cdf4766a51ebf2437b6f976
ms.sourcegitcommit: 73a3df90a51acf13416d6a049bddcd9aabc32441
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2020
ms.locfileid: "43069314"
---
# <a name="tutorial-create-an-excel-task-pane-add-in"></a><span data-ttu-id="5ae6e-103">教程：创建 Excel 任务窗格加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-103">Tutorial: Create an Excel task pane add-in</span></span>

<span data-ttu-id="5ae6e-104">在本教程中，将创建 Excel 任务窗格加载项，该加载项将：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-104">In this tutorial, you'll create an Excel task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="5ae6e-105">创建表格</span><span class="sxs-lookup"><span data-stu-id="5ae6e-105">Creates a table</span></span>
> * <span data-ttu-id="5ae6e-106">筛选和排序表格</span><span class="sxs-lookup"><span data-stu-id="5ae6e-106">Filters and sorts a table</span></span>
> * <span data-ttu-id="5ae6e-107">创建图表</span><span class="sxs-lookup"><span data-stu-id="5ae6e-107">Creates a chart</span></span>
> * <span data-ttu-id="5ae6e-108">冻结表格标题</span><span class="sxs-lookup"><span data-stu-id="5ae6e-108">Freezes a table header</span></span>
> * <span data-ttu-id="5ae6e-109">保护工作表</span><span class="sxs-lookup"><span data-stu-id="5ae6e-109">Protects a worksheet</span></span>
> * <span data-ttu-id="5ae6e-110">打开对话框</span><span class="sxs-lookup"><span data-stu-id="5ae6e-110">Opens a dialog</span></span>

> [!TIP]
> <span data-ttu-id="5ae6e-111">如果已完成了“[创建 Excel 任务窗格加载项](../quickstarts/excel-quickstart-jquery.md)”快速入门，并希望使用此项目作为该教程的起点，直接转至“[创建表](#create-a-table)”以开始此教程。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-111">If you've already completed the [Build an Excel task pane add-in](../quickstarts/excel-quickstart-jquery.md) quick start, and want to use that project as a starting point for this tutorial, go directly to the [Create a table](#create-a-table) section to start this tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5ae6e-112">先决条件</span><span class="sxs-lookup"><span data-stu-id="5ae6e-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="5ae6e-113">创建加载项项目</span><span class="sxs-lookup"><span data-stu-id="5ae6e-113">Create your add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="5ae6e-114">**选择项目类型:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="5ae6e-114">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="5ae6e-115">**选择脚本类型:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="5ae6e-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="5ae6e-116">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="5ae6e-116">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="5ae6e-117">**要支持哪一个 Office 客户端应用程序?**</span><span class="sxs-lookup"><span data-stu-id="5ae6e-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman 生成器](../images/yo-office-excel.png)

<span data-ttu-id="5ae6e-119">完成此向导后，生成器会创建项目，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="create-a-table"></a><span data-ttu-id="5ae6e-120">创建表</span><span class="sxs-lookup"><span data-stu-id="5ae6e-120">Create a table</span></span>

<span data-ttu-id="5ae6e-121">本教程的这一步是，以编程方式测试加载项是否支持用户的当前版本 Excel，向工作表中添加表格，使用数据填充表格，并设置格式。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-121">In this step of the tutorial, you'll programmatically test that your add-in supports the user's current version of Excel, add a table to a worksheet, populate the table with data, and format it.</span></span>

### <a name="code-the-add-in"></a><span data-ttu-id="5ae6e-122">编码加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-122">Code the add-in</span></span>

1. <span data-ttu-id="5ae6e-123">在代码编辑器中打开项目。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-123">Open the project in your code editor.</span></span>

2. <span data-ttu-id="5ae6e-124">打开 ./src/taskpane/taskpane.html\*\*\*\* 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-124">Open the file **./src/taskpane/taskpane.html**.</span></span>  <span data-ttu-id="5ae6e-125">此文件含有任务窗格的 HTML 标记。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-125">This file contains the HTML markup for the task pane.</span></span>

3. <span data-ttu-id="5ae6e-126">找到 `<main>` 元素并删除在开始 `<main>` 标记后和关闭 `</main>` 标记前出现的所有行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-126">Locate the `<main>` element and delete all lines that appear after the opening `<main>` tag and before the closing `</main>` tag.</span></span>

4. <span data-ttu-id="5ae6e-127">打开 `<main>` 标记后立即添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-127">Add the following markup immediately after the opening `<main>` tag:</span></span>

    ```html
    <button class="ms-Button" id="create-table">Create Table</button><br/><br/>
    ```

5. <span data-ttu-id="5ae6e-128">打开 ./src/taskpane/taskpane.js\*\*\*\* 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-128">Open the file **./src/taskpane/taskpane.js**.</span></span> <span data-ttu-id="5ae6e-129">此文件包含用于加快任务窗格与 Office 托管应用程序之间的交互的 Office JavaScript API 代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-129">This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

6. <span data-ttu-id="5ae6e-130">执行以下操作，删除对 `run` 按钮和 `run()` 函数的所有引用：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-130">Remove all references to the `run` button and the `run()` function by doing the following:</span></span>

    - <span data-ttu-id="5ae6e-131">查找并删除行 `document.getElementById("run").onclick = run;`。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-131">Locate and delete the line `document.getElementById("run").onclick = run;`.</span></span>

    - <span data-ttu-id="5ae6e-132">查找并删除整个 `run()` 函数。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-132">Locate and delete the entire `run()` function.</span></span>

7. <span data-ttu-id="5ae6e-133">在 `Office.onReady` 方法调用中，找到行 `if (info.host === Office.HostType.Excel) {` 并紧接着行添加下列代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-133">Within the `Office.onReady` method call, locate the line `if (info.host === Office.HostType.Excel) {` and add the following code immediately after that line.</span></span> <span data-ttu-id="5ae6e-134">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-134">Note:</span></span>

    - <span data-ttu-id="5ae6e-135">此代码的第一部分用于确定用户的 Excel 版本是否支持包含本系列教程将使用的所有 API 的 Excel.js 版本。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-135">The first part of this code determines whether the user's version of Excel supports a version of Excel.js that includes all the APIs that this series of tutorials will use.</span></span> <span data-ttu-id="5ae6e-136">在生产加载项中，若要隐藏或禁用调用不受支持的 API 的 UI，请使用条件块的主体。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-136">In a production add-in, use the body of the conditional block to hide or disable the UI that would call unsupported APIs.</span></span> <span data-ttu-id="5ae6e-137">这样一来，用户仍可以使用 Excel 版本支持的加载项部分。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-137">This will enable the user to still make use of the parts of the add-in that are supported by their version of Excel.</span></span>

    - <span data-ttu-id="5ae6e-138">此代码的第二部分为 `create-table` 按钮添加了事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-138">The second part of this code adds an event handler for the `create-table` button.</span></span>

    ```js
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    ```

8. <span data-ttu-id="5ae6e-139">将以下函数添加到文件结尾。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-139">Add the following function to the end of the file.</span></span> <span data-ttu-id="5ae6e-140">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-140">Note:</span></span>

    - <span data-ttu-id="5ae6e-p106">Excel.js 业务逻辑将添加到传递给 `Excel.run` 的函数。 此逻辑不立即执行。 相反，它会被添加到挂起的命令队列中。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p106">Your Excel.js business logic will be added to the function that is passed to `Excel.run`. This logic does not execute immediately. Instead, it is added to a queue of pending commands.</span></span>

    - <span data-ttu-id="5ae6e-144">`context.sync` 方法将所有已排入队列的命令发送到 Excel 以供执行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-144">The `context.sync` method sends all queued commands to Excel for execution.</span></span>

    - <span data-ttu-id="5ae6e-p107">`Excel.run` 后跟 `catch` 块。 这是应始终遵循的最佳做法。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p107">The `Excel.run` is followed by a `catch` block. This is a best practice that you should always follow.</span></span> 

    ```js
    function createTable() {
        Excel.run(function (context) {

            // TODO1: Queue table creation logic here.

            // TODO2: Queue commands to populate the table with data.

            // TODO3: Queue commands to format the table.

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

9. <span data-ttu-id="5ae6e-147">在 `createTable()` 函数中，将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-147">Within the `createTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5ae6e-148">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-148">Note:</span></span>

    - <span data-ttu-id="5ae6e-p109">此代码通过使用工作表的表格集合的 `add` 方法来创建表格，即使是空的，也始终存在。 这是创建 Excel.js 对象的标准方式。 没有类构造函数 API，切勿使用 `new` 运算符创建 Excel 对象。 相反，请添加到父集合对象。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p109">The code creates a table by using `add` method of a worksheet's table collection, which always exists even if it is empty. This is the standard way that Excel.js objects are created. There are no class constructor APIs, and you never use a `new` operator to create an Excel object. Instead, you add to a parent collection object.</span></span>

    - <span data-ttu-id="5ae6e-p110">`add` 方法的第一个参数仅是表格最上面一行的范围，而不是表格最终使用的整个范围。 这是因为当加载项填充数据行时（在下一步中），它将新行添加到表中，而不是将值写入现有行的单元格。 这是更为常见的模式，因为在创建表时表的行数通常是未知的。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p110">The first parameter of the `add` method is the range of only the top row of the table, not the entire range the table will ultimately use. This is because when the add-in populates the data rows (in the next step), it will add new rows to the table instead of writing values to the cells of existing rows. This is a more common pattern because the number of rows that a table will have is often not known when the table is created.</span></span>

    - <span data-ttu-id="5ae6e-156">表名称必须在整个工作簿中都是唯一的，而不仅仅是在工作表一级。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-156">Table names must be unique across the entire workbook, not just the worksheet.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";
    ```

10. <span data-ttu-id="5ae6e-157">在 `createTable()` 函数中，将 `TODO2` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-157">Within the `createTable()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5ae6e-158">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-158">Note:</span></span>

    - <span data-ttu-id="5ae6e-159">范围的单元格值是通过一组数组进行设置。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-159">The cell values of a range are set with an array of arrays.</span></span>

    - <span data-ttu-id="5ae6e-p112">表格中的新行是通过调用表格的行集合的 `add` 方法进行创建。 通过在作为第二个参数传递的父数组中添加多个单元格值数组，可以在一次 `add` 调用中添加多个行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p112">New rows are created in a table by calling the `add` method of the table's row collection. You can add multiple rows in a single call of `add` by including multiple cell value arrays in the parent array that is passed as the second parameter.</span></span>

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

11. <span data-ttu-id="5ae6e-162">在 `createTable()` 函数中，将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-162">Within the `createTable()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="5ae6e-163">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-163">Note:</span></span>

    - <span data-ttu-id="5ae6e-164">此代码将从零开始编制的索引传递给表格的列集合的 `getItemAt` 方法，以获取对“金额”\*\*\*\* 列的引用。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-164">The code gets a reference to the **Amount** column by passing its zero-based index to the `getItemAt` method of the table's column collection.</span></span>

        > [!NOTE]
        > <span data-ttu-id="5ae6e-165">Excel.js 集合对象（如 `TableCollection`、`WorksheetCollection` 和 `TableColumnCollection`）有 `items` 属性，此属性是子对象类型的数组（如 `Table`、`Worksheet` 或 `TableColumn`），但 `*Collection` 对象本身并不是数组。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-165">Excel.js collection objects, such as `TableCollection`, `WorksheetCollection`, and `TableColumnCollection` have an `items` property that is an array of the child object types, such as `Table` or `Worksheet` or `TableColumn`; but a `*Collection` object is not itself an array.</span></span>

    - <span data-ttu-id="5ae6e-166">然后，此代码将“金额”\*\*\*\* 列的范围格式化为欧元（精确到小数点后两位）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-166">The code then formats the range of the **Amount** column as Euros to the second decimal.</span></span> 

    - <span data-ttu-id="5ae6e-p114">最后，它确保了列宽和行高足以容纳最长（或最高）的数据项。 请注意，此代码必须获取要格式化的 `Range` 对象。 `TableColumn` 和 `TableRow` 对象没有格式属性。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p114">Finally, it ensures that the width of the columns and height of the rows is big enough to fit the longest (or tallest) data item. Notice that the code must get `Range` objects to format. `TableColumn` and `TableRow` objects do not have format properties.</span></span>

    ```js
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    ```

12. <span data-ttu-id="5ae6e-170">验证是否已保存了对项目所做的所有更改。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-170">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5ae6e-171">测试加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-171">Test the add-in</span></span>

1. <span data-ttu-id="5ae6e-172">完成以下步骤，以启动本地 Web 服务器并旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-172">Complete the following steps to start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5ae6e-173">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-173">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="5ae6e-174">如果系统在运行以下命令之一后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-174">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="5ae6e-175">如果在 Mac 上测试加载项，请先运行项目根目录中的以下命令，然后再继续。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-175">If you're testing your add-in on Mac, run the following command in the root directory of your project before proceeding.</span></span> <span data-ttu-id="5ae6e-176">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-176">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - <span data-ttu-id="5ae6e-177">若要在 Excel 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-177">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5ae6e-178">这将启动本地的 Web 服务器 (如果尚未运行的话), 并使用加载的加载项打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-178">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="5ae6e-179">若要在 Excel 网页版中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-179">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5ae6e-180">运行此命令时，本地 Web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-180">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="5ae6e-181">若要使用加载项，请在 Excel 网页版中打开新的文档，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-181">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

2. <span data-ttu-id="5ae6e-182">在 Excel 中，依次选择的“**开始**”选项卡和功能区中的“**显示任务窗格**”按钮，以打开加载项任务窗格。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-182">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel 加载项按钮](../images/excel-quickstart-addin-3b.png)

3. <span data-ttu-id="5ae6e-184">在任务窗格中，选择“**创建表**”按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-184">In the task pane, choose the **Create Table** button.</span></span>

    ![Excel 教程 - 创建表格](../images/excel-tutorial-create-table-2.png)

## <a name="filter-and-sort-a-table"></a><span data-ttu-id="5ae6e-186">筛选和排序表格</span><span class="sxs-lookup"><span data-stu-id="5ae6e-186">Filter and sort a table</span></span>

<span data-ttu-id="5ae6e-187">本教程的这一步是，筛选并排序之前创建的表。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-187">In this step of the tutorial, you'll filter and sort the table that you created previously.</span></span>

### <a name="filter-the-table"></a><span data-ttu-id="5ae6e-188">筛选表格</span><span class="sxs-lookup"><span data-stu-id="5ae6e-188">Filter the table</span></span>

1. <span data-ttu-id="5ae6e-189">打开 **./src/taskpane/taskpane.html** 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-189">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="5ae6e-190">找到 `<button>` 按钮的 `create-table` 元素，并在该行后添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-190">Locate the `<button>` element for the `create-table` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="filter-table">Filter Table</button><br/><br/>
    ```

3. <span data-ttu-id="5ae6e-191">打开 ./src/taskpane/taskpane.js\*\*\*\* 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-191">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="5ae6e-192">在 `Office.onReady` 方法调用中，找到用于将单击处理程序分配到 `create-table` 按钮的行，并在该行后添加以下代码：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-192">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("filter-table").onclick = filterTable;
    ```

5. <span data-ttu-id="5ae6e-193">将以下函数添加到文件末端：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-193">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5ae6e-194">在 `filterTable()` 函数中，将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-194">Within the `filterTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5ae6e-195">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-195">Note:</span></span>

   - <span data-ttu-id="5ae6e-p120">代码先将列名称传递给 `getItem` 方法（而不是像 `getItemAt` 方法一样将列索引传递给 `createTable` 方法），获取对需要筛选的列的引用。 由于用户可以移动表格列，因此给定索引处的列可能会在表格创建后更改。 所以，更安全的做法是，使用列名称获取对列的引用。 上一教程安全地使用了 `getItemAt`，因为是在与创建表格完全相同的方法中使用了它，所以用户没有机会移动列。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p120">The code first gets a reference to the column that needs filtering by passing the column name to the `getItem` method, instead of passing its index to the `getItemAt` method as the `createTable` method does. Since users can move table columns, the column at a given index might change after the table is created. Hence, it is safer to use the column name to get a reference to the column. We used `getItemAt` safely in the preceding tutorial, because we used it in the very same method that creates the table, so there is no chance that a user has moved the column.</span></span>

   - <span data-ttu-id="5ae6e-200">`applyValuesFilter` 方法是对 `Filter` 对象执行的多种筛选方法之一。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-200">The `applyValuesFilter` method is one of several filtering methods on the `Filter` object.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var categoryFilter = expensesTable.columns.getItem('Category').filter;
    categoryFilter.applyValuesFilter(['Education', 'Groceries']);
    ``` 

### <a name="sort-the-table"></a><span data-ttu-id="5ae6e-201">排序表格</span><span class="sxs-lookup"><span data-stu-id="5ae6e-201">Sort the table</span></span>

1. <span data-ttu-id="5ae6e-202">打开 **./src/taskpane/taskpane.html** 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-202">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="5ae6e-203">找到 `<button>` 按钮的 `filter-table` 元素，并在该行后添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-203">Locate the `<button>` element for the `filter-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="sort-table">Sort Table</button><br/><br/>
    ```

3. <span data-ttu-id="5ae6e-204">打开 ./src/taskpane/taskpane.js\*\*\*\* 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-204">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="5ae6e-205">在 `Office.onReady` 方法调用中，找到用于将单击处理程序分配到 `filter-table` 按钮的行，并在该行后添加以下代码：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-205">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `filter-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("sort-table").onclick = sortTable;
    ```

5. <span data-ttu-id="5ae6e-206">将以下函数添加到文件末端：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-206">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5ae6e-207">在 `sortTable()` 函数中，将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-207">Within the `sortTable()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5ae6e-208">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-208">Note:</span></span>

   - <span data-ttu-id="5ae6e-209">此代码创建一组 `SortField` 对象，其中只有一个成员，因为加载项只对“商家”列进行了排序。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-209">The code creates an array of `SortField` objects which has just one member since the add-in only sorts on the Merchant column.</span></span>

   - <span data-ttu-id="5ae6e-210">`SortField` 对象的 `key` 属性是要排序的列的从零开始编制索引。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-210">The `key` property of a `SortField` object is the zero-based index of the column to sort-on.</span></span>

   - <span data-ttu-id="5ae6e-211">`Table` 的 `sort` 成员是 `TableSort` 对象，并不是方法。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-211">The `sort` member of a `Table` is a `TableSort` object, not a method.</span></span> <span data-ttu-id="5ae6e-212">`SortField` 传递到 `TableSort` 对象的 `apply` 方法。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-212">The `SortField`s are passed to the `TableSort` object's `apply` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var sortFields = [
        {
            key: 1,            // Merchant column
            ascending: false,
        }
    ];

    expensesTable.sort.apply(sortFields);
    ```

7. <span data-ttu-id="5ae6e-213">验证是否已保存了对项目所做的所有更改。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-213">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5ae6e-214">测试加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-214">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5ae6e-215">如果加载项任务窗格已在 Excel 中打开，转至 **Home** 选项卡并选中功能区中的“**显示任务窗格**”以打开。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-215">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5ae6e-216">如果之前在此教程中添加的表格未在打开的工作表中出现，选择任务窗格中的“**创建表**”按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-216">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button in the task pane.</span></span>

4. <span data-ttu-id="5ae6e-217">选择“**筛选表**”按钮和“**排序表**”按钮（按顺序和倒序中的任一顺序排序皆可）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-217">Choose the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

    ![Excel 教程 - 筛选和排序表格](../images/excel-tutorial-filter-and-sort-table-2.png)

## <a name="create-a-chart"></a><span data-ttu-id="5ae6e-219">创建图表</span><span class="sxs-lookup"><span data-stu-id="5ae6e-219">Create a chart</span></span>

<span data-ttu-id="5ae6e-220">本教程的这一步是，使用先前创建的表中的数据创建图表，再设置图表格式。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-220">In this step of the tutorial, you'll create a chart using data from the table that you created previously, and then format the chart.</span></span>

### <a name="chart-a-chart-using-table-data"></a><span data-ttu-id="5ae6e-221">使用表格数据绘制图表</span><span class="sxs-lookup"><span data-stu-id="5ae6e-221">Chart a chart using table data</span></span>

1. <span data-ttu-id="5ae6e-222">打开 **./src/taskpane/taskpane.html** 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-222">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="5ae6e-223">找到 `<button>` 按钮的 `sort-table` 元素，并在该行后添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-223">Locate the `<button>` element for the `sort-table` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="create-chart">Create Chart</button><br/><br/>
    ```

3. <span data-ttu-id="5ae6e-224">打开 ./src/taskpane/taskpane.js\*\*\*\* 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-224">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="5ae6e-225">在 `Office.onReady` 方法调用中，找到用于将单击处理程序分配到 `sort-table` 按钮的行，并在该行后添加以下代码：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-225">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `sort-table` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("create-chart").onclick = createChart;
    ```

5. <span data-ttu-id="5ae6e-226">将以下函数添加到文件末端：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-226">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5ae6e-227">在 `createChart()` 函数中，将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-227">Within the `createChart()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5ae6e-228">请注意，为了排除标题行，此代码使用 `Table.getDataBodyRange` 方法（而不是 `getRange` 方法），获取要绘制成图表的数据的范围。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-228">Note that in order to exclude the header row, the code uses the `Table.getDataBodyRange` method to get the range of data you want to chart instead of the `getRange` method.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
    var dataRange = expensesTable.getDataBodyRange();
    ```

7. <span data-ttu-id="5ae6e-229">在 `createChart()` 函数中，将 `TODO2` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-229">Within the `createChart()` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5ae6e-230">请注意以下参数：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-230">Note the following parameters:</span></span>

   - <span data-ttu-id="5ae6e-p125">`add` 方法的第一个参数指定图表类型。有几十种类型。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p125">The first parameter to the `add` method specifies the type of chart. There are several dozen types.</span></span>

   - <span data-ttu-id="5ae6e-233">第二个参数指定要在图表中添加的数据的范围。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-233">The second parameter specifies the range of data to include in the chart.</span></span>

   - <span data-ttu-id="5ae6e-234">第三个参数确定是按行方向还是按列方向绘制表格中的一系列数据点。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-234">The third parameter determines whether a series of data points from the table should be charted row-wise or column-wise.</span></span> <span data-ttu-id="5ae6e-235">选项 `auto` 指示 Excel 确定最佳方法。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-235">The option `auto` tells Excel to decide the best method.</span></span>

    ```js
    var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'auto');
    ```

8. <span data-ttu-id="5ae6e-236">在 `createChart()` 函数中，将 `TODO3` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-236">Within the `createChart()` function, replace `TODO3` with the following code.</span></span> <span data-ttu-id="5ae6e-237">此代码的大部分内容非常直观明了。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-237">Most of this code is self-explanatory.</span></span> <span data-ttu-id="5ae6e-238">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-238">Note:</span></span>
   
   - <span data-ttu-id="5ae6e-p128">`setPosition` 方法的参数指定应包含图表的工作表区域的左上角和右下角单元格。 Excel 可以调整行宽等设置，以便图表能够适应所提供的空间。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p128">The parameters to the `setPosition` method specify the upper left and lower right cells of the worksheet area that should contain the chart. Excel can adjust things like line width to make the chart look good in the space it has been given.</span></span>
   
   - <span data-ttu-id="5ae6e-p129">“系列”是指表格列中的一组数据点。 因为表格中只有一个非字符串列，所以 Excel 推断此列就是要绘制成图表的唯一一列数据点。 它将其他列解释为图表标签。 因此，图表中只有一个系列，它的索引为 0。 这是要标记为“金额（&euro;）”的系列。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p129">A "series" is a set of data points from a column of the table. Since there is only one non-string column in the table, Excel infers that the column is the only column of data points to chart. It interprets the other columns as chart labels. So there will be just one series in the chart and it will have index 0. This is the one to label with "Value in &euro;".</span></span>

    ```js
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right"
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = 'Value in &euro;';
    ```

9. <span data-ttu-id="5ae6e-246">验证是否已保存了对项目所做的所有更改。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-246">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5ae6e-247">测试加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-247">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5ae6e-248">如果加载项任务窗格已在 Excel 中打开，转至 **Home** 选项卡并选中功能区中的“**显示任务窗格**”以打开。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-248">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5ae6e-249">如果之前在此教程中添加的表格未在打开的工作表中出现，选择任务窗格中的“**创建表**”按钮，随后选择“**筛选表**”按钮和“**排序表**”按钮（按顺序和倒序中的任一顺序排序皆可）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-249">If the table you added previously in this tutorial is not present in the open worksheet, choose the **Create Table** button, and then the **Filter Table** button and the **Sort Table** button, in either order.</span></span>

4. <span data-ttu-id="5ae6e-p130">选择 **“创建图表”** 按钮。 此时，图表创建完成，其中仅包含筛选出的行中的数据。 底部数据点上的标签按图表的排序顺序进行排序，即按商家名称的字母倒序排序。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p130">Choose the **Create Chart** button. A chart is created and only the data from the rows that have been filtered are included. The labels on the data points across the bottom are in the sort order of the chart; that is, merchant names in reverse alphabetical order.</span></span>

    ![Excel 教程 - 创建图表](../images/excel-tutorial-create-chart-2.png)

## <a name="freeze-a-table-header"></a><span data-ttu-id="5ae6e-254">冻结表格标题</span><span class="sxs-lookup"><span data-stu-id="5ae6e-254">Freeze a table header</span></span>

<span data-ttu-id="5ae6e-p131">如果表格很长，导致用户必须滚动才能看到一些行，那么标题行可能会在滚动时不可见。 本教程的这一步是，冻结以前创建的表格的标题行，让它在用户向下滚动工作表时依然可见。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p131">When a table is long enough that a user must scroll to see some rows, the header row can scroll out of sight. In this step of the tutorial, you'll freeze the header row of the table that you created previously, so that it remains visible even as the user scrolls down the worksheet.</span></span>

### <a name="freeze-the-tables-header-row"></a><span data-ttu-id="5ae6e-257">冻结表格的标题行</span><span class="sxs-lookup"><span data-stu-id="5ae6e-257">Freeze the table's header row</span></span>

1. <span data-ttu-id="5ae6e-258">打开 **./src/taskpane/taskpane.html** 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-258">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="5ae6e-259">找到 `<button>` 按钮的 `create-chart` 元素，并在该行后添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-259">Locate the `<button>` element for the `create-chart` button, and add the following markup after that line:</span></span> 

    ```html
    <button class="ms-Button" id="freeze-header">Freeze Header</button><br/><br/>
    ```

3. <span data-ttu-id="5ae6e-260">打开 ./src/taskpane/taskpane.js\*\*\*\* 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-260">Open the file **./src/taskpane/taskpane.js**.</span></span>

4. <span data-ttu-id="5ae6e-261">在 `Office.onReady` 方法调用中，找到用于将单击处理程序分配到 `create-chart` 按钮的行，并在该行后添加以下代码：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-261">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `create-chart` button, and add the following code after that line:</span></span>

    ```js
    document.getElementById("freeze-header").onclick = freezeHeader;
    ```

5. <span data-ttu-id="5ae6e-262">将以下函数添加到文件末端：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-262">Add the following function to the end of the file:</span></span>

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

6. <span data-ttu-id="5ae6e-263">在 `freezeHeader()` 函数中，将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-263">Within the `freezeHeader()` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5ae6e-264">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-264">Note:</span></span>

   - <span data-ttu-id="5ae6e-265">`Worksheet.freezePanes` 集合是工作表中的一组窗格，在工作表滚动时就地固定或冻结。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-265">The `Worksheet.freezePanes` collection is a set of panes in the worksheet that are pinned, or frozen, in place when the worksheet is scrolled.</span></span>

   - <span data-ttu-id="5ae6e-p133">`freezeRows` 方法需要使用要就地固定的行数（自顶部算起）作为参数。传递 `1` 可以就地固定第一行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p133">The `freezeRows` method takes as a parameter the number of rows, from the top that are to be pinned in place. We pass `1` to pin the first row in place.</span></span>

    ```js
    var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);
    ```

7. <span data-ttu-id="5ae6e-268">验证是否已保存了对项目所做的所有更改。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-268">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5ae6e-269">测试加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-269">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5ae6e-270">如果加载项任务窗格已在 Excel 中打开，转至 **Home** 选项卡并选中功能区中的“**显示任务窗格**”以打开。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-270">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5ae6e-271">如果之前在此教程中添加的表格出现在此工作表中，将其删除。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-271">If the table you added previously in this tutorial is present in the worksheet, delete it.</span></span>

4. <span data-ttu-id="5ae6e-272">在任务窗格中，选择“**创建表**”按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-272">In the task pane, choose the **Create Table** button.</span></span>

5. <span data-ttu-id="5ae6e-273">在任务窗格中，选择“**冻结标题**”按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-273">In the task pane, choose the **Freeze Header** button.</span></span>

6. <span data-ttu-id="5ae6e-274">尽量向下滚动工作表，直到在上面的行不可见时表格标题在顶部依然可见。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-274">Scroll down the worksheet far enough to see that the table header remains visible at the top even when the higher rows scroll out of sight.</span></span>

    ![Excel 教程 - 冻结标题](../images/excel-tutorial-freeze-header-2.png)

## <a name="protect-a-worksheet"></a><span data-ttu-id="5ae6e-276">保护工作表</span><span class="sxs-lookup"><span data-stu-id="5ae6e-276">Protect a worksheet</span></span>

<span data-ttu-id="5ae6e-277">本教程的这一步是，向功能区添加另一个按钮。如果用户选择此按钮，便会执行所定义的函数，从而启用和禁用工作表保护。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-277">In this step of the tutorial, you'll add another button to the ribbon that, when chosen, executes a function that you'll define to toggle worksheet protection on and off.</span></span>

### <a name="configure-the-manifest-to-add-a-second-ribbon-button"></a><span data-ttu-id="5ae6e-278">将清单配置为添加第二个功能区按钮</span><span class="sxs-lookup"><span data-stu-id="5ae6e-278">Configure the manifest to add a second ribbon button</span></span>

1. <span data-ttu-id="5ae6e-279">打开清单文件 **./manifest.xml**。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-279">Open the manifest file **./manifest.xml**.</span></span>

2. <span data-ttu-id="5ae6e-280">查找 `<Control>` 元素。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-280">Locate the `<Control>` element.</span></span> <span data-ttu-id="5ae6e-281">此元素定义了“主页”\*\*\*\* 功能区上一直用于启动加载项的“显示任务窗格”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-281">This element defines the **Show Taskpane** button on the **Home** ribbon you have been using to launch the add-in.</span></span> <span data-ttu-id="5ae6e-282">将向“主页”\*\*\*\* 功能区上的相同组添加第二个按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-282">We're going to add a second button to the same group on the **Home** ribbon.</span></span> <span data-ttu-id="5ae6e-283">在结束 Control 标记 (`</Control>`) 和结束 Group 标记 (`</Group>`) 之间，添加下列标记。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-283">In between the end Control tag (`</Control>`) and the end Group tag (`</Group>`), add the following markup.</span></span>

    ```xml
    <Control xsi:type="Button" id="<!--TODO1: Unique (in manifest) name for button -->">
        <Label resid="<!--TODO2: Button label -->" />
        <Supertip>            
            <Title resid="<!-- TODO3: Button tool tip title -->" />
            <Description resid="<!-- TODO4: Button tool tip description -->" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="<!-- TODO5: Specify the type of action-->">
            <!-- TODO6: Identify the function.-->
        </Action>
    </Control>
    ```

3. <span data-ttu-id="5ae6e-284">在刚添加至清单文件的 XML 内，将 `TODO1` 替换为字符串，以便向按钮提供在此清单文件内唯一的 ID。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-284">Within the XML you just added to the manifest file, replace `TODO1` with a string that gives the button an ID that is unique within this manifest file.</span></span> <span data-ttu-id="5ae6e-285">由于按钮将启用和禁用工作表保护，因此请使用“ToggleProtection”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-285">Since our button is going to toggle protection of the worksheet on and off, use "ToggleProtection".</span></span> <span data-ttu-id="5ae6e-286">完成时，`Control` 元素的开始标记如下所示：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-286">When you are done, the opening tag for the `Control` element should look like this:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
    ```

4. <span data-ttu-id="5ae6e-287">接下来的三个 `TODO` 设置“resid”（这是资源 ID 的简称）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-287">The next three `TODO`s set "resid"s, which is short for resource ID.</span></span> <span data-ttu-id="5ae6e-288">资源是字符串，这三个字符串将在后续步骤中创建。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-288">A resource is a string, and you'll create these three strings in a later step.</span></span> <span data-ttu-id="5ae6e-289">现在，需要向资源提供 ID。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-289">For now, you need to give IDs to the resources.</span></span> <span data-ttu-id="5ae6e-290">虽然按钮标签应名为“切换保护”，但此字符串的 *ID* 应为“ProtectionButtonLabel”。因此 `Label` 元素的样式如下：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-290">The button label should read "Toggle Protection", but the *ID* of this string should be "ProtectionButtonLabel", so the `Label` element should look like this:</span></span>

    ```xml
    <Label resid="ProtectionButtonLabel" />
    ```

5. <span data-ttu-id="5ae6e-291">`SuperTip` 元素定义了按钮的工具提示。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-291">The `SuperTip` element defines the tool tip for the button.</span></span> <span data-ttu-id="5ae6e-292">由于工具提示标题应与按钮标签相同，因此使用完全相同的资源 ID，即“ProtectionButtonLabel”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-292">The tool tip title should be the same as the button label, so we use the very same resource ID: "ProtectionButtonLabel".</span></span> <span data-ttu-id="5ae6e-293">工具提示说明为“单击即可启用和禁用工作表保护”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-293">The tool tip description will be "Click to turn protection of the worksheet on and off".</span></span> <span data-ttu-id="5ae6e-294">不过，`ID` 应为“ProtectionButtonToolTip”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-294">But the `ID` should be "ProtectionButtonToolTip".</span></span> <span data-ttu-id="5ae6e-295">完成后，`SuperTip` 元素如下所示：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-295">So, when you are done, the `SuperTip` element should look like this:</span></span> 

    ```xml
    <Supertip>            
        <Title resid="ProtectionButtonLabel" />
        <Description resid="ProtectionButtonToolTip" />
    </Supertip>
    ```

   > [!NOTE] 
   > <span data-ttu-id="5ae6e-p138">在生产加载项中，不建议对两个不同的按钮使用相同的图标；但为了简单起见，本教程将采用这样的做法。 因此，新 `Icon` 中的 `Control` 标记直接就是现有 `Icon` 中 `Control` 元素的副本。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p138">In a production add-in, you would not want to use the same icon for two different buttons; but to simplify this tutorial, we'll do that. So the `Icon` markup in our new `Control` is just a copy of the `Icon` element from the existing `Control`.</span></span> 

6. <span data-ttu-id="5ae6e-298">虽然清单中现有原始 `Control` 元素内的 `Action` 元素的类型设置为 `ShowTaskpane`，但新按钮不会要打开任务窗格，而是要运行在后续步骤中创建的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-298">The `Action` element inside the original `Control` element that was already present in the manifest, has its type set to `ShowTaskpane`, but our new button isn't going to open a task pane; it's going to run a custom function that you create in a later step.</span></span> <span data-ttu-id="5ae6e-299">因此，将 `TODO5` 替换为 `ExecuteFunction`，即触发自定义函数的按钮的操作类型。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-299">So replace `TODO5` with `ExecuteFunction` which is the action type for buttons that trigger custom functions.</span></span> <span data-ttu-id="5ae6e-300">`Action` 元素的开始标记如下所示：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-300">The opening tag for the `Action` element should look like this:</span></span>
 
    ```xml
    <Action xsi:type="ExecuteFunction">
    ```

7. <span data-ttu-id="5ae6e-p140">原始 `Action` 元素的子元素指定任务窗格 ID，以及应当在任务窗格中打开的页面 URL。 不过，`Action` 类型的 `ExecuteFunction` 元素只有一个子元素，用于命名控件执行的函数。 此函数（名为 `toggleProtection`）将在后续步骤中创建。 因此，将 `TODO6` 替换为以下标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p140">The original `Action` element has child elements that specify a task pane ID and a URL of the page that should be opened in the task pane. But an `Action` element of the `ExecuteFunction` type has a single child element that names the function that the control executes. You'll create that function in a later step, and it will be called `toggleProtection`. So, replace `TODO6` with the following markup:</span></span>
 
    ```xml
    <FunctionName>toggleProtection</FunctionName>
    ```

    <span data-ttu-id="5ae6e-305">此时，整个 `Control` 标记应如下所示：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-305">The entire `Control` markup should now look like the following:</span></span>

    ```xml
    <Control xsi:type="Button" id="ToggleProtection">
        <Label resid="ProtectionButtonLabel" />
        <Supertip>            
            <Title resid="ProtectionButtonLabel" />
            <Description resid="ProtectionButtonToolTip" />
        </Supertip>
        <Icon>
            <bt:Image size="16" resid="Icon.16x16"/>
            <bt:Image size="32" resid="Icon.32x32"/>
            <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
           <FunctionName>toggleProtection</FunctionName>
        </Action>
    </Control>
    ```

8. <span data-ttu-id="5ae6e-306">向下滚动到清单的 `Resources` 部分。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-306">Scroll down to the `Resources` section of the manifest.</span></span>

9. <span data-ttu-id="5ae6e-307">将下列标记添加为 `bt:ShortStrings` 元素的子级。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-307">Add the following markup as a child of the `bt:ShortStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonLabel" DefaultValue="Toggle Worksheet Protection" />
    ```

10. <span data-ttu-id="5ae6e-308">将下列标记添加为 `bt:LongStrings` 元素的子级。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-308">Add the following markup as a child of the `bt:LongStrings` element.</span></span>

    ```xml
    <bt:String id="ProtectionButtonToolTip" DefaultValue="Click to protect or unprotect the current worksheet." />
    ```

11. <span data-ttu-id="5ae6e-309">保存文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-309">Save the file.</span></span>

### <a name="create-the-function-that-protects-the-sheet"></a><span data-ttu-id="5ae6e-310">创建工作表保护函数</span><span class="sxs-lookup"><span data-stu-id="5ae6e-310">Create the function that protects the sheet</span></span>

1. <span data-ttu-id="5ae6e-311">打开文件 **.\commands\commands.js**。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-311">Open the file **.\commands\commands.js**.</span></span>

2. <span data-ttu-id="5ae6e-312">紧接着 `action` 函数添加下列函数。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-312">Add the following function immediately after the `action` function.</span></span> <span data-ttu-id="5ae6e-313">注意，我们向函数和函数调用 `args.completed` 的最后一行指定了 `args` 参数。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-313">Note that we specify an `args` parameter to the function and the very last line of the function calls `args.completed`.</span></span> <span data-ttu-id="5ae6e-314">**ExecuteFunction** 类型的所有加载项命令都必须满足这项要求。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-314">This is a requirement for all add-in commands of type **ExecuteFunction**.</span></span> <span data-ttu-id="5ae6e-315">它会指示 Office 主机应用，函数已完成，且 UI 可以再次变成响应式。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-315">It signals the Office host application that the function has finished and the UI can become responsive again.</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {
            
            // TODO1: Queue commands to reverse the protection status of the current worksheet.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        args.completed();
    }
    ```

3. <span data-ttu-id="5ae6e-316">添加下列行至文件结尾：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-316">Add the following line to the end of the file:</span></span>

    ```js
    g.toggleProtection = toggleProtection;
    ```

4. <span data-ttu-id="5ae6e-317">在 `toggleProtection` 函数中，将 `TODO1` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-317">Within the `toggleProtection` function, replace `TODO1` with the following code.</span></span> <span data-ttu-id="5ae6e-318">此代码使用处于标准切换模式的工作表对象 protection 属性。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-318">This code uses the worksheet object's protection property in a standard toggle pattern.</span></span> <span data-ttu-id="5ae6e-319">`TODO2` 将在下一部分中进行介绍。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-319">The `TODO2` will be explained in the next section.</span></span>

    ```js
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // TODO2: Queue command to load the sheet's "protection.protected" property from
    //        the document and re-synchronize the document and task pane.

    if (sheet.protection.protected) {
        sheet.protection.unprotect();
    } else {
        sheet.protection.protect();
    }
    ``` 

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a><span data-ttu-id="5ae6e-320">添加代码以将文档属性提取到任务窗格的脚本对象</span><span class="sxs-lookup"><span data-stu-id="5ae6e-320">Add code to fetch document properties into the task pane's script objects</span></span>

<span data-ttu-id="5ae6e-321">在此教程中创建的各函数内，通过对命令进行排队来*写入* Office 文档。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-321">In each function that you've created in this tutorial until now, you queued commands to *write* to the Office document.</span></span> <span data-ttu-id="5ae6e-322">每个函数结束时都会调用 `context.sync()` 方法，从而将排入队列的命令发送到文档，以供执行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-322">Each function ended with a call to the `context.sync()` method which sends the queued commands to the document to be executed.</span></span> <span data-ttu-id="5ae6e-323">不过，在上一步中添加的代码调用的是 `sheet.protection.protected` 属性，这与之前编写的函数明显不同，因为 `sheet` 对象只是任务窗格脚本中的代理对象。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-323">But the code you added in the last step calls the `sheet.protection.protected` property, and this is a significant difference from the earlier functions you wrote, because the `sheet` object is only a proxy object that exists in your task pane's script.</span></span> <span data-ttu-id="5ae6e-324">它并不了解文档的实际保护状态，因此它的 `protection.protected` 属性无法有实值。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-324">It doesn't know what the actual protection state of the document is, so its `protection.protected` property can't have a real value.</span></span> <span data-ttu-id="5ae6e-325">必须先从文档提取保护状态，再用它设置 `sheet.protection.protected` 值。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-325">It is necessary to first fetch the protection status from the document and use it set the value of `sheet.protection.protected`.</span></span> <span data-ttu-id="5ae6e-326">只有这样，才能调用 `sheet.protection.protected`，而不导致异常抛出。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-326">Only then can `sheet.protection.protected` be called without causing an exception to be thrown.</span></span> <span data-ttu-id="5ae6e-327">此提取过程分为三步：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-327">This fetching process has three steps:</span></span>

   1. <span data-ttu-id="5ae6e-328">将命令排入队列，以加载（即提取）代码需要读取的属性。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-328">Queue a command to load (that is; fetch) the properties that your code needs to read.</span></span>

   2. <span data-ttu-id="5ae6e-329">调用上下文对象的 `sync`方法，从而向文档发送已排入队列的命令以供执行，并返回请求获取的信息。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-329">Call the context object's `sync` method to send the queued command to the document for execution and return the requested information.</span></span>

   3. <span data-ttu-id="5ae6e-330">由于 `sync` 是异步方法，因此请先确保它已完成，然后代码才能调用已提取的属性。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-330">Because the `sync` method is asynchronous, ensure that it has completed before your code calls the properties that were fetched.</span></span>

<span data-ttu-id="5ae6e-331">只要代码需要从 Office 文档*读取*信息，就必须完成这些步骤。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-331">These steps must be completed whenever your code needs to *read* information from the Office document.</span></span>

1. <span data-ttu-id="5ae6e-332">在 `toggleProtection` 函数中，将 `TODO2` 替换为以下代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-332">Within the `toggleProtection` function, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5ae6e-333">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-333">Note:</span></span>
   
   - <span data-ttu-id="5ae6e-p145">每个 Excel 对象都有 `load` 方法。 对于要在参数中读取的对象属性，将它们指定为逗号分隔名称字符串。 在此示例中，需要读取的属性为 `protection` 属性的子属性。 引用子属性的方法与在代码中的其他任何地方引用属性几乎完全一样，不同之处在于使用的是正斜杠（“/”）字符，而不是“.”字符。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p145">Every Excel object has a `load` method. You specify the properties of the object that you want to read in the parameter as a string of comma-delimited names. In this case, the property you need to read is a subproperty of the `protection` property. You reference the subproperty almost exactly as you would anywhere else in your code, with the exception that you use a forward slash ('/') character instead of a "." character.</span></span>

   - <span data-ttu-id="5ae6e-338">为了确保切换逻辑 `sheet.protection.protected` 只在 `sync` 完成后且 `sheet.protection.protected` 分配有从文档提取的正确值后才运行，（在下一步中）它会被移到 `then` 函数中，此函数在 `sync` 完成前不会运行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-338">To ensure that the toggle logic, which reads `sheet.protection.protected`, does not run until after the `sync` is complete and the `sheet.protection.protected` has been assigned the correct value that is fetched from the document, it will be moved (in the next step) into a `then` function that won't run until the `sync` has completed.</span></span> 

    ```js
    sheet.load('protection/protected');
    return context.sync()
        .then(
            function() {
                // TODO3: Move the queued toggle logic here.
            }
        )
        // TODO4: Move the final call of `context.sync` here and ensure that it
        //        does not run until the toggle logic has been queued.
    ``` 

2. <span data-ttu-id="5ae6e-p146">由于不能在同一取消分支代码路径中有两个 `return` 语句，因此请删除 `return context.sync();` 末尾的最后一行代码 `Excel.run`。 新的最后一行代码 `context.sync`将在后续步骤中添加。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p146">You can't have two `return` statements in the same unbranching code path, so delete the final line `return context.sync();` at the end of the `Excel.run`. You will add a new final `context.sync`, in a later step.</span></span>

3. <span data-ttu-id="5ae6e-341">剪切并粘贴 `toggleProtection` 函数中的 `if ... else` 结构，以替换 `TODO3`。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-341">Cut the `if ... else` structure in the `toggleProtection` function and paste it in place of `TODO3`.</span></span>

4. <span data-ttu-id="5ae6e-p147">将 `TODO4` 替换为以下代码。注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p147">Replace `TODO4` with the following code. Note:</span></span>

   - <span data-ttu-id="5ae6e-344">将 `sync` 方法传递到 `then` 函数可确保它不会在 `sheet.protection.unprotect()` 或 `sheet.protection.protect()` 已排入队列前运行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-344">Passing the `sync` method to a `then` function ensures that it does not run until either `sheet.protection.unprotect()` or `sheet.protection.protect()` has been queued.</span></span>

   - <span data-ttu-id="5ae6e-345">由于 `then` 方法调用传递给它的任何函数，并且也不想调用 `sync` 两次，因此请从 `context.sync` 末尾省略掉“()”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-345">The `then` method invokes whatever function is passed to it, and you don't want `sync` to be invoked twice, so leave off the "()" from the end of `context.sync`.</span></span>

    ```js
    .then(context.sync);
    ```

   <span data-ttu-id="5ae6e-346">完成后，整个函数应如下所示：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-346">When you are done, the entire function should look like the following:</span></span>

    ```js
    function toggleProtection(args) {
        Excel.run(function (context) {            
          var sheet = context.workbook.worksheets.getActiveWorksheet();          
          sheet.load('protection/protected');

          return context.sync()
              .then(
                  function() {
                    if (sheet.protection.protected) {
                        sheet.protection.unprotect();
                    } else {
                        sheet.protection.protect();
                    }
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
        args.completed();
    }
    ```

5. <span data-ttu-id="5ae6e-347">验证是否已保存了对项目所做的所有更改。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-347">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5ae6e-348">测试加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-348">Test the add-in</span></span>

1. <span data-ttu-id="5ae6e-349">关闭包括 Excel 在内的所有 Office 应用。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-349">Close all Office applications, including Excel.</span></span> 

2. <span data-ttu-id="5ae6e-p148">通过删除缓存文件夹内容，删除 Office 缓存。 若要完全清除主机中的旧版加载项，必须这样做。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p148">Delete the Office cache by deleting the contents of the cache folder. This is necessary to completely clear the old version of the add-in from the host.</span></span> 

    - <span data-ttu-id="5ae6e-352">对于 Windows：`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-352">For Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

    - <span data-ttu-id="5ae6e-353">对于 Mac：`~/Library/Containers/com.Microsoft.OsfWebHost/Data/`。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-353">For Mac: `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span> 
    
        > [!NOTE]
        > <span data-ttu-id="5ae6e-354">如果文件夹不存在，查看下列文件夹，如果找到，删除文件夹的内容：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-354">If that folder doesn't exist, check for the following folders and if found, delete the contents of the folder:</span></span>
        >    - <span data-ttu-id="5ae6e-355">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/`此位置`{host}`是 Office 主机（如 `Excel` ）</span><span class="sxs-lookup"><span data-stu-id="5ae6e-355">`~/Library/Containers/com.microsoft.{host}/Data/Library/Caches/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
        >    - <span data-ttu-id="5ae6e-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`此位置`{host}`是 Office 主机（如 `Excel` ）</span><span class="sxs-lookup"><span data-stu-id="5ae6e-356">`~/Library/Containers/com.microsoft.{host}/Data/Library/Application Support/Microsoft/Office/16.0/Wef/` where `{host}` is the Office host (e.g., `Excel`)</span></span>
        >    - `com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/`

3. <span data-ttu-id="5ae6e-357">如果本地 web 服务器已在运行，通过关闭节点命令窗口以停止。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-357">If the local web server is already running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="5ae6e-358">因为清单文件已更新，必须使用已更新的清单文件，重新旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-358">Because your manifest file has been updated, you must sideload your add-in again, using the updated manifest file.</span></span> <span data-ttu-id="5ae6e-359">启动本地 Web 服务器并旁加载你的加载项：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-359">Start the local web server and sideload your add-in:</span></span> 

    - <span data-ttu-id="5ae6e-360">若要在 Excel 中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-360">To test your add-in in Excel, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5ae6e-361">这将启动本地的 Web 服务器 (如果尚未运行的话), 并使用加载的加载项打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-361">This starts the local web server (if it's not already running) and opens Excel with your add-in loaded.</span></span>

        ```command&nbsp;line
        npm start
        ```

    - <span data-ttu-id="5ae6e-362">若要在 Excel 网页版中测试加载项，请在项目的根目录中运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-362">To test your add-in in Excel on the web, run the following command in the root directory of your project.</span></span> <span data-ttu-id="5ae6e-363">运行此命令时，本地 Web 服务器将启动（如果尚未运行）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-363">When you run this command, the local web server will start (if it's not already running).</span></span>

        ```command&nbsp;line
        npm run start:web
        ```

        <span data-ttu-id="5ae6e-364">若要使用加载项，请在 Excel 网页版中打开新的文档，并按照[在 Office 网页版中旁加载 Office 加载项](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web)中的说明操作，以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-364">To use your add-in, open a new document in Excel on the web and then sideload your add-in by following the instructions in [Sideload Office Add-ins in Office on the web](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-on-the-web).</span></span>

5. <span data-ttu-id="5ae6e-365">在 Excel 的 **Home** 选项卡上，选择“**切换工作表保护**”按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-365">On the **Home** tab in Excel, choose the **Toggle Worksheet Protection** button.</span></span> <span data-ttu-id="5ae6e-366">请注意，功能区上的大部分控件都处于禁用状态（灰显），如下面的屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-366">Note that most of the controls on the ribbon are disabled (and visually grayed-out) as seen in the following screenshot.</span></span> 

    ![Excel 教程 - 在功能区上启用工作表保护](../images/excel-tutorial-ribbon-with-protection-on-2.png)

6. <span data-ttu-id="5ae6e-368">选择要更改其内容的单元格。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-368">Choose a cell as you would if you wanted to change its content.</span></span> <span data-ttu-id="5ae6e-369">Excel 将显示一条错误消息，指示工作表受保护。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-369">Excel displays an error message indicating that the worksheet is protected.</span></span>

7. <span data-ttu-id="5ae6e-370">再次选择“**切换工作表保护**”按钮，此时控件重新启用，可以再次更改单元格值。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-370">Choose the **Toggle Worksheet Protection** button again, and the controls are reenabled, and you can change cell values again.</span></span>

## <a name="open-a-dialog"></a><span data-ttu-id="5ae6e-371">打开对话框</span><span class="sxs-lookup"><span data-stu-id="5ae6e-371">Open a dialog</span></span>

<span data-ttu-id="5ae6e-p154">本教程的最后一步是，在加载项中打开对话框，将消息从对话框进程传递到任务窗格进程，再关闭对话框。 Office 加载项对话框是*非模式*窗口。也就是说，用户可以继续与主机 Office 应用中的文档，以及与任务窗格中的主机页进行交互。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p154">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog. Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

### <a name="create-the-dialog-page"></a><span data-ttu-id="5ae6e-374">创建对话框页面</span><span class="sxs-lookup"><span data-stu-id="5ae6e-374">Create the dialog page</span></span>

1. <span data-ttu-id="5ae6e-375">在项目根目录下的 **./src** 文件夹中，新建文件夹“**dialogs**”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-375">In the **./src** folder that's located at the root of the project, create a new folder named **dialogs**.</span></span>

2. <span data-ttu-id="5ae6e-376">在 **./src/dialogs** 文件夹中，新建文件“**popup.html**”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-376">In the **./src/dialogs** folder, create new file named **popup.html**.</span></span>

3. <span data-ttu-id="5ae6e-377">将下面的标记添加到 **popup.html** 中。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-377">Add the following markup to **popup.html**.</span></span> <span data-ttu-id="5ae6e-378">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-378">Note:</span></span>

   - <span data-ttu-id="5ae6e-379">此页面包含可供用户输入用户名的 `<input>`，并包含将用户名发送到任务窗格中用户名显示页面的按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-379">The page has a `<input>` where the user will enter their name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>

   - <span data-ttu-id="5ae6e-380">此标记加载在后续步骤中创建的 **popup.js** 脚本。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-380">The markup loads a script named **popup.js** that you will create in a later step.</span></span>

   - <span data-ttu-id="5ae6e-381">此标记还加载 Office.JS 库，因为 **popup.js** 将使用它们。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-381">It also loads the Office.js library because it will be used in **popup.js**.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">

            <!-- For more information on Office UI Fabric, visit https://developer.microsoft.com/fabric. -->
            <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
            <script type="text/javascript" src="popup.js"></script>

        </head>
        <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
            <input id="name-box" type="text"/><br/><br/>
            <button id="ok-button" class="ms-Button">OK</button>
        </body>
    </html>
    ```

4. <span data-ttu-id="5ae6e-382">在 **./src/dialogs** 文件夹中，新建文件“**popup.js**”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-382">In the **./src/dialogs** folder, create new file named **popup.js**.</span></span>

5. <span data-ttu-id="5ae6e-383">将下面的代码添加到 **popup.js** 中。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-383">Add the following code to **popup.js**.</span></span> <span data-ttu-id="5ae6e-384">关于此代码，请注意以下几点：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-384">Note the following about this code:</span></span>

   - <span data-ttu-id="5ae6e-385">*每个调用 Office.JS 库中的 API 的页面均必须首先确保该库已完成初始化。*</span><span class="sxs-lookup"><span data-stu-id="5ae6e-385">*Every page that calls APIs in the Office.js library must first ensure that the library is fully initialized.*</span></span> <span data-ttu-id="5ae6e-386">执行此操作的最佳方法是调用 `Office.onReady()` 方法。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-386">The best way to do that is to call the `Office.onReady()` method.</span></span> <span data-ttu-id="5ae6e-387">如果加载项具有其自己的初始化任务，则代码应位于链接至 `Office.onReady()` 调用的 `then()` 方法中。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-387">If your add-in has its own initialization tasks, the code should go in a `then()` method that is chained to the call of `Office.onReady()`.</span></span> <span data-ttu-id="5ae6e-388">必须在调用 Office.js 之前运行 `Office.onReady()` 调用；因此，作业位于由页面加载的脚本文件中，如同本示例中一样。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-388">The call of `Office.onReady()` must run before any calls to Office.js; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>

    ```js
    (function () {
    "use strict";

        Office.onReady()
            .then(function() {

                // TODO1: Assign handler to the OK button.

            });

        // TODO2: Create the OK button handler

    }());
    ```

6. <span data-ttu-id="5ae6e-p158">将 `TODO1` 替换为下列代码。 将在下一步中创建 `sendStringToParentPage` 函数。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p158">Replace `TODO1` with the following code. You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    document.getElementById("ok-button").onclick = sendStringToParentPage;
    ```

7. <span data-ttu-id="5ae6e-p159">将 `TODO2` 替换为以下代码。 `messageParent` 方法将它的参数传递到父页面（在此示例中，为任务窗格中的页面）。 参数可以是布尔值或字符串，其中包含可串行化为字符串的任何内容（如 XML 或 JSON）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p159">Replace `TODO2` with the following code. The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane. The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span>

    ```js
    function sendStringToParentPage() {
        var userName = document.getElementById("name-box").value;
        Office.context.ui.messageParent(userName);
    }
    ```

> [!NOTE]
> <span data-ttu-id="5ae6e-394">**popup.html** 文件及其加载的 **popup.js** 文件在完全独立于加载项任务窗格的 Microsoft Edge 或 Internet Explorer 11 进程中运行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-394">The **popup.html** file, and the **popup.js** file that it loads, run in an entirely separate Microsoft Edge or Internet Explorer 11 process from the add-in's task pane.</span></span> <span data-ttu-id="5ae6e-395">如果将 **popup.js** 转换为与 **app.js** 文件相同的 **bundle.js** 文件，加载项必须加载 **bundle.js** 文件的两个副本，这就违背了绑定目的。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-395">If **popup.js** was transpiled into the same **bundle.js** file as the **app.js** file, then the add-in would have to load two copies of the **bundle.js** file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="5ae6e-396">所以此加载项根本不会转换 **popup.js**。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-396">Therefore, this add-in does not transpile the **popup.js** file at all.</span></span>

### <a name="update-webpack-config-settings"></a><span data-ttu-id="5ae6e-397">更新 webpack 配置设置</span><span class="sxs-lookup"><span data-stu-id="5ae6e-397">Update webpack config settings</span></span>

<span data-ttu-id="5ae6e-398">打开项目根目录中的 **webpack.config.js** 文件，并完成以下步骤。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-398">Open the file **webpack.config.js** in the root directory of the project and complete the following steps.</span></span>

1. <span data-ttu-id="5ae6e-399">在 `config` 对象内找到 `entry` 对象并为 `popup` 添加新条目。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-399">Locate the `entry` object within the `config` object and add a new entry for `popup`.</span></span>

    ```js
    popup: "./src/dialogs/popup.js"
    ```

    <span data-ttu-id="5ae6e-400">完成此操作之后，新的 `entry` 对象将与此类似：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-400">After you've done this, the new `entry` object will look like this:</span></span>

    ```js
    entry: {
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      popup: "./src/dialogs/popup.js"
    },
    ```
  
2. <span data-ttu-id="5ae6e-401">在 `config` 对象中找到 `plugins` 数组，并添加下列对象至数组的结尾。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-401">Locate the `plugins` array within the `config` object and add the following object to the end of that array.</span></span>

    ```js
    new HtmlWebpackPlugin({
      filename: "popup.html",
      template: "./src/dialogs/popup.html",
      chunks: ["polyfill", "popup"]
    })
    ```

    <span data-ttu-id="5ae6e-402">完成此操作之后，新的 `plugins` 数组将与此类似：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-402">After you've done this, the new `plugins` array will look like this:</span></span>

    ```js
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin([
      {
        to: "taskpane.css",
        from: "./src/taskpane/taskpane.css"
      }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "popup.html",
        template: "./src/dialogs/popup.html",
        chunks: ["polyfill", "popup"]
      })
    ],
    ```

3. <span data-ttu-id="5ae6e-403">如果本地 web 服务器正在运行，通过关闭节点命令窗口以停止。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-403">If the local web server is running, stop it by closing the node command window.</span></span>

4. <span data-ttu-id="5ae6e-404">运行以下命令以重建项目。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-404">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

### <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="5ae6e-405">从任务窗格打开对话框</span><span class="sxs-lookup"><span data-stu-id="5ae6e-405">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="5ae6e-406">打开 **./src/taskpane/taskpane.html** 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-406">Open the file **./src/taskpane/taskpane.html**.</span></span>

2. <span data-ttu-id="5ae6e-407">定位 `<button>` 按钮的 `freeze-header` 元素，并在行后添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-407">Locate the `<button>` element for the `freeze-header` button, and add the following markup after that line:</span></span>

    ```html
    <button class="ms-Button" id="open-dialog">Open Dialog</button><br/><br/>
    ```

3. <span data-ttu-id="5ae6e-408">对话框会提示用户输入用户名，并将用户名传递到任务窗格。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-408">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="5ae6e-409">任务窗格将在标签中显示用户名。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-409">The task pane will display it in a label.</span></span> <span data-ttu-id="5ae6e-410">紧接着刚添加的 `button`，添加下列标记：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-410">Immediately after the `button` that you just added, add the following markup:</span></span>

    ```html
    <label id="user-name"></label><br/><br/>
    ```

4. <span data-ttu-id="5ae6e-411">打开 **./src/taskpane/taskpane.js** 文件。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-411">Open the file **./src/taskpane/taskpane.js**.</span></span>

5. <span data-ttu-id="5ae6e-412">在 `Office.onReady` 方法调用中，定位分配点击事件至 `freeze-header` 按钮的行，并在行后添加虾类代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-412">Within the `Office.onReady` method call, locate the line that assigns a click handler to the `freeze-header` button, and add the following code after that line.</span></span> <span data-ttu-id="5ae6e-413">将在后续步骤中创建 `openDialog` 方法。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-413">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    document.getElementById("open-dialog").onclick = openDialog;
    ```

6. <span data-ttu-id="5ae6e-414">添加下列声明至文件结尾。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-414">Add the following declaration to the end of the file.</span></span> <span data-ttu-id="5ae6e-415">此变量用于保留父页面执行上下文中的对象，以用作对话框页面执行上下文的中间对象。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-415">This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    var dialog = null;
    ```

7. <span data-ttu-id="5ae6e-416">添加下列函数至文件结尾（`dialog` 声明后）。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-416">Add the following function to the end of the file (after the declaration of `dialog`).</span></span> <span data-ttu-id="5ae6e-417">关于此代码，请务必注意它*不*包含的内容，即不含 `Excel.run` 调用。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-417">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="5ae6e-418">这是因为对话框打开 API 跨所有 Office 主机共享，所以它属于 Office JavaScript 公用 API，而不属于 Excel 专用 API。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-418">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Common API that opens a dialog
    }
    ```

8. <span data-ttu-id="5ae6e-p165">将 `TODO1` 替换为以下代码。注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p165">Replace `TODO1` with the following code. Note:</span></span>

   - <span data-ttu-id="5ae6e-421">`displayDialogAsync` 方法在屏幕中央打开对话框。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-421">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>

   - <span data-ttu-id="5ae6e-422">第一个参数是要打开的页面 URL。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-422">The first parameter is the URL of the page to open.</span></span>

   - <span data-ttu-id="5ae6e-p166">第二个参数用于传递选项。`height` 和 `width` 是 Office 应用程序窗口大小百分比。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p166">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span>

    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},

        // TODO2: Add callback parameter.
    );
    ```

### <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="5ae6e-425">处理对话框发送的消息并关闭对话框</span><span class="sxs-lookup"><span data-stu-id="5ae6e-425">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="5ae6e-426">在 **./src/taskpane/taskpane.js** 文件的 `openDialog` 函数内，将 `TODO2` 替换为下列代码。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-426">Within the `openDialog` function in the file **./src/taskpane/taskpane.js**, replace `TODO2` with the following code.</span></span> <span data-ttu-id="5ae6e-427">注意：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-427">Note:</span></span>

   - <span data-ttu-id="5ae6e-428">回调在对话框成功打开后和用户在对话框中执行任何操作前立即执行。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-428">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>

   - <span data-ttu-id="5ae6e-429">`result.value` 对象用作父页面执行上下文和对话框页面执行上下文的中间对象。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-429">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>

   - <span data-ttu-id="5ae6e-p168">`processMessage` 函数将在后续步骤中创建。 此处理程序将处理通过 `messageParent` 函数调用从对话框页面发送的任何值。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-p168">The `processMessage` function will be created in a later step. This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="5ae6e-432">在 `openDialog` 函数后面添加以下函数。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-432">Add the following function after the `openDialog` function.</span></span>

    ```js
    function processMessage(arg) {
        document.getElementById("user-name").innerHTML = arg.message;
        dialog.close();
    }
    ```

3. <span data-ttu-id="5ae6e-433">验证是否已保存了对项目所做的所有更改。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-433">Verify that you've saved all of the changes you've made to the project.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="5ae6e-434">测试加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-434">Test the add-in</span></span>

1. [!include[Start server and sideload add-in instructions](../includes/tutorial-excel-start-server.md)]

2. <span data-ttu-id="5ae6e-435">如果加载项任务窗格已在 Excel 中打开，转至 **Home** 选项卡并选中功能区中的“**显示任务窗格**”以打开。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-435">If the add-in task pane isn't already open in Excel, go to the **Home** tab and choose the **Show Taskpane** button in the ribbon to open it.</span></span>

3. <span data-ttu-id="5ae6e-436">选择任务窗格中的“打开对话框”\*\*\*\* 按钮。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-436">Choose the **Open Dialog** button in the task pane.</span></span>

4. <span data-ttu-id="5ae6e-437">对话框打开后，拖动它并重设大小。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-437">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="5ae6e-438">请注意，你可以与工作表进行交互并按任务窗格上的其他按钮，但无法从同一任务窗格页面启动第二个对话框。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-438">Note that you can interact with the worksheet and press other buttons on the task pane, but you cannot launch a second dialog from the same task pane page.</span></span>

5. <span data-ttu-id="5ae6e-439">在对话框中，输入名称并选择“**确定**”。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-439">In the dialog, enter a name and choose the **OK** button.</span></span> <span data-ttu-id="5ae6e-440">此时，用户名显示在任务窗格上，且对话框关闭。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-440">The name appears on the task pane and the dialog closes.</span></span>

6. <span data-ttu-id="5ae6e-441">（可选）注释掉 `processMessage` 函数中的代码行 `dialog.close();`。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-441">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="5ae6e-442">然后，重复执行此部分的步骤。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-442">Then repeat the steps of this section.</span></span> <span data-ttu-id="5ae6e-443">这样一来，对话框便会继续处于打开状态，可供用户更改用户名。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-443">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="5ae6e-444">按右上角的“X”\*\*\*\* 按钮，可手动关闭对话框。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-444">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Excel 教程 - 对话框](../images/excel-tutorial-dialog-open-2.png)

## <a name="next-steps"></a><span data-ttu-id="5ae6e-446">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5ae6e-446">Next steps</span></span>

<span data-ttu-id="5ae6e-447">在本教程中，你已创建与 Excel 工作簿中的表格、图表、工作表和对话框进行交互的 Excel 任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="5ae6e-447">In this tutorial, you've created an Excel task pane add-in that interacts with tables, charts, worksheets, and dialogs in an Excel workbook.</span></span> <span data-ttu-id="5ae6e-448">若要了解有关构建 Excel 加载项的详细信息，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="5ae6e-448">To learn more about building Excel add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5ae6e-449">Excel 加载项概述</span><span class="sxs-lookup"><span data-stu-id="5ae6e-449">Excel add-ins overview</span></span>](../excel/excel-add-ins-overview.md)

## <a name="see-also"></a><span data-ttu-id="5ae6e-450">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5ae6e-450">See also</span></span>

* [<span data-ttu-id="5ae6e-451">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="5ae6e-451">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="5ae6e-452">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-452">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="5ae6e-453">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="5ae6e-453">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="5ae6e-454">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="5ae6e-454">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
