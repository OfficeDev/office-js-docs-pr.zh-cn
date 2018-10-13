# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="d564f-101">教程：在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="d564f-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="d564f-102">简介</span><span class="sxs-lookup"><span data-stu-id="d564f-102">Introduction</span></span>

<span data-ttu-id="d564f-p101">自定义函数使你可以通过在 JavaScript 中定义这些函数作为加载项的一部分，将新函数添加到 Excel。然后，用户可以像使用 Excel 中的其他本机函数一样访问自定义函数，如 `SUM()`。你可以创建自定义函数，以执行简单任务，例如自定义计算或更复杂的任务，如将来自 Web 的实时数据l以流式处理方法插入工作表。</span><span class="sxs-lookup"><span data-stu-id="d564f-p101">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="d564f-106">在此教程中，你将：</span><span class="sxs-lookup"><span data-stu-id="d564f-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="d564f-107">使用 Yo Office 生成器创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="d564f-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="d564f-108">使用预建的自定义函数执行简单计算</span><span class="sxs-lookup"><span data-stu-id="d564f-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="d564f-109">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="d564f-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="d564f-110">创建以流式处理方法处理来自 Web 的实时数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="d564f-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="d564f-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="d564f-111">Prerequisites</span></span>

* [<span data-ttu-id="d564f-112">Node.js 和 npm</span><span class="sxs-lookup"><span data-stu-id="d564f-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="d564f-113">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="d564f-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="d564f-p102"> [Yeoman](http://yeoman.io/) 和 [Yo Office 生成器](https://www.npmjs.com/package/generator-office)的最新版本。若要全局安装这些工具，请通过命令提示符处运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="d564f-p102">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="d564f-116">Excel for Windows（内部版本号 10827 或更高版本）或 Excel Online</span><span class="sxs-lookup"><span data-stu-id="d564f-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* [<span data-ttu-id="d564f-117">加入 Office 预览体验计划</span><span class="sxs-lookup"><span data-stu-id="d564f-117">Join the Office Insider program</span></span>](https://products.office.com/office-insider)

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="d564f-118">创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="d564f-118">Create a custom functions project</span></span>

<span data-ttu-id="d564f-119">你将通过使用 Yo Office 生成器创建自定义函数项目所需的文件来开始此教程。</span><span class="sxs-lookup"><span data-stu-id="d564f-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="d564f-120">运行下面的命令，然后回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="d564f-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="d564f-121">选择项目类型： `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="d564f-121">Choose a project type  </span></span>
    * <span data-ttu-id="d564f-122">选择脚本类型： `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="d564f-122">Choose a script type  </span></span>
    * <span data-ttu-id="d564f-123">要如何命名加载项?</span><span class="sxs-lookup"><span data-stu-id="d564f-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![自定义函数的 Yo Office bash 提示](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="d564f-125">完成向导后，生成器会创建项目文件，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="d564f-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="d564f-126">项目文件来自[ Excel 自定义函数](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub 存储库。</span><span class="sxs-lookup"><span data-stu-id="d564f-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="d564f-127">导航到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="d564f-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="d564f-128">启动本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="d564f-128">Start the local web server.</span></span>

    * <span data-ttu-id="d564f-129">如果你将使用 Excel for Windows 测试自定义函数，请运行以下命令以启动本地 Web 服务器、启动 Excel 和旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="d564f-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="d564f-130">如果你将使用 Excel Online 测试自定义函数，请运行以下命令以启动本地 Web 服务器：</span><span class="sxs-lookup"><span data-stu-id="d564f-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="d564f-131">试用预建的自定义函数</span><span class="sxs-lookup"><span data-stu-id="d564f-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="d564f-132">使用 Yo Office 生成器创建的自定义函数项目包含一些在 **src/customfunction.js** 文件内定义的预建的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d564f-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="d564f-133">在项目根目录中的 **manifest.xml** 文件指定所有自定义函数属于 `CONTOSO` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="d564f-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="d564f-134">在你可以使用任何预建的自定义函数之前，必须在 Excel 中注册自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="d564f-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="d564f-135">通过完成适用于你将在本教程中使用的平台的步骤来实现。</span><span class="sxs-lookup"><span data-stu-id="d564f-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="d564f-136">如果你将使用 Excel for Windows 测试自定义函数：</span><span class="sxs-lookup"><span data-stu-id="d564f-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="d564f-137">在 Excel 中，依次选择**插入**选项卡和位于**我的加载项**右侧的向下箭头。带突出显示“我的加载项”箭头的 Excel for Windows 中的 ![“插入”功能区](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="d564f-138">在可用加载项列表中，找到**开发人员加载项**一节，然后选择 **Excel 自定义函数**加载项对其进行注册。</span><span class="sxs-lookup"><span data-stu-id="d564f-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="d564f-139">![“我的加载项”列表中突出显示的带“Excel 自定义函数”加载项的 Excel for Windows 中的“插入”功能区](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="d564f-140">如果你将使用 Excel Online 测试自定义函数：</span><span class="sxs-lookup"><span data-stu-id="d564f-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="d564f-141">在 Excel Online 中，依次选择**插入**选项卡和**加载项**。带突出显示“我的加载项”图标的 Excel Online 中的 ![“插入”功能区](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="d564f-142">依次选择**管理我的加载项**和**上传我的加载项**。</span><span class="sxs-lookup"><span data-stu-id="d564f-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="d564f-143">选择**浏览...** 并导航到 Yo Office 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="d564f-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="d564f-144">依次选择 **manifest.xml** 文件、**打开**和**上传**。</span><span class="sxs-lookup"><span data-stu-id="d564f-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="d564f-145">此时，在 Excel 内加载了项目中的预建自定义函数并可用。</span><span class="sxs-lookup"><span data-stu-id="d564f-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="d564f-146">通过在 Excel 中完成以下步骤来试用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="d564f-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="d564f-147">在单元格内，键入 **=CONTOSO**。</span><span class="sxs-lookup"><span data-stu-id="d564f-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="d564f-148">请注意，自动完成菜单显示了 `CONTOSO` 命名空间中的所有函数列表。</span><span class="sxs-lookup"><span data-stu-id="d564f-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="d564f-149">通过在单元格中指定下列值并按输入与作为输入参数的数字 `10` 和 `200` 一起来运行 `CONTOSO.ADD` 函数：</span><span class="sxs-lookup"><span data-stu-id="d564f-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="d564f-150">`ADD` 自定义函数计算指定作为输入参数的两个数字的总和。</span><span class="sxs-lookup"><span data-stu-id="d564f-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="d564f-151">在按 enter 后，键入 `=CONTOSO.ADD(10,200)` 应在单元格中生成结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="d564f-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="d564f-152">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="d564f-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="d564f-153">要是需要一个可以通过 API 请求股票价格并在工作表的单元格中显示结果的函数会怎么样呢？</span><span class="sxs-lookup"><span data-stu-id="d564f-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="d564f-154">设计自定义函数，以便可以轻松地从 Web 异步请求数据。</span><span class="sxs-lookup"><span data-stu-id="d564f-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="d564f-155">完成以下步骤以创建一个名为 `stockPrice` 的自定义函数，可接受股票代码（例如，**MSFT**），并返回该股票的价格。</span><span class="sxs-lookup"><span data-stu-id="d564f-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="d564f-156">此自定义函数使用 IEX 交易 API，这是免费的，不需要身份验证。</span><span class="sxs-lookup"><span data-stu-id="d564f-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="d564f-157">在 Yo Office 生成器创建的**股票代码**项目中，查找文件 **src/customfunctions.js** 并在代码编辑器中将其打开。</span><span class="sxs-lookup"><span data-stu-id="d564f-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="d564f-158">将以下代码添加到 **customfunctions.js**，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="d564f-158">Add the following code to **home.js** and save the file.</span></span>

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. <span data-ttu-id="d564f-159">在 Excel 可以使此新函数可供最终用户使用之前，必须指定介绍此函数的元数据。</span><span class="sxs-lookup"><span data-stu-id="d564f-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="d564f-160">在 Yo Office 生成器创建的**股票代码**项目中，查找文件 **config/customfunctions.json** 并在代码编辑器中将其打开。</span><span class="sxs-lookup"><span data-stu-id="d564f-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="d564f-161">将以下对象添加到 **config/customfunctions.json** 文件内的 `functions` 数组，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="d564f-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="d564f-162">此 JSON 介绍了 `stockPrice` 函数。</span><span class="sxs-lookup"><span data-stu-id="d564f-162">This JSON describes the `stockPrice` function.</span></span>

    ```json
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. <span data-ttu-id="d564f-163">必须在 Excel 中重新注册加载项，以便新函数可供最终用户使用。</span><span class="sxs-lookup"><span data-stu-id="d564f-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="d564f-164">完成适用于你将在本教程中使用的平台的步骤。</span><span class="sxs-lookup"><span data-stu-id="d564f-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="d564f-165">如果正在使用 Excel for Windows：</span><span class="sxs-lookup"><span data-stu-id="d564f-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="d564f-166">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="d564f-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="d564f-167">在 Excel 中，依次选择**插入**选项卡和位于**我的加载项**右侧的向下箭头。带突出显示“我的加载项”箭头的 Excel for Windows 中的 ![“插入”功能区](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="d564f-168">在可用加载项列表中，找到**开发人员加载项**一节，然后选择 **Excel 自定义函数**加载项对其进行注册。</span><span class="sxs-lookup"><span data-stu-id="d564f-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="d564f-169">![“我的加载项”列表中突出显示的带“Excel 自定义函数”加载项的 Excel for Windows 中的“插入”功能区](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="d564f-170">如果正在使用 Excel Online：</span><span class="sxs-lookup"><span data-stu-id="d564f-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="d564f-171">在 Excel Online 中，依次选择**插入**选项卡和**加载项**。带突出显示“我的加载项”图标的 Excel Online 中的 ![“插入”功能区](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="d564f-172">依次选择**管理我的加载项**和**上传我的加载项**。</span><span class="sxs-lookup"><span data-stu-id="d564f-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="d564f-173">选择**浏览...** 并导航到 Yo Office 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="d564f-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="d564f-174">依次选择 **manifest.xml** 文件、**打开**和**上传**。</span><span class="sxs-lookup"><span data-stu-id="d564f-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="d564f-175">现在，让我们试用新函数。</span><span class="sxs-lookup"><span data-stu-id="d564f-175">Now, let's try out the new function.</span></span> <span data-ttu-id="d564f-176">在单元格 **B1**中，键入文本 `=CONTOSO.STOCKPRICE("MSFT")` 并按 enter。</span><span class="sxs-lookup"><span data-stu-id="d564f-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="d564f-177">应该看到单元格 **B1** 中的结果是针对 Microsoft 股票中的一支的当前股票价格。</span><span class="sxs-lookup"><span data-stu-id="d564f-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="d564f-178">创建流式异步自定义函数</span><span class="sxs-lookup"><span data-stu-id="d564f-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="d564f-179">刚创建的 `stockPrice` 函数返回特定时刻的股票价格，但股票价格总是在变化。</span><span class="sxs-lookup"><span data-stu-id="d564f-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="d564f-180">让我们一起创建一个通过 API 以流式处理传输数据以获取股票价格实时更新的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d564f-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="d564f-181">完成以下步骤以创建一个名为 `stockPriceStream` 自定义函数，请求指定每 1000 毫秒的价格股票（前提是在上一个请求已完成）。</span><span class="sxs-lookup"><span data-stu-id="d564f-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="d564f-182">在进行的初始请求时，你可能会看到占位符值 **#GETTING_DATA** 单元格，函数正在这里调用。</span><span class="sxs-lookup"><span data-stu-id="d564f-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="d564f-183">该函数返回一个值后，单元格中的那个值将替换 **#GETTING_DATA**。</span><span class="sxs-lookup"><span data-stu-id="d564f-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="d564f-184">在 Yo Office 生成器创建的**股票代码**项目中，将以下代码添加到 **src/customfunctions.js** 并保存文件。</span><span class="sxs-lookup"><span data-stu-id="d564f-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. <span data-ttu-id="d564f-185">在 Excel 可以使此新函数可供最终用户使用之前，必须指定介绍此函数的元数据。</span><span class="sxs-lookup"><span data-stu-id="d564f-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="d564f-186">在 Yo Office 生成器创建的**股票代码**项目中，将以下对象添加到 **config/customfunctions.json** 文件内的 `functions` 数组，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="d564f-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="d564f-187">此 JSON 介绍了 `stockPriceStream` 函数。</span><span class="sxs-lookup"><span data-stu-id="d564f-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="d564f-188">对于任何流式函数，`stream` 属性和 `cancelable` 属性必须设置为 `options` 对象内的 `true`，如此代码示例中所示。</span><span class="sxs-lookup"><span data-stu-id="d564f-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

3. <span data-ttu-id="d564f-189">必须在 Excel 中重新注册加载项，以便新函数可供最终用户使用。</span><span class="sxs-lookup"><span data-stu-id="d564f-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="d564f-190">完成适用于你将在本教程中使用的平台的步骤。</span><span class="sxs-lookup"><span data-stu-id="d564f-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="d564f-191">如果正在使用 Excel for Windows：</span><span class="sxs-lookup"><span data-stu-id="d564f-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="d564f-192">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="d564f-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="d564f-193">在 Excel 中，依次选择**插入**选项卡和位于**我的加载项**右侧的向下箭头。带突出显示“我的加载项”箭头的 Excel for Windows 中的 ![“插入”功能区](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="d564f-194">在可用加载项列表中，找到**开发人员加载项**一节，然后选择 **Excel 自定义函数**加载项对其进行注册。</span><span class="sxs-lookup"><span data-stu-id="d564f-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="d564f-195">![“我的加载项”列表中突出显示的带“Excel 自定义函数”加载项的 Excel for Windows 中的“插入”功能区](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="d564f-196">如果正在使用 Excel Online：</span><span class="sxs-lookup"><span data-stu-id="d564f-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="d564f-197">在 Excel Online 中，依次选择**插入**选项卡和**加载项**。带突出显示“我的加载项”图标的 Excel Online 中的 ![“插入”功能区](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="d564f-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="d564f-198">依次选择**管理我的加载项**和**上传我的加载项**。</span><span class="sxs-lookup"><span data-stu-id="d564f-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="d564f-199">选择**浏览...** 并导航到 Yo Office 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="d564f-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="d564f-200">依次选择 **manifest.xml** 文件、**打开**和**上传**。</span><span class="sxs-lookup"><span data-stu-id="d564f-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="d564f-201">现在，让我们试用新函数。</span><span class="sxs-lookup"><span data-stu-id="d564f-201">Now, let's try out the new function.</span></span> <span data-ttu-id="d564f-202">在单元格 **C1**中，键入文本 `=CONTOSO.STOCKPRICESTREAM("MSFT")` 并按 enter。</span><span class="sxs-lookup"><span data-stu-id="d564f-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="d564f-203">假设股票市场处于打开状态，您应看到 **C1** 单元格中的结果不断地更新，以反映 Microsoft 股票中的一支的实时价格。</span><span class="sxs-lookup"><span data-stu-id="d564f-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d564f-204">后续步骤</span><span class="sxs-lookup"><span data-stu-id="d564f-204">Next steps</span></span>

<span data-ttu-id="d564f-205">在本教程中，你创建了一个新的自定义函数项目，尝试了预建的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 以流式处理方法传输实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="d564f-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="d564f-206">要了解在 Excel 中自定义函数的详细信息，请继续查看以下文章：</span><span class="sxs-lookup"><span data-stu-id="d564f-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="d564f-207">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="d564f-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="d564f-208">法律信息</span><span class="sxs-lookup"><span data-stu-id="d564f-208">Legal Information</span></span>

<span data-ttu-id="d564f-209">由 [IEX](https://iextrading.com/developer/) 免费提供的数据。</span><span class="sxs-lookup"><span data-stu-id="d564f-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="d564f-210">查看 [IEX 的使用条款](https://iextrading.com/api-exhibit-a/)。</span><span class="sxs-lookup"><span data-stu-id="d564f-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="d564f-211">本教程中 Microsoft 使用的 IEX API 只用于教学目的。</span><span class="sxs-lookup"><span data-stu-id="d564f-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
