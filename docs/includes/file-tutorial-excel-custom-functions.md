# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="1f6e0-101">教程：在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="1f6e0-101">Tutorial: Create custom functions in Excel</span></span>

## <a name="introduction"></a><span data-ttu-id="1f6e0-102">简介</span><span class="sxs-lookup"><span data-stu-id="1f6e0-102">Introduction</span></span>

<span data-ttu-id="1f6e0-103">用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="1f6e0-104">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="1f6e0-105">可以创建自定义函数，以执行简单的任务（如自定义计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="1f6e0-106">将在本教程中执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-106">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="1f6e0-107">通过使用 Yo Office 生成器创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="1f6e0-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="1f6e0-108">使用预生成的自定义函数来执行简单计算</span><span class="sxs-lookup"><span data-stu-id="1f6e0-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="1f6e0-109">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="1f6e0-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="1f6e0-110">创建从 Web 传送实时数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="1f6e0-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="1f6e0-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="1f6e0-111">Prerequisites</span></span>

* <span data-ttu-id="1f6e0-112">[Node.js](https://nodejs.org/en/)（版本 8.0.0 或更高版本）</span><span class="sxs-lookup"><span data-stu-id="1f6e0-112">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="1f6e0-113">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="1f6e0-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="1f6e0-114">最新版本的 [Yeoman](https://yeoman.io/) 和[适用于 Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-114">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command from the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="1f6e0-115">即便先前已安装 Yeoman 生成器，我们仍建议将包更新至最新的 npm 版本。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-115">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="1f6e0-116">Excel for Windows（64 位，版本 1810 或更高版本）或 Excel Online</span><span class="sxs-lookup"><span data-stu-id="1f6e0-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="1f6e0-117">加入 [Office 预览体验计划](https://products.office.com/office-insider)（**预览体验成员**级别 - 以前称为“预览体验成员 - 快”）</span><span class="sxs-lookup"><span data-stu-id="1f6e0-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="1f6e0-118">创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="1f6e0-118">Create a custom functions project</span></span>

 <span data-ttu-id="1f6e0-119">首先，使用 Yeoman 生成器创建自定义函数项目。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-119">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="1f6e0-120">这将为你的项目设置开始对自定义函数进行编码所需的正确文件夹结构、源文件和依存关系。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-120">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="1f6e0-121">运行下面的命令，再回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-121">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    * <span data-ttu-id="1f6e0-122">选择项目类型：`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="1f6e0-122">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    * <span data-ttu-id="1f6e0-123">选择脚本类型：`JavaScript`</span><span class="sxs-lookup"><span data-stu-id="1f6e0-123">Choose a script type: `JavaScript`</span></span>

    * <span data-ttu-id="1f6e0-124">要如何命名加载项？</span><span class="sxs-lookup"><span data-stu-id="1f6e0-124">What do you want to name your add-in?</span></span> `stock-ticker`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="1f6e0-126">Yeoman 生成器将创建项目文件并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-126">The generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="1f6e0-127">项目文件来自 [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub 存储库。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-127">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="1f6e0-128">转到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-128">Go to the project folder.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="1f6e0-129">信任运行此项目所需的自签名证书。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="1f6e0-130">有关适用于 Windows 或 Mac 的详细说明，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="1f6e0-131">生成项目。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-131">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="1f6e0-132">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-132">Start the local web server, which runs in Node.js.</span></span>

    * <span data-ttu-id="1f6e0-133">如果将使用 Excel for Windows 测试自定义函数，请运行以下命令来启动本地 Web 服务器，启动 Excel，并旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-133">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="1f6e0-134">运行此命令之后，命令提示符将显示与已完成项目相关的详细信息，打开的另一个 npm 窗口将显示与版本相关的详细信息，并且 Excel 将启动且加载项将会加载。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-134">After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded.</span></span> <span data-ttu-id="1f6e0-135">如果加载项未加载，请检查是否已正确完成步骤 3。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-135">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    * <span data-ttu-id="1f6e0-136">如果要使用 Excel Online 测试自定义函数，请运行以下命令来启动本地 Web 服务器：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-136">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="1f6e0-137">运行此命令之后，打开的另一个窗口将向你显示与版本相关的详细信息。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-137">After running this command, another window will open showing you the details of the build.</span></span> <span data-ttu-id="1f6e0-138">要使用函数，请在 Office Online 中打开一个新的工作簿。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-138">To use your functions, open a new workbook in Office Online.</span></span>

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="1f6e0-139">尝试预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="1f6e0-139">Try out a prebuilt custom function</span></span>

<span data-ttu-id="1f6e0-140">使用 Yeoman 生成器创建的自定义函数项目包含一些预生成的自定义函数，这些函数在 **src/customfunction.js** 文件中定义。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-140">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/functions/functions.js** file.</span></span> <span data-ttu-id="1f6e0-141">项目根目录中的 **manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 名称空间。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-141">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="1f6e0-142">在 Excel 工作簿中，通过在 Excel 中完成以下步骤来尝试使用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-142">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="1f6e0-143">在单元格内，键入 **=CONTOSO**。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-143">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="1f6e0-144">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-144">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="1f6e0-145">通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-145">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="1f6e0-146">`ADD` 自定义函数计算指定为输入参数的两个数字的总和。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-146">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="1f6e0-147">键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-147">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="1f6e0-148">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="1f6e0-148">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="1f6e0-149">如果需要一个可以从 API 请求股票价格并在工作表单元格中显示结果的函数，该怎么办？</span><span class="sxs-lookup"><span data-stu-id="1f6e0-149">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="1f6e0-150">自定义函数旨在使用户可以轻松地以异步方式从 Web 中请求数据。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-150">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="1f6e0-151">完成以下步骤，以创建一个名为 `stockPrice` 的自定义函数，该函数接受股票代码符号（例如，**MSFT**）并返回该股票的价格。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-151">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="1f6e0-152">此自定义函数使用 IEX Trading API，该 API 是免费的，并且不需要身份验证。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-152">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="1f6e0-153">在 Yeoman 生成器创建的 **stock-ticker** 项目中，找到文件 **src/customfunctions.js** 并在代码编辑器中打开它。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-153">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="1f6e0-154">在 **customfunctions.js** 中，找到 `increment` 函数并将以下代码添加到该函数后面。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-154">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

3. In **customfunctions.js**, locate the line`CustomFunctionMappings.INCREMENT = increment;`, add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

4. <span data-ttu-id="1f6e0-155">用户必须指定说明 Excel 函数的元数据，Excel 才能提供此新函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-155">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="1f6e0-156">打开 **config/customfunctions.json** 文件。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-156">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="1f6e0-157">将 JSON 对象添加到“函数”数组中，然后保存该文件。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-157">Add the following object to the  array within the src/functions/functions.json file and save the file.</span></span>

    <span data-ttu-id="1f6e0-158">此 JSON 说明了 `stockPrice` 函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-158">This JSON describes the `stockPrice` function.</span></span>

    ```JSON
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

5. <span data-ttu-id="1f6e0-159">必须在 Excel 中重新注册加载项，以便最终用户可以使用此新函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-159">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="1f6e0-160">完成针对本教程中将要使用的平台的下列相应步骤。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-160">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="1f6e0-161">如果使用的是 Excel for Windows，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-161">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="1f6e0-162">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-162">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="1f6e0-163">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="1f6e0-163">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="1f6e0-164">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-164">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="1f6e0-165">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="1f6e0-165">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="1f6e0-166">如果使用的是 Excel Online，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-166">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="1f6e0-167">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="1f6e0-167">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="1f6e0-168">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-168">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="1f6e0-169">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-169">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="1f6e0-170">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-170">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

6. <span data-ttu-id="1f6e0-171">现在，让我们尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-171">Now, let's try out the new function.</span></span> <span data-ttu-id="1f6e0-172">在单元格 **B1** 中，键入文本 `=CONTOSO.STOCKPRICE("MSFT")` 然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-172">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="1f6e0-173">应看到单元格 **B1** 中的结果是 Microsoft 一股股票的当前股票价格。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-173">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="1f6e0-174">创建流式处理异步自定义函数</span><span class="sxs-lookup"><span data-stu-id="1f6e0-174">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="1f6e0-175">刚刚创建的 `stockPrice` 函数返回特定时刻的股票价格，但股票价格一直在变化。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-175">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="1f6e0-176">让我们创建一个自定义函数，它从 API 传送数据，以获取股票价格的实时更新。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-176">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="1f6e0-177">完成以下步骤，创建一个名为 `stockPriceStream` 的自定义函数，该函数每 1000 毫秒请求指定股票的价格（假设之前的请求已经完成）。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-177">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="1f6e0-178">正在进行初始请求时，用户可能会在调用函数的单元格中看到占位符值 **#GETTING_DATA**。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-178">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="1f6e0-179">函数返回一个值后，**#GETTING_DATA** 将被替换为单元格中的该值。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-179">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="1f6e0-180">在 Yeoman 生成器创建的 **stock-ticker** 项目中，向 **src/customfunctions.js** 添加以下代码并保存文件。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-180">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="1f6e0-181">用户必须指定说明新函数的元数据，Excel 才能为用户提供此新函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-181">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="1f6e0-182">在 Yeoman 生成器创建的 **stock-ticker** 项目中，向 **config/customfunctions.json** 文件中的 `functions` 数组添加以下对象，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-182">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="1f6e0-183">此 JSON 说明了 `stockPriceStream` 函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-183">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="1f6e0-184">对于任何流式处理函数，必须在 `options` 对象中将 `stream` 属性和 `cancelable` 属性设置为 `true`，如本代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-184">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="1f6e0-185">必须在 Excel 中重新注册加载项，以便最终用户可以使用此新函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-185">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="1f6e0-186">完成针对本教程中将要使用的平台的下列相应步骤。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-186">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="1f6e0-187">如果使用的是 Excel for Windows，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-187">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="1f6e0-188">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-188">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="1f6e0-189">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="1f6e0-189">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="1f6e0-190">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-190">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="1f6e0-191">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="1f6e0-191">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="1f6e0-192">如果使用的是 Excel Online，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-192">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="1f6e0-193">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="1f6e0-193">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="1f6e0-194">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-194">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

        3. <span data-ttu-id="1f6e0-195">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-195">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span>

        4. <span data-ttu-id="1f6e0-196">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-196">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="1f6e0-197">现在，让我们尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-197">Now, let's try out the new function.</span></span> <span data-ttu-id="1f6e0-198">在单元格 **C1** 中，键入文本 `=CONTOSO.STOCKPRICESTREAM("MSFT")`，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-198">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="1f6e0-199">假设股票市场开盘，应该会看到单元格 **C1** 中的结果在不断更新，以反映 Microsoft 一股股票的实时价格。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-199">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="1f6e0-200">后续步骤</span><span class="sxs-lookup"><span data-stu-id="1f6e0-200">Next steps</span></span>

<span data-ttu-id="1f6e0-201">在本教程中，你已经创建新的自定义函数项目，尝试了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-201">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="1f6e0-202">若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="1f6e0-202">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="1f6e0-203">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="1f6e0-203">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="1f6e0-204">法律信息</span><span class="sxs-lookup"><span data-stu-id="1f6e0-204">Legal information</span></span>

<span data-ttu-id="1f6e0-205">[IEX](https://iextrading.com/developer/) 免费提供的数据。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-205">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="1f6e0-206">查看 [IEX 使用条款](https://iextrading.com/api-exhibit-a/)。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-206">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="1f6e0-207">Microsoft 在本教程中使用的 IEX API 仅供教学使用。</span><span class="sxs-lookup"><span data-stu-id="1f6e0-207">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
