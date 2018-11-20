# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="5a096-101">教程：在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="5a096-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="5a096-102">简介</span><span class="sxs-lookup"><span data-stu-id="5a096-102">Introduction</span></span>

<span data-ttu-id="5a096-103">用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="5a096-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="5a096-104">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="5a096-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="5a096-105">可以创建自定义函数，以执行简单的任务（如自定义计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。</span><span class="sxs-lookup"><span data-stu-id="5a096-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="5a096-106">将在本教程中执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5a096-106">In this tutorial you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="5a096-107">通过使用 Yo Office 生成器创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="5a096-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="5a096-108">使用预生成的自定义函数来执行简单计算</span><span class="sxs-lookup"><span data-stu-id="5a096-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="5a096-109">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="5a096-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="5a096-110">创建从 Web 传送实时数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="5a096-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="5a096-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="5a096-111">Prerequisites</span></span>

* [<span data-ttu-id="5a096-112">Node.js 和 npm</span><span class="sxs-lookup"><span data-stu-id="5a096-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="5a096-113">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="5a096-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="5a096-114">最新版本的 [Yeoman](http://yeoman.io/) 和 [Yo Office 生成器](https://www.npmjs.com/package/generator-office)。</span><span class="sxs-lookup"><span data-stu-id="5a096-114">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office).</span></span> <span data-ttu-id="5a096-115">若要全局安装这些工具，请通过命令提示符运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="5a096-115">To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="5a096-116">Excel for Windows（版本 1810 或更高版本）或 Excel Online</span><span class="sxs-lookup"><span data-stu-id="5a096-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="5a096-117">加入 [Office 预览体验计划](https://products.office.com/office-insider)（**预览体验成员**级别 - 以前称为“预览体验成员 - 快”）</span><span class="sxs-lookup"><span data-stu-id="5a096-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="5a096-118">创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="5a096-118">Create a custom functions project</span></span>

<span data-ttu-id="5a096-119">本教程首先使用 Yo Office 生成器创建自定义函数项目所需的文件。</span><span class="sxs-lookup"><span data-stu-id="5a096-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="5a096-120">运行下面的命令，再回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="5a096-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="5a096-121">选择项目类型：`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="5a096-121">Choose a project type  </span></span>
    * <span data-ttu-id="5a096-122">选择脚本类型：`JavaScript`</span><span class="sxs-lookup"><span data-stu-id="5a096-122">Choose a script type  </span></span>
    * <span data-ttu-id="5a096-123">要如何命名加载项？</span><span class="sxs-lookup"><span data-stu-id="5a096-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![自定义函数的 Yo Office bash 提示](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="5a096-125">完成此向导后，生成器将创建项目文件，并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="5a096-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="5a096-126">项目文件来自 [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub 存储库。</span><span class="sxs-lookup"><span data-stu-id="5a096-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="5a096-127">导航到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="5a096-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="5a096-128">启动本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="5a096-128">Start the local web server.</span></span>

    * <span data-ttu-id="5a096-129">如果要使用 Excel for Windows 测试自定义函数，请运行以下命令来启动本地 Web 服务器，启动 Excel，并旁加载加载项：</span><span class="sxs-lookup"><span data-stu-id="5a096-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="5a096-130">如果要使用 Excel Online 测试自定义函数，请运行以下命令来启动本地 Web 服务器：</span><span class="sxs-lookup"><span data-stu-id="5a096-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="5a096-131">尝试预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="5a096-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="5a096-132">使用 Yo Office 生成器创建的自定义函数项目包含一些预生成的自定义函数，这些函数在 **src/customfunction.js** 文件中定义。</span><span class="sxs-lookup"><span data-stu-id="5a096-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="5a096-133">项目根目录中的 **manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 名称空间。</span><span class="sxs-lookup"><span data-stu-id="5a096-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="5a096-134">在使用任何预生成的自定义函数之前，必须在 Excel 中注册自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="5a096-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="5a096-135">通过完成针对本教程中将要使用的平台的相应步骤来执行上述操作。</span><span class="sxs-lookup"><span data-stu-id="5a096-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="5a096-136">如果要使用 Excel for Windows 测试自定义函数，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5a096-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="5a096-137">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="5a096-138">在可用加载项列表中，找到“开发人员加载项”\*\*\*\* 部分，并选择“Excel 自定义函数”\*\*\*\* 加载项以进行注册。</span><span class="sxs-lookup"><span data-stu-id="5a096-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="5a096-139">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="5a096-140">如果要使用 Excel Online 测试自定义函数，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5a096-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="5a096-141">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="5a096-142">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5a096-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="5a096-143">选择“浏览...”\*\*\*\*，并导航到 Yo Office 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="5a096-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="5a096-144">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5a096-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="5a096-145">此时，在项目中预生成的自定义函数将在 Excel 中加载并在其中可用。</span><span class="sxs-lookup"><span data-stu-id="5a096-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="5a096-146">通过在 Excel 中完成以下步骤来尝试使用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="5a096-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="5a096-147">在单元格内，键入 **=CONTOSO**。</span><span class="sxs-lookup"><span data-stu-id="5a096-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="5a096-148">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="5a096-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="5a096-149">通过在单元格中指定以下值并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数：</span><span class="sxs-lookup"><span data-stu-id="5a096-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="5a096-150">`ADD` 自定义函数计算指定为输入参数的两个数字的总和。</span><span class="sxs-lookup"><span data-stu-id="5a096-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="5a096-151">键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="5a096-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="5a096-152">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="5a096-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="5a096-153">如果需要一个可以从 API 请求股票价格并在工作表单元格中显示结果的函数，该怎么办？</span><span class="sxs-lookup"><span data-stu-id="5a096-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="5a096-154">自定义函数旨在使用户可以轻松地以异步方式从 Web 中请求数据。</span><span class="sxs-lookup"><span data-stu-id="5a096-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="5a096-155">完成以下步骤，以创建一个名为 `stockPrice` 的自定义函数，该函数接受股票代码（例如，**MSFT**）并返回该股票的价格。</span><span class="sxs-lookup"><span data-stu-id="5a096-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="5a096-156">此自定义函数使用 IEX Trading API，该 API 是免费的，并且不需要身份验证。</span><span class="sxs-lookup"><span data-stu-id="5a096-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="5a096-157">在 Yo Office 生成器创建的 **stock-ticker** 项目中，找到文件 **src/customfunctions.js** 并在代码编辑器中打开它。</span><span class="sxs-lookup"><span data-stu-id="5a096-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="5a096-158">将以下代码添加到 **customfunctions.js**，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="5a096-158">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="5a096-159">用户必须指定说明新函数的元数据，Excel 才能为最终用户提供此新函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="5a096-160">在 Yo Office 生成器创建的 **stock-ticker** 项目中，找到文件 **config/customfunctions.json** 并在代码编辑器中打开它。</span><span class="sxs-lookup"><span data-stu-id="5a096-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="5a096-161">将以下对象添加到 **config/customfunctions.json** 文件中的 `functions` 数组，并保存该文件。</span><span class="sxs-lookup"><span data-stu-id="5a096-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="5a096-162">此 JSON 说明了 `stockPrice` 函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-162">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="5a096-163">必须在 Excel 中重新注册加载项，以便最终用户可以使用此新函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="5a096-164">完成针对本教程中将要使用的平台的下列相应步骤。</span><span class="sxs-lookup"><span data-stu-id="5a096-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="5a096-165">如果使用的是 Excel for Windows，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5a096-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="5a096-166">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="5a096-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="5a096-167">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="5a096-168">在可用加载项列表中，找到“开发人员加载项”\*\*\*\* 部分，并选择“Excel 自定义函数”\*\*\*\* 加载项以进行注册。</span><span class="sxs-lookup"><span data-stu-id="5a096-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="5a096-169">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="5a096-170">如果使用的是 Excel Online，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5a096-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="5a096-171">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="5a096-172">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5a096-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="5a096-173">选择“浏览...”\*\*\*\*，并导航到 Yo Office 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="5a096-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="5a096-174">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5a096-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="5a096-175">现在，让我们尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-175">Now, let's try out the new function.</span></span> <span data-ttu-id="5a096-176">在单元格 **B1** 中，键入文本 `=CONTOSO.STOCKPRICE("MSFT")` 然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="5a096-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="5a096-177">应看到单元格 **B1** 中的结果是 Microsoft 一股股票的当前股票价格。</span><span class="sxs-lookup"><span data-stu-id="5a096-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="5a096-178">创建流式处理异步自定义函数</span><span class="sxs-lookup"><span data-stu-id="5a096-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="5a096-179">刚刚创建的 `stockPrice` 函数返回特定时刻的股票价格，但股票价格一直在变化。</span><span class="sxs-lookup"><span data-stu-id="5a096-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="5a096-180">让我们创建一个自定义函数，它从 API 传送数据，以获取股票价格的实时更新。</span><span class="sxs-lookup"><span data-stu-id="5a096-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="5a096-181">完成以下步骤，创建一个名为 `stockPriceStream` 的自定义函数，该函数每 1000 毫秒请求指定股票的价格（假设之前的请求已经完成）。</span><span class="sxs-lookup"><span data-stu-id="5a096-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="5a096-182">正在进行初始请求时，用户可能会在调用函数的单元格中看到占位符值 **#GETTING_DATA**。</span><span class="sxs-lookup"><span data-stu-id="5a096-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="5a096-183">函数返回一个值后，**#GETTING_DATA** 将被替换为单元格中的该值。</span><span class="sxs-lookup"><span data-stu-id="5a096-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="5a096-184">在 Yo Office 生成器创建的 **stock-ticker** 项目中，向 **src/customfunctions.js** 添加以下代码并保存文件。</span><span class="sxs-lookup"><span data-stu-id="5a096-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="5a096-185">用户必须指定说明新函数的元数据，Excel 才能为最终用户提供此新函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="5a096-186">在 Yo Office 生成器创建的 **stock-ticker** 项目中，向 **config/customfunctions.json** 文件中的 `functions` 数组添加以下对象，并保存文件。</span><span class="sxs-lookup"><span data-stu-id="5a096-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="5a096-187">此 JSON 说明了 `stockPriceStream` 函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="5a096-188">对于任何流式处理函数，必须在 `options` 对象中将 `stream` 属性和 `cancelable` 属性设置为 `true`，如本代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="5a096-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="5a096-189">必须在 Excel 中重新注册加载项，以便最终用户可以使用此新函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="5a096-190">完成针对本教程中将要使用的平台的下列相应步骤。</span><span class="sxs-lookup"><span data-stu-id="5a096-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="5a096-191">如果使用的是 Excel for Windows，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5a096-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="5a096-192">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="5a096-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="5a096-193">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="5a096-194">在可用加载项列表中，找到“开发人员加载项”\*\*\*\* 部分，并选择“Excel 自定义函数”\*\*\*\* 加载项以进行注册。</span><span class="sxs-lookup"><span data-stu-id="5a096-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="5a096-195">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="5a096-196">如果使用的是 Excel Online，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="5a096-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="5a096-197">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="5a096-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="5a096-198">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5a096-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="5a096-199">选择“浏览...”\*\*\*\*，并导航到 Yo Office 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="5a096-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="5a096-200">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="5a096-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="5a096-201">现在，让我们尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-201">Now, let's try out the new function.</span></span> <span data-ttu-id="5a096-202">在单元格 **C1** 中，键入文本 `=CONTOSO.STOCKPRICESTREAM("MSFT")`，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="5a096-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="5a096-203">假设股票市场开盘，应该会看到单元格 **C1** 中的结果在不断更新，以反映 Microsoft 一股股票的实时价格。</span><span class="sxs-lookup"><span data-stu-id="5a096-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="5a096-204">后续步骤</span><span class="sxs-lookup"><span data-stu-id="5a096-204">Next steps</span></span>

<span data-ttu-id="5a096-205">在本教程中，你已经创建新的自定义函数项目，尝试了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="5a096-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="5a096-206">若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="5a096-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="5a096-207">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="5a096-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="5a096-208">法律信息</span><span class="sxs-lookup"><span data-stu-id="5a096-208">Legal information</span></span>

<span data-ttu-id="5a096-209">[IEX](https://iextrading.com/developer/) 免费提供的数据。</span><span class="sxs-lookup"><span data-stu-id="5a096-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="5a096-210">查看 [IEX 使用条款](https://iextrading.com/api-exhibit-a/)。</span><span class="sxs-lookup"><span data-stu-id="5a096-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="5a096-211">Microsoft 在本教程中使用的 IEX API 仅供教学使用。</span><span class="sxs-lookup"><span data-stu-id="5a096-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
