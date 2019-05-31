---
title: Excel 自定义函数教程
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流 Web 数据的自定义函数。
ms.date: 05/16/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 7d4d87a6bb3910c1b46698d5a2ff211ea1bbc6dd
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589172"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="11aa5-103">教程：在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="11aa5-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="11aa5-104">用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="11aa5-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="11aa5-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="11aa5-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="11aa5-106">可以创建自定义函数，以执行简单的任务（如计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。</span><span class="sxs-lookup"><span data-stu-id="11aa5-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="11aa5-107">在本教程中，你将：</span><span class="sxs-lookup"><span data-stu-id="11aa5-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="11aa5-108">使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="11aa5-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="11aa5-109">使用预生成的自定义函数来执行简单计算。</span><span class="sxs-lookup"><span data-stu-id="11aa5-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="11aa5-110">创建从 Web 获取数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="11aa5-111">创建从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="11aa5-112">先决条件</span><span class="sxs-lookup"><span data-stu-id="11aa5-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="11aa5-113">Windows 上的 Excel (64 位版本1810或更高版本) 或 Excel Online</span><span class="sxs-lookup"><span data-stu-id="11aa5-113">Excel on Windows (64-bit version 1810 or later) or Excel Online</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="11aa5-114">创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="11aa5-114">Create a custom functions project</span></span>

 <span data-ttu-id="11aa5-115">首先，创建代码项目以构建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="11aa5-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="11aa5-116">[Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)将使用一些预生成的自定义函数来设置您的项目, 您可以试用这些函数。如果已运行自定义函数 "快速启动" 并生成了一个项目, 请继续使用该项目, 然后跳到[此步骤](#create-a-custom-function-that-requests-data-from-the-web)。</span><span class="sxs-lookup"><span data-stu-id="11aa5-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. <span data-ttu-id="11aa5-117">运行下面的命令，再回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="11aa5-117">Run the following command and then answer the prompts as follows.</span></span>
    
    ```command&nbsp;line
    yo office
    ```
    
    * <span data-ttu-id="11aa5-118">**选择项目类型:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="11aa5-118">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="11aa5-119">**选择脚本类型:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="11aa5-119">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="11aa5-120">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="11aa5-120">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/UpdatedYoOfficePrompt.png)
    
    <span data-ttu-id="11aa5-122">Yeoman 生成器将创建项目文件并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="11aa5-122">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="11aa5-123">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="11aa5-123">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="11aa5-124">生成项目。</span><span class="sxs-lookup"><span data-stu-id="11aa5-124">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="11aa5-125">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="11aa5-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="11aa5-126">如果系统在运行 `npm run build` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="11aa5-126">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="11aa5-127">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="11aa5-127">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="11aa5-128">可以在 Windows 或 Excel Online 上试用 Excel 中的自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="11aa5-128">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="11aa5-129">Windows 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="11aa5-129">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="11aa5-130">若要在 Windows 中的 Excel 中测试外接程序, 请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="11aa5-130">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="11aa5-131">运行此命令时, 本地 web 服务器将启动, 并且 Excel 将在加载的外接程序中打开。</span><span class="sxs-lookup"><span data-stu-id="11aa5-131">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="11aa5-132">Excel Online</span><span class="sxs-lookup"><span data-stu-id="11aa5-132">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="11aa5-133">若要在 Excel Online 中测试外接程序, 请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="11aa5-133">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="11aa5-134">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="11aa5-134">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="11aa5-135">若要使用自定义函数外接程序, 请在 Excel Online 中打开一个新工作簿。</span><span class="sxs-lookup"><span data-stu-id="11aa5-135">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="11aa5-136">在此工作簿中, 完成以下步骤以旁加载您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="11aa5-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="11aa5-137">在 Excel Online 中，依次选择“插入”\*\*\*\* 选项卡和“加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="11aa5-137">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![在 Excel Online 中插入带突出显示 "我的外接程序" 图标的功能区](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="11aa5-139">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="11aa5-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="11aa5-140">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="11aa5-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="11aa5-141">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="11aa5-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="11aa5-142">尝试预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="11aa5-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="11aa5-143">您创建的自定义函数项目包含一些预生成的自定义函数, 这些函数是在 **/src/functions/functions.js**文件中定义的。</span><span class="sxs-lookup"><span data-stu-id="11aa5-143">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="11aa5-144">**./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="11aa5-144">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="11aa5-145">你将使用 CONTOSO 命名空间来访问 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-145">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="11aa5-146">接下来，通过完成以下步骤来尝试使用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="11aa5-146">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="11aa5-147">在 Excel 中，转至任意单元格并输入 `=CONTOSO`。</span><span class="sxs-lookup"><span data-stu-id="11aa5-147">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="11aa5-148">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="11aa5-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="11aa5-149">通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="11aa5-150">`ADD` 自定义函数将计算你提供的两个数字的总和，并返回结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="11aa5-150">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="11aa5-151">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="11aa5-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="11aa5-152">集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。</span><span class="sxs-lookup"><span data-stu-id="11aa5-152">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="11aa5-153">接下来，你将创建一个名为 `stockPrice` 的自定义函数，该函数从 Web API 获取股票报价并将结果返回到工作表的单元格。</span><span class="sxs-lookup"><span data-stu-id="11aa5-153">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> <span data-ttu-id="11aa5-154">你将使用使用 IEX Trading API，该 API 是免费的，并且不需要身份验证。</span><span class="sxs-lookup"><span data-stu-id="11aa5-154">You’ll use the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="11aa5-155">在**股票报价**项目中, 找到 **/src/functions/functions.js**并在代码编辑器中打开该文件。</span><span class="sxs-lookup"><span data-stu-id="11aa5-155">In the **stock-ticker** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="11aa5-156">在**函数 .js**中, 找到`increment`函数并在该函数后面添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="11aa5-156">In **functions.js**, locate the `increment` function and add the following code after that function.</span></span>

    ```js
    /**
    * Fetches current stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @returns {number} The current stock price.
    */
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
    CustomFunctions.associate("STOCKPRICE", stockPrice);
    ```

    <span data-ttu-id="11aa5-157">`CustomFunctions.associate` 代码会将函数的 `id` 与 JavaScript 中的 `stockPrice` 的函数地址相关联，以便 Excel 能够调用你的函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-157">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `stockPrice` in JavaScript so that Excel can call your function.</span></span>

3. <span data-ttu-id="11aa5-158">运行以下命令以重建项目。</span><span class="sxs-lookup"><span data-stu-id="11aa5-158">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="11aa5-159">完成以下步骤 (针对 Windows 或 Excel Online 上的 Excel), 以便在 Excel 中重新注册加载项。</span><span class="sxs-lookup"><span data-stu-id="11aa5-159">Complete the following steps (for either Excel on Windows or Excel Online) to re-register the add-in in Excel.</span></span> <span data-ttu-id="11aa5-160">您必须完成这些步骤, 新函数才可用。</span><span class="sxs-lookup"><span data-stu-id="11aa5-160">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="11aa5-161">Windows 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="11aa5-161">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="11aa5-162">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="11aa5-162">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="11aa5-163">在 Excel 中, 选择 "**插入**" 选项卡, 然后选择位于 **"我的外接程序**" 右侧的向下箭头。 ![在 Excel 中的 "我的外接程序" 箭头突出显示 Windows 中插入功能区](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="11aa5-163">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="11aa5-164">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="11aa5-164">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="11aa5-165">![在 Windows Excel 中插入带有 "我的外接程序" 列表中突出显示 Excel 自定义函数外接程序的功能区](../images/list-stock-ticker-red.png)</span><span class="sxs-lookup"><span data-stu-id="11aa5-165">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-stock-ticker-red.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="11aa5-166">Excel Online</span><span class="sxs-lookup"><span data-stu-id="11aa5-166">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="11aa5-167">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="11aa5-167">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="11aa5-168">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="11aa5-168">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

3. <span data-ttu-id="11aa5-169">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="11aa5-169">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

4. <span data-ttu-id="11aa5-170">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="11aa5-170">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="11aa5-171">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-171">Try out the new function.</span></span> <span data-ttu-id="11aa5-172">在单元格 <strong>B1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICE("MSFT")</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="11aa5-172">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="11aa5-173">应看到单元格 <strong>B1</strong> 中的结果是 Microsoft 一股股票的当前股票价格。</span><span class="sxs-lookup"><span data-stu-id="11aa5-173">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="11aa5-174">创建流式处理异步自定义函数</span><span class="sxs-lookup"><span data-stu-id="11aa5-174">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="11aa5-175">`stockPrice` 函数将返回特定时刻的股票价格，但股票价格一直在变化。</span><span class="sxs-lookup"><span data-stu-id="11aa5-175">The `stockPrice` function returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="11aa5-176">接下来，将创建一个名为 `stockPriceStream` 的自定义函数，该函数每隔 1000 毫秒获取一次股票价格。</span><span class="sxs-lookup"><span data-stu-id="11aa5-176">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="11aa5-177">在**股票报价**项目中, 将以下代码添加到 **./src/functions/functions.js**并保存文件。</span><span class="sxs-lookup"><span data-stu-id="11aa5-177">In the **stock-ticker** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

    ```js
    /**
    * Streams real time stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @param {CustomFunctions.StreamingInvocation<number>} invocation
    */
    function stockPriceStream(ticker, invocation) {
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
                    invocation.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    invocation.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        invocation.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
    ```
    
    <span data-ttu-id="11aa5-178">`CustomFunctions.associate` 代码会将函数的 `id` 与 JavaScript 中的 `stockPriceStream` 的函数地址相关联，以便 Excel 能够调用你的函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-178">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `stockPriceStream` in JavaScript so that Excel can call your function.</span></span>
    
2. <span data-ttu-id="11aa5-179">运行以下命令以重建项目。</span><span class="sxs-lookup"><span data-stu-id="11aa5-179">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="11aa5-180">完成以下步骤 (针对 Windows 或 Excel Online 上的 Excel), 以便在 Excel 中重新注册加载项。</span><span class="sxs-lookup"><span data-stu-id="11aa5-180">Complete the following steps (for either Excel on Windows or Excel Online) to re-register the add-in in Excel.</span></span> <span data-ttu-id="11aa5-181">您必须完成这些步骤, 新函数才可用。</span><span class="sxs-lookup"><span data-stu-id="11aa5-181">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="11aa5-182">Windows 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="11aa5-182">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="11aa5-183">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="11aa5-183">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="11aa5-184">在 Excel 中, 选择 "**插入**" 选项卡, 然后选择位于 **"我的外接程序**" 右侧的向下箭头。 ![在 Excel 中的 "我的外接程序" 箭头突出显示 Windows 中插入功能区](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="11aa5-184">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="11aa5-185">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="11aa5-185">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="11aa5-186">![在 Windows Excel 中插入带有 "我的外接程序" 列表中突出显示 Excel 自定义函数外接程序的功能区](../images/list-stock-ticker-red.png)</span><span class="sxs-lookup"><span data-stu-id="11aa5-186">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-stock-ticker-red.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="11aa5-187">Excel Online</span><span class="sxs-lookup"><span data-stu-id="11aa5-187">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="11aa5-188">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="11aa5-188">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="11aa5-189">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="11aa5-189">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="11aa5-190">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="11aa5-190">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="11aa5-191">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="11aa5-191">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="11aa5-192">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-192">Try out the new function.</span></span> <span data-ttu-id="11aa5-193">在单元格 <strong>C1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="11aa5-193">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="11aa5-194">假设股票市场开盘，应该会看到单元格 <strong>C1</strong> 中的结果在不断更新，以反映 Microsoft 一股股票的实时价格。</span><span class="sxs-lookup"><span data-stu-id="11aa5-194">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="11aa5-195">后续步骤</span><span class="sxs-lookup"><span data-stu-id="11aa5-195">Next steps</span></span>

<span data-ttu-id="11aa5-196">恭喜！</span><span class="sxs-lookup"><span data-stu-id="11aa5-196">Congratulations!</span></span> <span data-ttu-id="11aa5-197">你已经创建新的自定义函数项目，尝试了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-197">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="11aa5-198">您也可以尝试使用[自定义函数调试指令](../excel/custom-functions-debugging.md)来调试此函数。</span><span class="sxs-lookup"><span data-stu-id="11aa5-198">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="11aa5-199">若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="11aa5-199">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="11aa5-200">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="11aa5-200">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="11aa5-201">法律信息</span><span class="sxs-lookup"><span data-stu-id="11aa5-201">Legal information</span></span>

<span data-ttu-id="11aa5-202">[IEX](https://iextrading.com/developer/) 免费提供的数据。</span><span class="sxs-lookup"><span data-stu-id="11aa5-202">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="11aa5-203">查看 [IEX 使用条款](https://iextrading.com/api-exhibit-a/)。</span><span class="sxs-lookup"><span data-stu-id="11aa5-203">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="11aa5-204">Microsoft 在本教程中使用的 IEX API 仅供教学使用。</span><span class="sxs-lookup"><span data-stu-id="11aa5-204">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
