---
title: Excel 自定义函数教程（预览）
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流式传输 Web 数据的自定义函数。
ms.date: 01/08/2019
ms.topic: tutorial
ms.openlocfilehash: 46a9883e9dbc2e3bfbbe170665d82826bdfb26f9
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770642"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a><span data-ttu-id="82484-103">教程：在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="82484-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="82484-104">用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="82484-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="82484-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="82484-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="82484-106">可以创建自定义函数，以执行简单的任务（如计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。</span><span class="sxs-lookup"><span data-stu-id="82484-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="82484-107">在本教程中，你将：</span><span class="sxs-lookup"><span data-stu-id="82484-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="82484-108">使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="82484-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="82484-109">使用预生成的自定义函数来执行简单计算。</span><span class="sxs-lookup"><span data-stu-id="82484-109">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="82484-110">创建从 Web 获取数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="82484-110">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="82484-111">创建从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="82484-111">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="82484-112">先决条件</span><span class="sxs-lookup"><span data-stu-id="82484-112">Prerequisites</span></span>

* <span data-ttu-id="82484-113">[Node.js](https://nodejs.org/en/)（版本 8.0.0 或更高版本）</span><span class="sxs-lookup"><span data-stu-id="82484-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="82484-114">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="82484-114">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="82484-115">最新版本的 [Yeoman](https://yeoman.io/) 和[适用于 Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="82484-115">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="82484-116">即便先前已安装 Yeoman 生成器，我们仍建议将包更新至最新的 npm 版本。</span><span class="sxs-lookup"><span data-stu-id="82484-116">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="82484-117">Excel for Windows（64 位，版本 1810 或更高版本）或 Excel Online</span><span class="sxs-lookup"><span data-stu-id="82484-117">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="82484-118">加入 [Office 预览体验计划](https://products.office.com/office-insider)（**预览体验成员**级别 - 以前称为“预览体验成员 - 快”）</span><span class="sxs-lookup"><span data-stu-id="82484-118">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="82484-119">创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="82484-119">Create a custom functions project</span></span>

 <span data-ttu-id="82484-120">首先，创建代码项目以构建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="82484-120">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="82484-121">[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)将使用可供你试用的一些初始自定义函数来设置项目。</span><span class="sxs-lookup"><span data-stu-id="82484-121">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some initial custom functions that you can try out.</span></span>

1. <span data-ttu-id="82484-122">运行下面的命令，再回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="82484-122">Run the following command and then answer the prompts as follows.</span></span>
    
    ```
    yo office
    ```
    
    * <span data-ttu-id="82484-123">选择项目类型：`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="82484-123">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="82484-124">选择脚本类型：`JavaScript`</span><span class="sxs-lookup"><span data-stu-id="82484-124">Choose a script type: `JavaScript`</span></span>
    * <span data-ttu-id="82484-125">要如何命名加载项？</span><span class="sxs-lookup"><span data-stu-id="82484-125">What do you want to name your add-in?</span></span> `stock-ticker`
    
    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/12-10-fork-cf-pic.jpg)
    
    <span data-ttu-id="82484-127">Yeoman 生成器将创建项目文件并安装支持的 Node.js 组件。</span><span class="sxs-lookup"><span data-stu-id="82484-127">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="82484-128">转到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="82484-128">Go to the project folder.</span></span>
    
    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="82484-129">信任运行此项目所需的自签名证书。</span><span class="sxs-lookup"><span data-stu-id="82484-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="82484-130">有关适用于 Windows 或 Mac 的详细说明，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="82484-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="82484-131">生成项目。</span><span class="sxs-lookup"><span data-stu-id="82484-131">Build the project.</span></span>
    
    ```
    npm run build
    ```

5. <span data-ttu-id="82484-132">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="82484-132">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="82484-133">你可以在 Excel for Windows 或 Excel Online 中尝试使用自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="82484-133">You can try out the custom function add-in in Excel for Windows, or Excel Online.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="82484-134">Excel for Windows</span><span class="sxs-lookup"><span data-stu-id="82484-134">Excel for Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="82484-135">运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="82484-135">Run the following command:</span></span>

```
npm run start
```

<span data-ttu-id="82484-136">此命令将启动 Web 服务器，并将自定义函数加载项旁加载到 Excel for Windows 中。</span><span class="sxs-lookup"><span data-stu-id="82484-136">This command starts the web server, and sideloads your custom function add-in into Excel for Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="82484-137">如果加载项未加载，请检查是否已正确完成步骤 3。</span><span class="sxs-lookup"><span data-stu-id="82484-137">If you add-in does not load, check that you have completed step 3 properly.</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="82484-138">Excel Online</span><span class="sxs-lookup"><span data-stu-id="82484-138">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="82484-139">运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="82484-139">Run the following command:</span></span>

```
npm run start-web
```

<span data-ttu-id="82484-140">此命令将启动 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="82484-140">This command starts the web server.</span></span> <span data-ttu-id="82484-141">使用以下步骤来旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="82484-141">Use the following steps to sideload your add-in.</span></span>

<ol type="a">
   <li><span data-ttu-id="82484-142">在 Excel Online 中，依次选择“插入”<strong></strong>选项卡和“加载项”<strong></strong>。</span><span class="sxs-lookup"><span data-stu-id="82484-142">In Excel Online, choose the <strong>Insert</strong> tab and then choose <strong>Add-ins</strong>.  Insert ribbon in Excel Online with the My Add-ins icon highlighted</span></span><br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li><span data-ttu-id="82484-143">选择“管理我的加载项”<strong></strong>，然后选择“上载我的加载项”<strong></strong>。</span><span class="sxs-lookup"><span data-stu-id="82484-143">Choose <strong>Manage My Add-ins</strong> and select <strong>Upload My Add-in</strong>.</span></span></li> 
   <li><span data-ttu-id="82484-144">选择“浏览...”<strong></strong>，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="82484-144">Choose <strong>Browse...</strong> and navigate to the root directory of the project that the Yeoman generator created.</span></span></li> 
   <li><span data-ttu-id="82484-145">依次选择文件“manifest.xml”<strong></strong>，“打开”<strong></strong>，然后选择“上载”<strong></strong>。</span><span class="sxs-lookup"><span data-stu-id="82484-145">Select the file <strong>manifest.xml</strong> and choose <strong>Open</strong>, then choose <strong>Upload</strong>.</span></span></li>
</ol>

> [!NOTE]
> <span data-ttu-id="82484-146">如果加载项未加载，请检查是否已正确完成步骤 3。</span><span class="sxs-lookup"><span data-stu-id="82484-146">If you add-in does not load, check that you have completed step 3 properly.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="82484-147">尝试预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="82484-147">Try out a prebuilt custom function</span></span>

<span data-ttu-id="82484-148">你创建的自定义函数项目已经有两个预生成的自定义函数，名为 ADD 和 INCREMENT。</span><span class="sxs-lookup"><span data-stu-id="82484-148">The custom functions project that you created alrady has two prebuilt custom functions named ADD and INCREMENT.</span></span> <span data-ttu-id="82484-149">这些预生成的函数的代码位于 **src/customfunctions.js** 文件中。</span><span class="sxs-lookup"><span data-stu-id="82484-149">The code for these prebuilt functions is in the  **src/customfunctions.js** file.</span></span> <span data-ttu-id="82484-150">**./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="82484-150">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="82484-151">你将使用 CONTOSO 命名空间来访问 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="82484-151">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="82484-152">接下来，通过完成以下步骤来尝试使用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="82484-152">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="82484-153">在 Excel 中，转至任意单元格并输入 `=CONTOSO`。</span><span class="sxs-lookup"><span data-stu-id="82484-153">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="82484-154">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="82484-154">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="82484-155">通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="82484-155">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="82484-156">`ADD` 自定义函数将计算你提供的两个数字的总和，并返回结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="82484-156">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="82484-157">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="82484-157">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="82484-158">集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。</span><span class="sxs-lookup"><span data-stu-id="82484-158">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="82484-159">接下来，你将创建一个名为 `stockPrice` 的自定义函数，该函数从 Web API 获取股票报价并将结果返回到工作表的单元格。</span><span class="sxs-lookup"><span data-stu-id="82484-159">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> <span data-ttu-id="82484-160">你将使用使用 IEX Trading API，该 API 是免费的，并且不需要身份验证。</span><span class="sxs-lookup"><span data-stu-id="82484-160">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="82484-161">在 **stock-ticker** 项目中，找到文件 **src/customfunctions.js** 并在代码编辑器中打开它。</span><span class="sxs-lookup"><span data-stu-id="82484-161">In the **stock-ticker** project that the Yeoman generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="82484-162">在 **customfunctions.js** 中，找到 `increment` 函数并将以下代码添加到该函数后面。</span><span class="sxs-lookup"><span data-stu-id="82484-162">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

3. In **customfunctions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    <span data-ttu-id="82484-163">`CustomFunctions.associate` 代码会将函数的 `id` 与 JavaScript 中的 `increment` 的函数地址相关联，以便 Excel 能够调用你的函数。</span><span class="sxs-lookup"><span data-stu-id="82484-163">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `increment` in JavaScript so that Excel can call your function.</span></span>

    <span data-ttu-id="82484-164">在 Excel 能够使用你的自定义函数之前，你需要先使用元数据来描述它。</span><span class="sxs-lookup"><span data-stu-id="82484-164">Before Excel can use your custom function, you need to describe it using metadata.</span></span> <span data-ttu-id="82484-165">你需要先定义在 `associate` 方法中使用的 `id` 以及某些其他元数据。</span><span class="sxs-lookup"><span data-stu-id="82484-165">You need to define the `id` used in the `associate` method previously, along with some other metadata.</span></span>


4. <span data-ttu-id="82484-166">打开 **config/customfunctions.json** 文件。</span><span class="sxs-lookup"><span data-stu-id="82484-166">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="82484-167">将 JSON 对象添加到“函数”数组中，然后保存该文件。</span><span class="sxs-lookup"><span data-stu-id="82484-167">Add the following JSON object to the 'functions' array and save the file.</span></span>

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
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    <span data-ttu-id="82484-168">此 JSON 将描述 `stockPrice` 函数、其参数以及它返回的结果类型。</span><span class="sxs-lookup"><span data-stu-id="82484-168">This JSON describes the `stockPrice` function, its parameters, and the type of result it returns.</span></span>

5. <span data-ttu-id="82484-169">在 Excel 中重新注册加载项，以便新函数可用。</span><span class="sxs-lookup"><span data-stu-id="82484-169">Re-register the add-in in Excel so that the new function is available.</span></span> 

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="82484-170">Excel for Windows</span><span class="sxs-lookup"><span data-stu-id="82484-170">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="82484-171">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="82484-171">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="82484-172">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="82484-172">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="82484-173">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="82484-173">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="82484-174">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="82484-174">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="82484-175">Excel Online</span><span class="sxs-lookup"><span data-stu-id="82484-175">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="82484-176">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="82484-176">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="82484-177">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="82484-177">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

3. <span data-ttu-id="82484-178">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="82484-178">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

4. <span data-ttu-id="82484-179">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="82484-179">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="6">
<li> <span data-ttu-id="82484-180">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="82484-180">Try out the new function.</span></span> <span data-ttu-id="82484-181">在单元格 <strong>B1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICE("MSFT")</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="82484-181">In cell <strong>B1</strong>, type the text <strong></strong> and press enter.</span></span> <span data-ttu-id="82484-182">应看到单元格 <strong>B1</strong> 中的结果是 Microsoft 一股股票的当前股票价格。</span><span class="sxs-lookup"><span data-stu-id="82484-182">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="82484-183">创建流式处理异步自定义函数</span><span class="sxs-lookup"><span data-stu-id="82484-183">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="82484-184">`stockPrice` 函数将返回特定时刻的股票价格，但股票价格一直在变化。</span><span class="sxs-lookup"><span data-stu-id="82484-184">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="82484-185">接下来，将创建一个名为 `stockPriceStream` 的自定义函数，该函数每隔 1000 毫秒获取一次股票价格。</span><span class="sxs-lookup"><span data-stu-id="82484-185">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="82484-186">在 **stock-ticker** 项目中，将以下代码添加到 **src/customfunctions.js** 并保存该文件。</span><span class="sxs-lookup"><span data-stu-id="82484-186">In the **stock-ticker** project that the Yeoman generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    <span data-ttu-id="82484-187">在 Excel 能够使用你的自定义函数之前，你需要先使用元数据来描述它。</span><span class="sxs-lookup"><span data-stu-id="82484-187">Before Excel can use your custom function, you need to describe it using metadata.</span></span>
    
2. <span data-ttu-id="82484-188">在 **stock-ticker** 项目中，将以下对象添加到 **config/customfunctions.json** 文件中的 `functions` 数组，并保存该文件。</span><span class="sxs-lookup"><span data-stu-id="82484-188">In the **stock-ticker** project that the Yeoman generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>
    
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
                "description": "stock symbol",
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

    <span data-ttu-id="82484-189">此 JSON 说明了 `stockPriceStream` 函数。</span><span class="sxs-lookup"><span data-stu-id="82484-189">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="82484-190">对于任何流式处理函数，必须在 `options` 对象中将 `stream` 属性和 `cancelable` 属性设置为 `true`，如本代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="82484-190">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

3. <span data-ttu-id="82484-191">在 Excel 中重新注册加载项，以便新函数可用。</span><span class="sxs-lookup"><span data-stu-id="82484-191">Re-register the add-in in Excel so that the new function is available.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="82484-192">Excel for Windows</span><span class="sxs-lookup"><span data-stu-id="82484-192">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="82484-193">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="82484-193">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="82484-194">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="82484-194">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="82484-195">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="82484-195">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="82484-196">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="82484-196">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="82484-197">Excel Online</span><span class="sxs-lookup"><span data-stu-id="82484-197">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="82484-198">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="82484-198">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="82484-199">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="82484-199">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="82484-200">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="82484-200">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="82484-201">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="82484-201">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="82484-202">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="82484-202">Try out the new function.</span></span> <span data-ttu-id="82484-203">在单元格 <strong>C1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="82484-203">In cell <strong>C1</strong>, type the text <strong></strong> and press enter.</span></span> <span data-ttu-id="82484-204">假设股票市场开盘，应该会看到单元格 <strong>C1</strong> 中的结果在不断更新，以反映 Microsoft 一股股票的实时价格。</span><span class="sxs-lookup"><span data-stu-id="82484-204">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="82484-205">后续步骤</span><span class="sxs-lookup"><span data-stu-id="82484-205">Next steps</span></span>

<span data-ttu-id="82484-206">恭喜！</span><span class="sxs-lookup"><span data-stu-id="82484-206">Congratulations!</span></span> <span data-ttu-id="82484-207">你已经创建新的自定义函数项目，尝试了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="82484-207">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="82484-208">若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="82484-208">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="82484-209">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="82484-209">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="82484-210">法律信息</span><span class="sxs-lookup"><span data-stu-id="82484-210">Legal information</span></span>

<span data-ttu-id="82484-211">[IEX](https://iextrading.com/developer/) 免费提供的数据。</span><span class="sxs-lookup"><span data-stu-id="82484-211">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="82484-212">查看 [IEX 使用条款](https://iextrading.com/api-exhibit-a/)。</span><span class="sxs-lookup"><span data-stu-id="82484-212">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="82484-213">Microsoft 在本教程中使用的 IEX API 仅供教学使用。</span><span class="sxs-lookup"><span data-stu-id="82484-213">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>


