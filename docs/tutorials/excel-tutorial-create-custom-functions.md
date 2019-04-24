---
title: Excel 自定义函数教程（预览）
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流式传输 Web 数据的自定义函数。
ms.date: 03/19/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 76f4d88b9da39a4d71927982836ee061b329a9b3
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451404"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a><span data-ttu-id="3ce5b-103">教程：在 Excel 中创建自定义函数（预览）</span><span class="sxs-lookup"><span data-stu-id="3ce5b-103">Tutorial: Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="3ce5b-104">用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="3ce5b-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="3ce5b-106">可以创建自定义函数，以执行简单的任务（如计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="3ce5b-107">在本教程中，你将：</span><span class="sxs-lookup"><span data-stu-id="3ce5b-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="3ce5b-108">使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="3ce5b-109">使用预生成的自定义函数来执行简单计算。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="3ce5b-110">创建从 Web 获取数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="3ce5b-111">创建从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-111">Create a custom function that streams real-time data from the web.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="3ce5b-112">先决条件</span><span class="sxs-lookup"><span data-stu-id="3ce5b-112">Prerequisites</span></span>

* <span data-ttu-id="3ce5b-113">[Node.js](https://nodejs.org/en/)（版本 8.0.0 或更高版本）</span><span class="sxs-lookup"><span data-stu-id="3ce5b-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="3ce5b-114">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="3ce5b-114">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="3ce5b-115">最新版本的 [Yeoman](https://yeoman.io/) 和[适用于 Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="3ce5b-115">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="3ce5b-116">即便先前已安装 Yeoman 生成器，我们仍建议将包更新至最新的 npm 版本。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-116">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="3ce5b-117">Excel for Windows（64 位，版本 1810 或更高版本）或 Excel Online</span><span class="sxs-lookup"><span data-stu-id="3ce5b-117">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="3ce5b-118">加入 [Office 预览体验计划](https://products.office.com/office-insider)（**预览体验成员**级别 - 以前称为“预览体验成员 - 快”）</span><span class="sxs-lookup"><span data-stu-id="3ce5b-118">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="3ce5b-119">创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="3ce5b-119">Create a custom functions project</span></span>

 <span data-ttu-id="3ce5b-120">首先，创建代码项目以构建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-120">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="3ce5b-121">[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)将使用可供你试用的一些初始自定义函数来设置项目。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-121">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some initial custom functions that you can try out.</span></span>

1. <span data-ttu-id="3ce5b-122">运行下面的命令，再回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-122">Run the following command and then answer the prompts as follows.</span></span>
    
    ```
    yo office
    ```
    
    * <span data-ttu-id="3ce5b-123">选择项目类型：`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="3ce5b-123">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="3ce5b-124">选择脚本类型：`JavaScript`</span><span class="sxs-lookup"><span data-stu-id="3ce5b-124">Choose a script type: `JavaScript`</span></span>
    * <span data-ttu-id="3ce5b-125">要如何命名加载项？</span><span class="sxs-lookup"><span data-stu-id="3ce5b-125">What do you want to name your add-in?</span></span> `stock-ticker`
    
    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/12-10-fork-cf-pic.jpg)
    
    <span data-ttu-id="3ce5b-127">Yeoman 生成器将创建项目文件并安装支持的 Node.js 组件。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-127">The Yeoman generator creates the project files and installs supporting Node.js components.</span></span>

2. <span data-ttu-id="3ce5b-128">转到项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-128">Go to the project folder.</span></span>
    
    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="3ce5b-129">信任运行此项目所需的自签名证书。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="3ce5b-130">有关适用于 Windows 或 Mac 的详细说明，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="3ce5b-131">生成项目。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-131">Build the project.</span></span>
    
    ```
    npm run build
    ```

5. <span data-ttu-id="3ce5b-132">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-132">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="3ce5b-133">你可以在 Excel for Windows 或 Excel Online 中尝试使用自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-133">You can try out the custom function add-in in Excel for Windows, or Excel Online.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="3ce5b-134">Excel for Windows</span><span class="sxs-lookup"><span data-stu-id="3ce5b-134">Excel for Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="3ce5b-135">运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-135">Run the following command.</span></span>

```
npm start desktop
```

<span data-ttu-id="3ce5b-136">此命令将启动 Web 服务器，并将自定义函数加载项旁加载到 Excel for Windows 中。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-136">This command starts the web server, and sideloads your custom function add-in into Excel for Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="3ce5b-137">如果加载项未加载，请检查是否已正确完成步骤 3。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-137">If your add-in does not load, check that you have completed step 3 properly.</span></span> <span data-ttu-id="3ce5b-138">您还可以启用**[运行时日志记录](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** 以解决外接程序的 XML 清单文件中的问题, 以及任何安装或运行时问题。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-138">You can also enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as any installation or runtime problems.</span></span> <span data-ttu-id="3ce5b-139">运行时日志`console.log`记录将语句写入日志文件, 以帮助您查找和修复问题。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-139">Runtime logging writes `console.log` statements to a log file to help you find and fix issues.</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="3ce5b-140">Excel Online</span><span class="sxs-lookup"><span data-stu-id="3ce5b-140">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="3ce5b-141">运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-141">Run the following command.</span></span>

```
npm start web
```

<span data-ttu-id="3ce5b-142">此命令将启动 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-142">This command starts the web server.</span></span> <span data-ttu-id="3ce5b-143">使用以下步骤来旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-143">Use the following steps to sideload your add-in.</span></span>

<ol type="a">
   <li><span data-ttu-id="3ce5b-144">在 Excel Online 中，依次选择“插入”<strong></strong>选项卡和“加载项”<strong></strong>。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-144">In Excel Online, choose the <strong>Insert</strong> tab and then choose <strong>Add-ins</strong>.</span></span><br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li><span data-ttu-id="3ce5b-145">选择“管理我的加载项”<strong></strong>，然后选择“上载我的加载项”<strong></strong>。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-145">Choose <strong>Manage My Add-ins</strong> and select <strong>Upload My Add-in</strong>.</span></span></li> 
   <li><span data-ttu-id="3ce5b-146">选择“浏览...”<strong></strong>，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-146">Choose <strong>Browse...</strong> and navigate to the root directory of the project that the Yeoman generator created.</span></span></li> 
   <li><span data-ttu-id="3ce5b-147">依次选择文件“manifest.xml”<strong></strong>，“打开”<strong></strong>，然后选择“上载”<strong></strong>。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-147">Select the file <strong>manifest.xml</strong> and choose <strong>Open</strong>, then choose <strong>Upload</strong>.</span></span></li>
</ol>

> [!NOTE]
> <span data-ttu-id="3ce5b-148">如果加载项未加载，请检查是否已正确完成步骤 3。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-148">If your add-in does not load, check that you have completed step 3 properly.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="3ce5b-149">尝试预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="3ce5b-149">Try out a prebuilt custom function</span></span>

<span data-ttu-id="3ce5b-150">你创建的自定义函数项目已经有两个预生成的自定义函数，名为 ADD 和 INCREMENT。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-150">The custom functions project that you created alrady has two prebuilt custom functions named ADD and INCREMENT.</span></span> <span data-ttu-id="3ce5b-151">这些预构建函数的代码位于**src/函数/函数 .js**文件中。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-151">The code for these prebuilt functions is in the  **src/functions/functions.js** file.</span></span> <span data-ttu-id="3ce5b-152">**./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-152">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="3ce5b-153">你将使用 CONTOSO 命名空间来访问 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-153">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="3ce5b-154">接下来，通过完成以下步骤来尝试使用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="3ce5b-154">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="3ce5b-155">在 Excel 中，转至任意单元格并输入 `=CONTOSO`。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-155">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="3ce5b-156">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-156">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="3ce5b-157">通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-157">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="3ce5b-158">`ADD` 自定义函数将计算你提供的两个数字的总和，并返回结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-158">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="3ce5b-159">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="3ce5b-159">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="3ce5b-160">集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-160">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="3ce5b-161">接下来，你将创建一个名为 `stockPrice` 的自定义函数，该函数从 Web API 获取股票报价并将结果返回到工作表的单元格。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-161">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> <span data-ttu-id="3ce5b-162">你将使用使用 IEX Trading API，该 API 是免费的，并且不需要身份验证。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-162">You’ll use the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="3ce5b-163">在**股票报价**项目中, 找到文件**src/函数/函数 .js**并在代码编辑器中打开它。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-163">In the **stock-ticker** project, find the file **src/functions/functions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="3ce5b-164">在**函数 .js**中, 找到`increment`函数并在该函数后面紧接着添加以下代码。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-164">In **functions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **functions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    <span data-ttu-id="3ce5b-165">`CustomFunctions.associate` 代码会将函数的 `id` 与 JavaScript 中的 `increment` 的函数地址相关联，以便 Excel 能够调用你的函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-165">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `increment` in JavaScript so that Excel can call your function.</span></span>

    <span data-ttu-id="3ce5b-166">在 Excel 能够使用你的自定义函数之前，你需要先使用元数据来描述它。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-166">Before Excel can use your custom function, you need to describe it using metadata.</span></span> <span data-ttu-id="3ce5b-167">你需要先定义在 `associate` 方法中使用的 `id` 以及某些其他元数据。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-167">You need to define the `id` used in the `associate` method previously, along with some other metadata.</span></span>


4. <span data-ttu-id="3ce5b-168">打开 " **src/函数/函数. json** " 文件。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-168">Open the **src/functions/functions.json** file.</span></span> <span data-ttu-id="3ce5b-169">将 JSON 对象添加到“函数”数组中，然后保存该文件。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-169">Add the following JSON object to the 'functions' array and save the file.</span></span>

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

    <span data-ttu-id="3ce5b-170">此 JSON 将描述 `stockPrice` 函数、其参数以及它返回的结果类型。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-170">This JSON describes the `stockPrice` function, its parameters, and the type of result it returns.</span></span>

5. <span data-ttu-id="3ce5b-171">在 Excel 中重新注册加载项，以便新函数可用。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-171">Re-register the add-in in Excel so that the new function is available.</span></span> 

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="3ce5b-172">Excel for Windows</span><span class="sxs-lookup"><span data-stu-id="3ce5b-172">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="3ce5b-173">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-173">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="3ce5b-174">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="3ce5b-174">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="3ce5b-175">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-175">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="3ce5b-176">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="3ce5b-176">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="3ce5b-177">Excel Online</span><span class="sxs-lookup"><span data-stu-id="3ce5b-177">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="3ce5b-178">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="3ce5b-178">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="3ce5b-179">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-179">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

3. <span data-ttu-id="3ce5b-180">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-180">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

4. <span data-ttu-id="3ce5b-181">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-181">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="6">
<li> <span data-ttu-id="3ce5b-182">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-182">Try out the new function.</span></span> <span data-ttu-id="3ce5b-183">在单元格 <strong>B1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICE("MSFT")</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-183">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="3ce5b-184">应看到单元格 <strong>B1</strong> 中的结果是 Microsoft 一股股票的当前股票价格。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-184">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="3ce5b-185">创建流式处理异步自定义函数</span><span class="sxs-lookup"><span data-stu-id="3ce5b-185">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="3ce5b-186">`stockPrice` 函数将返回特定时刻的股票价格，但股票价格一直在变化。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-186">The `stockPrice` function returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="3ce5b-187">接下来，将创建一个名为 `stockPriceStream` 的自定义函数，该函数每隔 1000 毫秒获取一次股票价格。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-187">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="3ce5b-188">在**股票报价**项目中, 将以下代码添加到**src/函数/函数 .js**中并保存文件。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-188">In the **stock-ticker** project, add the following code to **src/functions/functions.js** and save the file.</span></span>

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
    
    <span data-ttu-id="3ce5b-189">在 Excel 能够使用你的自定义函数之前，你需要先使用元数据来描述它。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-189">Before Excel can use your custom function, you need to describe it using metadata.</span></span>
    
2. <span data-ttu-id="3ce5b-190">在**股票报价**项目中, 将以下对象添加到`functions` **src/函数/函数-json**文件中的数组, 并保存文件。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-190">In the **stock-ticker** project add the following object to the `functions` array within the **src/functions/functions.json** file and save the file.</span></span>
    
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

    <span data-ttu-id="3ce5b-191">此 JSON 说明了 `stockPriceStream` 函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-191">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="3ce5b-192">对于任何流式处理函数，必须在 `options` 对象中将 `stream` 属性和 `cancelable` 属性设置为 `true`，如本代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-192">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

3. <span data-ttu-id="3ce5b-193">在 Excel 中重新注册加载项，以便新函数可用。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-193">Re-register the add-in in Excel so that the new function is available.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="3ce5b-194">Excel for Windows</span><span class="sxs-lookup"><span data-stu-id="3ce5b-194">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="3ce5b-195">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-195">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="3ce5b-196">在 Excel 中，选择“插入”\*\*\*\* 选项卡，然后选择位于“我的加载项”\*\*\*\* 右侧的向下箭头。![Excel for Windows 中的“插入”功能区，同时突出显示“我的加载项”箭头](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="3ce5b-196">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="3ce5b-197">在可用加载项列表中，找到“**开发人员加载项**”部分并选择 **stock-ticker** 加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-197">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="3ce5b-198">![Excel for Windows 中的“插入”功能区，同时在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="3ce5b-198">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="3ce5b-199">Excel Online</span><span class="sxs-lookup"><span data-stu-id="3ce5b-199">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="3ce5b-200">在 Excel Online 中，选择“插入”\*\*\*\* 选项卡，然后选择“加载项”\*\*\*\*。![Excel Online 中的“插入”功能区，同时突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="3ce5b-200">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="3ce5b-201">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-201">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="3ce5b-202">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-202">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="3ce5b-203">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-203">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="3ce5b-204">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-204">Try out the new function.</span></span> <span data-ttu-id="3ce5b-205">在单元格 <strong>C1</strong> 中，键入文本 <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-205">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="3ce5b-206">假设股票市场开盘，应该会看到单元格 <strong>C1</strong> 中的结果在不断更新，以反映 Microsoft 一股股票的实时价格。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-206">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="3ce5b-207">后续步骤</span><span class="sxs-lookup"><span data-stu-id="3ce5b-207">Next steps</span></span>

<span data-ttu-id="3ce5b-208">恭喜！</span><span class="sxs-lookup"><span data-stu-id="3ce5b-208">Congratulations!</span></span> <span data-ttu-id="3ce5b-209">你已经创建新的自定义函数项目，尝试了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-209">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="3ce5b-210">您也可以尝试使用[自定义函数调试指令](../excel/custom-functions-debugging.md)来调试此函数。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-210">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="3ce5b-211">若要详细了解 Excel 中的自定义函数，请继续阅读以下文章：</span><span class="sxs-lookup"><span data-stu-id="3ce5b-211">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3ce5b-212">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="3ce5b-212">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="3ce5b-213">法律信息</span><span class="sxs-lookup"><span data-stu-id="3ce5b-213">Legal information</span></span>

<span data-ttu-id="3ce5b-214">[IEX](https://iextrading.com/developer/) 免费提供的数据。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-214">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="3ce5b-215">查看 [IEX 使用条款](https://iextrading.com/api-exhibit-a/)。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-215">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="3ce5b-216">Microsoft 在本教程中使用的 IEX API 仅供教学使用。</span><span class="sxs-lookup"><span data-stu-id="3ce5b-216">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>


