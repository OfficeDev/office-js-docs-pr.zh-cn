---
title: Excel 自定义函数教程
description: 在本教程中，你将创建一个 Excel 外接程序，其中包含可执行计算、请求 Web 数据或流式传输 Web 数据的自定义函数。
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 4e2ad0276690d0b427a6788adc89ba09a274e203
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611083"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="762db-103">教程：在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="762db-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="762db-104">用户可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="762db-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="762db-105">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="762db-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="762db-106">可以创建自定义函数，以执行简单的任务（如计算）或更复杂的任务（如将实时数据从 Web 传送到工作表中）。</span><span class="sxs-lookup"><span data-stu-id="762db-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="762db-107">在本教程中，你将：</span><span class="sxs-lookup"><span data-stu-id="762db-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="762db-108">使用[适用于 Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)来创建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="762db-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="762db-109">使用预生成的自定义函数来执行简单计算。</span><span class="sxs-lookup"><span data-stu-id="762db-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="762db-110">创建从 Web 获取数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="762db-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="762db-111">创建从 Web 传送实时数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="762db-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="762db-112">先决条件</span><span class="sxs-lookup"><span data-stu-id="762db-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="762db-113">Windows 版 Excel（版本 1904 或更高版本，关联至 Office 365 订阅）或 Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="762db-113">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="762db-114">创建自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="762db-114">Create a custom functions project</span></span>

 <span data-ttu-id="762db-115">首先，创建代码项目以构建自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="762db-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="762db-116">[Office 加载项的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)将使用一些预生成的自定义函数（你可以试用这些函数）来设置你的项目。如果已运行自定义函数快速启动并生成了项目，请继续使用该项目，然后改为跳到[此步骤](#create-a-custom-function-that-requests-data-from-the-web)。</span><span class="sxs-lookup"><span data-stu-id="762db-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]
    
    * <span data-ttu-id="762db-117">**选择项目类型:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="762db-117">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="762db-118">**选择脚本类型:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="762db-118">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="762db-119">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="762db-119">**What do you want to name your add-in?**</span></span> `starcount`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/starcountPrompt.png)
    
    <span data-ttu-id="762db-121">Yeoman 生成器将创建项目文件并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="762db-121">The Yeoman generator will create the project files and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

2. <span data-ttu-id="762db-122">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="762db-122">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="762db-123">生成项目。</span><span class="sxs-lookup"><span data-stu-id="762db-123">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="762db-124">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="762db-124">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="762db-125">如果系统在运行 `npm run build` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="762db-125">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="762db-126">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="762db-126">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="762db-127">你可以在 Excel 网页版或 Windows 版 Excel 中尝试使用自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="762db-127">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="762db-128">Windows 版或 Mac 版 Excel</span><span class="sxs-lookup"><span data-stu-id="762db-128">Excel on Windows or Mac</span></span>](#tab/excel-windows)

<span data-ttu-id="762db-129">若要在 Windows 版或 Mac 版 Excel 中测试加载项，请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="762db-129">To test your add-in in Excel on Windows or Mac, run the following command.</span></span> <span data-ttu-id="762db-130">运行此命令时，本地 Web 服务器将启动，Excel 将打开并载入加载项。</span><span class="sxs-lookup"><span data-stu-id="762db-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[<span data-ttu-id="762db-131">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="762db-131">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="762db-132">若要在浏览器中的 Excel 中测试加载项，请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="762db-132">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="762db-133">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="762db-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="762db-134">若要使用自定义函数加载项，请在 Excel 网页版中打开一个新工作簿。</span><span class="sxs-lookup"><span data-stu-id="762db-134">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="762db-135">在此工作簿中，完成以下步骤以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="762db-135">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="762db-136">在 Excel 中，选择“**插入**”选项卡，然后选择“**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="762db-136">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Excel 网页版中的“插入”功能区，突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="762db-138">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="762db-138">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="762db-139">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="762db-139">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="762db-140">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="762db-140">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="762db-141">尝试使用预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="762db-141">Try out a prebuilt custom function</span></span>

<span data-ttu-id="762db-142">创建的自定义函数项目中包含一些预生成的自定义函数，这些函数在 **./src/functions/functions.js** 文件中定义。</span><span class="sxs-lookup"><span data-stu-id="762db-142">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="762db-143">**./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="762db-143">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="762db-144">你将使用 CONTOSO 命名空间来访问 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="762db-144">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="762db-145">接下来，通过完成以下步骤来尝试使用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="762db-145">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="762db-146">在 Excel 中，转至任意单元格并输入 `=CONTOSO`。</span><span class="sxs-lookup"><span data-stu-id="762db-146">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="762db-147">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="762db-147">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="762db-148">通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="762db-148">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="762db-149">`ADD` 自定义函数将计算你提供的两个数字的总和，并返回结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="762db-149">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="762db-150">创建从 Web 请求数据的自定义函数</span><span class="sxs-lookup"><span data-stu-id="762db-150">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="762db-151">集成来自 Web 的数据是通过自定义函数来扩展 Excel 的好方法。</span><span class="sxs-lookup"><span data-stu-id="762db-151">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="762db-152">接下来，需要创建一个名为“`getStarCount`”的自定义函数，显示给定 Github 存储库所拥有的星星数量。</span><span class="sxs-lookup"><span data-stu-id="762db-152">Next you'll create a custom function named `getStarCount` that shows how many stars a given Github repository possesses.</span></span>

1. <span data-ttu-id="762db-153">在 **starcount** 项目中，找到 **./src/functions/functions.js** 文件，然后在代码编辑器中将其打开。</span><span class="sxs-lookup"><span data-stu-id="762db-153">In the **starcount** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span> 

2. <span data-ttu-id="762db-154">在 **function.js** 中，添加以下代码：</span><span class="sxs-lookup"><span data-stu-id="762db-154">In **function.js**, add the following code:</span></span> 

```JS
/**
  * Gets the star count for a given Github repository.
  * @customfunction 
  * @param {string} userName string name of Github user or organization.
  * @param {string} repoName string name of the Github repository.
  * @return {number} number of stars given to a Github repository.
  */
  async function getStarCount(userName, repoName) {
    try {
      //You can change this URL to any web request you want to work with.
      const url = "https://api.github.com/repos/" + userName + "/" + repoName;
      const response = await fetch(url);
      //Expect that status code is in 200-299 range
      if (!response.ok) {
        throw new Error(response.statusText)
      }
        const jsonResponse = await response.json();
        return jsonResponse.watchers_count;
    }
    catch (error) {
      return error;
    }
  }
```

3. <span data-ttu-id="762db-155">运行以下命令以重新生成项目。</span><span class="sxs-lookup"><span data-stu-id="762db-155">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="762db-156">完成以下步骤（针对 Excel 网页版或者 Windows 版或 Mac 版 Excel），以在 Excel 中重新注册加载项。</span><span class="sxs-lookup"><span data-stu-id="762db-156">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="762db-157">必须完成这些步骤，才能使用新函数。</span><span class="sxs-lookup"><span data-stu-id="762db-157">You must complete these steps before the new function will be available.</span></span>

### <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="762db-158">Windows 版或 Mac 版 Excel</span><span class="sxs-lookup"><span data-stu-id="762db-158">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="762db-159">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="762db-159">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="762db-160">在 Excel 中，选择“**插入**”选项卡，然后选择位于“**我的加载项**”右侧的向下箭头。![Windows 版 Excel 中的“插入”功能区，突出显示“我的加载项”箭头](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="762db-160">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="762db-161">在可用加载项列表中，找到“**开发人员加载项**”部分并选择“**starcount**”加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="762db-161">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="762db-162">![Windows 版 Excel 中的“插入”功能区，在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="762db-162">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>


# <a name="excel-on-the-web"></a>[<span data-ttu-id="762db-163">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="762db-163">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="762db-164">在 Excel 中，选择“**插入**”选项卡，然后选择“**加载项**”。![Excel 网页版中的“插入”功能区，突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="762db-164">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="762db-165">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="762db-165">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="762db-166">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="762db-166">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="762db-167">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="762db-167">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="762db-168">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="762db-168">Try out the new function.</span></span> <span data-ttu-id="762db-169">在单元格 <strong>B1</strong> 中，键入文本 <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="762db-169">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> and press enter.</span></span> <span data-ttu-id="762db-170">你会看到，单元格 <strong>B1</strong> 中的结果便是 [Excel-Custom-Functions Github 存储库](https://github.com/OfficeDev/Excel-Custom-Functions)所获得的星星的当前数目。</span><span class="sxs-lookup"><span data-stu-id="762db-170">You should see that the result in cell <strong>B1</strong> is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="762db-171">创建流式处理异步自定义函数</span><span class="sxs-lookup"><span data-stu-id="762db-171">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="762db-172">`getStarCount` 函数返回存储库在特定时刻所拥有的星星数量。</span><span class="sxs-lookup"><span data-stu-id="762db-172">The `getStarCount` function returns the number of stars a repository has at a specific moment in time.</span></span> <span data-ttu-id="762db-173">自定义函数也可以返回不断变化的数据。</span><span class="sxs-lookup"><span data-stu-id="762db-173">Custom functions can also return data that is continuously changing.</span></span> <span data-ttu-id="762db-174">这些函数称为流式处理函数。</span><span class="sxs-lookup"><span data-stu-id="762db-174">These functions are called streaming functions.</span></span> <span data-ttu-id="762db-175">它们必须包含一个 `invocation` 参数，该参数指向从中调用函数的单元格。</span><span class="sxs-lookup"><span data-stu-id="762db-175">They must include an `invocation` parameter which refers to the cell where the function was called from.</span></span> <span data-ttu-id="762db-176">`invocation` 参数用于随时更新该单元格的内容。</span><span class="sxs-lookup"><span data-stu-id="762db-176">The `invocation` parameter is used to update the contents of the cell at any time.</span></span>  

<span data-ttu-id="762db-177">在下面的代码示例中，你会注意到有两个函数：`currentTime` 和 `clock`。</span><span class="sxs-lookup"><span data-stu-id="762db-177">In the following code sample, you'll notice that there are two functions, `currentTime` and `clock`.</span></span> <span data-ttu-id="762db-178">`currentTime` 函数是不使用流式处理的静态函数。</span><span class="sxs-lookup"><span data-stu-id="762db-178">The `currentTime` function is a static function that does not use streaming.</span></span> <span data-ttu-id="762db-179">它将以字符串形式返回日期。</span><span class="sxs-lookup"><span data-stu-id="762db-179">It returns the date as a string.</span></span> <span data-ttu-id="762db-180">`clock` 函数使用 `currentTime` 函数每秒向 Excel 中的单元格提供一次新时间。</span><span class="sxs-lookup"><span data-stu-id="762db-180">The `clock` function uses the `currentTime` function to provide the new time every second to a cell in Excel.</span></span> <span data-ttu-id="762db-181">它使用 `invocation.setResult` 将时间传递到 Excel 单元格，使用 `invocation.onCanceled` 处理取消该函数时发生的情况。</span><span class="sxs-lookup"><span data-stu-id="762db-181">It uses `invocation.setResult` to deliver the time to the Excel cell and `invocation.onCanceled` to handle what occurs when the function is canceled.</span></span>

1. <span data-ttu-id="762db-182">在 **starcount** 项目中，将以下代码添加到 **./src/functions/functions.js** 并保存该文件。</span><span class="sxs-lookup"><span data-stu-id="762db-182">In the **starcount** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

 /**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

2. <span data-ttu-id="762db-183">运行以下命令以重新生成项目。</span><span class="sxs-lookup"><span data-stu-id="762db-183">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="762db-184">完成以下步骤（针对 Excel 网页版或者 Windows 版或 Mac 版 Excel），以在 Excel 中重新注册加载项。</span><span class="sxs-lookup"><span data-stu-id="762db-184">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="762db-185">必须完成这些步骤，才能使用新函数。</span><span class="sxs-lookup"><span data-stu-id="762db-185">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="762db-186">Windows 版或 Mac 版 Excel</span><span class="sxs-lookup"><span data-stu-id="762db-186">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="762db-187">关闭 Excel，然后重新打开 Excel。</span><span class="sxs-lookup"><span data-stu-id="762db-187">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="762db-188">在 Excel 中，选择“**插入**”选项卡，然后选择位于“**我的加载项**”右侧的向下箭头。![Windows 版 Excel 中的“插入”功能区，突出显示“我的加载项”箭头](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="762db-188">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="762db-189">在可用加载项列表中，找到“**开发人员加载项**”部分并选择“**starcount**”加载项进行注册。</span><span class="sxs-lookup"><span data-stu-id="762db-189">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="762db-190">![Windows 版 Excel 中的“插入”功能区，在“我的加载项”列表中突出显示“Excel 自定义函数”加载项](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="762db-190">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>

# <a name="excel-on-the-web"></a>[<span data-ttu-id="762db-191">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="762db-191">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="762db-192">在 Excel 中，选择“**插入**”选项卡，然后选择“**加载项**”。![Excel 网页版中的“插入”功能区，突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="762db-192">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="762db-193">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="762db-193">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="762db-194">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="762db-194">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="762db-195">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="762db-195">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="762db-196">尝试使用新函数。</span><span class="sxs-lookup"><span data-stu-id="762db-196">Try out the new function.</span></span> <span data-ttu-id="762db-197">在单元格 <strong>C1</strong> 中，键入文本 <strong>=CONTOSO.CLOCK()</strong>，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="762db-197">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.CLOCK()</strong> and press enter.</span></span> <span data-ttu-id="762db-198">此时会显示当前日期，该日期每秒更新一次。</span><span class="sxs-lookup"><span data-stu-id="762db-198">You should see the current date, which streams an update every second.</span></span> <span data-ttu-id="762db-199">虽然此时钟只是一个循环计时器，但利用这一理念，你可以在更复杂的函数上设置计时器，以便执行对实时数据的 Web 请求。</span><span class="sxs-lookup"><span data-stu-id="762db-199">While this clock is just a timer on a loop, you can use the same idea of setting a timer on more complex functions that make web requests for real-time data.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="762db-200">后续步骤</span><span class="sxs-lookup"><span data-stu-id="762db-200">Next steps</span></span>

<span data-ttu-id="762db-201">恭喜！</span><span class="sxs-lookup"><span data-stu-id="762db-201">Congratulations!</span></span> <span data-ttu-id="762db-202">你已经创建新的自定义函数项目，试用了预生成的函数，创建了从 Web 请求数据的自定义函数，并创建了流式传输数据的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="762db-202">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data.</span></span> <span data-ttu-id="762db-203">接下来，可以修改项目以使用共享运行时，使您的函数更易于与任务窗格交互。</span><span class="sxs-lookup"><span data-stu-id="762db-203">Next, you can modify your project to use a shared runtime, making it easier for your function to interact with the task pane.</span></span> <span data-ttu-id="762db-204">按照以下文章中的步骤操作：</span><span class="sxs-lookup"><span data-stu-id="762db-204">Follow the steps in the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="762db-205">配置加载项以使用共享运行时</span><span class="sxs-lookup"><span data-stu-id="762db-205">Configure your add-in to use a shared runtime</span></span>](../excel/configure-your-add-in-to-use-a-shared-runtime.md)
