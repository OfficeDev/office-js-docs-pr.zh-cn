---
ms.date: 01/16/2020
description: 在 Excel 中开发自定义函数快速入门指南。
title: 自定义函数快速入门
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 126349f316cc923349f3f42e719c0017bbd7d7c0
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094146"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="3c1b5-103">开始开发 Excel 自定义函数</span><span class="sxs-lookup"><span data-stu-id="3c1b5-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="3c1b5-104">借助自定义函数，开发人员现在可以在 Excel 中添加新函数，方法是在 JavaScript 或 Typescript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="3c1b5-105">Excel 用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="3c1b5-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="3c1b5-106">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="3c1b5-107">Windows 版 Excel（版本 1904 或更高版本，关联至 Microsoft 365 订阅）或 Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="3c1b5-107">Excel on Windows (version 1904 or later, connected to Microsoft 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="3c1b5-108">Mac 版 Office（关联至 Microsoft 365 订阅）支持 Excel 自定义函数，并且本教程即将推出相关更新。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-108">Excel custom functions are supported in Office on Mac (connected to Microsoft 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="3c1b5-109">Office 2019（一次性购买）不支持 Excel 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="3c1b5-110">生成首个自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="3c1b5-110">Build your first custom functions project</span></span>

<span data-ttu-id="3c1b5-111">首先，使用 Yeoman 生成器创建自定义函数项目。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="3c1b5-112">这将为你的项目设置开始对自定义函数进行编码所需的正确文件夹结构、源文件和依存关系。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="3c1b5-113">**选择项目类型:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="3c1b5-113">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="3c1b5-114">**选择脚本类型:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="3c1b5-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="3c1b5-115">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="3c1b5-115">**What do you want to name your add-in?**</span></span> `starcount`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/starcountPrompt.png)

    <span data-ttu-id="3c1b5-117">Yeoman 生成器将创建项目文件并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="3c1b5-118">Yeoman 生成器将在命令行中为你提供有关如何处理项目的说明，但请忽略它们并继续按照我们的说明进行操作。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-118">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="3c1b5-119">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-119">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="3c1b5-120">生成项目。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-120">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="3c1b5-121">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-121">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="3c1b5-122">如果系统在运行 `npm run build` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-122">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="3c1b5-123">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-123">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="3c1b5-124">你可以在 Excel 网页版或 Windows 版 Excel 中尝试使用自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-124">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="3c1b5-125">系统可能会提示你打开加载项的任务窗格，不过这是可选的。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-125">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="3c1b5-126">你仍可在不打开加载项的任务窗格的情况下运行自定义函数。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-126">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windows"></a>[<span data-ttu-id="3c1b5-127">Windows 版 Excel</span><span class="sxs-lookup"><span data-stu-id="3c1b5-127">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="3c1b5-128">若要在 Windows 版 Excel 中测试加载项，请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-128">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="3c1b5-129">运行此命令时，本地 Web 服务器将启动，Excel 将打开并载入加载项。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-129">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[<span data-ttu-id="3c1b5-130">Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="3c1b5-130">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="3c1b5-131">若要在Excel 网页版中测试加载项，请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-131">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="3c1b5-132">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-132">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="3c1b5-133">若要使用自定义函数加载项，请在 Excel 网页版中打开一个新工作簿。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-133">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="3c1b5-134">在此工作簿中，完成以下步骤以旁加载你的加载项。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-134">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="3c1b5-135">在 Excel 中，选择“**插入**”选项卡，然后选择“**加载项**”。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-135">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Excel 网页版中的“插入”功能区，突出显示“我的加载项”图标](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="3c1b5-137">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-137">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="3c1b5-138">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-138">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="3c1b5-139">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-139">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="3c1b5-140">尝试使用预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="3c1b5-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="3c1b5-141">使用 Yeoman 生成器创建的自定义函数项目包含一些预生成的自定义函数，这些函数在 **./src/functions/functions.js** 文件中定义。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="3c1b5-142">项目根目录中的 **./manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 命名空间。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-142">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="3c1b5-143">在 Excel 工作簿中，通过完成以下步骤来尝试使用 `ADD` 自定义函数：</span><span class="sxs-lookup"><span data-stu-id="3c1b5-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="3c1b5-144">选择一个单元格，然后键入 `=CONTOSO`。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-144">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="3c1b5-145">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="3c1b5-146">通过在单元格中指定值 `=CONTOSO.ADD(10,200)` 并按 Enter 来运行 `CONTOSO.ADD` 函数，并将数字 `10` 和 `200` 作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-146">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="3c1b5-147">`ADD` 自定义函数计算指定为输入参数的两个数字的总和。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="3c1b5-148">键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="3c1b5-149">后续步骤</span><span class="sxs-lookup"><span data-stu-id="3c1b5-149">Next steps</span></span>

<span data-ttu-id="3c1b5-150">祝贺你，你已成功在 Excel 加载项中创建自定义函数！</span><span class="sxs-lookup"><span data-stu-id="3c1b5-150">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="3c1b5-151">接下来，可生成具有流式数据功能的更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-151">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="3c1b5-152">通过以下链接，可了解 Excel 自定义函数加载项教程中的后续步骤。</span><span class="sxs-lookup"><span data-stu-id="3c1b5-152">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="3c1b5-153">Excel 自定义函数加载项教程</span><span class="sxs-lookup"><span data-stu-id="3c1b5-153">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="3c1b5-154">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3c1b5-154">See also</span></span>

* [<span data-ttu-id="3c1b5-155">自定义函数概述</span><span class="sxs-lookup"><span data-stu-id="3c1b5-155">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="3c1b5-156">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="3c1b5-156">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="3c1b5-157">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="3c1b5-157">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)