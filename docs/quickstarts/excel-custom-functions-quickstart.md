---
ms.date: 05/15/2019
description: 在 Excel 快速入门指南中开发自定义函数。
title: 自定义功能快速入门
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2ef7e05390750173c1679cdb53ed96ff19020e67
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589109"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="53854-103">开始开发 Excel 自定义函数</span><span class="sxs-lookup"><span data-stu-id="53854-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="53854-104">通过自定义函数, 开发人员现在可以通过在 JavaScript 或 Typescript 中将新函数定义为外接程序的一部分, 将它们添加到 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="53854-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="53854-105">Excel 用户可以像对待 Excel 中的任何本机函数一样访问自定义函数, 例如`SUM()`。</span><span class="sxs-lookup"><span data-stu-id="53854-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="53854-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="53854-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="53854-107">Windows 上的 Excel (64 位版本1810或更高版本) 或 Excel Online</span><span class="sxs-lookup"><span data-stu-id="53854-107">Excel on Windows (64-bit version 1810 or later) or Excel Online</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="53854-108">生成第一个自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="53854-108">Build your first custom functions project</span></span>

<span data-ttu-id="53854-109">首先，使用 Yeoman 生成器创建自定义函数项目。</span><span class="sxs-lookup"><span data-stu-id="53854-109">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="53854-110">这将为你的项目设置开始对自定义函数进行编码所需的正确文件夹结构、源文件和依存关系。</span><span class="sxs-lookup"><span data-stu-id="53854-110">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="53854-111">在所选的文件夹中, 运行以下命令, 然后按如下所示回答提示。</span><span class="sxs-lookup"><span data-stu-id="53854-111">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="53854-112">**选择项目类型:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="53854-112">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="53854-113">**选择脚本类型:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="53854-113">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="53854-114">**要如何命名加载项?**</span><span class="sxs-lookup"><span data-stu-id="53854-114">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/UpdatedYoOfficePrompt.png)

    <span data-ttu-id="53854-116">Yeoman 生成器将创建项目文件并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="53854-116">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="53854-117">Yeoman 生成器将为您提供有关如何处理项目的命令行中的一些说明, 但忽略它们并继续按照我们的说明操作。</span><span class="sxs-lookup"><span data-stu-id="53854-117">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="53854-118">导航到项目的根文件夹。</span><span class="sxs-lookup"><span data-stu-id="53854-118">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="53854-119">生成项目。</span><span class="sxs-lookup"><span data-stu-id="53854-119">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="53854-120">Office 加载项应使用 HTTPS，而不是 HTTP（即便是在开发时也是如此）。</span><span class="sxs-lookup"><span data-stu-id="53854-120">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="53854-121">如果系统在运行 `npm run build` 后提示你安装证书，请接受提示以安装 Yeoman 生成器提供的证书。</span><span class="sxs-lookup"><span data-stu-id="53854-121">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="53854-122">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="53854-122">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="53854-123">可以在 Windows 或 Excel Online 上试用 Excel 中的自定义函数加载项。</span><span class="sxs-lookup"><span data-stu-id="53854-123">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span> <span data-ttu-id="53854-124">系统可能会提示您打开加载项的任务窗格, 但这是可选的。</span><span class="sxs-lookup"><span data-stu-id="53854-124">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="53854-125">您仍可以运行自定义函数, 而无需打开加载项的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="53854-125">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="53854-126">Windows 上的 Excel</span><span class="sxs-lookup"><span data-stu-id="53854-126">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="53854-127">若要在 Windows 中的 Excel 中测试外接程序, 请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="53854-127">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="53854-128">运行此命令时, 本地 web 服务器将启动, 并且 Excel 将在加载的外接程序中打开。</span><span class="sxs-lookup"><span data-stu-id="53854-128">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="53854-129">Excel Online</span><span class="sxs-lookup"><span data-stu-id="53854-129">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="53854-130">若要在 Excel Online 中测试外接程序, 请运行以下命令。</span><span class="sxs-lookup"><span data-stu-id="53854-130">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="53854-131">运行此命令时，本地 Web 服务器将启动。</span><span class="sxs-lookup"><span data-stu-id="53854-131">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="53854-132">若要使用自定义函数外接程序, 请在 Excel Online 中打开一个新工作簿。</span><span class="sxs-lookup"><span data-stu-id="53854-132">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="53854-133">在此工作簿中, 完成以下步骤以旁加载您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="53854-133">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="53854-134">在 Excel Online 中，依次选择“插入”\*\*\*\* 选项卡和“加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="53854-134">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![在 Excel Online 中插入带突出显示 "我的外接程序" 图标的功能区](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="53854-136">选择“管理我的加载项”\*\*\*\*，然后选择“上载我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="53854-136">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="53854-137">选择“浏览...”\*\*\*\*，并导航到 Yeoman 生成器创建的项目的根目录。</span><span class="sxs-lookup"><span data-stu-id="53854-137">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="53854-138">依次选择文件“manifest.xml”\*\*\*\*，“打开”\*\*\*\*，然后选择“上载”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="53854-138">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="53854-139">尝试预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="53854-139">Try out a prebuilt custom function</span></span>

<span data-ttu-id="53854-140">使用 Yeoman 生成器创建的自定义函数项目包含一些预生成的自定义函数, 这些函数是在 **/src/functions/functions.js**文件中定义的。</span><span class="sxs-lookup"><span data-stu-id="53854-140">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="53854-141">项目根目录中的 **/manifest.xml**文件指定所有自定义函数均属于该`CONTOSO`命名空间。</span><span class="sxs-lookup"><span data-stu-id="53854-141">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="53854-142">在 Excel 工作簿中, 通过完成`ADD`以下步骤来尝试使用自定义函数:</span><span class="sxs-lookup"><span data-stu-id="53854-142">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="53854-143">选择一个单元格并`=CONTOSO`键入。</span><span class="sxs-lookup"><span data-stu-id="53854-143">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="53854-144">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="53854-144">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="53854-145">通过在`CONTOSO.ADD`单元格中键入值`10` `=CONTOSO.ADD(10,200)`并`200`按 enter 来运行函数, 并使用数字和作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="53854-145">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="53854-146">`ADD` 自定义函数计算指定为输入参数的两个数字的总和。</span><span class="sxs-lookup"><span data-stu-id="53854-146">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="53854-147">键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="53854-147">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="53854-148">后续步骤</span><span class="sxs-lookup"><span data-stu-id="53854-148">Next steps</span></span>

<span data-ttu-id="53854-149">恭喜! 你已成功在 Excel 加载项中创建了自定义函数!</span><span class="sxs-lookup"><span data-stu-id="53854-149">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="53854-150">接下来, 使用流式数据功能生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="53854-150">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="53854-151">下面的链接将指导您完成 Excel 加载项的自定义函数教程中的后续步骤。</span><span class="sxs-lookup"><span data-stu-id="53854-151">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="53854-152">Excel 自定义函数加载项教程</span><span class="sxs-lookup"><span data-stu-id="53854-152">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="53854-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="53854-153">See also</span></span>

* [<span data-ttu-id="53854-154">自定义函数概述</span><span class="sxs-lookup"><span data-stu-id="53854-154">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="53854-155">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="53854-155">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="53854-156">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="53854-156">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="53854-157">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="53854-157">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
