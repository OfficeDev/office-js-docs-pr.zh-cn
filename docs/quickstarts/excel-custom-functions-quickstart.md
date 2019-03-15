---
ms.date: 03/06/2019
description: 在 Excel 快速入门指南中开发自定义函数。
title: 自定义函数快速入门 (预览)
localization_priority: Normal
ms.openlocfilehash: 9dd3e5a99f08ce0b931e705fac3312ab10c19e18
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632700"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="a3c23-103">开始开发 Excel 自定义函数</span><span class="sxs-lookup"><span data-stu-id="a3c23-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="a3c23-104">通过自定义函数, 开发人员现在可以通过在 JavaScript 或 Typescript 中将新函数定义为外接程序的一部分, 将它们添加到 Excel 中。</span><span class="sxs-lookup"><span data-stu-id="a3c23-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="a3c23-105">excel 用户可以像对待 excel 中的任何本机函数一样访问自定义函数, 例如`SUM()`。</span><span class="sxs-lookup"><span data-stu-id="a3c23-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a3c23-106">先决条件</span><span class="sxs-lookup"><span data-stu-id="a3c23-106">Prerequisites</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="a3c23-107">您需要以下工具和相关资源来开始创建自定义函数。</span><span class="sxs-lookup"><span data-stu-id="a3c23-107">You'll need the following tools and related resources to begin creating custom functions.</span></span>

- <span data-ttu-id="a3c23-108">[Node.js](https://nodejs.org/en/)（版本 8.0.0 或更高版本）</span><span class="sxs-lookup"><span data-stu-id="a3c23-108">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

- <span data-ttu-id="a3c23-109">[Git Bash](https://git-scm.com/downloads)（或其他 Git 客户端）</span><span class="sxs-lookup"><span data-stu-id="a3c23-109">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="a3c23-110">最新版本的 [Yeoman](https://yeoman.io/) 和[适用于 Office 外接程序的 Yeoman 生成器](https://www.npmjs.com/package/generator-office)。若要全局安装这些工具，请从命令提示符处运行以下命令：</span><span class="sxs-lookup"><span data-stu-id="a3c23-110">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="a3c23-111">即使以前安装了 Yeoman 生成器, 我们也建议您将程序包从 npm 更新到最新版本。</span><span class="sxs-lookup"><span data-stu-id="a3c23-111">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="a3c23-112">生成第一个自定义函数项目</span><span class="sxs-lookup"><span data-stu-id="a3c23-112">Build your first custom functions project</span></span>

<span data-ttu-id="a3c23-113">首先，使用 Yeoman 生成器创建自定义函数项目。</span><span class="sxs-lookup"><span data-stu-id="a3c23-113">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="a3c23-114">这将为你的项目设置开始对自定义函数进行编码所需的正确文件夹结构、源文件和依存关系。</span><span class="sxs-lookup"><span data-stu-id="a3c23-114">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="a3c23-115">运行下面的命令，再回答如下所示的提示问题。</span><span class="sxs-lookup"><span data-stu-id="a3c23-115">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    - <span data-ttu-id="a3c23-116">选择项目类型：`Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="a3c23-116">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    - <span data-ttu-id="a3c23-117">选择脚本类型：`JavaScript`</span><span class="sxs-lookup"><span data-stu-id="a3c23-117">Choose a script type: `JavaScript`</span></span>

    - <span data-ttu-id="a3c23-118">要如何命名加载项？</span><span class="sxs-lookup"><span data-stu-id="a3c23-118">What do you want to name your add-in?</span></span> `stock-ticker`

    ![自定义函数的 Office 外接程序提示的 Yeoman 生成器](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="a3c23-120">Yeoman 生成器将创建项目文件并安装支持的 Node 组件。</span><span class="sxs-lookup"><span data-stu-id="a3c23-120">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="a3c23-121">导航到刚创建的项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="a3c23-121">Navigate to the project folder you just created.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="a3c23-122">信任自签名证书, 您需要运行此项目。</span><span class="sxs-lookup"><span data-stu-id="a3c23-122">Trust the self-signed certificate you need to run this project.</span></span> <span data-ttu-id="a3c23-123">有关适用于 Windows 或 Mac 的详细说明，请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="a3c23-123">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="a3c23-124">生成项目。</span><span class="sxs-lookup"><span data-stu-id="a3c23-124">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="a3c23-125">启动在 Node.js 中运行的本地 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="a3c23-125">Start the local web server, which runs in Node.js.</span></span>

    - <span data-ttu-id="a3c23-126">如果使用 Excel for Windows 测试自定义函数, 请运行以下命令来启动本地 web 服务器, 启动 Excel, 并旁加载外接程序:</span><span class="sxs-lookup"><span data-stu-id="a3c23-126">If you use Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="a3c23-127">运行此命令后, 命令提示符将显示有关启动 web 服务器的详细信息。</span><span class="sxs-lookup"><span data-stu-id="a3c23-127">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="a3c23-128">Excel 将从加载的加载项开始。</span><span class="sxs-lookup"><span data-stu-id="a3c23-128">Excel will start with your add-in loaded.</span></span> <span data-ttu-id="a3c23-129">如果加载项未加载，请检查是否已正确完成步骤 3。</span><span class="sxs-lookup"><span data-stu-id="a3c23-129">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    - <span data-ttu-id="a3c23-130">如果使用 Excel Online 测试自定义函数, 请运行以下命令来启动本地 web 服务器:</span><span class="sxs-lookup"><span data-stu-id="a3c23-130">If you use Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="a3c23-131">运行此命令后, 命令提示符将显示有关启动 web 服务器的详细信息。</span><span class="sxs-lookup"><span data-stu-id="a3c23-131">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="a3c23-132">若要使用您的函数, 请在 Excel Online 中打开一个新工作簿。</span><span class="sxs-lookup"><span data-stu-id="a3c23-132">To use your functions, open a new workbook in Excel Online.</span></span> <span data-ttu-id="a3c23-133">在此工作簿中, 需要加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="a3c23-133">In this workbook, you'll need to load your add-in.</span></span> 

        <span data-ttu-id="a3c23-134">若要执行此操作, 请选择功能区上的 "**插入**" 选项卡, 然后选择 "**获取外接程序**"。在生成的新窗口中, 确保您在 "**我的外接程序**" 选项卡上。接下来, 选择 "**管理我的外接程序" > 上传我的外接程序**。</span><span class="sxs-lookup"><span data-stu-id="a3c23-134">To do this, select the **Insert** tab on the ribbon and select **Get Add-ins**. In the resulting new window, ensure you are on the **My Add-ins** tab. Next, select **Manage My Add-ins > Upload My Add-in**.</span></span> <span data-ttu-id="a3c23-135">浏览清单文件并将其上传。</span><span class="sxs-lookup"><span data-stu-id="a3c23-135">Browse for your manifest file and upload it.</span></span> <span data-ttu-id="a3c23-136">如果加载项未加载, 请检查是否已正确完成步骤3。</span><span class="sxs-lookup"><span data-stu-id="a3c23-136">If your add-in does not load, check you've completed step 3 correctly.</span></span>

## <a name="try-out-the-prebuilt-custom-functions"></a><span data-ttu-id="a3c23-137">尝试预生成的自定义函数</span><span class="sxs-lookup"><span data-stu-id="a3c23-137">Try out the prebuilt custom functions</span></span>

<span data-ttu-id="a3c23-138">使用 Yeoman 生成器创建的自定义函数项目包含一些预生成的自定义函数，这些函数在 **src/customfunction.js** 文件中定义。</span><span class="sxs-lookup"><span data-stu-id="a3c23-138">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file.</span></span> <span data-ttu-id="a3c23-139">项目根目录中的 **manifest.xml** 文件指定所有自定义函数均属于 `CONTOSO` 名称空间。</span><span class="sxs-lookup"><span data-stu-id="a3c23-139">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="a3c23-140">在 Excel 工作簿中, 通过完成`ADD`以下步骤来尝试使用自定义函数:</span><span class="sxs-lookup"><span data-stu-id="a3c23-140">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="a3c23-141">选择一个单元格并`=CONTOSO`键入。</span><span class="sxs-lookup"><span data-stu-id="a3c23-141">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="a3c23-142">请注意，自动完成菜单将显示 `CONTOSO` 命名空间中所有函数的列表。</span><span class="sxs-lookup"><span data-stu-id="a3c23-142">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="a3c23-143">通过在`CONTOSO.ADD`单元格中键入值`10` `=CONTOSO.ADD(10,200)`并`200`按 enter 来运行函数, 并使用数字和作为输入参数。</span><span class="sxs-lookup"><span data-stu-id="a3c23-143">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="a3c23-144">`ADD` 自定义函数计算指定为输入参数的两个数字的总和。</span><span class="sxs-lookup"><span data-stu-id="a3c23-144">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="a3c23-145">键入 `=CONTOSO.ADD(10,200)` 应在按下 Enter 后在单元格中生成结果 **210**。</span><span class="sxs-lookup"><span data-stu-id="a3c23-145">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a3c23-146">后续步骤</span><span class="sxs-lookup"><span data-stu-id="a3c23-146">Next steps</span></span>

<span data-ttu-id="a3c23-147">恭喜! 你已成功在 Excel 加载项中创建了自定义函数!</span><span class="sxs-lookup"><span data-stu-id="a3c23-147">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="a3c23-148">接下来, 使用流式数据功能生成更复杂的加载项。</span><span class="sxs-lookup"><span data-stu-id="a3c23-148">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="a3c23-149">下面的链接将指导您完成 Excel 加载项的自定义函数教程中的后续步骤。</span><span class="sxs-lookup"><span data-stu-id="a3c23-149">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a3c23-150">Excel 自定义函数加载项教程</span><span class="sxs-lookup"><span data-stu-id="a3c23-150">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="a3c23-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a3c23-151">See also</span></span>

* [<span data-ttu-id="a3c23-152">自定义函数概述</span><span class="sxs-lookup"><span data-stu-id="a3c23-152">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="a3c23-153">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="a3c23-153">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="a3c23-154">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="a3c23-154">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="a3c23-155">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="a3c23-155">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
