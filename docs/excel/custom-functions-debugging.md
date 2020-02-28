---
ms.date: 07/10/2019
description: 在 Excel 中调试自定义函数。
title: 自定义函数调试
localization_priority: Normal
ms.openlocfilehash: dc620d8bab50c5efb3b9d9ec4f79f6532605f48b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324622"
---
# <a name="custom-functions-debugging"></a><span data-ttu-id="db7af-103">自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="db7af-103">Custom functions debugging</span></span>

<span data-ttu-id="db7af-104">自定义函数的调试可以通过多种方式来完成，具体取决于您使用的平台。</span><span class="sxs-lookup"><span data-stu-id="db7af-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="db7af-105">在 Windows 上：</span><span class="sxs-lookup"><span data-stu-id="db7af-105">On Windows:</span></span>
- [<span data-ttu-id="db7af-106">Excel Desktop 和 Visual Studio Code （VS Code）调试器</span><span class="sxs-lookup"><span data-stu-id="db7af-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="db7af-107">Web 上的 Excel 和 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="db7af-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="db7af-108">Web 和浏览器工具上的 Excel</span><span class="sxs-lookup"><span data-stu-id="db7af-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="db7af-109">命令行</span><span class="sxs-lookup"><span data-stu-id="db7af-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="db7af-110">在 Mac 上：</span><span class="sxs-lookup"><span data-stu-id="db7af-110">On Mac:</span></span>
- [<span data-ttu-id="db7af-111">Web 和浏览器工具上的 Excel</span><span class="sxs-lookup"><span data-stu-id="db7af-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="db7af-112">命令行</span><span class="sxs-lookup"><span data-stu-id="db7af-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="db7af-113">为简单起见，本文介绍了如何在使用 Visual Studio Code 编辑、运行任务以及某些情况下使用调试视图的上下文中进行调试。</span><span class="sxs-lookup"><span data-stu-id="db7af-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="db7af-114">如果使用的是其他编辑器或命令行工具，请参阅本文末尾的[命令行说明](#commands-for-building-and-running-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="db7af-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="db7af-115">Requirements</span><span class="sxs-lookup"><span data-stu-id="db7af-115">Requirements</span></span>

<span data-ttu-id="db7af-116">开始调试之前，应使用[Office 外接程序的 Yeoman 生成器](https://github.com/OfficeDev/generator-office)创建自定义函数项目。</span><span class="sxs-lookup"><span data-stu-id="db7af-116">Before starting to debug, you should use the [Yeoman generator for Office add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project.</span></span> <span data-ttu-id="db7af-117">有关如何创建自定义函数项目的指南，请参阅[自定义函数教程](../tutorials/excel-tutorial-create-custom-functions.md)。</span><span class="sxs-lookup"><span data-stu-id="db7af-117">For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="db7af-118">对 Excel 桌面使用 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="db7af-118">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="db7af-119">您可以使用 VS 代码在桌面上调试 Office Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="db7af-119">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="db7af-120">对 Mac 的桌面调试不可用，但可通过[使用浏览器工具和命令行来调试 web 上的 Excel 来](#use-the-command-line-tools-to-debug)实现。</span><span class="sxs-lookup"><span data-stu-id="db7af-120">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="db7af-121">从 VS 代码运行外接程序</span><span class="sxs-lookup"><span data-stu-id="db7af-121">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="db7af-122">打开[VS 代码](https://code.visualstudio.com/)中的自定义函数根项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="db7af-122">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="db7af-123">选择 "**终端 > 运行任务**" 并键入或选择 "**监视**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-123">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="db7af-124">这将监视和重建任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="db7af-124">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="db7af-125">选择 "**终端 > 运行任务**"，然后键入或选择 " **Dev Server**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-125">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="db7af-126">启动 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="db7af-126">Start the VS Code debugger</span></span>

4. <span data-ttu-id="db7af-127">选择 "**查看 > 调试**" 或输入**Ctrl + Shift + D**以切换到 "调试" 视图。</span><span class="sxs-lookup"><span data-stu-id="db7af-127">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="db7af-128">从 "调试" 选项中，选择 " **Excel 桌面**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-128">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="db7af-129">选择**F5** （或从菜单中选择**Debug-> 启动调试**）开始调试。</span><span class="sxs-lookup"><span data-stu-id="db7af-129">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="db7af-130">将打开一个新的 Excel 工作簿，您的外接程序已旁加载并可供使用。</span><span class="sxs-lookup"><span data-stu-id="db7af-130">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="db7af-131">开始调试</span><span class="sxs-lookup"><span data-stu-id="db7af-131">Start debugging</span></span>

1. <span data-ttu-id="db7af-132">在 VS 代码中，打开源代码脚本文件（**函数 .js**或**函数**）。</span><span class="sxs-lookup"><span data-stu-id="db7af-132">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="db7af-133">在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。</span><span class="sxs-lookup"><span data-stu-id="db7af-133">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="db7af-134">在 Excel 工作簿中，输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="db7af-134">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="db7af-135">此时执行将在您设置断点的代码行处停止。</span><span class="sxs-lookup"><span data-stu-id="db7af-135">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="db7af-136">现在，您可以逐步完成您的代码、设置监视和使用所需的任何与代码调试功能。</span><span class="sxs-lookup"><span data-stu-id="db7af-136">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="db7af-137">在 Microsoft Edge 中将 VS 代码调试程序与 Excel 一起使用</span><span class="sxs-lookup"><span data-stu-id="db7af-137">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="db7af-138">您可以使用 VS 代码在 Microsoft Edge 浏览器上调试 Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="db7af-138">You can use VS Code to debug custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="db7af-139">若要将 VS 代码与 Microsoft Edge 结合使用，必须[为 Microsoft edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)扩展安装调试器。</span><span class="sxs-lookup"><span data-stu-id="db7af-139">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="db7af-140">从 VS 代码运行外接程序</span><span class="sxs-lookup"><span data-stu-id="db7af-140">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="db7af-141">打开[VS 代码](https://code.visualstudio.com/)中的自定义函数根项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="db7af-141">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="db7af-142">选择 "**终端 > 运行任务**" 并键入或选择 "**监视**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-142">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="db7af-143">这将监视和重建任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="db7af-143">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="db7af-144">选择 "**终端 > 运行任务**"，然后键入或选择 " **Dev Server**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-144">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="db7af-145">启动 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="db7af-145">Start the VS Code debugger</span></span>

4. <span data-ttu-id="db7af-146">选择 "**查看 > 调试**" 或输入**Ctrl + Shift + D**以切换到 "调试" 视图。</span><span class="sxs-lookup"><span data-stu-id="db7af-146">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="db7af-147">从 "调试" 选项中，选择 " **Office Online （Microsoft Edge）**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-147">From the Debug options, choose **Office Online (Microsoft Edge)**.</span></span>
6. <span data-ttu-id="db7af-148">在 Microsoft Edge 浏览器中打开 Excel 并创建一个新的工作簿。</span><span class="sxs-lookup"><span data-stu-id="db7af-148">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="db7af-149">在功能区中选择 "**共享**"，然后复制此新工作簿的 URL 的链接。</span><span class="sxs-lookup"><span data-stu-id="db7af-149">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="db7af-150">选择**F5** （或从菜单中选择 "**调试" > 启动调试**）以开始调试。</span><span class="sxs-lookup"><span data-stu-id="db7af-150">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="db7af-151">将显示提示，询问您的文档的 URL。</span><span class="sxs-lookup"><span data-stu-id="db7af-151">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="db7af-152">粘贴到工作簿的 URL 中，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="db7af-152">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="db7af-153">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="db7af-153">Sideload your add-in</span></span>

1. <span data-ttu-id="db7af-154">选择功能区上的 "**插入**" 选项卡，然后在 "**外接程序**" 部分，选择 " **Office 外接程序**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-154">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="db7af-155">在 " **Office 外接程序**" 对话框中，选择 "**我的外**接程序" 选项卡，选择 "**管理我的外接**程序"，然后**上传我的外接程序**。</span><span class="sxs-lookup"><span data-stu-id="db7af-155">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="db7af-157">**浏览**到加载项清单文件，然后选择 "**上传**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-157">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="db7af-159">设置断点</span><span class="sxs-lookup"><span data-stu-id="db7af-159">Set breakpoints</span></span>
1. <span data-ttu-id="db7af-160">在 VS 代码中，打开源代码脚本文件（**函数 .js**或**函数**）。</span><span class="sxs-lookup"><span data-stu-id="db7af-160">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="db7af-161">在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。</span><span class="sxs-lookup"><span data-stu-id="db7af-161">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="db7af-162">在 Excel 工作簿中，输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="db7af-162">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="db7af-163">使用浏览器开发人员工具在 Excel 网页版中调试自定义函数</span><span class="sxs-lookup"><span data-stu-id="db7af-163">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="db7af-164">您可以使用浏览器开发人员工具在 Excel 网页版中调试自定义函数。</span><span class="sxs-lookup"><span data-stu-id="db7af-164">You can use the browser developer tools to debug custom functions in Excel on the web.</span></span> <span data-ttu-id="db7af-165">以下步骤适用于 Windows 和 macOS。</span><span class="sxs-lookup"><span data-stu-id="db7af-165">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="db7af-166">从 Visual Studio Code 运行外接程序</span><span class="sxs-lookup"><span data-stu-id="db7af-166">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="db7af-167">在[Visual Studio Code （VS code）](https://code.visualstudio.com/)中打开您的自定义函数根项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="db7af-167">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="db7af-168">选择 "**终端 > 运行任务**" 并键入或选择 "**监视**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-168">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="db7af-169">这将监视和重建任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="db7af-169">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="db7af-170">选择 "**终端 > 运行任务**"，然后键入或选择 " **Dev Server**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-170">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="db7af-171">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="db7af-171">Sideload your add-in</span></span>

1. <span data-ttu-id="db7af-172">打开 [Microsoft Office 网页版](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="db7af-172">Open [Microsoft Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="db7af-173">打开一个新的 Excel 工作簿。</span><span class="sxs-lookup"><span data-stu-id="db7af-173">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="db7af-174">打开功能区上的 "**插入**" 选项卡，然后在 "**外接程序**" 部分中，选择 " **Office 外接程序**"。</span><span class="sxs-lookup"><span data-stu-id="db7af-174">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="db7af-175">在 " **Office 外接程序**" 对话框中，选择 "**我的外**接程序" 选项卡，选择 "**管理我的外接**程序"，然后**上传我的外接程序**。</span><span class="sxs-lookup"><span data-stu-id="db7af-175">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="db7af-177">**转到**加载项清单文件，再选择“上传”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="db7af-177">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="db7af-179">在旁加载文档后，每次打开文档时它都将保留旁加载。</span><span class="sxs-lookup"><span data-stu-id="db7af-179">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="db7af-180">开始调试</span><span class="sxs-lookup"><span data-stu-id="db7af-180">Start debugging</span></span>

1. <span data-ttu-id="db7af-181">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="db7af-181">Open developer tools in the browser.</span></span> <span data-ttu-id="db7af-182">对于 Chrome 和大多数浏览器 F12 将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="db7af-182">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="db7af-183">在开发人员工具中，使用**Cmd + p**或**Ctrl + p**打开源代码脚本文件（**函数 .js**或**函数**）。</span><span class="sxs-lookup"><span data-stu-id="db7af-183">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="db7af-184">在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。</span><span class="sxs-lookup"><span data-stu-id="db7af-184">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="db7af-185">如果您需要更改代码，您可以在 VS 代码中进行编辑并保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="db7af-185">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="db7af-186">刷新浏览器以查看加载的更改。</span><span class="sxs-lookup"><span data-stu-id="db7af-186">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="db7af-187">使用命令行工具进行调试</span><span class="sxs-lookup"><span data-stu-id="db7af-187">Use the command line tools to debug</span></span>

<span data-ttu-id="db7af-188">如果未使用 VS 代码，则可以使用命令行（如 bash 或 PowerShell）运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="db7af-188">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="db7af-189">您需要使用浏览器开发人员工具在 Excel 网页版中调试代码。</span><span class="sxs-lookup"><span data-stu-id="db7af-189">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="db7af-190">无法使用命令行调试桌面版本的 Excel。</span><span class="sxs-lookup"><span data-stu-id="db7af-190">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="db7af-191">在命令行中运行`npm run watch` ，以便在发生代码更改时监视和重建。</span><span class="sxs-lookup"><span data-stu-id="db7af-191">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="db7af-192">打开第二个命令行窗口（运行监视时将阻止第一个命令行窗口。）</span><span class="sxs-lookup"><span data-stu-id="db7af-192">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="db7af-193">如果要在 Excel 的桌面版本中启动外接程序，请运行以下命令</span><span class="sxs-lookup"><span data-stu-id="db7af-193">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="db7af-194">或者，如果您更愿意在 Excel 网页上启动您的外接程序，请运行以下命令</span><span class="sxs-lookup"><span data-stu-id="db7af-194">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="db7af-195">对于 web 上的 Excel，您还需要旁加载您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="db7af-195">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="db7af-196">按照[旁加载您的外接程序](#sideload-your-add-in)中的步骤，旁加载你的外接程序。</span><span class="sxs-lookup"><span data-stu-id="db7af-196">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="db7af-197">然后继续转到下一节以开始调试。</span><span class="sxs-lookup"><span data-stu-id="db7af-197">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="db7af-198">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="db7af-198">Open developer tools in the browser.</span></span> <span data-ttu-id="db7af-199">对于 Chrome 和大多数浏览器 F12 将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="db7af-199">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="db7af-200">在开发人员工具中，打开源代码脚本文件（**函数 .js**或**函数**）。</span><span class="sxs-lookup"><span data-stu-id="db7af-200">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="db7af-201">您的自定义函数代码可能位于文件末尾附近。</span><span class="sxs-lookup"><span data-stu-id="db7af-201">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="db7af-202">在自定义函数源代码中，通过选择一行代码来应用断点。</span><span class="sxs-lookup"><span data-stu-id="db7af-202">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="db7af-203">如果您需要更改代码，您可以在 Visual Studio 中进行编辑并保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="db7af-203">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="db7af-204">刷新浏览器以查看加载的更改。</span><span class="sxs-lookup"><span data-stu-id="db7af-204">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="db7af-205">用于生成和运行外接程序的命令</span><span class="sxs-lookup"><span data-stu-id="db7af-205">Commands for building and running your add-in</span></span>

<span data-ttu-id="db7af-206">有几个可用的生成任务：</span><span class="sxs-lookup"><span data-stu-id="db7af-206">There are several build tasks available:</span></span>
- <span data-ttu-id="db7af-207">`npm run watch`：用于开发的构建，在保存源文件时自动重建</span><span class="sxs-lookup"><span data-stu-id="db7af-207">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="db7af-208">`npm run build-dev`：开发一次开发版本</span><span class="sxs-lookup"><span data-stu-id="db7af-208">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="db7af-209">`npm run build`：生产的内部版本</span><span class="sxs-lookup"><span data-stu-id="db7af-209">`npm run build`: builds for production</span></span>
- <span data-ttu-id="db7af-210">`npm run dev-server`：运行用于开发的 web 服务器</span><span class="sxs-lookup"><span data-stu-id="db7af-210">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="db7af-211">您可以使用以下任务在桌面或联机时开始调试。</span><span class="sxs-lookup"><span data-stu-id="db7af-211">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="db7af-212">`npm run start:desktop`：在桌面上启动 Excel 并将您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="db7af-212">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="db7af-213">`npm run start:web`：在 web 上启动 Excel 并将您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="db7af-213">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="db7af-214">`npm run stop`：停止 Excel 和调试。</span><span class="sxs-lookup"><span data-stu-id="db7af-214">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="db7af-215">后续步骤</span><span class="sxs-lookup"><span data-stu-id="db7af-215">Next steps</span></span>
<span data-ttu-id="db7af-216">了解[自定义函数中的身份验证方法](custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="db7af-216">Learn about [authentication practices in custom functions](custom-functions-authentication.md).</span></span> <span data-ttu-id="db7af-217">或者，查看[自定义函数的独特体系结构](custom-functions-architecture.md)。</span><span class="sxs-lookup"><span data-stu-id="db7af-217">Or, review [custom function's unique architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="db7af-218">另请参阅</span><span class="sxs-lookup"><span data-stu-id="db7af-218">See also</span></span>

* [<span data-ttu-id="db7af-219">自定义函数疑难解答</span><span class="sxs-lookup"><span data-stu-id="db7af-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="db7af-220">在 Excel 中处理自定义函数时出错</span><span class="sxs-lookup"><span data-stu-id="db7af-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="db7af-221">让自定义函数与 XLL 用户定义的函数兼容</span><span class="sxs-lookup"><span data-stu-id="db7af-221">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="db7af-222">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="db7af-222">Create custom functions in Excel</span></span>](custom-functions-overview.md)
