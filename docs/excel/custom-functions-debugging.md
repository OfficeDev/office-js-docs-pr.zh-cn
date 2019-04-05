---
ms.date: 03/13/2019
description: 在 Excel 中调试自定义函数。
title: 自定义函数调试 (预览)
localization_priority: Normal
ms.openlocfilehash: 66b55855fdbdc3b3cfc7a316cb8fd7e06f073213
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478952"
---
# <a name="custom-functions-debugging-preview"></a><span data-ttu-id="0f58b-103">自定义函数调试 (预览)</span><span class="sxs-lookup"><span data-stu-id="0f58b-103">Custom functions debugging (preview)</span></span>

<span data-ttu-id="0f58b-104">自定义函数的调试可以通过多种方式来完成, 具体取决于您使用的平台。</span><span class="sxs-lookup"><span data-stu-id="0f58b-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

<span data-ttu-id="0f58b-105">在 Windows 上:</span><span class="sxs-lookup"><span data-stu-id="0f58b-105">On Windows:</span></span>
- [<span data-ttu-id="0f58b-106">Excel Desktop 和 Visual Studio Code (VS Code) 调试器</span><span class="sxs-lookup"><span data-stu-id="0f58b-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="0f58b-107">Excel Online 和 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="0f58b-107">Excel Online and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [<span data-ttu-id="0f58b-108">Excel Online 和浏览器工具</span><span class="sxs-lookup"><span data-stu-id="0f58b-108">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="0f58b-109">命令行</span><span class="sxs-lookup"><span data-stu-id="0f58b-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="0f58b-110">在 Mac 上:</span><span class="sxs-lookup"><span data-stu-id="0f58b-110">On Mac:</span></span>
- [<span data-ttu-id="0f58b-111">Excel Online 和浏览器工具</span><span class="sxs-lookup"><span data-stu-id="0f58b-111">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="0f58b-112">命令行</span><span class="sxs-lookup"><span data-stu-id="0f58b-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> <span data-ttu-id="0f58b-113">为简单起见, 本文介绍了如何在使用 Visual Studio Code 编辑、运行任务以及某些情况下使用调试视图的上下文中进行调试。</span><span class="sxs-lookup"><span data-stu-id="0f58b-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="0f58b-114">如果使用的是其他编辑器或命令行工具, 请参阅本文末尾的[命令行说明](#Use-the-command-line-tools-to-debug)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-114">If you are using a different editor or command line tool, see the [command line instructions](#Use-the-command-line-tools-to-debug) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="0f58b-115">Requirements</span><span class="sxs-lookup"><span data-stu-id="0f58b-115">Requirements</span></span>

<span data-ttu-id="0f58b-116">开始调试之前, 应使用 Yo Office 生成器创建自定义函数外接程序项目, 并确保您的项目具有受信任的自签名证书。</span><span class="sxs-lookup"><span data-stu-id="0f58b-116">Before starting to debug, you should create a custom functions add-in project using the Yo Office generator and ensured that you have trusted self-signed certificates for your project.</span></span> <span data-ttu-id="0f58b-117">有关创建项目的说明, 请参阅[自定义函数教程](https://review.docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-117">For instructions to create a project, see the [custom functions tutorial](https://review.docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions).</span></span> <span data-ttu-id="0f58b-118">有关信任证书的说明, 请参阅[将自签名证书添加为受信任的根证书](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-118">For instructions on trusting certificates, see [Adding self-signed certificates as trusted root certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="0f58b-119">对 Excel 桌面使用 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="0f58b-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="0f58b-120">您可以使用 VS 代码在桌面上调试 Office Excel 中的自定义函数。</span><span class="sxs-lookup"><span data-stu-id="0f58b-120">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="0f58b-121">对 Mac 的桌面调试不可用, 但可通过[使用浏览器工具来调试 Excel Online](#debug-in-excel-online-by-using-the-browser-developer-tools)来实现。</span><span class="sxs-lookup"><span data-stu-id="0f58b-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools to debug Excel Online](#debug-in-excel-online-by-using-the-browser-developer-tools).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="0f58b-122">从 VS 代码运行外接程序</span><span class="sxs-lookup"><span data-stu-id="0f58b-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="0f58b-123">打开[VS 代码](https://code.visualstudio.com/)中的自定义函数根项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="0f58b-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="0f58b-124">选择 "**终端 > 运行任务**", 然后键入或选择 "**监视**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="0f58b-125">这将监视和重建任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="0f58b-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="0f58b-126">选择 "**终端 > 运行任务**", 然后键入或选择 " **Dev Server**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="0f58b-127">启动 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="0f58b-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="0f58b-128">选择 "**查看 > 调试**" 或输入**Ctrl + Shift + D**以切换到 "调试" 视图。</span><span class="sxs-lookup"><span data-stu-id="0f58b-128">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="0f58b-129">从 "调试" 选项中, 选择 " **Excel 桌面**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-129">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="0f58b-130">选择**F5** (或从菜单中选择 **> 启动调试**) 以开始调试。</span><span class="sxs-lookup"><span data-stu-id="0f58b-130">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="0f58b-131">将打开一个新的 Excel 工作簿, 您的外接程序已旁加载并可供使用。</span><span class="sxs-lookup"><span data-stu-id="0f58b-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="0f58b-132">开始调试</span><span class="sxs-lookup"><span data-stu-id="0f58b-132">Start debugging</span></span>

1. <span data-ttu-id="0f58b-133">在 VS 代码中, 打开源代码脚本文件 (函数 .js 或函数)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-133">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="0f58b-134">在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="0f58b-135">在 Excel 工作簿中, 输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="0f58b-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="0f58b-136">此时执行将在您设置断点的代码行处停止。</span><span class="sxs-lookup"><span data-stu-id="0f58b-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="0f58b-137">现在, 您可以逐步完成您的代码、设置监视和使用所需的任何与代码调试功能。</span><span class="sxs-lookup"><span data-stu-id="0f58b-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a><span data-ttu-id="0f58b-138">在 Microsoft Edge 中将 VS 代码调试程序与 Excel Online 一起使用</span><span class="sxs-lookup"><span data-stu-id="0f58b-138">Use the VS Code debugger for Excel Online in Microsoft Edge</span></span>

<span data-ttu-id="0f58b-139">您可以使用 VS 代码在 Microsoft Edge 浏览器的 Excel Online 中调试自定义函数。</span><span class="sxs-lookup"><span data-stu-id="0f58b-139">You can use VS Code to debug custom functions in Excel Online in the Microsoft Edge browser.</span></span> <span data-ttu-id="0f58b-140">若要将 VS 代码与 microsoft edge 结合使用, 必须[为 microsoft edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)扩展安装调试器。</span><span class="sxs-lookup"><span data-stu-id="0f58b-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="0f58b-141">从 VS 代码运行外接程序</span><span class="sxs-lookup"><span data-stu-id="0f58b-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="0f58b-142">打开[VS 代码](https://code.visualstudio.com/)中的自定义函数根项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="0f58b-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="0f58b-143">选择 "**终端 > 运行任务**", 然后键入或选择 "**监视**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="0f58b-144">这将监视和重建任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="0f58b-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="0f58b-145">选择 "**终端 > 运行任务**", 然后键入或选择 " **Dev Server**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="0f58b-146">启动 VS 代码调试器</span><span class="sxs-lookup"><span data-stu-id="0f58b-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="0f58b-147">选择 "**查看 > 调试**" 或输入**Ctrl + Shift + D**以切换到 "调试" 视图。</span><span class="sxs-lookup"><span data-stu-id="0f58b-147">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="0f58b-148">从 "调试" 选项中, 选择 " **Office Online (边缘)**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-148">From the Debug options, choose **Office Online (Edge)**.</span></span>
6. <span data-ttu-id="0f58b-149">使用 Microsoft Edge 浏览器打开 excel online, 打开 excel Online, 并创建新的工作簿。</span><span class="sxs-lookup"><span data-stu-id="0f58b-149">Open Excel Online using the Microsoft Edge browser, open Excel Online, and create a new workbook.</span></span>
7. <span data-ttu-id="0f58b-150">在功能区中选择 "**共享**", 然后复制此新工作簿的 URL 的链接。</span><span class="sxs-lookup"><span data-stu-id="0f58b-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="0f58b-151">选择**F5** (或从菜单中选择 "**调试" > "启动调试**") 开始调试。</span><span class="sxs-lookup"><span data-stu-id="0f58b-151">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="0f58b-152">将显示提示, 询问您的文档的 URL。</span><span class="sxs-lookup"><span data-stu-id="0f58b-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="0f58b-153">粘贴到工作簿的 URL 中, 然后按 enter。</span><span class="sxs-lookup"><span data-stu-id="0f58b-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="0f58b-154">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="0f58b-154">Sideload your add-in</span></span>   

1. <span data-ttu-id="0f58b-155">选择功能区上的 "**插入**" 选项卡, 然后在 "**外接程序**" 部分, 选择 " **Office 外接程序**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-155">Select the  **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="0f58b-156">在“Office 加载项”\*\*\*\* 对话框中，依次选择“我的加载项”\*\*\*\* 选项卡、“管理我的加载项”\*\*\*\* 和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="0f58b-156">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

3.  <span data-ttu-id="0f58b-158">**浏览**到加载项清单文件, 然后选择 "**上传**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="0f58b-160">设置断点</span><span class="sxs-lookup"><span data-stu-id="0f58b-160">Set breakpoints</span></span>
1. <span data-ttu-id="0f58b-161">在 VS 代码中, 打开源代码脚本文件 (函数 .js 或函数)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-161">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="0f58b-162">在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="0f58b-163">在 Excel 工作簿中, 输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="0f58b-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a><span data-ttu-id="0f58b-164">使用浏览器开发人员工具在 Excel Online 中调试自定义函数</span><span class="sxs-lookup"><span data-stu-id="0f58b-164">Use the browser developer tools to debug custom functions in Excel Online</span></span>

<span data-ttu-id="0f58b-165">您可以使用浏览器开发人员工具在 Excel Online 中调试自定义函数。</span><span class="sxs-lookup"><span data-stu-id="0f58b-165">You can use the browser developer tools to debug custom functions in Excel Online.</span></span> <span data-ttu-id="0f58b-166">以下步骤适用于 Windows 和 macOS。</span><span class="sxs-lookup"><span data-stu-id="0f58b-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="0f58b-167">从 Visual Studio Code 运行外接程序</span><span class="sxs-lookup"><span data-stu-id="0f58b-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="0f58b-168">在[Visual Studio Code (VS code)](https://code.visualstudio.com/)中打开您的自定义函数根项目文件夹。</span><span class="sxs-lookup"><span data-stu-id="0f58b-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="0f58b-169">选择 "**终端 > 运行任务**", 然后键入或选择 "**监视**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="0f58b-170">这将监视和重建任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="0f58b-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="0f58b-171">选择 "**终端 > 运行任务**", 然后键入或选择 " **Dev Server**"。</span><span class="sxs-lookup"><span data-stu-id="0f58b-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="sideload-your-add-in"></a><span data-ttu-id="0f58b-172">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="0f58b-172">Sideload your add-in</span></span>   

1. <span data-ttu-id="0f58b-173">打开 [Microsoft Office Online](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-173">Open [Microsoft Office Online](https://office.live.com/).</span></span>
2. <span data-ttu-id="0f58b-174">打开一个新的 Excel 工作簿。</span><span class="sxs-lookup"><span data-stu-id="0f58b-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="0f58b-175">打开功能区上的“**插入**”选项卡，然后在“**外接程序**”部分中，选择“**Office 外接程序**”。</span><span class="sxs-lookup"><span data-stu-id="0f58b-175">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="0f58b-176">在“Office 加载项”\*\*\*\* 对话框中，依次选择“我的加载项”\*\*\*\* 选项卡、“管理我的加载项”\*\*\*\* 和“上传我的加载项”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="0f58b-176">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="0f58b-178">**转到**加载项清单文件，再选择“上传”\*\*\*\*。</span><span class="sxs-lookup"><span data-stu-id="0f58b-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="0f58b-180">在旁加载文档后, 每次打开文档时它都将保留旁加载。</span><span class="sxs-lookup"><span data-stu-id="0f58b-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="0f58b-181">开始调试</span><span class="sxs-lookup"><span data-stu-id="0f58b-181">Start debugging</span></span>

1. <span data-ttu-id="0f58b-182">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="0f58b-182">Open developer tools in the browser.</span></span> <span data-ttu-id="0f58b-183">对于 Chrome 和大多数浏览器 F12 将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="0f58b-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="0f58b-184">在开发人员工具中, 使用**Cmd + p**或**Ctrl + p**打开源代码脚本文件 (函数 .js 或函数)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (functions.js or functions.ts).</span></span>
3. <span data-ttu-id="0f58b-185">在自定义函数源代码中[设置断点](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="0f58b-186">如果您需要更改代码, 您可以在 VS 代码中进行编辑并保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="0f58b-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="0f58b-187">刷新浏览器以查看加载的更改。</span><span class="sxs-lookup"><span data-stu-id="0f58b-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="0f58b-188">使用命令行工具进行调试</span><span class="sxs-lookup"><span data-stu-id="0f58b-188">Use the command line tools to debug</span></span>

<span data-ttu-id="0f58b-189">如果未使用 VS 代码, 则可以使用命令行 (如 bash 或 PowerShell) 运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f58b-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="0f58b-190">您需要使用浏览器开发人员工具在 Excel Online 中调试代码。</span><span class="sxs-lookup"><span data-stu-id="0f58b-190">You'll need to use the browser developer tools to debug your code in Excel Online.</span></span> <span data-ttu-id="0f58b-191">无法使用命令行调试桌面版本的 Excel。</span><span class="sxs-lookup"><span data-stu-id="0f58b-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="0f58b-192">在命令行中运行`npm run watch` , 以便在发生代码更改时监视和重建。</span><span class="sxs-lookup"><span data-stu-id="0f58b-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="0f58b-193">打开第二个命令行窗口 (运行监视时将阻止第一个命令行窗口。)</span><span class="sxs-lookup"><span data-stu-id="0f58b-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="0f58b-194">如果要在 Excel 的桌面版本中启动外接程序, 请运行以下命令</span><span class="sxs-lookup"><span data-stu-id="0f58b-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start desktop`
    
    <span data-ttu-id="0f58b-195">或者, 如果您更愿意在 Excel Online 中启动加载项, 请运行以下命令</span><span class="sxs-lookup"><span data-stu-id="0f58b-195">Or if you prefer to start your add-in in Excel Online run the following command</span></span>
    
    `npm run start web`
    
    <span data-ttu-id="0f58b-196">对于 Excel Online, 还需要旁加载您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f58b-196">For Excel Online you also need to sideload your add-in.</span></span> <span data-ttu-id="0f58b-197">按照[旁加载您的外接程序](#Sideload-your-add-in)中的步骤, 旁加载你的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f58b-197">Follow the steps in [Sideload your add-in](#Sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="0f58b-198">然后继续转到下一节以开始调试。</span><span class="sxs-lookup"><span data-stu-id="0f58b-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="0f58b-199">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="0f58b-199">Open developer tools in the browser.</span></span> <span data-ttu-id="0f58b-200">对于 Chrome 和大多数浏览器 F12 将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="0f58b-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="0f58b-201">在开发人员工具中, 打开源代码脚本文件 (函数 .js 或函数)。</span><span class="sxs-lookup"><span data-stu-id="0f58b-201">In developer tools, open your source code script file (functions.js or functions.ts).</span></span> <span data-ttu-id="0f58b-202">您的自定义函数代码可能位于文件末尾附近。</span><span class="sxs-lookup"><span data-stu-id="0f58b-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="0f58b-203">在自定义函数源代码中, 通过选择一行代码来应用断点。</span><span class="sxs-lookup"><span data-stu-id="0f58b-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="0f58b-204">如果您需要更改代码, 您可以在 Visual Studio 中进行编辑并保存所做的更改。</span><span class="sxs-lookup"><span data-stu-id="0f58b-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="0f58b-205">刷新浏览器以查看加载的更改。</span><span class="sxs-lookup"><span data-stu-id="0f58b-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="0f58b-206">用于生成和运行外接程序的命令</span><span class="sxs-lookup"><span data-stu-id="0f58b-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="0f58b-207">有几个可用的生成任务:</span><span class="sxs-lookup"><span data-stu-id="0f58b-207">There are several build tasks available:</span></span>
- `npm run watch`<span data-ttu-id="0f58b-208">: 用于开发的构建, 在保存源文件时自动重建</span><span class="sxs-lookup"><span data-stu-id="0f58b-208">: builds for development and automatically rebuilds when a source file is saved</span></span>
- `npm run build-dev`<span data-ttu-id="0f58b-209">: 开发一次开发版本</span><span class="sxs-lookup"><span data-stu-id="0f58b-209">: builds for development once</span></span>
- `npm run build`<span data-ttu-id="0f58b-210">: 生产的内部版本</span><span class="sxs-lookup"><span data-stu-id="0f58b-210">: builds for production</span></span>
- `npm run dev-server`<span data-ttu-id="0f58b-211">: 运行用于开发的 web 服务器</span><span class="sxs-lookup"><span data-stu-id="0f58b-211">: runs the web server used for development</span></span>

<span data-ttu-id="0f58b-212">您可以使用以下任务在桌面或联机时开始调试。</span><span class="sxs-lookup"><span data-stu-id="0f58b-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- `npm run start desktop`<span data-ttu-id="0f58b-213">: 在桌面上启动 Excel 并将您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f58b-213">: Starts Excel on desktop and sideloads your add-in.</span></span>
- `npm run start web`<span data-ttu-id="0f58b-214">: 启动 Excel Online 并将您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="0f58b-214">: Starts Excel Online and sideloads your add-in.</span></span>
- `npm run stop`<span data-ttu-id="0f58b-215">: 停止 Excel 和调试。</span><span class="sxs-lookup"><span data-stu-id="0f58b-215">: Stops Excel and debugging.</span></span>

## <a name="see-also"></a><span data-ttu-id="0f58b-216">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0f58b-216">See also</span></span>

* [<span data-ttu-id="0f58b-217">自定义函数元数据</span><span class="sxs-lookup"><span data-stu-id="0f58b-217">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0f58b-218">Excel 自定义函数的运行时</span><span class="sxs-lookup"><span data-stu-id="0f58b-218">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0f58b-219">自定义函数最佳实践</span><span class="sxs-lookup"><span data-stu-id="0f58b-219">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0f58b-220">自定义函数更改日志</span><span class="sxs-lookup"><span data-stu-id="0f58b-220">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="0f58b-221">Excel 自定义函数教程</span><span class="sxs-lookup"><span data-stu-id="0f58b-221">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
