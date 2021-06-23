---
ms.date: 04/12/2021
description: 了解如何调试不使用Excel窗格的自定义函数。
title: 无 UI 自定义函数调试
localization_priority: Normal
ms.openlocfilehash: a692f376cb5c874fa4d510d3459469d803e643f7
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075934"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="13b12-103">无 UI 自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="13b12-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="13b12-104">本文仅讨论不使用任务窗格或其他用户界面元素的自定义函数的调试 (无 UI 自定义函数) 。</span><span class="sxs-lookup"><span data-stu-id="13b12-104">This article discusses debugging *only* for custom functions that don't use a task pane or other user interface elements (UI-less custom functions).</span></span> 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="13b12-105">在Windows：</span><span class="sxs-lookup"><span data-stu-id="13b12-105">On Windows:</span></span>
- [<span data-ttu-id="13b12-106">Excel桌面和Visual Studio Code (VS Code) 调试器</span><span class="sxs-lookup"><span data-stu-id="13b12-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="13b12-107">Excel web 版调试VS Code和调试器</span><span class="sxs-lookup"><span data-stu-id="13b12-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="13b12-108">Excel web 版和浏览器工具</span><span class="sxs-lookup"><span data-stu-id="13b12-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="13b12-109">命令行</span><span class="sxs-lookup"><span data-stu-id="13b12-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="13b12-110">在 Mac 上：</span><span class="sxs-lookup"><span data-stu-id="13b12-110">On Mac:</span></span>
- [<span data-ttu-id="13b12-111">Excel web 版和浏览器工具</span><span class="sxs-lookup"><span data-stu-id="13b12-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="13b12-112">命令行</span><span class="sxs-lookup"><span data-stu-id="13b12-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="13b12-113">为简单起见，本文介绍在使用 Visual Studio Code编辑、运行任务的情况下进行调试，在某些情况下，还使用调试视图。</span><span class="sxs-lookup"><span data-stu-id="13b12-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="13b12-114">如果使用的是其他编辑器或命令行工具，请参阅本文末尾的命令行说明[](#commands-for-building-and-running-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="13b12-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="13b12-115">Requirements</span><span class="sxs-lookup"><span data-stu-id="13b12-115">Requirements</span></span>

<span data-ttu-id="13b12-116">此调试过程 **仅适用于不使用** 任务窗格或其他 UI 元素的无 UI 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="13b12-116">This debugging process works **only** for UI-less custom functions, which don't use a task pane or other UI elements.</span></span> <span data-ttu-id="13b12-117">可以按照在 Excel 中创建自定义函数教程中的步骤创建无 UI 自定义函数，然后删除由适用于[Office](../tutorials/excel-tutorial-create-custom-functions.md)加载项的[Yeoman](https://www.npmjs.com/package/generator-office)生成器安装的所有任务窗格和 UI 元素。</span><span class="sxs-lookup"><span data-stu-id="13b12-117">A UI-less custom function can be created by following the steps in the [Create custom functions in Excel](../tutorials/excel-tutorial-create-custom-functions.md) tutorial, and then removing all of the task pane and UI elements that are installed by the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span>

<span data-ttu-id="13b12-118">请注意，此调试过程与使用共享运行时 的自定义函数 [项目不兼容](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="13b12-118">Note that this debugging process is not compatible with custom functions projects using a [shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="13b12-119">使用 VS Code 桌面版Excel调试程序</span><span class="sxs-lookup"><span data-stu-id="13b12-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="13b12-120">可以使用 VS Code调试桌面上的 Office Excel UI 无 UI 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="13b12-120">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="13b12-121">适用于 Mac 的桌面调试不可用，但可以使用浏览器工具和命令行[](#use-the-command-line-tools-to-debug)来调试Excel web 版) 。</span><span class="sxs-lookup"><span data-stu-id="13b12-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="13b12-122">从应用程序运行VS Code</span><span class="sxs-lookup"><span data-stu-id="13b12-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="13b12-123">在 中打开自定义函数根项目[VS Code。](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="13b12-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="13b12-124">选择 **"终端>运行任务**"，然后键入或选择"**监视"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="13b12-125">这将监视并重新生成任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="13b12-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="13b12-126">选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="13b12-127">启动VS Code调试器</span><span class="sxs-lookup"><span data-stu-id="13b12-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="13b12-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span><span class="sxs-lookup"><span data-stu-id="13b12-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="13b12-129">From the Run drop-down menu， choose **Excel Desktop (Custom Functions)**.</span><span class="sxs-lookup"><span data-stu-id="13b12-129">From the Run drop-down menu, choose **Excel Desktop (Custom Functions)**.</span></span>
6. <span data-ttu-id="13b12-130">选择 **F5** (，或者从 **>开始调试** "菜单中选择") 开始调试"。</span><span class="sxs-lookup"><span data-stu-id="13b12-130">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="13b12-131">新Excel工作簿将打开，同时加载项已旁加载并可供使用。</span><span class="sxs-lookup"><span data-stu-id="13b12-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="13b12-132">开始调试</span><span class="sxs-lookup"><span data-stu-id="13b12-132">Start debugging</span></span>

1. <span data-ttu-id="13b12-133">在VS Code中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。 </span><span class="sxs-lookup"><span data-stu-id="13b12-133">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="13b12-134">[在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="13b12-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="13b12-135">在Excel工作簿中，输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="13b12-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="13b12-136">此时，将在设置断点的代码行上停止执行。</span><span class="sxs-lookup"><span data-stu-id="13b12-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="13b12-137">现在，你可以逐步调试代码、设置监视以及使用VS Code调试功能。</span><span class="sxs-lookup"><span data-stu-id="13b12-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="13b12-138">在 VS Code 中Excel调试Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="13b12-138">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="13b12-139">可以使用VS Code在浏览器上的 Excel 调试无 UI Microsoft Edge函数。</span><span class="sxs-lookup"><span data-stu-id="13b12-139">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="13b12-140">若要将 VS Code 与 Microsoft Edge 一起使用，必须安装适用于 Microsoft Edge 扩展[的调试](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)程序。</span><span class="sxs-lookup"><span data-stu-id="13b12-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="13b12-141">从应用程序运行VS Code</span><span class="sxs-lookup"><span data-stu-id="13b12-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="13b12-142">在 中打开自定义函数根项目[VS Code。](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="13b12-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="13b12-143">选择 **"终端>运行任务**"，然后键入或选择"**监视"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="13b12-144">这将监视并重新生成任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="13b12-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="13b12-145">选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="13b12-146">启动VS Code调试器</span><span class="sxs-lookup"><span data-stu-id="13b12-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="13b12-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span><span class="sxs-lookup"><span data-stu-id="13b12-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="13b12-148">从"调试"选项中，选择 **"Office Online (Edge Chromium) "。**</span><span class="sxs-lookup"><span data-stu-id="13b12-148">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="13b12-149">在Excel中打开Microsoft Edge新建工作簿。</span><span class="sxs-lookup"><span data-stu-id="13b12-149">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="13b12-150">在 **功能** 区中选择"共享"，并复制此新工作簿的 URL 链接。</span><span class="sxs-lookup"><span data-stu-id="13b12-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="13b12-151">选择 **F5** (**或从>** 开始调试"菜单中选择") 开始调试"。</span><span class="sxs-lookup"><span data-stu-id="13b12-151">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="13b12-152">将出现一个提示，询问文档的 URL。</span><span class="sxs-lookup"><span data-stu-id="13b12-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="13b12-153">粘贴工作簿的 URL，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="13b12-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="13b12-154">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="13b12-154">Sideload your add-in</span></span>

1. <span data-ttu-id="13b12-155">选择功能 **区** 上的"插入"选项卡，在"外接程序"部分，选择"Office **外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-155">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="13b12-156">在 **"Office** 外接程序"对话框中，选择"**我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，Upload"**我的外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-156">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![the Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="13b12-158">**浏览** 到外接程序清单文件，然后选择 **"Upload"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="13b12-160">设置断点</span><span class="sxs-lookup"><span data-stu-id="13b12-160">Set breakpoints</span></span>
1. <span data-ttu-id="13b12-161">在VS Code中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。 </span><span class="sxs-lookup"><span data-stu-id="13b12-161">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="13b12-162">[在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="13b12-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="13b12-163">在Excel工作簿中，输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="13b12-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="13b12-164">使用浏览器开发人员工具在浏览器中调试自定义Excel web 版</span><span class="sxs-lookup"><span data-stu-id="13b12-164">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="13b12-165">可以使用浏览器开发人员工具在浏览器中调试无 UI Excel web 版。</span><span class="sxs-lookup"><span data-stu-id="13b12-165">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="13b12-166">以下步骤适用于 Windows 和 macOS。</span><span class="sxs-lookup"><span data-stu-id="13b12-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="13b12-167">从应用程序运行Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="13b12-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="13b12-168">在 中打开自定义函数根项目[Visual Studio Code (VS Code) 。 ](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="13b12-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="13b12-169">选择 **"终端>运行任务**"，然后键入或选择"**监视"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="13b12-170">这将监视并重新生成任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="13b12-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="13b12-171">选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="13b12-172">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="13b12-172">Sideload your add-in</span></span>

1. <span data-ttu-id="13b12-173">打开[Office web 版](https://office.live.com/)。</span><span class="sxs-lookup"><span data-stu-id="13b12-173">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="13b12-174">打开一个新的Excel工作簿。</span><span class="sxs-lookup"><span data-stu-id="13b12-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="13b12-175">打开功能 **区** 上的"插入"选项卡，在"外接程序"部分，选择"Office **外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-175">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="13b12-176">在 **"Office** 外接程序"对话框中，选择"**我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，Upload"**我的外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="13b12-176">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![the Office Add-ins dialog with a drop-down in the upper right reading "Manage my add-ins" and a drop-down below it with the option "Upload My Add-in".](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="13b12-178">**转到** 加载项清单文件，再选择“上传”。</span><span class="sxs-lookup"><span data-stu-id="13b12-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="13b12-180">旁加载文档后，每次打开文档时，文档都会保持旁加载状态。</span><span class="sxs-lookup"><span data-stu-id="13b12-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="13b12-181">开始调试</span><span class="sxs-lookup"><span data-stu-id="13b12-181">Start debugging</span></span>

1. <span data-ttu-id="13b12-182">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="13b12-182">Open developer tools in the browser.</span></span> <span data-ttu-id="13b12-183">对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="13b12-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="13b12-184">在开发人员工具中，使用 **Cmd+P** 或 **Ctrl+P** (functions.js或 **functions.ts**) 。</span><span class="sxs-lookup"><span data-stu-id="13b12-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="13b12-185">[在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="13b12-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="13b12-186">如果需要更改代码，可以在 VS Code并保存更改。</span><span class="sxs-lookup"><span data-stu-id="13b12-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="13b12-187">刷新浏览器以查看已加载的更改。</span><span class="sxs-lookup"><span data-stu-id="13b12-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="13b12-188">使用命令行工具进行调试</span><span class="sxs-lookup"><span data-stu-id="13b12-188">Use the command line tools to debug</span></span>

<span data-ttu-id="13b12-189">如果未使用 VS Code，可以使用命令行 (如 Bash 或 PowerShell) 运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="13b12-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="13b12-190">你将需要使用浏览器开发人员工具在 Excel web 版 中调试代码。</span><span class="sxs-lookup"><span data-stu-id="13b12-190">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="13b12-191">不能使用命令行调试桌面Excel版本。</span><span class="sxs-lookup"><span data-stu-id="13b12-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="13b12-192">从命令行运行 `npm run watch` 以观察代码发生更改时并重新生成代码。</span><span class="sxs-lookup"><span data-stu-id="13b12-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="13b12-193">打开第二个命令行窗口 (运行 watch.) </span><span class="sxs-lookup"><span data-stu-id="13b12-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="13b12-194">如果要在桌面版本的外接程序中启动Excel，请运行以下命令</span><span class="sxs-lookup"><span data-stu-id="13b12-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="13b12-195">或者，如果你想要在加载项中启动Excel web 版运行以下命令</span><span class="sxs-lookup"><span data-stu-id="13b12-195">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="13b12-196">例如Excel web 版你还需要旁加载你的外接程序。</span><span class="sxs-lookup"><span data-stu-id="13b12-196">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="13b12-197">按照旁加载 [加载项中的步骤](#sideload-your-add-in) 旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="13b12-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="13b12-198">然后继续下一部分以开始调试。</span><span class="sxs-lookup"><span data-stu-id="13b12-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="13b12-199">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="13b12-199">Open developer tools in the browser.</span></span> <span data-ttu-id="13b12-200">对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="13b12-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="13b12-201">在开发人员工具中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。 </span><span class="sxs-lookup"><span data-stu-id="13b12-201">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="13b12-202">自定义函数代码可能位于文件的末尾附近。</span><span class="sxs-lookup"><span data-stu-id="13b12-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="13b12-203">在自定义函数源代码中，通过选择一行代码来应用断点。</span><span class="sxs-lookup"><span data-stu-id="13b12-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="13b12-204">如果需要更改代码，可以在 Visual Studio并保存更改。</span><span class="sxs-lookup"><span data-stu-id="13b12-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="13b12-205">刷新浏览器以查看已加载的更改。</span><span class="sxs-lookup"><span data-stu-id="13b12-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="13b12-206">用于生成和运行加载项的命令</span><span class="sxs-lookup"><span data-stu-id="13b12-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="13b12-207">有几种可用的生成任务：</span><span class="sxs-lookup"><span data-stu-id="13b12-207">There are several build tasks available:</span></span>
- <span data-ttu-id="13b12-208">`npm run watch`：用于开发内部版本，在保存源文件时自动重新生成</span><span class="sxs-lookup"><span data-stu-id="13b12-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="13b12-209">`npm run build-dev`：生成一次用于开发</span><span class="sxs-lookup"><span data-stu-id="13b12-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="13b12-210">`npm run build`：用于生产内部版本</span><span class="sxs-lookup"><span data-stu-id="13b12-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="13b12-211">`npm run dev-server`：运行用于开发的 Web 服务器</span><span class="sxs-lookup"><span data-stu-id="13b12-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="13b12-212">可以使用以下任务在桌面或联机上开始调试。</span><span class="sxs-lookup"><span data-stu-id="13b12-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="13b12-213">`npm run start:desktop`：Excel启动加载项，并旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="13b12-213">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="13b12-214">`npm run start:web`：Excel web 版加载项并旁加载。</span><span class="sxs-lookup"><span data-stu-id="13b12-214">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="13b12-215">`npm run stop`：停止Excel调试。</span><span class="sxs-lookup"><span data-stu-id="13b12-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="13b12-216">后续步骤</span><span class="sxs-lookup"><span data-stu-id="13b12-216">Next steps</span></span>
<span data-ttu-id="13b12-217">了解 [无 UI 自定义函数的身份验证做法](custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="13b12-217">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="13b12-218">另请参阅</span><span class="sxs-lookup"><span data-stu-id="13b12-218">See also</span></span>

* [<span data-ttu-id="13b12-219">自定义函数疑难解答</span><span class="sxs-lookup"><span data-stu-id="13b12-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="13b12-220">在 Excel 中处理自定义函数时出错</span><span class="sxs-lookup"><span data-stu-id="13b12-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="13b12-221">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="13b12-221">Create custom functions in Excel</span></span>](custom-functions-overview.md)
