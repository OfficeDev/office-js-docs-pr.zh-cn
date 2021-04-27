---
ms.date: 04/12/2021
description: 了解如何调试不使用任务窗格的 Excel 自定义函数。
title: 无 UI 自定义函数调试
localization_priority: Normal
ms.openlocfilehash: c6954af4638ae416c789af339d35187467e37b7f
ms.sourcegitcommit: 78fb861afe7d7c3ee7fe3186150b3fed20994222
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/26/2021
ms.locfileid: "52024323"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="23172-103">无 UI 自定义函数调试</span><span class="sxs-lookup"><span data-stu-id="23172-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="23172-104">本文仅讨论不使用任务窗格或其他用户界面元素的自定义函数的调试 (无 UI 自定义函数) 。</span><span class="sxs-lookup"><span data-stu-id="23172-104">This article discusses debugging *only* for custom functions that don't use a task pane or other user interface elements (UI-less custom functions).</span></span> 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="23172-105">在 Windows 上：</span><span class="sxs-lookup"><span data-stu-id="23172-105">On Windows:</span></span>
- [<span data-ttu-id="23172-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span><span class="sxs-lookup"><span data-stu-id="23172-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="23172-107">Excel 网页和 VS 代码调试程序</span><span class="sxs-lookup"><span data-stu-id="23172-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="23172-108">Excel 网页和浏览器工具</span><span class="sxs-lookup"><span data-stu-id="23172-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="23172-109">命令行</span><span class="sxs-lookup"><span data-stu-id="23172-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="23172-110">在 Mac 上：</span><span class="sxs-lookup"><span data-stu-id="23172-110">On Mac:</span></span>
- [<span data-ttu-id="23172-111">Excel 网页和浏览器工具</span><span class="sxs-lookup"><span data-stu-id="23172-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="23172-112">命令行</span><span class="sxs-lookup"><span data-stu-id="23172-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="23172-113">为简单起见，本文介绍在使用 Visual Studio 代码编辑、运行任务的情况下进行调试，在某些情况下，还使用调试视图。</span><span class="sxs-lookup"><span data-stu-id="23172-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="23172-114">如果使用的是其他编辑器或命令行工具，请参阅本文末尾的命令行说明[](#commands-for-building-and-running-your-add-in)。</span><span class="sxs-lookup"><span data-stu-id="23172-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="23172-115">要求</span><span class="sxs-lookup"><span data-stu-id="23172-115">Requirements</span></span>

<span data-ttu-id="23172-116">此调试过程 **仅适用于不使用** 任务窗格或其他 UI 元素的无 UI 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="23172-116">This debugging process works **only** for UI-less custom functions, which don't use a task pane or other UI elements.</span></span> <span data-ttu-id="23172-117">可以按照在 [Excel](../tutorials/excel-tutorial-create-custom-functions.md) 中创建自定义函数教程中的步骤创建无 UI 自定义函数，然后删除由适用于 Office 外接程序的 [Yeoman](https://www.npmjs.com/package/generator-office)生成器安装的所有任务窗格和 UI 元素。</span><span class="sxs-lookup"><span data-stu-id="23172-117">A UI-less custom function can be created by following the steps in the [Create custom functions in Excel](../tutorials/excel-tutorial-create-custom-functions.md) tutorial, and then removing all of the task pane and UI elements that are installed by the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span>

<span data-ttu-id="23172-118">请注意，此调试过程与使用共享运行时 的自定义函数 [项目不兼容](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="23172-118">Note that this debugging process is not compatible with custom functions projects using a [shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="23172-119">使用适用于 Excel Desktop 的 VS 代码调试程序</span><span class="sxs-lookup"><span data-stu-id="23172-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="23172-120">您可以使用 VS Code 在桌面上的 Office Excel 中调试无 UI 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="23172-120">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="23172-121">适用于 Mac 的桌面调试不可用，但可以使用浏览器工具和命令行来调试 [Excel 网页](#use-the-command-line-tools-to-debug) 版) 。</span><span class="sxs-lookup"><span data-stu-id="23172-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="23172-122">从 VS Code 运行加载项</span><span class="sxs-lookup"><span data-stu-id="23172-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="23172-123">在 VS Code 中打开自定义函数根项目 [文件夹](https://code.visualstudio.com/)。</span><span class="sxs-lookup"><span data-stu-id="23172-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="23172-124">选择 **"终端>运行任务**"，然后键入或选择"**监视"。**</span><span class="sxs-lookup"><span data-stu-id="23172-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="23172-125">这将监视并重新生成任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="23172-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="23172-126">选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**</span><span class="sxs-lookup"><span data-stu-id="23172-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="23172-127">启动 VS 代码调试程序</span><span class="sxs-lookup"><span data-stu-id="23172-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="23172-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span><span class="sxs-lookup"><span data-stu-id="23172-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="23172-129">From the Run drop-down menu， choose **Excel Desktop (Custom Functions)**.</span><span class="sxs-lookup"><span data-stu-id="23172-129">From the Run drop-down menu, choose **Excel Desktop (Custom Functions)**.</span></span>
6. <span data-ttu-id="23172-130">选择 **F5** (，或者从 **>开始调试** "菜单中选择") 开始调试"。</span><span class="sxs-lookup"><span data-stu-id="23172-130">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="23172-131">新的 Excel 工作簿将在外接程序已旁加载且可供使用时打开。</span><span class="sxs-lookup"><span data-stu-id="23172-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="23172-132">开始调试</span><span class="sxs-lookup"><span data-stu-id="23172-132">Start debugging</span></span>

1. <span data-ttu-id="23172-133">在 VS Code 中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。</span><span class="sxs-lookup"><span data-stu-id="23172-133">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="23172-134">[在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="23172-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="23172-135">在 Excel 工作簿中，输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="23172-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="23172-136">此时，将在设置断点的代码行上停止执行。</span><span class="sxs-lookup"><span data-stu-id="23172-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="23172-137">现在，你可以逐步调试代码、设置监视以及使用所需的任何 VS 代码调试功能。</span><span class="sxs-lookup"><span data-stu-id="23172-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="23172-138">在 Microsoft Edge 中为 Excel 使用 VS 代码调试程序</span><span class="sxs-lookup"><span data-stu-id="23172-138">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="23172-139">您可以使用 VS Code 在 Microsoft Edge 浏览器的 Excel 中调试无 UI 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="23172-139">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="23172-140">若要将 VS Code 与 Microsoft Edge 一同使用，必须安装 [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) 扩展。</span><span class="sxs-lookup"><span data-stu-id="23172-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="23172-141">从 VS Code 运行加载项</span><span class="sxs-lookup"><span data-stu-id="23172-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="23172-142">在 VS Code 中打开自定义函数根项目 [文件夹](https://code.visualstudio.com/)。</span><span class="sxs-lookup"><span data-stu-id="23172-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="23172-143">选择 **"终端>运行任务**"，然后键入或选择"**监视"。**</span><span class="sxs-lookup"><span data-stu-id="23172-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="23172-144">这将监视并重新生成任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="23172-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="23172-145">选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**</span><span class="sxs-lookup"><span data-stu-id="23172-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="23172-146">启动 VS 代码调试程序</span><span class="sxs-lookup"><span data-stu-id="23172-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="23172-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span><span class="sxs-lookup"><span data-stu-id="23172-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="23172-148">从"调试"选项中，选择 **"Office Online (Edge Chromium) "。**</span><span class="sxs-lookup"><span data-stu-id="23172-148">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="23172-149">在 Microsoft Edge 浏览器中打开 Excel 并创建新的工作簿。</span><span class="sxs-lookup"><span data-stu-id="23172-149">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="23172-150">在 **功能** 区中选择"共享"，并复制此新工作簿的 URL 链接。</span><span class="sxs-lookup"><span data-stu-id="23172-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="23172-151">选择 **F5** (**或从>** 开始调试"菜单中选择") 开始调试"。</span><span class="sxs-lookup"><span data-stu-id="23172-151">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="23172-152">将出现一个提示，询问文档的 URL。</span><span class="sxs-lookup"><span data-stu-id="23172-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="23172-153">粘贴工作簿的 URL，然后按 Enter。</span><span class="sxs-lookup"><span data-stu-id="23172-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="23172-154">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="23172-154">Sideload your add-in</span></span>

1. <span data-ttu-id="23172-155">选择功能 **区** 上的"插入"选项卡，在 **"外接程序"** 部分，选择 **"Office 外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="23172-155">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="23172-156">在 **"Office 外接程序"** 对话框中，选择 **"我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，然后选择"**上载我的外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="23172-156">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="23172-158">**浏览** 到外接程序清单文件， **然后选择上载**。</span><span class="sxs-lookup"><span data-stu-id="23172-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="23172-160">设置断点</span><span class="sxs-lookup"><span data-stu-id="23172-160">Set breakpoints</span></span>
1. <span data-ttu-id="23172-161">在 VS Code 中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。</span><span class="sxs-lookup"><span data-stu-id="23172-161">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="23172-162">[在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="23172-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="23172-163">在 Excel 工作簿中，输入使用自定义函数的公式。</span><span class="sxs-lookup"><span data-stu-id="23172-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="23172-164">使用浏览器开发人员工具调试 Excel 网页版中的自定义函数</span><span class="sxs-lookup"><span data-stu-id="23172-164">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="23172-165">可以使用浏览器开发人员工具在 Excel 网页版中调试无 UI 自定义函数。</span><span class="sxs-lookup"><span data-stu-id="23172-165">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="23172-166">以下步骤适用于 Windows 和 macOS。</span><span class="sxs-lookup"><span data-stu-id="23172-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="23172-167">从代码运行Visual Studio加载项</span><span class="sxs-lookup"><span data-stu-id="23172-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="23172-168">打开自定义函数根项目文件夹，Visual Studio [代码 ](https://code.visualstudio.com/) (VS Code) 。</span><span class="sxs-lookup"><span data-stu-id="23172-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="23172-169">选择 **"终端>运行任务**"，然后键入或选择"**监视"。**</span><span class="sxs-lookup"><span data-stu-id="23172-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="23172-170">这将监视并重新生成任何文件更改。</span><span class="sxs-lookup"><span data-stu-id="23172-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="23172-171">选择 **"终端>运行任务**"，然后键入或选择 **"开发人员服务器"。**</span><span class="sxs-lookup"><span data-stu-id="23172-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="23172-172">旁加载加载项</span><span class="sxs-lookup"><span data-stu-id="23172-172">Sideload your add-in</span></span>

1. <span data-ttu-id="23172-173">在[Web 上打开 Office。](https://office.live.com/)</span><span class="sxs-lookup"><span data-stu-id="23172-173">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="23172-174">打开一个新的 Excel 工作簿。</span><span class="sxs-lookup"><span data-stu-id="23172-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="23172-175">打开功能 **区** 上的"插入"选项卡，在"**外接程序**"部分，选择 **"Office 外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="23172-175">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="23172-176">在 **"Office 外接程序"** 对话框中，选择 **"我的** 外接程序"选项卡，选择"管理 **我的** 外接程序"，然后选择"**上载我的外接程序"。**</span><span class="sxs-lookup"><span data-stu-id="23172-176">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![“Office 加载项”对话框，右上方有“管理我的加载项”下拉列表，其中有下拉选项“上传我的加载项”](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="23172-178">**转到** 加载项清单文件，再选择“上传”。</span><span class="sxs-lookup"><span data-stu-id="23172-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![带浏览、上载和取消按钮的上载外接程序对话框。](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="23172-180">旁加载文档后，每次打开文档时，文档都会保持旁加载状态。</span><span class="sxs-lookup"><span data-stu-id="23172-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="23172-181">开始调试</span><span class="sxs-lookup"><span data-stu-id="23172-181">Start debugging</span></span>

1. <span data-ttu-id="23172-182">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="23172-182">Open developer tools in the browser.</span></span> <span data-ttu-id="23172-183">对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="23172-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="23172-184">在开发人员工具中，使用 **Cmd+P** 或 **Ctrl+P** (functions.js或 **functions.ts**) 。</span><span class="sxs-lookup"><span data-stu-id="23172-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="23172-185">[在自定义函数](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) 源代码中设置断点。</span><span class="sxs-lookup"><span data-stu-id="23172-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="23172-186">如果需要更改代码，可以在 VS Code 中编辑并保存更改。</span><span class="sxs-lookup"><span data-stu-id="23172-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="23172-187">刷新浏览器以查看已加载的更改。</span><span class="sxs-lookup"><span data-stu-id="23172-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="23172-188">使用命令行工具进行调试</span><span class="sxs-lookup"><span data-stu-id="23172-188">Use the command line tools to debug</span></span>

<span data-ttu-id="23172-189">如果不使用 VS Code，可以使用命令行命令 (Bash 或 PowerShell) 运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="23172-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="23172-190">你需要使用浏览器开发人员工具在 Excel 网页版中调试代码。</span><span class="sxs-lookup"><span data-stu-id="23172-190">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="23172-191">不能使用命令行调试桌面版 Excel。</span><span class="sxs-lookup"><span data-stu-id="23172-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="23172-192">从命令行运行 `npm run watch` 以观察代码发生更改时并重新生成代码。</span><span class="sxs-lookup"><span data-stu-id="23172-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="23172-193">打开第二个命令行窗口 (运行 watch.) </span><span class="sxs-lookup"><span data-stu-id="23172-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="23172-194">如果要在桌面版 Excel 中启动加载项，请运行以下命令</span><span class="sxs-lookup"><span data-stu-id="23172-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="23172-195">或者，如果你想要在 Excel 网页中启动加载项，请运行以下命令</span><span class="sxs-lookup"><span data-stu-id="23172-195">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="23172-196">对于 Excel 网页应用，还需要旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="23172-196">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="23172-197">按照旁加载 [加载项中的步骤](#sideload-your-add-in) 旁加载加载项。</span><span class="sxs-lookup"><span data-stu-id="23172-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="23172-198">然后继续下一部分以开始调试。</span><span class="sxs-lookup"><span data-stu-id="23172-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="23172-199">在浏览器中打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="23172-199">Open developer tools in the browser.</span></span> <span data-ttu-id="23172-200">对于 Chrome 和大多数浏览器 F12，将打开开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="23172-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="23172-201">在开发人员工具中，打开源代码脚本文件 (functions.js **或 functions.ts**) 。 </span><span class="sxs-lookup"><span data-stu-id="23172-201">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="23172-202">自定义函数代码可能位于文件的末尾附近。</span><span class="sxs-lookup"><span data-stu-id="23172-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="23172-203">在自定义函数源代码中，通过选择一行代码来应用断点。</span><span class="sxs-lookup"><span data-stu-id="23172-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="23172-204">如果需要更改代码，可以在该代码中进行Visual Studio并保存更改。</span><span class="sxs-lookup"><span data-stu-id="23172-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="23172-205">刷新浏览器以查看已加载的更改。</span><span class="sxs-lookup"><span data-stu-id="23172-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="23172-206">用于生成和运行加载项的命令</span><span class="sxs-lookup"><span data-stu-id="23172-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="23172-207">有几种可用的生成任务：</span><span class="sxs-lookup"><span data-stu-id="23172-207">There are several build tasks available:</span></span>
- <span data-ttu-id="23172-208">`npm run watch`：用于开发内部版本，在保存源文件时自动重新生成</span><span class="sxs-lookup"><span data-stu-id="23172-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="23172-209">`npm run build-dev`：生成一次用于开发</span><span class="sxs-lookup"><span data-stu-id="23172-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="23172-210">`npm run build`：用于生产内部版本</span><span class="sxs-lookup"><span data-stu-id="23172-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="23172-211">`npm run dev-server`：运行用于开发的 Web 服务器</span><span class="sxs-lookup"><span data-stu-id="23172-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="23172-212">可以使用以下任务在桌面或联机上开始调试。</span><span class="sxs-lookup"><span data-stu-id="23172-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="23172-213">`npm run start:desktop`：在桌面上启动 Excel 并旁加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="23172-213">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="23172-214">`npm run start:web`：在 Web 上启动 Excel 并旁加载外接程序。</span><span class="sxs-lookup"><span data-stu-id="23172-214">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="23172-215">`npm run stop`：停止 Excel 和调试。</span><span class="sxs-lookup"><span data-stu-id="23172-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="23172-216">后续步骤</span><span class="sxs-lookup"><span data-stu-id="23172-216">Next steps</span></span>
<span data-ttu-id="23172-217">了解 [无 UI 自定义函数的身份验证做法](custom-functions-authentication.md)。</span><span class="sxs-lookup"><span data-stu-id="23172-217">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="23172-218">另请参阅</span><span class="sxs-lookup"><span data-stu-id="23172-218">See also</span></span>

* [<span data-ttu-id="23172-219">自定义函数疑难解答</span><span class="sxs-lookup"><span data-stu-id="23172-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="23172-220">在 Excel 中处理自定义函数时出错</span><span class="sxs-lookup"><span data-stu-id="23172-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="23172-221">在 Excel 中创建自定义函数</span><span class="sxs-lookup"><span data-stu-id="23172-221">Create custom functions in Excel</span></span>](custom-functions-overview.md)
