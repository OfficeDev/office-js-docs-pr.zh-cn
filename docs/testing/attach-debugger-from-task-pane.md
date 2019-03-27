---
title: 从任务窗格附加调试器
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 941a67695f573cbf0955038916cb1e7ada9cc08f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870427"
---
# <a name="attach-a-debugger-from-the-task-pane"></a><span data-ttu-id="6e084-102">从任务窗格附加调试器</span><span class="sxs-lookup"><span data-stu-id="6e084-102">Attach a debugger from the task pane</span></span>

<span data-ttu-id="6e084-p101">在 Office 2016 for Windows 生成号 77xx.xxxx 或更高版本中，可以从任务窗格附加调试器。使用附加调试器功能，可直接将调试器附加到正确的 Internet Explorer 进程中。无论你使用的是 Yeoman 生成器、Visual Studio Code、node.js、Angular 还是其他任何工具，都可以附加调试器。</span><span class="sxs-lookup"><span data-stu-id="6e084-p101">In Office 2016 for Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, node.js, Angular, or another tool.</span></span> 

<span data-ttu-id="6e084-106">若要启动“**附加调试器**”工具，选择任务窗格右上角来激活“**个性**”菜单，如下图红圈所示。</span><span class="sxs-lookup"><span data-stu-id="6e084-106">To launch the **Attach Debugger** tool, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).</span></span>   

> [!NOTE]
> - <span data-ttu-id="6e084-p102">目前，唯一受支持的调试器工具是 [Visual Studio 2015](https://www.visualstudio.com/downloads/) [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) 或更高版本。如果没有安装 Visual Studio，选择“附加调试器”\*\*\*\* 选项不会有任何结果。</span><span class="sxs-lookup"><span data-stu-id="6e084-p102">Currently the only supported debugger tool is [Visual Studio 2015](https://www.visualstudio.com/downloads/) with [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) or later. If you don't have Visual Studio installed, selecting the **Attach Debugger** option doesn’t result in any action.</span></span>   
> - <span data-ttu-id="6e084-109">只能使用“附加调试器”\*\*\*\* 工具调试客户端 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="6e084-109">You can only debug client-side JavaScript with the **Attach Debugger** tool.</span></span> <span data-ttu-id="6e084-110">要调试服务器端代码（例如 Node.js 服务器），可选择多种方式。</span><span class="sxs-lookup"><span data-stu-id="6e084-110">To debug server-side code, such as with a Node.js server, you have many options.</span></span> <span data-ttu-id="6e084-111">有关如何使用 Visual Studio Code 进行调试的信息，请参阅 [VS Code 中的 Node.js 调试](https://code.visualstudio.com/docs/nodejs/nodejs-debugging)。</span><span class="sxs-lookup"><span data-stu-id="6e084-111">For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging).</span></span> <span data-ttu-id="6e084-112">如果没有使用 Visual Studio Code，请搜索“debug Node.js”或“debug {name-of-server}”。</span><span class="sxs-lookup"><span data-stu-id="6e084-112">If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".</span></span>

![“附加调试器”菜单屏幕截图](../images/attach-debugger.png)

<span data-ttu-id="6e084-p104">选择“**附加调试器**”。此操作将启动“**Visual Studio 实时调试器**”对话框，如下图所示。</span><span class="sxs-lookup"><span data-stu-id="6e084-p104">Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image.</span></span> 

![“Visual Studio JIT 调试器”对话框屏幕截图](../images/visual-studio-debugger.png)

<span data-ttu-id="6e084-p105">Visual Studio 中的“解决方案资源管理器”\*\*\*\* 会显示代码文件。   可以在要使用 Visual Studio 调试的代码行处设置断点。</span><span class="sxs-lookup"><span data-stu-id="6e084-p105">In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="6e084-119">如果你没有看到“个性”菜单，则可以使用 Visual Studio 调试加载项。</span><span class="sxs-lookup"><span data-stu-id="6e084-119">If you don't see the Personality menu, you can debug your add-in using Visual Studio.</span></span> <span data-ttu-id="6e084-120">确保你的任务窗格加载项已在 Office 中打开，然后按照以下步骤操作：</span><span class="sxs-lookup"><span data-stu-id="6e084-120">Ensure your task pane add-in is open in Office, and then follow these steps:</span></span>

> 1. <span data-ttu-id="6e084-121">在 Visual Studio 中，依次选择“**调试**” > “**附加到进程**”。</span><span class="sxs-lookup"><span data-stu-id="6e084-121">In Visual Studio, choose **DEBUG** > **Attach to Process**.</span></span>
> 2. <span data-ttu-id="6e084-122">在“**附加到进程**”对话框中，选择所有可用的 Iexplore.exe 进程，然后选择“**附加**”按钮。</span><span class="sxs-lookup"><span data-stu-id="6e084-122">In **Attach to Process**, choose all of the available Iexplore.exe processes, and then choose the **Attach** button.</span></span>

<span data-ttu-id="6e084-123">若要详细了解如何在 Visual Studio 中进行调试，请参阅以下内容：</span><span class="sxs-lookup"><span data-stu-id="6e084-123">For more information about debugging in Visual Studio, see the following:</span></span>

-   <span data-ttu-id="6e084-124">若要在 Visual Studio 中启动并使用 DOM 资源管理器，请参阅 [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates)（使用新项目模板为 Office 生成漂亮应用）博客文章中[提示和技巧](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks)部分的提示 4。</span><span class="sxs-lookup"><span data-stu-id="6e084-124">To launch and use the DOM Explorer in Visual Studio, see Tip 4 in the [Tips and Tricks](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) section of the [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) blog post.</span></span>
-   <span data-ttu-id="6e084-125">若要设置断点，请参阅[使用断点](/visualstudio/debugger/using-breakpoints?view=vs-2015)。</span><span class="sxs-lookup"><span data-stu-id="6e084-125">To set breakpoints, see [Using Breakpoints](/visualstudio/debugger/using-breakpoints?view=vs-2015).</span></span>
-   <span data-ttu-id="6e084-126">若要使用 F12，请参阅[使用 F12 开发人员工具](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))。</span><span class="sxs-lookup"><span data-stu-id="6e084-126">To use F12, see [Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span></span>

## <a name="see-also"></a><span data-ttu-id="6e084-127">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6e084-127">See also</span></span>

- [<span data-ttu-id="6e084-128">在 Visual Studio 中创建和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6e084-128">Create and debug Office Add-ins in Visual Studio</span></span>](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [<span data-ttu-id="6e084-129">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6e084-129">Publish your Office Add-in</span></span>](../publish/publish.md)
