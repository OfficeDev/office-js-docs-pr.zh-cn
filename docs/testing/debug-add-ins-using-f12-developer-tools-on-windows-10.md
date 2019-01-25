---
title: 在 Windows 10 上使用 F12 开发人员工具调试外接程序
description: ''
ms.date: 10/16/2018
localization_priority: Priority
ms.openlocfilehash: e2378a0449ea33551051b9c3788b84b23a51feb8
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386902"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="9c373-102">在 Windows 10 上使用 F12 开发人员工具调试外接程序</span><span class="sxs-lookup"><span data-stu-id="9c373-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="9c373-103">Windows 10 中随附的 F12 开发人员工具可帮助您调试、测试和加速您的网页。</span><span class="sxs-lookup"><span data-stu-id="9c373-103">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages.</span></span> <span data-ttu-id="9c373-104">如果您未使用 IDE（如 Visual Studio），或者如果您需要调查在 IDE 外部运行外接程序时出现的问题，您还可以使用该工具开发和调试您的 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="9c373-104">You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE.</span></span> <span data-ttu-id="9c373-105">本文介绍如何在 Windows 10 中使用 F12 开发人员工具中的调试器工具来测试你的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="9c373-105">This article describes how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9c373-106">本文中的说明不能用于调试使用 Execute 函数的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="9c373-106">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="9c373-107">若要调试使用 Execute 函数的 Outlook 加载项，我们建议你在脚本模式下附加到 Visual Studio 或附加到某些其他脚本调试器。</span><span class="sxs-lookup"><span data-stu-id="9c373-107">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9c373-108">先决条件</span><span class="sxs-lookup"><span data-stu-id="9c373-108">Prerequisites</span></span>

<span data-ttu-id="9c373-109">您需要安装以下软件：</span><span class="sxs-lookup"><span data-stu-id="9c373-109">You need the following software:</span></span>

- <span data-ttu-id="9c373-110">Windows 10 随附的 F12 开发人员工具</span><span class="sxs-lookup"><span data-stu-id="9c373-110">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="9c373-111">托管您的外接程序的 Office 客户端应用程序。 </span><span class="sxs-lookup"><span data-stu-id="9c373-111">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="9c373-112">您的外接程序。 </span><span class="sxs-lookup"><span data-stu-id="9c373-112">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="9c373-113">使用调试器</span><span class="sxs-lookup"><span data-stu-id="9c373-113">Using the Debugger</span></span>

<span data-ttu-id="9c373-114">本文介绍了如何在 Windows 10 上使用 F12 开发人员工具中的调试器工具，测试 Office 加载项。可以测试从 AppSource 获取的加载项，也可以测试从其他位置添加的加载项。F12 工具在单独的窗口中显示，并不使用 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="9c373-114">You can use the Debugger from the F12 developer tools in Windows 10 to test add-ins from AppSource or add-ins that you have added from other locations.</span></span> <span data-ttu-id="9c373-115">运行加载项后，可以启动 F12 开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="9c373-115">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="9c373-116">F12 工具显示在单独的窗口中，并不使用 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="9c373-116">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="9c373-p104">调试器属于 Windows 10 和 Internet Explorer 上的 F12 开发人员工具。旧版 Windows 不包含调试器。</span><span class="sxs-lookup"><span data-stu-id="9c373-p104">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="9c373-119">此示例使用 Word 和从 AppSource 获取的免费加载项。</span><span class="sxs-lookup"><span data-stu-id="9c373-119">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="9c373-120">打开 Word 并选择空白文档。</span><span class="sxs-lookup"><span data-stu-id="9c373-120">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="9c373-121">在“**插入**”选项卡上的“加载项”组中，依次选择“**存储**”和 **QR4Office** 加载项。</span><span class="sxs-lookup"><span data-stu-id="9c373-121">On the  Insert tab, in the Add-ins group, Store and select the QR4Office add-in.</span></span> <span data-ttu-id="9c373-122">（你可以从应用商店或加载项目录中加载任何加载项。）</span><span class="sxs-lookup"><span data-stu-id="9c373-122">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="9c373-123">启动与 Office 版本相对应的 F12 开发工具：</span><span class="sxs-lookup"><span data-stu-id="9c373-123">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="9c373-124">对于 32 位版 Office，请使用 C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="9c373-124">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="9c373-125">对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="9c373-125">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="9c373-126">当你启动 IEChooser 时，一个单独的窗口（名为“选择要调试的目标”）显示要调试的可能的应用程序。</span><span class="sxs-lookup"><span data-stu-id="9c373-126">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="9c373-127">选择你感兴趣的应用程序。</span><span class="sxs-lookup"><span data-stu-id="9c373-127">Select the application that you are interested in.</span></span> <span data-ttu-id="9c373-128">如果你正在编写自己的加载项，请选择你已在其中部署加载项的网站，这可能是本地主机 URL。</span><span class="sxs-lookup"><span data-stu-id="9c373-128">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="9c373-129">例如，选择 **home.html**。</span><span class="sxs-lookup"><span data-stu-id="9c373-129">For example, select **home.html**.</span></span> 
    
   ![IEChooser 屏幕，指向圈出的加载项](../images/choose-target-to-debug.png)

4. <span data-ttu-id="9c373-131">在 F12 窗口中，选择你想要调试的文件。</span><span class="sxs-lookup"><span data-stu-id="9c373-131">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="9c373-132">若要在 F12 窗口中选择文件，请选择“**脚本**”（左）窗格上方的文件夹图标。</span><span class="sxs-lookup"><span data-stu-id="9c373-132">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="9c373-133">从下拉列表中显示的可用文件列表中，选择 **Home.js**。</span><span class="sxs-lookup"><span data-stu-id="9c373-133">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="9c373-134">设置断点。</span><span class="sxs-lookup"><span data-stu-id="9c373-134">Set the breakpoint.</span></span>
    
   <span data-ttu-id="9c373-135">若要在 **Home.js** 中设置断点，请选择第 144 行，它位于 `textChanged` 函数中。</span><span class="sxs-lookup"><span data-stu-id="9c373-135">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="9c373-136">你将在该行左侧和“调用堆栈和断点”（右下角）窗格中的对应行左侧看到一个红点。</span><span class="sxs-lookup"><span data-stu-id="9c373-136">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="9c373-137">有关设置断点的其他方法，请参阅[使用调试器检查正在运行的 JavaScript](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))。</span><span class="sxs-lookup"><span data-stu-id="9c373-137">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![断点位于 home.js 文件中的调试程序](../images/debugger-home-js-02.png)

6. <span data-ttu-id="9c373-139">运行加载项，以触发断点。</span><span class="sxs-lookup"><span data-stu-id="9c373-139">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="9c373-140">在 Word 中，选择 **QR4Office** 窗格上部的 URL 文本框，然后尝试输入一些文本。</span><span class="sxs-lookup"><span data-stu-id="9c373-140">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="9c373-141">在调试器的“**调用堆栈和断点**”窗格中，你将看到该断点已触发，并显示了各种信息。</span><span class="sxs-lookup"><span data-stu-id="9c373-141">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="9c373-142">你可能需要刷新调试器以查看结果。</span><span class="sxs-lookup"><span data-stu-id="9c373-142">You might need to refresh the Debugger to see the results.</span></span>
    
   ![调试器，包含已触发的断点生成的结果](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="9c373-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9c373-144">See also</span></span>

- <span data-ttu-id="9c373-145">[使用调试器检查正在运行的 JavaScript](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="9c373-145">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="9c373-146">[使用 F12 开发人员工具](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="9c373-146">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
