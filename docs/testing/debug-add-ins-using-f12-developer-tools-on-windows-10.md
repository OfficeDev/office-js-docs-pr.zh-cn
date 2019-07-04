---
title: 使用 Windows 10 上的开发人员工具调试加载项
description: ''
ms.date: 07/01/2019
localization_priority: Priority
ms.openlocfilehash: a2090eca41f59f0e7fab1a172aff96cbbca28ed7
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454879"
---
# <a name="debug-add-ins-using-developer-tools-on-windows-10"></a><span data-ttu-id="3d00b-102">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="3d00b-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="3d00b-103">在 IDE 之外，还有一些开发人员工具可用于帮助你在 Windows 10 上调试加载项。</span><span class="sxs-lookup"><span data-stu-id="3d00b-103">There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10.</span></span> <span data-ttu-id="3d00b-104">当你在 IDE 之外运行加载项的同时，需要调查问题时，这些工具非常有用。</span><span class="sxs-lookup"><span data-stu-id="3d00b-104">These are useful when you need to investigate a problem while running your add-in outside the IDE.</span></span>

<span data-ttu-id="3d00b-105">所使用的工具取决于加载项是在 Microsoft Edge 还是在 Internet Explorer 中运行。</span><span class="sxs-lookup"><span data-stu-id="3d00b-105">The tool that you use depends on whether the add-in is running in Edge or Internet Explorer.</span></span> <span data-ttu-id="3d00b-106">这取决于计算机上安装的 Windows 10 版本和 Office 版本。</span><span class="sxs-lookup"><span data-stu-id="3d00b-106">This is determined by the version of Windows 10 and the version of Office that are installed on the computer.</span></span> <span data-ttu-id="3d00b-107">若要确定开发计算机上使用的浏览器，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="3d00b-107">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span> 


> [!NOTE]
> <span data-ttu-id="3d00b-108">本文中的说明不能用于调试使用 Execute 函数的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="3d00b-108">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="3d00b-109">若要调试使用 Execute 函数的 Outlook 加载项，我们建议你在脚本模式下附加到 Visual Studio 或附加到某些其他脚本调试器。</span><span class="sxs-lookup"><span data-stu-id="3d00b-109">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="when-the-add-in-is-running-in-edge"></a><span data-ttu-id="3d00b-110">当加载项在 Microsoft Edge 中运行时</span><span class="sxs-lookup"><span data-stu-id="3d00b-110">When the add-in is running in Edge</span></span>

<span data-ttu-id="3d00b-111">当加载项在 Microsoft Edge 中运行时，可使用 [Microsoft Edge 开发人员工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab)。</span><span class="sxs-lookup"><span data-stu-id="3d00b-111">When the add-in is running in Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span> 

1. <span data-ttu-id="3d00b-112">运行加载项。</span><span class="sxs-lookup"><span data-stu-id="3d00b-112">Run the add-in</span></span> 

2. <span data-ttu-id="3d00b-113">运行 Microsoft Edge 开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="3d00b-113">Run the Microsoft Edge DevTools.</span></span>

3. <span data-ttu-id="3d00b-114">在工具中，打开“**本地**”选项卡。加载项将按其名称列出。</span><span class="sxs-lookup"><span data-stu-id="3d00b-114">In the tools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

4. <span data-ttu-id="3d00b-115">单击加载项名称，将其在工具中打开。</span><span class="sxs-lookup"><span data-stu-id="3d00b-115">Click the add-in name to open it in the tools.</span></span>

5. <span data-ttu-id="3d00b-116">打开“**调试器**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="3d00b-116">Open the **Permissions** tab.</span></span> 

6. <span data-ttu-id="3d00b-117">选择“**脚本**”（左）窗格上方的文件夹图标。</span><span class="sxs-lookup"><span data-stu-id="3d00b-117">To select the file, choose the folder icon above the  **script** (left) pane.</span></span> <span data-ttu-id="3d00b-118">从下拉列表中显示的可用文件列表中，选择要调试的 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="3d00b-118">From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.</span></span>

7. <span data-ttu-id="3d00b-119">若要设置断点，请选择该行。</span><span class="sxs-lookup"><span data-stu-id="3d00b-119">To set a breakpoint, select the line.</span></span> <span data-ttu-id="3d00b-120">你将在该行左侧和“**调用堆栈**”（右下角）窗格中的对应行左侧看到一个红点。</span><span class="sxs-lookup"><span data-stu-id="3d00b-120">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span>

8. <span data-ttu-id="3d00b-121">根据需要在加载项中执行函数以触发断点。</span><span class="sxs-lookup"><span data-stu-id="3d00b-121">Execute functions in the add-in as needed to trigger the breakpoint.</span></span>

## <a name="when-the-add-in-is-running-in-internet-explorer"></a><span data-ttu-id="3d00b-122">当加载项在 Internet Explorer 中运行时</span><span class="sxs-lookup"><span data-stu-id="3d00b-122">When the add-in is running in Internet Explorer</span></span>

<span data-ttu-id="3d00b-123">当加载项在 Internet Explorer 中运行时，可以使用 Windows 10 中 F12 开发人员工具中的调试器来测试加载项。</span><span class="sxs-lookup"><span data-stu-id="3d00b-123">When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in.</span></span> <span data-ttu-id="3d00b-124">运行加载项后，可以启动 F12 开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="3d00b-124">You can start the F12 developer tools after your add-in is running.</span></span> <span data-ttu-id="3d00b-125">F12 工具显示在单独的窗口中，并不使用 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="3d00b-125">The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="3d00b-p107">调试器属于 Windows 10 和 Internet Explorer 上的 F12 开发人员工具。旧版 Windows 不包含调试器。</span><span class="sxs-lookup"><span data-stu-id="3d00b-p107">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="3d00b-128">此示例使用 Word 和从 AppSource 获取的免费加载项。</span><span class="sxs-lookup"><span data-stu-id="3d00b-128">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="3d00b-129">打开 Word 并选择空白文档。</span><span class="sxs-lookup"><span data-stu-id="3d00b-129">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="3d00b-130">在“**插入**”选项卡上的“加载项”组中，依次选择“**存储**”和 **QR4Office** 加载项。</span><span class="sxs-lookup"><span data-stu-id="3d00b-130">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="3d00b-131">（你可以从应用商店或加载项目录中加载任何加载项。）</span><span class="sxs-lookup"><span data-stu-id="3d00b-131">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="3d00b-132">启动与 Office 版本相对应的 F12 开发工具：</span><span class="sxs-lookup"><span data-stu-id="3d00b-132">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="3d00b-133">对于 32 位版 Office，请使用 C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="3d00b-133">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="3d00b-134">对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="3d00b-134">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="3d00b-135">当你启动 IEChooser 时，一个单独的窗口（名为“选择要调试的目标”）显示要调试的可能的应用程序。</span><span class="sxs-lookup"><span data-stu-id="3d00b-135">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="3d00b-136">选择你感兴趣的应用程序。</span><span class="sxs-lookup"><span data-stu-id="3d00b-136">Select the application that you are interested in.</span></span> <span data-ttu-id="3d00b-137">如果你正在编写自己的加载项，请选择你已在其中部署加载项的网站，这可能是本地主机 URL。</span><span class="sxs-lookup"><span data-stu-id="3d00b-137">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="3d00b-138">例如，选择 **home.html**。</span><span class="sxs-lookup"><span data-stu-id="3d00b-138">For example, select **home.html**.</span></span> 
    
   ![IEChooser 屏幕，指向圈出的加载项](../images/choose-target-to-debug.png)

4. <span data-ttu-id="3d00b-140">在 F12 窗口中，选择你想要调试的文件。</span><span class="sxs-lookup"><span data-stu-id="3d00b-140">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="3d00b-141">若要在 F12 窗口中选择文件，请选择“**脚本**”（左）窗格上方的文件夹图标。</span><span class="sxs-lookup"><span data-stu-id="3d00b-141">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="3d00b-142">从下拉列表中显示的可用文件列表中，选择 **Home.js**。</span><span class="sxs-lookup"><span data-stu-id="3d00b-142">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="3d00b-143">设置断点。</span><span class="sxs-lookup"><span data-stu-id="3d00b-143">Set the breakpoint.</span></span>
    
   <span data-ttu-id="3d00b-144">若要在 **Home.js** 中设置断点，请选择第 144 行，它位于 `textChanged` 函数中。</span><span class="sxs-lookup"><span data-stu-id="3d00b-144">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="3d00b-145">你将在该行左侧和“调用堆栈和断点”（右下角）窗格中的对应行左侧看到一个红点。</span><span class="sxs-lookup"><span data-stu-id="3d00b-145">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="3d00b-146">有关设置断点的其他方法，请参阅[使用调试器检查正在运行的 JavaScript](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))。</span><span class="sxs-lookup"><span data-stu-id="3d00b-146">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![断点位于 home.js 文件中的调试程序](../images/debugger-home-js-02.png)

6. <span data-ttu-id="3d00b-148">运行加载项，以触发断点。</span><span class="sxs-lookup"><span data-stu-id="3d00b-148">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="3d00b-149">在 Word 中，选择 **QR4Office** 窗格上部的 URL 文本框，然后尝试输入一些文本。</span><span class="sxs-lookup"><span data-stu-id="3d00b-149">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="3d00b-150">在调试器的“**调用堆栈和断点**”窗格中，你将看到该断点已触发，并显示了各种信息。</span><span class="sxs-lookup"><span data-stu-id="3d00b-150">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="3d00b-151">你可能需要刷新调试器以查看结果。</span><span class="sxs-lookup"><span data-stu-id="3d00b-151">You might need to refresh the Debugger to see the results.</span></span>
    
   ![调试器，包含已触发的断点生成的结果](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="3d00b-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3d00b-153">See also</span></span>

- <span data-ttu-id="3d00b-154">[使用调试器检查正在运行的 JavaScript](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="3d00b-154">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="3d00b-155">[使用 F12 开发人员工具](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="3d00b-155">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
