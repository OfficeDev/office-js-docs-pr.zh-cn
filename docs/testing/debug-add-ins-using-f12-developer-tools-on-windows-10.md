---
title: 使用 Windows 10 上的开发人员工具调试加载项
description: 使用 Windows 10 上的 Microsoft Edge 开发人员工具调试加载项
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 16964b69f144d30c4ac5a9616ee4fdba2d536fd3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950521"
---
# <a name="debug-add-ins-using-developer-tools-on-windows-10"></a><span data-ttu-id="e51bd-103">使用 Windows 10 上的开发人员工具调试加载项</span><span class="sxs-lookup"><span data-stu-id="e51bd-103">Debug add-ins using developer tools on Windows 10</span></span>

<span data-ttu-id="e51bd-104">在 IDE 之外，还有一些开发人员工具可用于帮助你在 Windows 10 上调试加载项。</span><span class="sxs-lookup"><span data-stu-id="e51bd-104">There are developer tools outside of IDEs available to help you debug your add-ins on Windows 10.</span></span> <span data-ttu-id="e51bd-105">当你在 IDE 之外运行加载项的同时，需要调查问题时，这些工具非常有用。</span><span class="sxs-lookup"><span data-stu-id="e51bd-105">These are useful when you need to investigate a problem while running your add-in outside the IDE.</span></span>

<span data-ttu-id="e51bd-106">所使用的工具取决于加载项是在 Microsoft Edge 还是在 Internet Explorer 中运行。</span><span class="sxs-lookup"><span data-stu-id="e51bd-106">The tool that you use depends on whether the add-in is running in Microsoft Edge or Internet Explorer.</span></span> <span data-ttu-id="e51bd-107">这取决于计算机上安装的 Windows 10 版本和 Office 版本。</span><span class="sxs-lookup"><span data-stu-id="e51bd-107">This is determined by the version of Windows 10 and the version of Office that are installed on the computer.</span></span> <span data-ttu-id="e51bd-108">若要确定开发计算机上使用的浏览器，请参阅 [Office 加载项使用的浏览器](../concepts/browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="e51bd-108">To determine which browser is being used on your development computer, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!NOTE]
> <span data-ttu-id="e51bd-109">本文中的说明不能用于调试使用 Execute 函数的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="e51bd-109">The instructions in this article cannot be used to debug an Outlook add-in that uses Execute Functions.</span></span> <span data-ttu-id="e51bd-110">若要调试使用 Execute 函数的 Outlook 加载项，我们建议你在脚本模式下附加到 Visual Studio 或附加到某些其他脚本调试器。</span><span class="sxs-lookup"><span data-stu-id="e51bd-110">To debug an Outlook add-in that uses Execute Functions, we recommend that you attach to Visual Studio in script mode or to some other script debugger.</span></span>

## <a name="when-the-add-in-is-running-in-microsoft-edge"></a><span data-ttu-id="e51bd-111">当加载项在 Microsoft Edge 中运行时</span><span class="sxs-lookup"><span data-stu-id="e51bd-111">When the add-in is running in Microsoft Edge</span></span>

[!include[Enable debugging on Microsoft Edge DevTools](../includes/enable-debugging-on-edge-devtools.md)]

### <a name="debug-using-microsoft-edge-devtools"></a><span data-ttu-id="e51bd-112">使用 Microsoft Edge DevTools 进行调试</span><span class="sxs-lookup"><span data-stu-id="e51bd-112">Debug using Microsoft Edge DevTools</span></span>

<span data-ttu-id="e51bd-113">当加载项在 Microsoft Edge 中运行时，可使用 [Microsoft Edge 开发人员工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab)。</span><span class="sxs-lookup"><span data-stu-id="e51bd-113">When the add-in is running in Microsoft Edge, you can use the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab).</span></span>

1. <span data-ttu-id="e51bd-114">运行加载项。</span><span class="sxs-lookup"><span data-stu-id="e51bd-114">Run the add-in.</span></span>

2. <span data-ttu-id="e51bd-115">运行 Microsoft Edge 开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="e51bd-115">Run the Microsoft Edge DevTools.</span></span>

3. <span data-ttu-id="e51bd-116">在工具中，打开“**本地**”选项卡。加载项将按其名称列出。</span><span class="sxs-lookup"><span data-stu-id="e51bd-116">In the tools, open the **Local** tab. Your add-in will be listed by its name.</span></span>

4. <span data-ttu-id="e51bd-117">单击加载项名称，将其在工具中打开。</span><span class="sxs-lookup"><span data-stu-id="e51bd-117">Click the add-in name to open it in the tools.</span></span>

5. <span data-ttu-id="e51bd-118">打开“**调试器**”选项卡。</span><span class="sxs-lookup"><span data-stu-id="e51bd-118">Open the **Debugger** tab.</span></span> 

6. <span data-ttu-id="e51bd-119">选择“**脚本**”（左）窗格上方的文件夹图标。</span><span class="sxs-lookup"><span data-stu-id="e51bd-119">Choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="e51bd-120">从下拉列表中显示的可用文件列表中，选择要调试的 JavaScript 文件。</span><span class="sxs-lookup"><span data-stu-id="e51bd-120">From the list of available files shown in the dropdown list, select the JavaScript file that you want to debug.</span></span>

7. <span data-ttu-id="e51bd-121">若要设置断点，请选择该行。</span><span class="sxs-lookup"><span data-stu-id="e51bd-121">To set a breakpoint, select the line.</span></span> <span data-ttu-id="e51bd-122">你将在该行左侧和“**调用堆栈**”（右下角）窗格中的对应行左侧看到一个红点。</span><span class="sxs-lookup"><span data-stu-id="e51bd-122">You will see a red dot to the left of the line and a corresponding line in the **Call stack** (bottom right) pane.</span></span>

8. <span data-ttu-id="e51bd-123">根据需要在加载项中执行函数以触发断点。</span><span class="sxs-lookup"><span data-stu-id="e51bd-123">Execute functions in the add-in as needed to trigger the breakpoint.</span></span>

## <a name="when-the-add-in-is-running-in-internet-explorer"></a><span data-ttu-id="e51bd-124">当加载项在 Internet Explorer 中运行时</span><span class="sxs-lookup"><span data-stu-id="e51bd-124">When the add-in is running in Internet Explorer</span></span>

<span data-ttu-id="e51bd-125">当加载项在 Internet Explorer 中运行时，可以使用 Windows 10 中 F12 开发人员工具中的调试器来测试加载项。</span><span class="sxs-lookup"><span data-stu-id="e51bd-125">When the add-in is running in Internet Explorer, you can use the debugger from the F12 developer tools in Windows 10 to test your add-in.</span></span> <span data-ttu-id="e51bd-126">运行加载项后，可以启动 F12 开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="e51bd-126">You can start the F12 developer tools after the add-in is running.</span></span> <span data-ttu-id="e51bd-127">F12 工具显示在单独的窗口中，并不使用 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="e51bd-127">The F12 tools are displayed in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="e51bd-p107">调试器属于 Windows 10 和 Internet Explorer 上的 F12 开发人员工具。旧版 Windows 不包含调试器。</span><span class="sxs-lookup"><span data-stu-id="e51bd-p107">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

<span data-ttu-id="e51bd-130">此示例使用 Word 和从 AppSource 获取的免费加载项。</span><span class="sxs-lookup"><span data-stu-id="e51bd-130">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="e51bd-131">打开 Word 并选择空白文档。</span><span class="sxs-lookup"><span data-stu-id="e51bd-131">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="e51bd-132">在“**插入**”选项卡上的“加载项”组中，依次选择“**存储**”和 **QR4Office** 加载项。</span><span class="sxs-lookup"><span data-stu-id="e51bd-132">On the **Insert** tab, in the Add-ins group, choose **Store** and select the **QR4Office** Add-in.</span></span> <span data-ttu-id="e51bd-133">（你可以从应用商店或加载项目录中加载任何加载项。）</span><span class="sxs-lookup"><span data-stu-id="e51bd-133">(You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="e51bd-134">启动与 Office 版本相对应的 F12 开发工具：</span><span class="sxs-lookup"><span data-stu-id="e51bd-134">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="e51bd-135">对于 32 位版 Office，请使用 C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="e51bd-135">For the 32-bit version of Office, use C:\Windows\System32\F12\IEChooser.exe</span></span>
    
   - <span data-ttu-id="e51bd-136">对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="e51bd-136">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\IEChooser.exe</span></span>
    
   <span data-ttu-id="e51bd-137">当你启动 IEChooser 时，一个单独的窗口（名为“选择要调试的目标”）显示要调试的可能的应用程序。</span><span class="sxs-lookup"><span data-stu-id="e51bd-137">When you launch IEChooser, a separate window named "Choose target to debug" displays the possible applications to debug.</span></span> <span data-ttu-id="e51bd-138">选择你感兴趣的应用程序。</span><span class="sxs-lookup"><span data-stu-id="e51bd-138">Select the application that you are interested in.</span></span> <span data-ttu-id="e51bd-139">如果你正在编写自己的加载项，请选择你已在其中部署加载项的网站，这可能是本地主机 URL。</span><span class="sxs-lookup"><span data-stu-id="e51bd-139">If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="e51bd-140">例如，选择 **home.html**。</span><span class="sxs-lookup"><span data-stu-id="e51bd-140">For example, select **home.html**.</span></span> 
    
   ![IEChooser 屏幕，指向圈出的加载项](../images/choose-target-to-debug.png)

4. <span data-ttu-id="e51bd-142">在 F12 窗口中，选择你想要调试的文件。</span><span class="sxs-lookup"><span data-stu-id="e51bd-142">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="e51bd-143">若要在 F12 窗口中选择文件，请选择“**脚本**”（左）窗格上方的文件夹图标。</span><span class="sxs-lookup"><span data-stu-id="e51bd-143">To select the file in the F12 window, choose the folder icon above the **script** (left) pane.</span></span> <span data-ttu-id="e51bd-144">从下拉列表中显示的可用文件列表中，选择 **Home.js**。</span><span class="sxs-lookup"><span data-stu-id="e51bd-144">From the list of available files shown in the dropdown list, select **Home.js**.</span></span>
    
5. <span data-ttu-id="e51bd-145">设置断点。</span><span class="sxs-lookup"><span data-stu-id="e51bd-145">Set the breakpoint.</span></span>
    
   <span data-ttu-id="e51bd-146">若要在 **Home.js** 中设置断点，请选择第 144 行，它位于 `textChanged` 函数中。</span><span class="sxs-lookup"><span data-stu-id="e51bd-146">To set the breakpoint in **Home.js**, choose line 144, which is in the  `textChanged` function.</span></span> <span data-ttu-id="e51bd-147">你将在该行左侧和“调用堆栈和断点”（右下角）窗格中的对应行左侧看到一个红点。</span><span class="sxs-lookup"><span data-stu-id="e51bd-147">You will see a red dot to the left of the line and a corresponding line in the **Call stack and Breakpoints** (bottom right) pane.</span></span> <span data-ttu-id="e51bd-148">有关设置断点的其他方法，请参阅[使用调试器检查正在运行的 JavaScript](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))。</span><span class="sxs-lookup"><span data-stu-id="e51bd-148">For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![断点位于 home.js 文件中的调试程序](../images/debugger-home-js-02.png)

6. <span data-ttu-id="e51bd-150">运行加载项，以触发断点。</span><span class="sxs-lookup"><span data-stu-id="e51bd-150">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="e51bd-151">在 Word 中，选择 **QR4Office** 窗格上部的 URL 文本框，然后尝试输入一些文本。</span><span class="sxs-lookup"><span data-stu-id="e51bd-151">In Word, choose the URL textbox in the upper part of the **QR4Office** pane and attempt to enter some text.</span></span> <span data-ttu-id="e51bd-152">在调试器的“**调用堆栈和断点**”窗格中，你将看到该断点已触发，并显示了各种信息。</span><span class="sxs-lookup"><span data-stu-id="e51bd-152">In the Debugger, in the **Call stack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information.</span></span> <span data-ttu-id="e51bd-153">你可能需要刷新调试器以查看结果。</span><span class="sxs-lookup"><span data-stu-id="e51bd-153">You might need to refresh the Debugger to see the results.</span></span>
    
   ![调试器，包含已触发的断点生成的结果](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="e51bd-155">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e51bd-155">See also</span></span>

- <span data-ttu-id="e51bd-156">[使用调试器检查正在运行的 JavaScript](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="e51bd-156">[Inspect running JavaScript with the Debugger](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="e51bd-157">[使用 F12 开发人员工具](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="e51bd-157">[Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
