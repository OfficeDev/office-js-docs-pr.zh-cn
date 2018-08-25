---
title: 在 Windows 10 上使用 F12 开发人员工具调试外接程序
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 226773962fb1777a3a1f0e09445721ae2b8b5f5b
ms.sourcegitcommit: e1c92ba882e6eb03a165867c6021a6aa742aa310
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/20/2018
ms.locfileid: "22925603"
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="bf207-102">在 Windows 10 上使用 F12 开发人员工具调试外接程序</span><span class="sxs-lookup"><span data-stu-id="bf207-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="bf207-p101">Windows 10 中随附的 F12 开发人员工具可帮助您调试、测试和加速您的网页。如果您未使用 IDE（如 Visual Studio），或者如果您需要调查在 IDE 外部运行外接程序时出现的问题，您还可以使用该工具开发和调试您的 Office 外接程序。运行外接程序后，可以启动 F12 开发人员工具。</span><span class="sxs-lookup"><span data-stu-id="bf207-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

<span data-ttu-id="bf207-p102">本文介绍了如何在 Windows 10 上使用 F12 开发人员工具中的调试器工具，测试 Office 加载项。可以测试从 AppSource 获取的加载项，也可以测试从其他位置添加的加载项。F12 工具在单独的窗口中显示，并不使用 Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="bf207-p102">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="bf207-p103">调试器属于 Windows 10 和 Internet Explorer 上的 F12 开发人员工具。旧版 Windows 不包含调试器。</span><span class="sxs-lookup"><span data-stu-id="bf207-p103">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="bf207-111">先决条件</span><span class="sxs-lookup"><span data-stu-id="bf207-111">Prerequisites</span></span>

<span data-ttu-id="bf207-112">您需要安装以下软件：</span><span class="sxs-lookup"><span data-stu-id="bf207-112">You need the following software:</span></span>

- <span data-ttu-id="bf207-113">Windows 10 随附的 F12 开发人员工具</span><span class="sxs-lookup"><span data-stu-id="bf207-113">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="bf207-114">托管您的外接程序的 Office 客户端应用程序。</span><span class="sxs-lookup"><span data-stu-id="bf207-114">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="bf207-115">您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="bf207-115">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="bf207-116">使用调试器</span><span class="sxs-lookup"><span data-stu-id="bf207-116">Using the Debugger</span></span>

<span data-ttu-id="bf207-117">此示例使用 Word 和从 AppSource 获取的免费加载项。</span><span class="sxs-lookup"><span data-stu-id="bf207-117">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="bf207-118">打开 Word 并选择空白文档。</span><span class="sxs-lookup"><span data-stu-id="bf207-118">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="bf207-p104">在“插入”**** 选项卡上的“加载项”组中，依次选择“Microsoft Store”**** 和 QR4Office 加载项。（可以从 Microsoft Store 或加载项目录中加载任何加载项。）</span><span class="sxs-lookup"><span data-stu-id="bf207-p104">On the **Insert** tab, in the Add-ins group, choose **Store** and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="bf207-121">启动与 Office 版本相对应的 F12 开发工具：</span><span class="sxs-lookup"><span data-stu-id="bf207-121">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="bf207-122">对于 32 位版 Office，请使用 C:\Windows\System32\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="bf207-122">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="bf207-123">对于 64 位版 Office，请使用 C:\Windows\SysWOW64\F12\IEChooser.exe</span><span class="sxs-lookup"><span data-stu-id="bf207-123">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="bf207-p105">当你启动 IEChooser 时，一个单独的窗口（名为“选择要调试的目标”）显示要调试的可能的应用程序。选择你感兴趣的应用程序。如果你正在编写自己的外接程序，请选择你已在其中部署外接程序的网站，这可以是本地主机 URL。</span><span class="sxs-lookup"><span data-stu-id="bf207-p105">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug. Select the application that you are interested in. If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="bf207-127">例如，选择“home.html”****。</span><span class="sxs-lookup"><span data-stu-id="bf207-127">For example, select **home.html**.</span></span> 
    
   ![IEChooser 屏幕，指向气泡加载项](../images/choose-target-to-debug.png)

4. <span data-ttu-id="bf207-129">在 F12 窗口中，选择您想要调试的文件。</span><span class="sxs-lookup"><span data-stu-id="bf207-129">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="bf207-p106">若要选择文件，请选择“**脚本**”（左）窗格上方的文件夹图标。下拉列表显示了可用文件。选择 home.js。</span><span class="sxs-lookup"><span data-stu-id="bf207-p106">To select the file, choose the folder icon above the  **script** (left) pane. The dropdown list shows the available files. Select home.js.</span></span>
    
5. <span data-ttu-id="bf207-133">设置断点。</span><span class="sxs-lookup"><span data-stu-id="bf207-133">Set the breakpoint.</span></span>
    
   <span data-ttu-id="bf207-p107">要在 home.js 中设置断点，请选择第 144 行（位于 _ textChanged_  函数中）。您将在该行的左侧看到一个红点，并在 ** “调用堆栈和断点”** （右下角）窗格中看到相应的行。有关设置断点的其他方法，请参阅[ “使用调试器检查正在运行的 JavaScript”](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))。</span><span class="sxs-lookup"><span data-stu-id="bf207-p107">To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85)).</span></span> 
    
   ![断点位于 home.js 文件中的调试程序](../images/debugger-home-js-02.png)

6. <span data-ttu-id="bf207-138">运行加载项，以触发断点。</span><span class="sxs-lookup"><span data-stu-id="bf207-138">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="bf207-p108">选择 QR4Office 窗格上半部分中的 URL 文本框，以更改文本。在“调试器”的“调用堆栈和断点”**** 窗格中，将看到断点已触发，以及显示的各种信息。建议刷新 F12 工具来查看结果。</span><span class="sxs-lookup"><span data-stu-id="bf207-p108">Choose the URL textbox in the upper part of the QR4Office pane to change the text. In the Debugger, in the **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the F12 tool to see the results.</span></span>
    
   ![调试器，包含已触发的断点生成的结果](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="bf207-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bf207-143">See also</span></span>

- <span data-ttu-id="bf207-144">[使用调试器检查正在运行的 JavaScript](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="bf207-144">[Inspect running JavaScript with the Debugger](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))</span></span>
- <span data-ttu-id="bf207-145">[使用 F12 开发人员工具](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span><span class="sxs-lookup"><span data-stu-id="bf207-145">[Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))</span></span>
    
