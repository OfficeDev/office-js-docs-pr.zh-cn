---
title: '? Windows 10 ??? F12 ????????????'
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: e1e4cde4a1a0fe27058346b93e8aaa39dd75a4e3
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="debug-add-ins-using-f12-developer-tools-on-windows-10"></a><span data-ttu-id="87336-102">? Windows 10 ??? F12 ????????????</span><span class="sxs-lookup"><span data-stu-id="87336-102">Debug add-ins using F12 developer tools on Windows 10</span></span>

<span data-ttu-id="87336-p101">Windows 10 ???? F12 ????????????????????????????? IDE?? Visual Studio???????????? IDE ??????????????????????????????? Office ????????????????? F12 ???????</span><span class="sxs-lookup"><span data-stu-id="87336-p101">The F12 developer tools included in Windows 10 help you debug, test, and speed up your webpages. You can also use them to develop and debug Office Add-ins, if you are not using an IDE like Visual Studio, or if you need to investigate a problem while running your add-in outside the IDE. You can start the F12 developer tools after your add-in is running.</span></span>

<span data-ttu-id="87336-p102">???????? Windows 10 ??? F12 ???????????????? Office ????????? AppSource ????????????????????????F12 ???????????????? Visual Studio?</span><span class="sxs-lookup"><span data-stu-id="87336-p102">This article shows how you how to use the Debugger tool from the F12 developer tools in Windows 10 to test your Office Add-in. You can test add-ins from AppSource or add-ins that you have added from other locations. The F12 tools display in a separate window and do not use Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="87336-p103">????? Windows 10 ? Internet Explorer ?? F12 ????????? Windows ???????</span><span class="sxs-lookup"><span data-stu-id="87336-p103">The Debugger is part of the F12 developer tools in Windows 10 and Internet Explorer. Earlier versions of Windows do not include the Debugger.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="87336-111">????</span><span class="sxs-lookup"><span data-stu-id="87336-111">Prerequisites</span></span>

<span data-ttu-id="87336-112">??????????</span><span class="sxs-lookup"><span data-stu-id="87336-112">You need the following software:</span></span>

- <span data-ttu-id="87336-113">Windows 10 ??? F12 ??????</span><span class="sxs-lookup"><span data-stu-id="87336-113">The F12 developer tools, which are included in Windows 10.</span></span> 
    
- <span data-ttu-id="87336-114">????????? Office ????????</span><span class="sxs-lookup"><span data-stu-id="87336-114">The Office client application that hosts your add-in.</span></span> 
    
- <span data-ttu-id="87336-115">???????</span><span class="sxs-lookup"><span data-stu-id="87336-115">Your add-in.</span></span> 

## <a name="using-the-debugger"></a><span data-ttu-id="87336-116">?????</span><span class="sxs-lookup"><span data-stu-id="87336-116">Using the Debugger</span></span>

<span data-ttu-id="87336-117">????? Word ?? AppSource ?????????</span><span class="sxs-lookup"><span data-stu-id="87336-117">This example uses Word and a free add-in from AppSource.</span></span>

1. <span data-ttu-id="87336-118">?? Word ????????</span><span class="sxs-lookup"><span data-stu-id="87336-118">Open Word and choose a blank document.</span></span> 
    
2. <span data-ttu-id="87336-p104">?????****??????????????????Microsoft Store?****? QR4Office ???????? Microsoft Store ????????????????</span><span class="sxs-lookup"><span data-stu-id="87336-p104">On the **Insert** tab, in the Add-ins group, choose **Store** and select the QR4Office add-in. (You can load any add-in from the Store or your add-in catalog.)</span></span>
    
3. <span data-ttu-id="87336-121">??? Office ?????? F12 ?????</span><span class="sxs-lookup"><span data-stu-id="87336-121">Launch the F12 development tools that corresponds to your version of Office:</span></span>
    
   - <span data-ttu-id="87336-122">?? 32 ?? Office???? C:\Windows\System32\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="87336-122">For the 32-bit version of Office, use C:\Windows\System32\F12\F12Chooser.exe</span></span>
    
   - <span data-ttu-id="87336-123">?? 64 ?? Office???? C:\Windows\SysWOW64\F12\F12Chooser.exe</span><span class="sxs-lookup"><span data-stu-id="87336-123">For the 64-bit version of Office, use C:\Windows\SysWOW64\F12\F12Chooser.exe</span></span>
    
   <span data-ttu-id="87336-p105">???? F12Chooser ?????????????????????????????????????????????????????????????????????????????????????????? URL?</span><span class="sxs-lookup"><span data-stu-id="87336-p105">When you launch F12Chooser, a separate window named "Choose target to debug" displays the possible applications to debug. Select the application that you are interested in. If you are writing your own add-in, select the website where you have the add-in deployed, which might be a localhost URL.</span></span> 
    
   <span data-ttu-id="87336-127">??????home.html?****?</span><span class="sxs-lookup"><span data-stu-id="87336-127">For example, select **home.html**.</span></span> 
    
   ![F12Chooser ???????????](../images/choose-target-to-debug.png)

4. <span data-ttu-id="87336-129">? F12 ???????????????</span><span class="sxs-lookup"><span data-stu-id="87336-129">In the F12 window, select the file you want to debug.</span></span>
    
   <span data-ttu-id="87336-p106">???????????**??**????????????????????????????? home.js?</span><span class="sxs-lookup"><span data-stu-id="87336-p106">To select the file, choose the folder icon above the  **script** (left) pane. The dropdown list shows the available files. Select home.js.</span></span>
    
5. <span data-ttu-id="87336-133">?????</span><span class="sxs-lookup"><span data-stu-id="87336-133">Set the breakpoint.</span></span>
    
   <span data-ttu-id="87336-p107">?? home.js ?????????? 144 ???? _ textChanged_  ?????????????????????? ** ?????????** ??????????????????????????????[ ????????????? JavaScript?](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)?</span><span class="sxs-lookup"><span data-stu-id="87336-p107">To set the breakpoint in home.js, choose line 144, which is in the  _textChanged_ function. You will see a red dot to the left of the line and a corresponding line in the **Callstack and Breakpoints** (bottom right) pane. For other ways to set a breakpoint, see [Inspect running JavaScript with the Debugger](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx).</span></span> 
    
   ![???? home.js ????????](../images/debugger-home-js-02.png)

6. <span data-ttu-id="87336-138">????????????</span><span class="sxs-lookup"><span data-stu-id="87336-138">Run your add-in to trigger the breakpoint.</span></span>
    
   <span data-ttu-id="87336-p108">?? QR4Office ???????? URL ??????????????????????????****??????????????????????????? F12 ????????</span><span class="sxs-lookup"><span data-stu-id="87336-p108">Choose the URL textbox in the upper part of the QR4Office pane to change the text. In the Debugger, in the **Callstack and Breakpoints** pane, you'll see that the breakpoint has triggered and shows various information. You might need to refresh the F12 tool to see the results.</span></span>
    
   ![?????????????????](../images/debugger-home-js-01.png)


## <a name="see-also"></a><span data-ttu-id="87336-143">????</span><span class="sxs-lookup"><span data-stu-id="87336-143">See also</span></span>

- [<span data-ttu-id="87336-144">???????????? JavaScript</span><span class="sxs-lookup"><span data-stu-id="87336-144">Inspect running JavaScript with the Debugger</span></span>](https://msdn.microsoft.com/library/dn255007%28v=vs.85%29.aspx)
- [<span data-ttu-id="87336-145">?? F12 ??????</span><span class="sxs-lookup"><span data-stu-id="87336-145">Using the F12 developer tools</span></span>](https://msdn.microsoft.com/en-us/library/bg182326%28v=vs.85%29.aspx)
    
