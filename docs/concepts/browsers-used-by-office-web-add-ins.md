---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 09/25/2019
localization_priority: Priority
ms.openlocfilehash: b5d7198e556f020bccdf7ba1e0a0fcffa3a9171b
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235293"
---
# <a name="browsers-used-by-office-add-ins"></a><span data-ttu-id="254c3-103">Office 加载项使用的浏览器</span><span class="sxs-lookup"><span data-stu-id="254c3-103">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="254c3-104">Office 加载项是使用 iFrames（在 Office 网页版中运行时）和使用 Office 桌面版和移动版客户端中的嵌入式浏览器控件显示的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="254c3-104">Office add-ins are web applications that are displayed using iFrames when running in Office on the web and using embedded browser controls in Office for desktop and mobile clients.</span></span> <span data-ttu-id="254c3-105">加载项还需要使用 JavaScript 引擎来运行 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="254c3-105">Add-ins also need a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="254c3-106">嵌入式浏览器和引擎均由用户计算机上安装的浏览器提供。</span><span class="sxs-lookup"><span data-stu-id="254c3-106">Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="254c3-107">要使用的浏览器取决于：</span><span class="sxs-lookup"><span data-stu-id="254c3-107">Which browser is used depends on:</span></span>

- <span data-ttu-id="254c3-108">计算机的操作系统。</span><span class="sxs-lookup"><span data-stu-id="254c3-108">The computer’s operating system.</span></span>
- <span data-ttu-id="254c3-109">加载项是在 Office 网页版、Office 365 还是非订阅版 Office 2013 或更高版本中运行。</span><span class="sxs-lookup"><span data-stu-id="254c3-109">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="254c3-110">下表显示在不同平台和操作系统中使用的浏览器。</span><span class="sxs-lookup"><span data-stu-id="254c3-110">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="254c3-111">**操作系统/平台**</span><span class="sxs-lookup"><span data-stu-id="254c3-111">**OS / Platform**</span></span>|<span data-ttu-id="254c3-112">**Browser**</span><span class="sxs-lookup"><span data-stu-id="254c3-112">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="254c3-113">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="254c3-113">Office on the web</span></span>|<span data-ttu-id="254c3-114">在其中打开 Office 的浏览器。</span><span class="sxs-lookup"><span data-stu-id="254c3-114">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="254c3-115">Mac</span><span class="sxs-lookup"><span data-stu-id="254c3-115">Mac</span></span>|<span data-ttu-id="254c3-116">Safari</span><span class="sxs-lookup"><span data-stu-id="254c3-116">Safari</span></span>|
|<span data-ttu-id="254c3-117">iOS</span><span class="sxs-lookup"><span data-stu-id="254c3-117">iOS</span></span>|<span data-ttu-id="254c3-118">Safari</span><span class="sxs-lookup"><span data-stu-id="254c3-118">Safari</span></span>|
|<span data-ttu-id="254c3-119">Android</span><span class="sxs-lookup"><span data-stu-id="254c3-119">Android</span></span>|<span data-ttu-id="254c3-120">Chrome</span><span class="sxs-lookup"><span data-stu-id="254c3-120">Chrome</span></span>|
|<span data-ttu-id="254c3-121">Windows/非订阅版 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="254c3-121">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="254c3-122">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="254c3-122">Internet Explorer 11</span></span>|
|<span data-ttu-id="254c3-123">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="254c3-123">Windows 10 ver.</span></span> <span data-ttu-id="254c3-124">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="254c3-124">< 1903 / Office 365</span></span>|<span data-ttu-id="254c3-125">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="254c3-125">Internet Explorer 11</span></span>|
|<span data-ttu-id="254c3-126">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="254c3-126">Windows 10 ver.</span></span> <span data-ttu-id="254c3-127">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="254c3-127">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="254c3-128">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="254c3-128">Internet Explorer 11</span></span>|
|<span data-ttu-id="254c3-129">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="254c3-129">Windows 10 ver.</span></span> <span data-ttu-id="254c3-130">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="254c3-130">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="254c3-131">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="254c3-131">Microsoft Edge\*</span></span>|

<span data-ttu-id="254c3-132">\*使用 Microsoft Edge 时，Windows 10 讲述人（有时称为“屏幕阅读器”）会读出页面中在任务窗格中打开的 `<title>` 标记。</span><span class="sxs-lookup"><span data-stu-id="254c3-132">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="254c3-133">如果使用的是 Internet Explorer 11，则Narrator 将会读取任务窗格的标题栏，它来自加载项清单中的 `<DisplayName>` 值。</span><span class="sxs-lookup"><span data-stu-id="254c3-133">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="254c3-134">Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。</span><span class="sxs-lookup"><span data-stu-id="254c3-134">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="254c3-135">如果任何加载项用户安装的是使用 Internet Explorer 11 的平台，若要使用 ECMAScript 2015 或更高版本的语法和功能，则必须将 JavaScript 转换为 ES5 或使用填充代码。</span><span class="sxs-lookup"><span data-stu-id="254c3-135">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="254c3-136">此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。</span><span class="sxs-lookup"><span data-stu-id="254c3-136">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="254c3-137">在它们公开发布之前，你需要是 Windows 预览体验成员才能获得 Windows 版本 1903 或更高版本，并且需要是 Office 预览体验成员才能获得 Office 版本 16.0.11629 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="254c3-137">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="254c3-138">若要加入 Windows 预览体验成员：</span><span class="sxs-lookup"><span data-stu-id="254c3-138">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="254c3-139">转至 [Windows 预览体验成员](https://insider.windows.com)并单击链接以加入 Windows 预览体验成员。</span><span class="sxs-lookup"><span data-stu-id="254c3-139">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="254c3-140">系统会将你引导至包含有关如何使用 Windows 设置支持 Windows 预览内部版本说明的页面。</span><span class="sxs-lookup"><span data-stu-id="254c3-140">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="254c3-141">按照说明执行操作。</span><span class="sxs-lookup"><span data-stu-id="254c3-141">Follow the instructions.</span></span> <span data-ttu-id="254c3-142">选择更新频率时，请选择时间最短的选项。</span><span class="sxs-lookup"><span data-stu-id="254c3-142">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="254c3-143">若要加入 Office 预览体验成员：</span><span class="sxs-lookup"><span data-stu-id="254c3-143">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="254c3-144">转至 [Office 预览体验成员入门](https://insider.office.com/join)。</span><span class="sxs-lookup"><span data-stu-id="254c3-144">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="254c3-145">按照加入页面上的说明操作。</span><span class="sxs-lookup"><span data-stu-id="254c3-145">Follow the instruction on that page to join.</span></span> <span data-ttu-id="254c3-146">系统要求你指定频道时，请选择预览体验成员。</span><span class="sxs-lookup"><span data-stu-id="254c3-146">When asked to specify a channel, select Insider.</span></span>

## <a name="troubleshooting-microsoft-edge-issues"></a><span data-ttu-id="254c3-147">Microsoft Edge 问题疑难解答</span><span class="sxs-lookup"><span data-stu-id="254c3-147">Troubleshooting Microsoft Edge Issues</span></span>

### <a name="scroll-bar-does-not-appear-in-task-pane"></a><span data-ttu-id="254c3-148">任务窗格中不显示滚动条</span><span class="sxs-lookup"><span data-stu-id="254c3-148">Scroll bar does not appear in task pane</span></span>

<span data-ttu-id="254c3-149">默认情况下，Microsoft Edge 中的滚动条是隐藏的，直到在其上悬停时。</span><span class="sxs-lookup"><span data-stu-id="254c3-149">By default, scroll bars in Microsoft Edge are hidden until hovered over.</span></span> <span data-ttu-id="254c3-150">适用于任务窗格中页面的 `<body>` 元素的 CSS 样式应包含 [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) 属性，且应将其设置为 `scrollbar`。</span><span class="sxs-lookup"><span data-stu-id="254c3-150">To ensure that the scroll bar is always visible, the CSS styling that applies to the `<body>` element of the pages in the task pane should include the [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) property and it should be set to `scrollbar`.</span></span> 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a><span data-ttu-id="254c3-151">使用 Microsoft Edge 开发工具进行调试时，加载项会崩溃或重新加载</span><span class="sxs-lookup"><span data-stu-id="254c3-151">When debugging with the Microsoft Edge DevTools, the add-in crashes or reloads</span></span>

<span data-ttu-id="254c3-152">[Microsoft Edge 开发工具](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab)中的设置断点可能导致 Office 认为该加载项已挂起。</span><span class="sxs-lookup"><span data-stu-id="254c3-152">Setting breakpoints in the [Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) can cause Office to think that the add-in is hung.</span></span> <span data-ttu-id="254c3-153">发生这种情况时，它将自动重新加载该加载项。</span><span class="sxs-lookup"><span data-stu-id="254c3-153">It will automatically reload the add-in when this happens.</span></span> <span data-ttu-id="254c3-154">为防止这种情况，请将以下注册表项和值添加到开发计算机：`[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。</span><span class="sxs-lookup"><span data-stu-id="254c3-154">To prevent this, add the following Registry key and value to the development computer: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`.</span></span>

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a><span data-ttu-id="254c3-155">加载项尝试打开时，出现“加载项错误 我们无法从 localhost 打开此加载项”错误</span><span class="sxs-lookup"><span data-stu-id="254c3-155">When the add-in tries to open, get "ADD-IN ERROR We can't open this add-in from the localhost" error</span></span>

<span data-ttu-id="254c3-156">一个已知的原因是 Microsoft Edge 要求在开发计算机上为本地主机提供环回豁免。</span><span class="sxs-lookup"><span data-stu-id="254c3-156">One known cause is that Microsoft Edge requires that localhost be given a loopback exemption on the development computer.</span></span> <span data-ttu-id="254c3-157">按照[无法从 localhost 打开加载项](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)中的说明操作。</span><span class="sxs-lookup"><span data-stu-id="254c3-157">Follow the instructions at [Cannot open add-in from localhost](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost).</span></span>


## <a name="see-also"></a><span data-ttu-id="254c3-158">另请参阅</span><span class="sxs-lookup"><span data-stu-id="254c3-158">See also</span></span>

- [<span data-ttu-id="254c3-159">Office 加载项的运行要求</span><span class="sxs-lookup"><span data-stu-id="254c3-159">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
