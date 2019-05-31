---
title: Office 加载项使用的浏览器
description: 指定操作系统和 Office 版本如何确定 Office 加载项使用的浏览器。
ms.date: 05/28/2019
localization_priority: Priority
ms.openlocfilehash: 92218bb012ae9031ebfc429606885a0ec0ea85b3
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/30/2019
ms.locfileid: "34592127"
---
# <a name="browsers-used-by-office-add-ins"></a><span data-ttu-id="5ee6c-103">Office 加载项使用的浏览器</span><span class="sxs-lookup"><span data-stu-id="5ee6c-103">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="5ee6c-104">Office 加载项是使用 iFrames（在 Office Online 中运行时）和使用 Office 桌面版和移动版客户端中的嵌入式浏览器控件显示的 Web 应用程序。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-104">Office add-ins are web applications that are displayed using iFrames when running in Office Online and using embedded browser controls in Office for desktop and mobile clients.</span></span> <span data-ttu-id="5ee6c-105">加载项还需要使用 JavaScript 引擎来运行 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-105">Add-ins also need a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="5ee6c-106">嵌入式浏览器和引擎均由用户计算机上安装的浏览器提供。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-106">Both the embedded browser and the engine are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="5ee6c-107">要使用的浏览器取决于：</span><span class="sxs-lookup"><span data-stu-id="5ee6c-107">Which browser is used depends on:</span></span>

- <span data-ttu-id="5ee6c-108">计算机的操作系统。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-108">The computer’s operating system.</span></span>
- <span data-ttu-id="5ee6c-109">加载项是在 Office Online、Office 365 还是非订阅版 Office 2013 或更高版本中运行。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-109">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="5ee6c-110">下表显示在不同平台和操作系统中使用的浏览器。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-110">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="5ee6c-111">**操作系统/平台**</span><span class="sxs-lookup"><span data-stu-id="5ee6c-111">**OS / Platform**</span></span>|<span data-ttu-id="5ee6c-112">**浏览器**</span><span class="sxs-lookup"><span data-stu-id="5ee6c-112">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5ee6c-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="5ee6c-113">Office Online</span></span>|<span data-ttu-id="5ee6c-114">在其中打开 Office Online 的浏览器。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-114">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="5ee6c-115">Mac</span><span class="sxs-lookup"><span data-stu-id="5ee6c-115">Mac</span></span>|<span data-ttu-id="5ee6c-116">Safari</span><span class="sxs-lookup"><span data-stu-id="5ee6c-116">Safari</span></span>|
|<span data-ttu-id="5ee6c-117">iOS</span><span class="sxs-lookup"><span data-stu-id="5ee6c-117">iOS</span></span>|<span data-ttu-id="5ee6c-118">Safari</span><span class="sxs-lookup"><span data-stu-id="5ee6c-118">Safari</span></span>|
|<span data-ttu-id="5ee6c-119">Android</span><span class="sxs-lookup"><span data-stu-id="5ee6c-119">Android</span></span>|<span data-ttu-id="5ee6c-120">Chrome</span><span class="sxs-lookup"><span data-stu-id="5ee6c-120">Chrome</span></span>|
|<span data-ttu-id="5ee6c-121">Windows/非订阅版 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="5ee6c-121">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="5ee6c-122">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="5ee6c-122">Internet Explorer 11</span></span>|
|<span data-ttu-id="5ee6c-123">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="5ee6c-123">Windows 10 ver.</span></span> <span data-ttu-id="5ee6c-124">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="5ee6c-124">< 1903 / Office 365</span></span>|<span data-ttu-id="5ee6c-125">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="5ee6c-125">Internet Explorer 11</span></span>|
|<span data-ttu-id="5ee6c-126">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="5ee6c-126">Windows 10 ver.</span></span> <span data-ttu-id="5ee6c-127">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="5ee6c-127">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="5ee6c-128">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="5ee6c-128">Internet Explorer 11</span></span>|
|<span data-ttu-id="5ee6c-129">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="5ee6c-129">Windows 10 ver.</span></span> <span data-ttu-id="5ee6c-130">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="5ee6c-130">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="5ee6c-131">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="5ee6c-131">Microsoft Edge\*</span></span>|

<span data-ttu-id="5ee6c-132">\*使用 Microsoft Edge 时，Windows 10 讲述人（有时称为“屏幕阅读器”）会读出页面中在任务窗格中打开的 `<title>` 标记。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-132">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="5ee6c-133">如果使用的是 Internet Explorer 11，则Narrator 将会读取任务窗格的标题栏，它来自加载项清单中的 `<DisplayName>` 值。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-133">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="5ee6c-134">Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-134">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="5ee6c-135">如果任何加载项用户安装的是使用 Internet Explorer 11 的平台，若要使用 ECMAScript 2015 或更高版本的语法和功能，则必须将 JavaScript 转换为 ES5 或使用填充代码。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-135">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="5ee6c-136">此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-136">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="5ee6c-137">在它们公开发布之前，你需要是 Windows 预览体验成员才能获得 Windows 版本 1903 或更高版本，并且需要是 Office 预览体验成员才能获得 Office 版本 16.0.11629 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-137">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="5ee6c-138">若要加入 Windows 预览体验成员：</span><span class="sxs-lookup"><span data-stu-id="5ee6c-138">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="5ee6c-139">转至 [Windows 预览体验成员](https://insider.windows.com)并单击链接以加入 Windows 预览体验成员。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-139">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="5ee6c-140">系统会将你引导至包含有关如何使用 Windows 设置支持 Windows 预览内部版本说明的页面。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-140">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="5ee6c-141">按照说明执行操作。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-141">Follow the instructions.</span></span> <span data-ttu-id="5ee6c-142">选择更新频率时，请选择时间最短的选项。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-142">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="5ee6c-143">若要加入 Office 预览体验成员：</span><span class="sxs-lookup"><span data-stu-id="5ee6c-143">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="5ee6c-144">转至 [Office 预览体验成员入门](https://insider.office.com/join)。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-144">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="5ee6c-145">按照加入页面上的说明操作。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-145">Follow the instruction on that page to join.</span></span> <span data-ttu-id="5ee6c-146">系统要求你指定频道时，请选择预览体验成员。</span><span class="sxs-lookup"><span data-stu-id="5ee6c-146">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="5ee6c-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="5ee6c-147">See also</span></span>

- [<span data-ttu-id="5ee6c-148">Office 加载项的运行要求</span><span class="sxs-lookup"><span data-stu-id="5ee6c-148">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
