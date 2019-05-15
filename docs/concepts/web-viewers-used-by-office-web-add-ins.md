---
title: Office 加载项使用的 Web 查看器
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 6cb0d6e97dd559727b6a1e140d8417e1146e479a
ms.sourcegitcommit: 944cbb5c6ce055f6db1833182b24d490d1dce01d
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/14/2019
ms.locfileid: "33992124"
---
# <a name="web-viewers-used-by-office-add-ins"></a><span data-ttu-id="ccaeb-102">Office 加载项使用的 Web 查看器</span><span class="sxs-lookup"><span data-stu-id="ccaeb-102">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="ccaeb-103">Office 加载项为 Web 应用程序，因此，它们需要通过 Web 页面查看器才能显示 Web 应用程序的 HTML 页面，并且需要通过 JavaScript 引擎才能运行 JavaScript。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-103">Since Office Add-ins are web applications, they need a web page viewer to display the HTML pages of the web application and a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="ccaeb-104">两者均由用户计算机上安装的浏览器提供。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-104">Both are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="ccaeb-105">要使用的浏览器取决于：</span><span class="sxs-lookup"><span data-stu-id="ccaeb-105">Which browser is used depends on:</span></span>

- <span data-ttu-id="ccaeb-106">计算机的操作系统。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-106">The computer’s operating system.</span></span>
- <span data-ttu-id="ccaeb-107">加载项是在 Office Online、Office 365 还是非订阅版 Office 2013 或更高版本中运行。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-107">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="ccaeb-108">下表显示在不同平台和操作系统中使用的浏览器。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-108">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="ccaeb-109">**操作系统/平台**</span><span class="sxs-lookup"><span data-stu-id="ccaeb-109">**OS / Platform**</span></span>|<span data-ttu-id="ccaeb-110">**浏览器**</span><span class="sxs-lookup"><span data-stu-id="ccaeb-110">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ccaeb-111">Office Online</span><span class="sxs-lookup"><span data-stu-id="ccaeb-111">Office Online</span></span>|<span data-ttu-id="ccaeb-112">在其中打开 Office Online 的浏览器。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-112">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="ccaeb-113">Mac</span><span class="sxs-lookup"><span data-stu-id="ccaeb-113">Mac</span></span>|<span data-ttu-id="ccaeb-114">Safari</span><span class="sxs-lookup"><span data-stu-id="ccaeb-114">Safari</span></span>|
|<span data-ttu-id="ccaeb-115">iOS</span><span class="sxs-lookup"><span data-stu-id="ccaeb-115">iOS</span></span>|<span data-ttu-id="ccaeb-116">Safari</span><span class="sxs-lookup"><span data-stu-id="ccaeb-116">Safari</span></span>|
|<span data-ttu-id="ccaeb-117">Android</span><span class="sxs-lookup"><span data-stu-id="ccaeb-117">Android</span></span>|<span data-ttu-id="ccaeb-118">Chrome</span><span class="sxs-lookup"><span data-stu-id="ccaeb-118">Chrome</span></span>|
|<span data-ttu-id="ccaeb-119">Windows/非订阅版 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="ccaeb-119">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="ccaeb-120">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="ccaeb-120">Internet Explorer 11</span></span>|
|<span data-ttu-id="ccaeb-121">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="ccaeb-121">Windows 10 ver.</span></span> <span data-ttu-id="ccaeb-122">< 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="ccaeb-122">< 1903 / Office 365</span></span>|<span data-ttu-id="ccaeb-123">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="ccaeb-123">Internet Explorer 11</span></span>|
|<span data-ttu-id="ccaeb-124">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="ccaeb-124">Windows 10 ver.</span></span> <span data-ttu-id="ccaeb-125">>= 1903 / Office 365 ver < 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="ccaeb-125">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="ccaeb-126">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="ccaeb-126">Internet Explorer 11</span></span>|
|<span data-ttu-id="ccaeb-127">Windows 10 版本</span><span class="sxs-lookup"><span data-stu-id="ccaeb-127">Windows 10 ver.</span></span> <span data-ttu-id="ccaeb-128">>= 1903 / Office 365 ver >= 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="ccaeb-128">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="ccaeb-129">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="ccaeb-129">Microsoft Edge\*</span></span>|

<span data-ttu-id="ccaeb-130">\*使用 Microsoft Edge 时，Windows 10 讲述人（有时称为“屏幕阅读器”）会读出页面中在任务窗格中打开的 `<title>` 标记。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-130">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="ccaeb-131">如果使用的是 Internet Explorer 11，则Narrator 将会读取任务窗格的标题栏，它来自加载项清单中的 `<DisplayName>` 值。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-131">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ccaeb-132">Internet Explorer 11 不支持高于 ES5 的 JavaScript 版本。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-132">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="ccaeb-133">如果任何加载项用户安装的是使用 Internet Explorer 11 的平台，若要使用 ECMAScript 2015 或更高版本的语法和功能，则必须将 JavaScript 转换为 ES5 或使用填充代码。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-133">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="ccaeb-134">此外，Internet Explorer 11 不支持媒体、录制和位置等部分 HTML5 功能。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-134">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="ccaeb-135">在它们公开发布之前，你需要是 Windows 预览体验成员才能获得 Windows 版本 1903 或更高版本，并且需要是 Office 预览体验成员才能获得 Office 版本 16.0.11629 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-135">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="ccaeb-136">若要加入 Windows 预览体验成员：</span><span class="sxs-lookup"><span data-stu-id="ccaeb-136">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="ccaeb-137">转至 [Windows 预览体验成员](https://insider.windows.com)并单击链接以加入 Windows 预览体验成员。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-137">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="ccaeb-138">系统会将你引导至包含有关如何使用 Windows 设置支持 Windows 预览内部版本说明的页面。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-138">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="ccaeb-139">按照说明执行操作。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-139">Follow the instructions.</span></span> <span data-ttu-id="ccaeb-140">选择更新频率时，请选择时间最短的选项。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-140">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="ccaeb-141">若要加入 Office 预览体验成员：</span><span class="sxs-lookup"><span data-stu-id="ccaeb-141">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="ccaeb-142">转至 [Office 预览体验成员入门](https://insider.office.com/join)。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-142">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="ccaeb-143">按照加入页面上的说明操作。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-143">Follow the instruction on that page to join.</span></span> <span data-ttu-id="ccaeb-144">系统要求你指定频道时，请选择预览体验成员。</span><span class="sxs-lookup"><span data-stu-id="ccaeb-144">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="ccaeb-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ccaeb-145">See also</span></span>

- [<span data-ttu-id="ccaeb-146">Office 加载项的运行要求</span><span class="sxs-lookup"><span data-stu-id="ccaeb-146">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
