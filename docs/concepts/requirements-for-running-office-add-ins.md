---
title: ?? Office ??????
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: a4859af73d8e9cf041990a3533894b24f1cbde6f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
---
# <a name="requirements-for-running-office-add-ins"></a><span data-ttu-id="cea6e-102">?? Office ??????</span><span class="sxs-lookup"><span data-stu-id="cea6e-102">Requirements for running Office Add-ins</span></span>

<span data-ttu-id="cea6e-103">??????? Office ????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-103">This article describes the software and device requirements for running Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="cea6e-p101">????????[??](../publish/publish.md)? AppSource ???? Office ???????? [AppSource ????](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)??????????????????????????????????????????[? 4.12 ??](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)?? [Office ?????????](../overview/office-add-in-availability.md)????</span><span class="sxs-lookup"><span data-stu-id="cea6e-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/en-us/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

<span data-ttu-id="cea6e-106">???? Office ??????????????? [Office ???????????](../overview/office-add-in-availability.md)?</span><span class="sxs-lookup"><span data-stu-id="cea6e-106">For a high-level view of where Office Add-ins are currently supported, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="server-requirements"></a><span data-ttu-id="cea6e-107">?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-107">Server requirements</span></span>

<span data-ttu-id="cea6e-108">??????????? Office ??????????? UI ??????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-108">To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.</span></span>

<span data-ttu-id="cea6e-109">???????????????Outlook ????????????????????????????????????? Web ???? Web ?????? [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md)?</span><span class="sxs-lookup"><span data-stu-id="cea6e-109">For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> <span data-ttu-id="cea6e-110">? Visual Studio ???????????Visual Studio ?? IIS Express ??????????????????????? Web ????</span><span class="sxs-lookup"><span data-stu-id="cea6e-110">When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server.</span></span> 

<span data-ttu-id="cea6e-111">??????????????????? Office ???????Access Web App?Word?Excel?PowerPoint ? Project??????? SharePoint ???? [??????](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)????????? XML ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-111">For content and task pane add-ins, in the supported Office host applications - Access web apps, Word, Excel, PowerPoint, or Project - you also need an [add-in catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file.</span></span>

<span data-ttu-id="cea6e-p102">?????? Outlook ???????? Outlook ?????????? Exchange 2013 ?????????? Office 365?Exchange Online ????????????????????????? Outlook ??????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-p102">To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Office 365, Exchange Online, or through an on-premises installation. The user or administrator installs manifest files for Outlook add-ins on that server.</span></span>

> [!NOTE]
> <span data-ttu-id="cea6e-114">Outlook ?? POP ? IMAP ????????? Office ????</span><span class="sxs-lookup"><span data-stu-id="cea6e-114">POP and IMAP email accounts in Outlook don't support Office Add-ins.</span></span>

## <a name="client-requirements-windows-desktop-and-tablet"></a><span data-ttu-id="cea6e-115">??????Windows ????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-115">Client requirements: Windows desktop and tablet</span></span>

<span data-ttu-id="cea6e-116">??? Windows ???????????????????????? Office ?????? Web ????? Office ????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-116">The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices:</span></span>


- <span data-ttu-id="cea6e-117">?? Windows x86 ? x64 ?????????? Surface Pro??</span><span class="sxs-lookup"><span data-stu-id="cea6e-117">For Windows x86 and x64 desktops, and tablets such as Surface Pro:</span></span>
    - <span data-ttu-id="cea6e-118">? Windows 7 ????????? 32 ?? 64 ??? Office 2013?</span><span class="sxs-lookup"><span data-stu-id="cea6e-118">The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.</span></span>
    - <span data-ttu-id="cea6e-p103">Excel 2013?Outlook 2013?PowerPoint 2013?Project Professional 2013?Project 2013 SP1?Word 2013 ?????? Office ??????????????? Office ?????????? Office ??????Office ??????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-p103">Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.</span></span>
    
  <span data-ttu-id="cea6e-121">?? Office 365 ?????????? Office 2013??????? CDN ?????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-121">If you have a valid Office 365 subscription and you do not have access to Office 2013, you can download it via one the following CDN links:</span></span>       
    - [<span data-ttu-id="cea6e-122">Office 2013 ??? (.exe)</span><span class="sxs-lookup"><span data-stu-id="cea6e-122">Office 2013 for Business (.exe)</span></span>](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 
    - [<span data-ttu-id="cea6e-123">Office 2013 ??? (.exe)</span><span class="sxs-lookup"><span data-stu-id="cea6e-123">Office 2013 for Home (.exe)</span></span>](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 

- <span data-ttu-id="cea6e-p104">????? Internet Explorer 11 ?????????????????? Office ?????????? Office ????????????? Internet Explorer 11 ??????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-p104">Internet Explorer 11 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 11 or later.</span></span>
- <span data-ttu-id="cea6e-126">????????????????Internet Explorer 11 ??????? Microsoft Edge?Chrome?Firefox ? Safari (Mac OS) ??????</span><span class="sxs-lookup"><span data-stu-id="cea6e-126">One of the following as the default browser: Internet Explorer 11 or later, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>
- <span data-ttu-id="cea6e-127">HTML ? JavaScript ??????????[Visual Studio ? Microsoft ??????](https://www.visualstudio.com/features/office-tools-vs) ???? Web ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-127">An HTML and JavaScript editor such as Notepad, [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or a third-party web development tool.</span></span>

## <a name="client-requirements-os-x-desktop"></a><span data-ttu-id="cea6e-128">??????OS X ??</span><span class="sxs-lookup"><span data-stu-id="cea6e-128">Client requirements: OS X desktop</span></span>

<span data-ttu-id="cea6e-p105">?? Office 365 ??????? ??? Mac ? Outlook ?? Outlook ?????? ??? Mac ? Outlook ??? Outlook ????? ??? Mac ? Outlook ????????????????? OS X v10.10"Yosemite"??? ??? Mac ? Outlook ?? WebKit ????????????????????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-p105">Outlook for Mac, which is distributed as part of Office 365, supports Outlook add-ins. Running Outlook add-ins on Outlook for Mac has the same requirements as Outlook for Mac itself: the operating system must be at least OS X v10.10 "Yosemite". Because Outlook for Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.</span></span>

<span data-ttu-id="cea6e-131">????? Office ????? Office for Mac ?????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-131">The following are the minimum client versions of Office for Mac that support Office Add-ins:</span></span>

- <span data-ttu-id="cea6e-132">Word for Mac ?? 15.18 (160109)</span><span class="sxs-lookup"><span data-stu-id="cea6e-132">Word for Mac version 15.18 (160109)</span></span> 
- <span data-ttu-id="cea6e-133">Excel for Mac ?? 15.19 (160206)</span><span class="sxs-lookup"><span data-stu-id="cea6e-133">Excel for Mac version 15.19 (160206)</span></span> 
- <span data-ttu-id="cea6e-134">PowerPoint for Mac ?? 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="cea6e-134">PowerPoint for Mac version 15.24 (160614)</span></span>

## <a name="client-requirements-browser-support-for-office-online-web-clients-and-sharepoint"></a><span data-ttu-id="cea6e-135">???????? Office Online Web ???? SharePoint ??????</span><span class="sxs-lookup"><span data-stu-id="cea6e-135">Client requirements: Browser support for Office Online web clients and SharePoint</span></span>

<span data-ttu-id="cea6e-136">?? ECMAScript 5.1?HTML5 ? CSS3 ???????? Internet Explorer 11 ??????? Microsoft Edge?Chrome?Firefox ? Safari (Mac OS) ??????</span><span class="sxs-lookup"><span data-stu-id="cea6e-136">Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 11 or later, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a><span data-ttu-id="cea6e-137">??????? Windows ?????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-137">Client requirements: non-Windows smartphone and tablet</span></span>

<span data-ttu-id="cea6e-138">???????????? Windows ??????????????? ?????? OWA ? Outlook Web App?????? Outlook ???????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-138">Specifically for OWA for Devices, and Outlook Web App running in a browser on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.</span></span>


| <span data-ttu-id="cea6e-139">??????</span><span class="sxs-lookup"><span data-stu-id="cea6e-139">Host application</span></span> | <span data-ttu-id="cea6e-140">??</span><span class="sxs-lookup"><span data-stu-id="cea6e-140">Device</span></span> | <span data-ttu-id="cea6e-141">????</span><span class="sxs-lookup"><span data-stu-id="cea6e-141">Operating system</span></span> | <span data-ttu-id="cea6e-142">Exchange ??</span><span class="sxs-lookup"><span data-stu-id="cea6e-142">Exchange account</span></span> | <span data-ttu-id="cea6e-143">?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-143">Mobile browser</span></span> |
|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="cea6e-144">OWA for Android</span><span class="sxs-lookup"><span data-stu-id="cea6e-144">OWA for Android</span></span>|<span data-ttu-id="cea6e-p106">Android ??????????? [Android OS](https://developer.android.com/guide/practices/screens_support.html) ???????"??"?"??"?</span><span class="sxs-lookup"><span data-stu-id="cea6e-p106">Android smartphones. Technically, those devices considered as "small" or "normal" by [Android OS](https://developer.android.com/guide/practices/screens_support.html).</span></span>|<span data-ttu-id="cea6e-147">Android 4.4 KitKat ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-147">Android 4.4 KitKat or later</span></span>|<span data-ttu-id="cea6e-148">?? Office 365 ??? ? Exchange Online ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-148">On the latest update of Office 365 for business or Exchange Online</span></span>|<span data-ttu-id="cea6e-149">?? Android ?????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-149">Native add-in for Android, browser not applicable</span></span>|
|<span data-ttu-id="cea6e-150">OWA for iPad</span><span class="sxs-lookup"><span data-stu-id="cea6e-150">OWA for iPad</span></span>|<span data-ttu-id="cea6e-151">iPad 2 ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-151">iPad 2 or later</span></span>|<span data-ttu-id="cea6e-152">iOS 6 ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-152">iOS 6 or later</span></span>|<span data-ttu-id="cea6e-153">?? Office 365 ??? ? Exchange Online ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-153">On the latest update of Office 365 for business or Exchange Online</span></span>|<span data-ttu-id="cea6e-154">?? iOS ?????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-154">Native add-in for iOS, browser not applicable</span></span>|
|<span data-ttu-id="cea6e-155">OWA for iPhone</span><span class="sxs-lookup"><span data-stu-id="cea6e-155">OWA for iPhone</span></span>|<span data-ttu-id="cea6e-156">iPhone 4S ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-156">iPhone 4S or later</span></span>|<span data-ttu-id="cea6e-157">iOS 6 ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-157">iOS 6 or later</span></span>|<span data-ttu-id="cea6e-158">?? Office 365 ??? ? Exchange Online ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-158">On the latest update of Office 365 for business or Exchange Online</span></span>|<span data-ttu-id="cea6e-159">?? iOS ?????????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-159">Native add-in for iOS, browser not applicable</span></span>|
|<span data-ttu-id="cea6e-160">Outlook Web App</span><span class="sxs-lookup"><span data-stu-id="cea6e-160">Outlook Web App</span></span>|<span data-ttu-id="cea6e-161">iPhone 4 ??????iPad 2 ??????iPod Touch 4 ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-161">iPhone 4 or later, iPad 2 or later, iPod Touch 4 or later</span></span>|<span data-ttu-id="cea6e-162">iOS 5 ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-162">iOS 5 or later</span></span>|<span data-ttu-id="cea6e-163">?? Office 365?Exchange Online ???? Exchange Server 2013 ?????</span><span class="sxs-lookup"><span data-stu-id="cea6e-163">On Office 365, Exchange Online, or on premise on Exchange Server 2013 or later</span></span>|<span data-ttu-id="cea6e-164">Safari</span><span class="sxs-lookup"><span data-stu-id="cea6e-164">Safari</span></span>|


## <a name="see-also"></a><span data-ttu-id="cea6e-165">????</span><span class="sxs-lookup"><span data-stu-id="cea6e-165">See also</span></span>

- [<span data-ttu-id="cea6e-166">Office ???????</span><span class="sxs-lookup"><span data-stu-id="cea6e-166">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="cea6e-167">Office ???????????</span><span class="sxs-lookup"><span data-stu-id="cea6e-167">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
