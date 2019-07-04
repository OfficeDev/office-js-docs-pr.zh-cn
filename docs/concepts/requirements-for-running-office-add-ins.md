---
title: 运行 Office 加载项的要求
description: ''
ms.date: 07/01/2019
localization_priority: Priority
ms.openlocfilehash: 5a33af6a3dc23739642a4ad0f6e3d29bff247f4d
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454613"
---
# <a name="requirements-for-running-office-add-ins"></a><span data-ttu-id="d2664-102">运行 Office 加载项的要求</span><span class="sxs-lookup"><span data-stu-id="d2664-102">Requirements for running Office Add-ins</span></span>

<span data-ttu-id="d2664-103">本文介绍了运行 Office 加载项的软件和设备要求。</span><span class="sxs-lookup"><span data-stu-id="d2664-103">This article describes the software and device requirements for running Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="d2664-p101">如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。</span><span class="sxs-lookup"><span data-stu-id="d2664-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

<span data-ttu-id="d2664-106">若要概览 Office 加载项的当前受支持情况，请参阅 [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="d2664-106">For a high-level view of where Office Add-ins are currently supported, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="server-requirements"></a><span data-ttu-id="d2664-107">服务器要求</span><span class="sxs-lookup"><span data-stu-id="d2664-107">Server requirements</span></span>

<span data-ttu-id="d2664-108">为了能够安装和运行任何 Office 外接程序，您需要首先将 UI 的清单和网页文件以及外接程序的代码部署到服务器上的合适位置。</span><span class="sxs-lookup"><span data-stu-id="d2664-108">To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.</span></span>

<span data-ttu-id="d2664-109">对于所有类型的外接程序（内容、Outlook 和任务窗格外接程序以及外接程序命令），你需要将你的外接程序的网页文件部署到 Web 服务器或 Web 托管服务，如 [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md)。</span><span class="sxs-lookup"><span data-stu-id="d2664-109">For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> <span data-ttu-id="d2664-110">在 Visual Studio 中开发和调试加载项时，Visual Studio 使用 IIS Express 在本地部署并运行加载项的网页文件，无需使用其他 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="d2664-110">When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server.</span></span>

<span data-ttu-id="d2664-111">对于内容和任务窗格加载项，在受支持的 Office 主机应用程序（Excel、PowerPoint、Project 或 Word）中，你还需要 SharePoint 上的一个[应用目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)才能上载加载项的 XML 清单文件。</span><span class="sxs-lookup"><span data-stu-id="d2664-111">For content and task pane add-ins, in the supported Office host applications - Access web apps, Word, Excel, PowerPoint, or Project - you also need an [add-in catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file.</span></span>

<span data-ttu-id="d2664-p102">要测试和运行 Outlook 外接程序，用户的 Outlook 电子邮件帐户必须位于 Exchange 2013 或更高版本上，可通过 Office 365、Exchange Online 或本地安装获得此软件。用户或管理员在该服务器上安装 Outlook 外接程序的清单文件。</span><span class="sxs-lookup"><span data-stu-id="d2664-p102">To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Office 365, Exchange Online, or through an on-premises installation. The user or administrator installs manifest files for Outlook add-ins on that server.</span></span>

> [!NOTE]
> <span data-ttu-id="d2664-114">Outlook 中的 POP 和 IMAP 电子邮件帐户不支持 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="d2664-114">POP and IMAP email accounts in Outlook don't support Office Add-ins.</span></span>

## <a name="client-requirements-windows-desktop-and-tablet"></a><span data-ttu-id="d2664-115">客户端要求：Windows 台式机和平板电脑</span><span class="sxs-lookup"><span data-stu-id="d2664-115">Client requirements: Windows desktop and tablet</span></span>

<span data-ttu-id="d2664-116">为基于 Windows 的台式机、笔记本电脑或平板电脑设备上运行的受支持 Office 桌面客户端或 Web 客户端开发 Office 外接程序，需要以下软件：</span><span class="sxs-lookup"><span data-stu-id="d2664-116">The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices:</span></span>


- <span data-ttu-id="d2664-117">对于 Windows x86 和 x64 台式机与平板电脑（如 Surface Pro）：</span><span class="sxs-lookup"><span data-stu-id="d2664-117">For Windows x86 and x64 desktops, and tablets such as Surface Pro:</span></span>
    - <span data-ttu-id="d2664-118">在 Windows 7 或更高版本上运行的 32 位或 64 位版本 Office 2013。</span><span class="sxs-lookup"><span data-stu-id="d2664-118">The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.</span></span>
    - <span data-ttu-id="d2664-p103">Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013 或更高版本的 Office 客户端，（如果您正在专门为这些 Office 桌面客户端测试或运行 Office 外接程序）。Office 桌面客户端可以在本地安装或通过即点即用安装在客户端计算机上。</span><span class="sxs-lookup"><span data-stu-id="d2664-p103">Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.</span></span>

  <span data-ttu-id="d2664-121">如果拥有有效的 Office 365 订阅但无权访问 Office 客户端，则可[下载并安装最新版的 Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658)。</span><span class="sxs-lookup"><span data-stu-id="d2664-121">If you have a valid Office 365 subscription and you do not have access to the Office client, you can [download and install the latest version of Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).</span></span>

- <span data-ttu-id="d2664-122">必须安装 Internet Explorer 11 或 Microsoft Edge（由 Windows 和 Office 版本而定），但它们不能是默认浏览器。</span><span class="sxs-lookup"><span data-stu-id="d2664-122">Internet Explorer 11 or Microsoft Edge (depending on the Windows and Office versions) must be installed but doesn't have to be the default browser.</span></span> <span data-ttu-id="d2664-123">为支持 Office 加载项，充当主机的 Office 客户端使用了 Internet Explorer 11 或 Microsoft Edge 所包含的浏览器组件。</span><span class="sxs-lookup"><span data-stu-id="d2664-123">To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 11 or later.</span></span> <span data-ttu-id="d2664-124">有关更多详细信息，请参阅 [Office加载项使用的浏览器](browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="d2664-124">See [Browsers used by Office Add-ins](browsers-used-by-office-web-add-ins.md) for more details.</span></span>

  > [!NOTE]
  > <span data-ttu-id="d2664-125">必须关闭 Internet Explorer 的增强安全配置 (ESC) 才能使 Office Web 加载项正常工作。</span><span class="sxs-lookup"><span data-stu-id="d2664-125">Internet Explorer's Enhanced Security Configuration (ESC) must be turned off for Office Web Add-ins to work.</span></span> <span data-ttu-id="d2664-126">如果在开发加载项时使用 Windows Server 计算机作为客户端，请注意 Windows Server 中会默认打开 ESC。</span><span class="sxs-lookup"><span data-stu-id="d2664-126">If you are using a Windows Server computer as your client when developing add-ins, note that ESC is turned on by default in Windows Server.</span></span>

- <span data-ttu-id="d2664-127">默认浏览器是下述软件之一：Internet Explorer 11，或者 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版。</span><span class="sxs-lookup"><span data-stu-id="d2664-127">One of the following as the default browser: Internet Explorer 11 or later, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>
- <span data-ttu-id="d2664-128">HTML 和 JavaScript 编辑器（如记事本）、[Visual Studio 和 Microsoft 开发人员工具](https://www.visualstudio.com/features/office-tools-vs) 或第三方 Web 开发工具。</span><span class="sxs-lookup"><span data-stu-id="d2664-128">An HTML and JavaScript editor such as Notepad, [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or a third-party web development tool.</span></span>

## <a name="client-requirements-os-x-desktop"></a><span data-ttu-id="d2664-129">客户端要求：OS X 桌面</span><span class="sxs-lookup"><span data-stu-id="d2664-129">Client requirements: OS X desktop</span></span>

<span data-ttu-id="d2664-130">作为 Office 365 的一部分分发的 Mac 版 Outlook 支持 Outlook 加载项。在 Mac 版 Outlook 中运行 Outlook 加载项的要求与 Mac 版 Outlook 本身的要求相同：操作系统必须至少为 OS X v10.10 “Yosemite”。</span><span class="sxs-lookup"><span data-stu-id="d2664-130">Outlook for Mac, which is distributed as part of Office 365, supports Outlook add-ins. Running Outlook add-ins on Outlook for Mac has the same requirements as Outlook for Mac itself: the operating system must be at least OS X v10.10 "Yosemite". Because Outlook for Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.</span></span> <span data-ttu-id="d2664-131">由于 Mac 版 Outlook 使用 WebKit 作为布局引擎以呈现加载项页，因此没有其他浏览器依赖项。</span><span class="sxs-lookup"><span data-stu-id="d2664-131">Because olmac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.</span></span>

<span data-ttu-id="d2664-132">以下是支持 Office 加载项的 Mac 版 Office 的最低客户端版本。</span><span class="sxs-lookup"><span data-stu-id="d2664-132">The following are the minimum client versions of Office for Mac that support Office Add-ins:</span></span>

- <span data-ttu-id="d2664-133">Word 版本 15.18 (160109)</span><span class="sxs-lookup"><span data-stu-id="d2664-133">Word for Mac version 15.18 (160109)</span></span>
- <span data-ttu-id="d2664-134">Excel 版本 15.19 (160206)</span><span class="sxs-lookup"><span data-stu-id="d2664-134">Excel for Mac version 15.19 (160206)</span></span>
- <span data-ttu-id="d2664-135">PowerPoint 版本 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="d2664-135">PowerPoint for Mac version 15.24 (160614)</span></span>

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a><span data-ttu-id="d2664-136">客户端要求：针对 Office Web 客户端和 SharePoint 的浏览器支持</span><span class="sxs-lookup"><span data-stu-id="d2664-136">Client requirements: Browser support for Office Online web clients and SharePoint</span></span>

<span data-ttu-id="d2664-137">支持 ECMAScript 5.1、HTML5 和 CSS3（例如 Internet Explorer 11，或者 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版）的任意浏览器。</span><span class="sxs-lookup"><span data-stu-id="d2664-137">Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 11 or later, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a><span data-ttu-id="d2664-138">客户端要求：非 Windows 智能手机和平板电脑</span><span class="sxs-lookup"><span data-stu-id="d2664-138">Client requirements: non-Windows smartphone and tablet</span></span>

<span data-ttu-id="d2664-139">尤其对于在智能手机或非 Windows 平板电脑设备上的浏览器中运行的 Outlook，需要以下软件才能测试和运行 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="d2664-139">Specifically for Outlook Web App running in a browser on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.</span></span>


| <span data-ttu-id="d2664-140">主机应用程序</span><span class="sxs-lookup"><span data-stu-id="d2664-140">Host application</span></span> | <span data-ttu-id="d2664-141">设备</span><span class="sxs-lookup"><span data-stu-id="d2664-141">Device</span></span> | <span data-ttu-id="d2664-142">操作系统</span><span class="sxs-lookup"><span data-stu-id="d2664-142">Operating system</span></span> | <span data-ttu-id="d2664-143">Exchange 帐户</span><span class="sxs-lookup"><span data-stu-id="d2664-143">Exchange account</span></span> | <span data-ttu-id="d2664-144">移动浏览器</span><span class="sxs-lookup"><span data-stu-id="d2664-144">Mobile browser</span></span> |
|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d2664-145">Android 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="d2664-145">Re-access Outlook on the Android device.</span></span>|<span data-ttu-id="d2664-146">Android 平板电脑和智能手机</span><span class="sxs-lookup"><span data-stu-id="d2664-146">Android tablets and smartphones</span></span>|<span data-ttu-id="d2664-147">Android 4.4 KitKat 及更高版本</span><span class="sxs-lookup"><span data-stu-id="d2664-147">Android 4.4 KitKat later</span></span>|<span data-ttu-id="d2664-148">在 Office 365 for Business 或 Exchange Online 的最新更新上</span><span class="sxs-lookup"><span data-stu-id="d2664-148">On the latest update of Office 365 for business or Exchange Online</span></span>|<span data-ttu-id="d2664-149">Android 版本机应用（不适用于浏览器）</span><span class="sxs-lookup"><span data-stu-id="d2664-149">Native app for Android, browser not applicable</span></span>|
|<span data-ttu-id="d2664-150">iOS 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="d2664-150">Re-access Outlook on the iOS device.</span></span>|<span data-ttu-id="d2664-151">iPad 平板电脑，iPhone 智能手机</span><span class="sxs-lookup"><span data-stu-id="d2664-151">iPad tablets, iPhone smartphones</span></span>|<span data-ttu-id="d2664-152">iOS 11 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d2664-152">iOS 11 or later</span></span>|<span data-ttu-id="d2664-153">在 Office 365 for Business 或 Exchange Online 的最新更新上</span><span class="sxs-lookup"><span data-stu-id="d2664-153">On the latest update of Office 365 for business or Exchange Online</span></span>|<span data-ttu-id="d2664-154">iOS 版本机应用（不适用于浏览器）</span><span class="sxs-lookup"><span data-stu-id="d2664-154">Native app for iOS, browser not applicable</span></span>|
|<span data-ttu-id="d2664-155">Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="d2664-155">Outlook on the web</span></span>|<span data-ttu-id="d2664-156">iPhone 4 或更高版本、iPad 2 或更高版本、iPod Touch 4 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d2664-156">iPhone 4 or later, iPad 2 or later, iPod Touch 4 or later</span></span>|<span data-ttu-id="d2664-157">iOS 5 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d2664-157">iOS 5 or later</span></span>|<span data-ttu-id="d2664-158">在 Office 365、Exchange Online、或者本地 Exchange Server 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d2664-158">On Office 365, Exchange Online, or on premises on Exchange Server 2013 or later</span></span>|<span data-ttu-id="d2664-159">Safari</span><span class="sxs-lookup"><span data-stu-id="d2664-159">Safari</span></span>|

> [!NOTE]
> <span data-ttu-id="d2664-160">Android 版本机应用 OWA、iPad 版 OWA 和 iPhone 版 OWA 现已[弃用](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b)且之后无需这些软件即可测试 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="d2664-160">The native apps OWA for Android, OWA for iPad, and OWA for iPhone have been [deprecated](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) and are no longer required or available for testing Outlook add-ins.</span></span>


## <a name="see-also"></a><span data-ttu-id="d2664-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d2664-161">See also</span></span>

- [<span data-ttu-id="d2664-162">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="d2664-162">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="d2664-163">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="d2664-163">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="d2664-164">Office 加载项使用的浏览器</span><span class="sxs-lookup"><span data-stu-id="d2664-164">Web viewers used by Office Add-ins</span></span>](browsers-used-by-office-web-add-ins.md)
