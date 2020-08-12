---
title: 运行 Office 加载项的要求
description: 了解最终用户运行 Office 外接程序所需的客户端和服务器要求。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 49e1799961a0367d9eaf00415375c98a42534ba9
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641226"
---
# <a name="requirements-for-running-office-add-ins"></a><span data-ttu-id="aad66-103">运行 Office 加载项的要求</span><span class="sxs-lookup"><span data-stu-id="aad66-103">Requirements for running Office Add-ins</span></span>

<span data-ttu-id="aad66-104">本文介绍了运行 Office 加载项的软件和设备要求。</span><span class="sxs-lookup"><span data-stu-id="aad66-104">This article describes the software and device requirements for running Office Add-ins.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

<span data-ttu-id="aad66-105">若要概览 Office 加载项的当前受支持情况，请参阅 [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)。</span><span class="sxs-lookup"><span data-stu-id="aad66-105">For a high-level view of where Office Add-ins are currently supported, see [Office Add-in host and platform availability](../overview/office-add-in-availability.md).</span></span>

## <a name="server-requirements"></a><span data-ttu-id="aad66-106">服务器要求</span><span class="sxs-lookup"><span data-stu-id="aad66-106">Server requirements</span></span>

<span data-ttu-id="aad66-107">为了能够安装和运行任何 Office 外接程序，您需要首先将 UI 的清单和网页文件以及外接程序的代码部署到服务器上的合适位置。</span><span class="sxs-lookup"><span data-stu-id="aad66-107">To be able to install and run any Office Add-in, you first need to deploy the manifest and webpage files for the UI and code of your add-in to the appropriate server locations.</span></span>

<span data-ttu-id="aad66-108">对于所有类型的外接程序（内容、Outlook 和任务窗格外接程序以及外接程序命令），你需要将你的外接程序的网页文件部署到 Web 服务器或 Web 托管服务，如 [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md)。</span><span class="sxs-lookup"><span data-stu-id="aad66-108">For all types of add-ins (content, Outlook, and task pane add-ins and add-in commands), you need to deploy your add-in's webpage files to a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).</span></span>

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> <span data-ttu-id="aad66-109">在 Visual Studio 中开发和调试加载项时，Visual Studio 使用 IIS Express 在本地部署并运行加载项的网页文件，无需使用其他 Web 服务器。</span><span class="sxs-lookup"><span data-stu-id="aad66-109">When you develop and debug an add-in in Visual Studio, Visual Studio deploys and runs your add-in's webpage files locally with IIS Express, and doesn't require an additional web server.</span></span>

<span data-ttu-id="aad66-110">对于内容和任务窗格外接程序，在受支持的 Office 主机应用程序（Excel、PowerPoint、Project 或 Word）中，您还需要 SharePoint 上的[应用程序目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)来上载外接程序的 XML 清单文件，或者需要使用[集中部署](../publish/centralized-deployment.md)来部署加载项。</span><span class="sxs-lookup"><span data-stu-id="aad66-110">For content and task pane add-ins, in the supported Office host applications - Excel, PowerPoint, Project, or Word - you also need either an [app catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) on SharePoint to upload the add-in's XML manifest file, or you need to deploy the add-in using [Centralized Deployment](../publish/centralized-deployment.md).</span></span>

<span data-ttu-id="aad66-111">若要测试和运行 Outlook 外接程序，用户的 Outlook 电子邮件帐户必须驻留在 Exchange 2013 或更高版本中，可通过 Microsoft 365、Exchange Online 或本地安装获得。</span><span class="sxs-lookup"><span data-stu-id="aad66-111">To test and run an Outlook add-in, the user's Outlook email account must reside on Exchange 2013 or later, which is available through Microsoft 365, Exchange Online, or through an on-premises installation.</span></span> <span data-ttu-id="aad66-112">用户或管理员在该服务器上安装 Outlook 外接程序的清单文件。</span><span class="sxs-lookup"><span data-stu-id="aad66-112">The user or administrator installs manifest files for Outlook add-ins on that server.</span></span>

> [!NOTE]
> <span data-ttu-id="aad66-113">Outlook 中的 POP 和 IMAP 电子邮件帐户不支持 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="aad66-113">POP and IMAP email accounts in Outlook don't support Office Add-ins.</span></span>

## <a name="client-requirements-windows-desktop-and-tablet"></a><span data-ttu-id="aad66-114">客户端要求：Windows 台式机和平板电脑</span><span class="sxs-lookup"><span data-stu-id="aad66-114">Client requirements: Windows desktop and tablet</span></span>

<span data-ttu-id="aad66-115">为基于 Windows 的台式机、笔记本电脑或平板电脑设备上运行的受支持 Office 桌面客户端或 Web 客户端开发 Office 外接程序，需要以下软件：</span><span class="sxs-lookup"><span data-stu-id="aad66-115">The following software is required for developing an Office Add-in for the supported Office desktop clients or web clients that run on Windows-based desktop, laptop, or tablet devices:</span></span>


- <span data-ttu-id="aad66-116">对于 Windows x86 和 x64 台式机与平板电脑（如 Surface Pro）：</span><span class="sxs-lookup"><span data-stu-id="aad66-116">For Windows x86 and x64 desktops, and tablets such as Surface Pro:</span></span>
    - <span data-ttu-id="aad66-117">在 Windows 7 或更高版本上运行的 32 位或 64 位版本 Office 2013。</span><span class="sxs-lookup"><span data-stu-id="aad66-117">The 32- or 64-bit version of Office 2013 or a later version, running on Windows 7 or a later version.</span></span>
    - <span data-ttu-id="aad66-p102">Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013 或更高版本的 Office 客户端，（如果您正在专门为这些 Office 桌面客户端测试或运行 Office 外接程序）。Office 桌面客户端可以在本地安装或通过即点即用安装在客户端计算机上。</span><span class="sxs-lookup"><span data-stu-id="aad66-p102">Excel 2013, Outlook 2013, PowerPoint 2013, Project Professional 2013, Project 2013 SP1, Word 2013, or a later version of the Office client, if you are testing or running an Office Add-in specifically for one of these Office desktop clients. Office desktop clients can be installed on premises or via Click-to-Run on the client computer.</span></span>

  <span data-ttu-id="aad66-120">如果您拥有有效的 Microsoft 365 订阅，并且您没有对 Office 客户端的访问权限，则可以[下载并安装最新版本的 office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658)。</span><span class="sxs-lookup"><span data-stu-id="aad66-120">If you have a valid Microsoft 365 subscription and you do not have access to the Office client, you can [download and install the latest version of Office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658).</span></span>

- <span data-ttu-id="aad66-121">必须安装 Internet Explorer 11 或 Microsoft Edge（由 Windows 和 Office 版本而定），但它们不能是默认浏览器。</span><span class="sxs-lookup"><span data-stu-id="aad66-121">Internet Explorer 11 or Microsoft Edge (depending on the Windows and Office versions) must be installed but doesn't have to be the default browser.</span></span> <span data-ttu-id="aad66-122">为支持 Office 加载项，充当主机的 Office 客户端使用了 Internet Explorer 11 或 Microsoft Edge 所包含的浏览器组件。</span><span class="sxs-lookup"><span data-stu-id="aad66-122">To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 11 or Microsoft Edge.</span></span> <span data-ttu-id="aad66-123">有关更多详细信息，请参阅 [Office加载项使用的浏览器](browsers-used-by-office-web-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="aad66-123">See [Browsers used by Office Add-ins](browsers-used-by-office-web-add-ins.md) for more details.</span></span>

  > [!NOTE]
  > <span data-ttu-id="aad66-124">必须关闭 Internet Explorer 的增强安全配置 (ESC) 才能使 Office Web 加载项正常工作。</span><span class="sxs-lookup"><span data-stu-id="aad66-124">Internet Explorer's Enhanced Security Configuration (ESC) must be turned off for Office Web Add-ins to work.</span></span> <span data-ttu-id="aad66-125">如果在开发加载项时使用 Windows Server 计算机作为客户端，请注意 Windows Server 中会默认打开 ESC。</span><span class="sxs-lookup"><span data-stu-id="aad66-125">If you are using a Windows Server computer as your client when developing add-ins, note that ESC is turned on by default in Windows Server.</span></span>

- <span data-ttu-id="aad66-126">默认浏览器是下述软件之一：Internet Explorer 11，或者 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版。</span><span class="sxs-lookup"><span data-stu-id="aad66-126">One of the following as the default browser: Internet Explorer 11, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>
- <span data-ttu-id="aad66-127">HTML 和 JavaScript 编辑器（如记事本）、[Visual Studio 和 Microsoft 开发人员工具](https://www.visualstudio.com/features/office-tools-vs) 或第三方 Web 开发工具。</span><span class="sxs-lookup"><span data-stu-id="aad66-127">An HTML and JavaScript editor such as Notepad, [Visual Studio and the Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs), or a third-party web development tool.</span></span>

## <a name="client-requirements-os-x-desktop"></a><span data-ttu-id="aad66-128">客户端要求：OS X 桌面</span><span class="sxs-lookup"><span data-stu-id="aad66-128">Client requirements: OS X desktop</span></span>

<span data-ttu-id="aad66-129">Mac 上作为 Microsoft 365 的一部分分发的 outlook 支持 Outlook 外接程序。在 Mac 上运行 Outlook 外接程序与 Mac 本身具有相同的要求：操作系统必须至少是 OS X v 10.10 "Yosemite"。</span><span class="sxs-lookup"><span data-stu-id="aad66-129">Outlook on Mac, which is distributed as part of Microsoft 365, supports Outlook add-ins. Running Outlook add-ins in Outlook on Mac has the same requirements as Outlook on Mac itself: the operating system must be at least OS X v10.10 "Yosemite".</span></span> <span data-ttu-id="aad66-130">由于 Mac 版 Outlook 使用 WebKit 作为布局引擎以呈现加载项页，因此没有其他浏览器依赖项。</span><span class="sxs-lookup"><span data-stu-id="aad66-130">Because Outlook on Mac uses WebKit as a layout engine to render the add-in pages, there is no additional browser dependency.</span></span>

<span data-ttu-id="aad66-131">以下是支持 Office 加载项的 Mac 版 Office 的最低客户端版本。</span><span class="sxs-lookup"><span data-stu-id="aad66-131">The following are the minimum client versions of Office on Mac that support Office Add-ins.</span></span>

- <span data-ttu-id="aad66-132">Word 版本 15.18 (160109)</span><span class="sxs-lookup"><span data-stu-id="aad66-132">Word version 15.18 (160109)</span></span>
- <span data-ttu-id="aad66-133">Excel 版本 15.19 (160206)</span><span class="sxs-lookup"><span data-stu-id="aad66-133">Excel version 15.19 (160206)</span></span>
- <span data-ttu-id="aad66-134">PowerPoint 版本 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="aad66-134">PowerPoint version 15.24 (160614)</span></span>

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a><span data-ttu-id="aad66-135">客户端要求：针对 Office Web 客户端和 SharePoint 的浏览器支持</span><span class="sxs-lookup"><span data-stu-id="aad66-135">Client requirements: Browser support for Office web clients and SharePoint</span></span>

<span data-ttu-id="aad66-136">支持 ECMAScript 5.1、HTML5 和 CSS3（例如 Internet Explorer 11，或者 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版）的任意浏览器。</span><span class="sxs-lookup"><span data-stu-id="aad66-136">Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 11, or the latest version of Microsoft Edge, Chrome, Firefox, or Safari (Mac OS).</span></span>


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a><span data-ttu-id="aad66-137">客户端要求：非 Windows 智能手机和平板电脑</span><span class="sxs-lookup"><span data-stu-id="aad66-137">Client requirements: non-Windows smartphone and tablet</span></span>

<span data-ttu-id="aad66-138">尤其对于在智能手机或非 Windows 平板电脑设备上的浏览器中运行的 Outlook，需要以下软件才能测试和运行 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="aad66-138">Specifically for Outlook running in a browser on smartphones and non-Windows tablet devices, the following software is required for testing and running Outlook add-ins.</span></span>


| <span data-ttu-id="aad66-139">主机应用程序</span><span class="sxs-lookup"><span data-stu-id="aad66-139">Host application</span></span> | <span data-ttu-id="aad66-140">设备</span><span class="sxs-lookup"><span data-stu-id="aad66-140">Device</span></span> | <span data-ttu-id="aad66-141">操作系统</span><span class="sxs-lookup"><span data-stu-id="aad66-141">Operating system</span></span> | <span data-ttu-id="aad66-142">Exchange 帐户</span><span class="sxs-lookup"><span data-stu-id="aad66-142">Exchange account</span></span> | <span data-ttu-id="aad66-143">移动浏览器</span><span class="sxs-lookup"><span data-stu-id="aad66-143">Mobile browser</span></span> |
|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="aad66-144">Android 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="aad66-144">Outlook on Android</span></span>|<span data-ttu-id="aad66-145">Android 平板电脑和智能手机</span><span class="sxs-lookup"><span data-stu-id="aad66-145">Android tablets and smartphones</span></span>|<span data-ttu-id="aad66-146">Android 4.4 KitKat 及更高版本</span><span class="sxs-lookup"><span data-stu-id="aad66-146">Android 4.4 KitKat later</span></span>|<span data-ttu-id="aad66-147">针对商业版或 Exchange Online 的 Microsoft 365 应用程序的最新更新</span><span class="sxs-lookup"><span data-stu-id="aad66-147">On the latest update of Microsoft 365 Apps for business or Exchange Online</span></span>|<span data-ttu-id="aad66-148">Android 版本机应用（不适用于浏览器）</span><span class="sxs-lookup"><span data-stu-id="aad66-148">Native app for Android, browser not applicable</span></span>|
|<span data-ttu-id="aad66-149">iOS 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="aad66-149">Outlook on iOS</span></span>|<span data-ttu-id="aad66-150">iPad 平板电脑，iPhone 智能手机</span><span class="sxs-lookup"><span data-stu-id="aad66-150">iPad tablets, iPhone smartphones</span></span>|<span data-ttu-id="aad66-151">iOS 11 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aad66-151">iOS 11 or later</span></span>|<span data-ttu-id="aad66-152">针对商业版或 Exchange Online 的 Microsoft 365 应用程序的最新更新</span><span class="sxs-lookup"><span data-stu-id="aad66-152">On the latest update of Microsoft 365 Apps for business or Exchange Online</span></span>|<span data-ttu-id="aad66-153">iOS 版本机应用（不适用于浏览器）</span><span class="sxs-lookup"><span data-stu-id="aad66-153">Native app for iOS, browser not applicable</span></span>|
|<span data-ttu-id="aad66-154">Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="aad66-154">Outlook on the web</span></span>|<span data-ttu-id="aad66-155">iPhone 4 或更高版本、iPad 2 或更高版本、iPod Touch 4 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aad66-155">iPhone 4 or later, iPad 2 or later, iPod Touch 4 or later</span></span>|<span data-ttu-id="aad66-156">iOS 5 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aad66-156">iOS 5 or later</span></span>|<span data-ttu-id="aad66-157">在 Microsoft 365、Exchange Online 或 Exchange Server 2013 或更高版本上的本地部署</span><span class="sxs-lookup"><span data-stu-id="aad66-157">On Microsoft 365, Exchange Online, or on premises on Exchange Server 2013 or later</span></span>|<span data-ttu-id="aad66-158">Safari</span><span class="sxs-lookup"><span data-stu-id="aad66-158">Safari</span></span>|

> [!NOTE]
> <span data-ttu-id="aad66-159">Android 版本机应用 OWA、iPad 版 OWA 和 iPhone 版 OWA 现已[弃用](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b)且之后无需这些软件即可测试 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="aad66-159">The native apps OWA for Android, OWA for iPad, and OWA for iPhone have been [deprecated](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b) and are no longer required or available for testing Outlook add-ins.</span></span>


## <a name="see-also"></a><span data-ttu-id="aad66-160">另请参阅</span><span class="sxs-lookup"><span data-stu-id="aad66-160">See also</span></span>

- [<span data-ttu-id="aad66-161">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="aad66-161">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="aad66-162">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="aad66-162">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
- [<span data-ttu-id="aad66-163">Office 加载项使用的浏览器</span><span class="sxs-lookup"><span data-stu-id="aad66-163">Browsers used by Office Add-ins</span></span>](browsers-used-by-office-web-add-ins.md)
