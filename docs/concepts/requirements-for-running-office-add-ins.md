---
title: 运行 Office 加载项的要求
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: a4859af73d8e9cf041990a3533894b24f1cbde6f
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437421"
---
# <a name="requirements-for-running-office-add-ins"></a>运行 Office 加载项的要求

本文介绍了运行 Office 加载项的软件和设备要求。

> [!NOTE]
> 如果计划将加载项[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/en-us/office/dev/store/validation-policies)。例如，加载项必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/en-us/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 加载项主机和可用性](../overview/office-add-in-availability.md)页面）。 

若要概览 Office 加载项的当前受支持情况，请参阅 [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)。

## <a name="server-requirements"></a>服务器要求

为了能够安装和运行任何 Office 外接程序，您需要首先将 UI 的清单和网页文件以及外接程序的代码部署到服务器上的合适位置。

对于所有类型的外接程序（内容、Outlook 和任务窗格外接程序以及外接程序命令），你需要将你的外接程序的网页文件部署到 Web 服务器或 Web 托管服务，如 [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md)。

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> 在 Visual Studio 中开发和调试加载项时，Visual Studio 使用 IIS Express 在本地部署并运行加载项的网页文件，无需使用其他 Web 服务器。 

对于内容和任务窗格外接程序，在受支持的 Office 主机应用程序（Access Web App、Word、Excel、PowerPoint 或 Project）中，你还需要 SharePoint 上的一个 [外接程序目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)才能上载外接程序的 XML 清单文件。

要测试和运行 Outlook 外接程序，用户的 Outlook 电子邮件帐户必须位于 Exchange 2013 或更高版本上，可通过 Office 365、Exchange Online 或本地安装获得此软件。用户或管理员在该服务器上安装 Outlook 外接程序的清单文件。

> [!NOTE]
> Outlook 中的 POP 和 IMAP 电子邮件帐户不支持 Office 加载项。

## <a name="client-requirements-windows-desktop-and-tablet"></a>客户端要求：Windows 台式机和平板电脑

为基于 Windows 的台式机、笔记本电脑或平板电脑设备上运行的受支持 Office 桌面客户端或 Web 客户端开发 Office 外接程序，需要以下软件：


- 对于 Windows x86 和 x64 台式机与平板电脑（如 Surface Pro）：
    - 在 Windows 7 或更高版本上运行的 32 位或 64 位版本 Office 2013。
    - Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013 或更高版本的 Office 客户端，（如果您正在专门为这些 Office 桌面客户端测试或运行 Office 外接程序）。Office 桌面客户端可以在本地安装或通过即点即用安装在客户端计算机上。
    
  如果 Office 365 订阅有效，但无权访问 Office 2013，可以通过以下 CDN 链接之一进行下载：       
    - [Office 2013 企业版 (.exe)](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365BusinessRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 
    - [Office 2013 家庭版 (.exe)](https://c2rsetup.officeapps.live.com/c2r/download.aspx?productReleaseID=O365HomePremRetail&platform=X86&language=en-us&version=O15GA&source=O15OLSO365) 

- 必须安装的 Internet Explorer 11 或更高版本无需是默认浏览器。为了支持 Office 外接程序，充当主机的 Office 客户端所使用的浏览器组件是 Internet Explorer 11 或更高版本的一部分。
- 将以下任一浏览器作为默认浏览器：Internet Explorer 11 或更高版本，或 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版本。
- HTML 和 JavaScript 编辑器（如记事本）、[Visual Studio 和 Microsoft 开发人员工具](https://www.visualstudio.com/features/office-tools-vs) 或第三方 Web 开发工具。

## <a name="client-requirements-os-x-desktop"></a>客户端要求：OS X 桌面

作为 Office 365 的一部分分发的 适用于 Mac 的 Outlook 支持 Outlook 外接程序。在 适用于 Mac 的 Outlook 上运行 Outlook 外接程序与 适用于 Mac 的 Outlook 本身的要求相同：操作系统必须至少为 OS X v10.10"Yosemite"。由于 适用于 Mac 的 Outlook 使用 WebKit 作为布局引擎以呈现外接程序页，因此没有其他浏览器依赖项。

以下是支持 Office 外接程序的 Office for Mac 的最低客户端版本：

- Word for Mac 版本 15.18 (160109) 
- Excel for Mac 版本 15.19 (160206) 
- PowerPoint for Mac 版本 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-online-web-clients-and-sharepoint"></a>客户端要求：针对 Office Online Web 客户端和 SharePoint 的浏览器支持

支持 ECMAScript 5.1、HTML5 和 CSS3 的任何浏览器，如 Internet Explorer 11 或更高版本、或 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版本。


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>客户端要求：非 Windows 智能手机和平板电脑

特别是对于在智能手机和非 Windows 平板电脑设备上的浏览器中运行的 适用于设备的 OWA 和 Outlook Web App，测试和运行 Outlook 外接程序需要下列软件。


| 主机应用程序 | 设备 | 操作系统 | Exchange 帐户 | 移动浏览器 |
|:-----|:-----|:-----|:-----|:-----|
|OWA for Android|Android 智能手机。从技术上讲， [Android OS](https://developer.android.com/guide/practices/screens_support.html) 将这些设备视为"小型"或"正常"。|Android 4.4 KitKat 或更高版本|有关 Office 365 企业版 或 Exchange Online 的最新更新|用于 Android 的本机加载项，浏览器不适用|
|OWA for iPad|iPad 2 或更高版本|iOS 6 或更高版本|有关 Office 365 企业版 或 Exchange Online 的最新更新|用于 iOS 的本机加载项，浏览器不适用|
|OWA for iPhone|iPhone 4S 或更高版本|iOS 6 或更高版本|有关 Office 365 企业版 或 Exchange Online 的最新更新|用于 iOS 的本机加载项，浏览器不适用|
|Outlook Web App|iPhone 4 或更高版本、iPad 2 或更高版本、iPod Touch 4 或更高版本|iOS 5 或更高版本|有关 Office 365、Exchange Online 或本地的 Exchange Server 2013 或更高版本|Safari|


## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)
