---
title: 运行 Office 加载项的要求
description: 了解最终用户运行 Office 外接程序所需的客户端和服务器要求。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 3c4ed1116fa83de694f390cfc26fde9a01858460
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293007"
---
# <a name="requirements-for-running-office-add-ins"></a>运行 Office 加载项的要求

本文介绍了运行 Office 加载项的软件和设备要求。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

有关 Office 外接程序目前受支持的高级别视图，请参阅 office [客户端应用程序和 Office 外接程序的平台可用性](../overview/office-add-in-availability.md)。

## <a name="server-requirements"></a>服务器要求

为了能够安装和运行任何 Office 外接程序，您需要首先将 UI 的清单和网页文件以及外接程序的代码部署到服务器上的合适位置。

对于所有类型的外接程序（内容、Outlook 和任务窗格外接程序以及外接程序命令），你需要将你的外接程序的网页文件部署到 Web 服务器或 Web 托管服务，如 [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md)。

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> 在 Visual Studio 中开发和调试加载项时，Visual Studio 使用 IIS Express 在本地部署并运行加载项的网页文件，无需使用其他 Web 服务器。

对于内容和任务窗格加载项，在受支持的 Office 客户端应用程序（Excel、PowerPoint、Project 或 Word）中，您还需要 SharePoint 上的 [应用程序目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) 以上载加载项的 XML 清单文件，或者需要使用 [集中部署](../publish/centralized-deployment.md)来部署加载项。

若要测试和运行 Outlook 外接程序，用户的 Outlook 电子邮件帐户必须驻留在 Exchange 2013 或更高版本中，可通过 Microsoft 365、Exchange Online 或本地安装获得。 用户或管理员在该服务器上安装 Outlook 外接程序的清单文件。

> [!NOTE]
> Outlook 中的 POP 和 IMAP 电子邮件帐户不支持 Office 加载项。

## <a name="client-requirements-windows-desktop-and-tablet"></a>客户端要求：Windows 台式机和平板电脑

为基于 Windows 的台式机、笔记本电脑或平板电脑设备上运行的受支持 Office 桌面客户端或 Web 客户端开发 Office 外接程序，需要以下软件：


- 对于 Windows x86 和 x64 台式机与平板电脑（如 Surface Pro）：
    - 在 Windows 7 或更高版本上运行的 32 位或 64 位版本 Office 2013。
    - Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013 或更高版本的 Office 客户端，（如果您正在专门为这些 Office 桌面客户端测试或运行 Office 外接程序）。Office 桌面客户端可以在本地安装或通过即点即用安装在客户端计算机上。

  如果您拥有有效的 Microsoft 365 订阅，并且您没有对 Office 客户端的访问权限，则可以 [下载并安装最新版本的 office](https://support.office.com/article/download-and-install-or-reinstall-office-365-or-office-2019-on-a-pc-or-mac-4414eaaf-0478-48be-9c42-23adc4716658)。

- 必须安装 Internet Explorer 11 或 Microsoft Edge（由 Windows 和 Office 版本而定），但它们不能是默认浏览器。 为支持 Office 加载项，充当主机的 Office 客户端使用了 Internet Explorer 11 或 Microsoft Edge 所包含的浏览器组件。 有关更多详细信息，请参阅 [Office加载项使用的浏览器](browsers-used-by-office-web-add-ins.md)。

  > [!NOTE]
  > 必须关闭 Internet Explorer 的增强安全配置 (ESC) 才能使 Office Web 加载项正常工作。 如果在开发加载项时使用 Windows Server 计算机作为客户端，请注意 Windows Server 中会默认打开 ESC。

- 默认浏览器是下述软件之一：Internet Explorer 11，或者 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版。
- HTML 和 JavaScript 编辑器（如记事本）、[Visual Studio 和 Microsoft 开发人员工具](https://www.visualstudio.com/features/office-tools-vs) 或第三方 Web 开发工具。

## <a name="client-requirements-os-x-desktop"></a>客户端要求：OS X 桌面

Mac 上的 outlook （作为 Microsoft 365 的一部分分发）支持 Outlook 外接程序。在 Outlook for Mac 上运行 Outlook 外接程序与 Outlook 在 Mac 本身具有相同的要求：操作系统必须至少是 OS X v 10.10 "Yosemite"。 由于 Mac 版 Outlook 使用 WebKit 作为布局引擎以呈现加载项页，因此没有其他浏览器依赖项。

以下是支持 Office 加载项的 Mac 版 Office 的最低客户端版本。

- Word 版本 15.18 (160109)
- Excel 版本 15.19 (160206)
- PowerPoint 版本 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>客户端要求：针对 Office Web 客户端和 SharePoint 的浏览器支持

支持 ECMAScript 5.1、HTML5 和 CSS3（例如 Internet Explorer 11，或者 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版）的任意浏览器。


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>客户端要求：非 Windows 智能手机和平板电脑

尤其对于在智能手机或非 Windows 平板电脑设备上的浏览器中运行的 Outlook，需要以下软件才能测试和运行 Outlook 加载项。


| Office 应用程序 | 设备 | 操作系统 | Exchange 帐户 | 移动浏览器 |
|:-----|:-----|:-----|:-----|:-----|
|Android 版 Outlook|Android 平板电脑和智能手机|Android 4.4 KitKat 及更高版本|针对商业版或 Exchange Online 的 Microsoft 365 应用程序的最新更新|Android 版本机应用（不适用于浏览器）|
|iOS 版 Outlook|iPad 平板电脑，iPhone 智能手机|iOS 11 或更高版本|针对商业版或 Exchange Online 的 Microsoft 365 应用程序的最新更新|iOS 版本机应用（不适用于浏览器）|
|Outlook 网页版|iPhone 4 或更高版本、iPad 2 或更高版本、iPod Touch 4 或更高版本|iOS 5 或更高版本|在 Microsoft 365、Exchange Online 或 Exchange Server 2013 或更高版本上的本地部署|Safari|

> [!NOTE]
> Android 版本机应用 OWA、iPad 版 OWA 和 iPhone 版 OWA 现已[弃用](https://support.office.com/article/Microsoft-OWA-mobile-apps-are-being-retired-076ec122-4576-4900-bc26-937f84d25a4b)且之后无需这些软件即可测试 Outlook 加载项。


## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [Office 外接程序的 office 客户端应用程序和平台可用性](../overview/office-add-in-availability.md)
- [Office 加载项使用的浏览器](browsers-used-by-office-web-add-ins.md)
