---
title: 运行 Office 加载项的要求
description: 了解最终用户需要运行Office外接程序的客户端和服务器要求。
ms.date: 04/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9bc093b3e04dd1a67ba63bebbe2e44acf5137a07
ms.sourcegitcommit: 9795f671cacaa0a9b03431ecdfff996f690e30ed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/20/2022
ms.locfileid: "64963489"
---
# <a name="requirements-for-running-office-add-ins"></a>运行 Office 加载项的要求

本文介绍了运行 Office 加载项的软件和设备要求。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

有关当前支持Office外接程序的高级别视图，请[参阅Office外接程序Office客户端应用程序和平台可用性](/javascript/api/requirement-sets)。

## <a name="server-requirements"></a>服务器要求

为了能够安装和运行任何 Office 外接程序，您需要首先将 UI 的清单和网页文件以及外接程序的代码部署到服务器上的合适位置。

对于所有类型的外接程序（内容、Outlook 和任务窗格外接程序以及外接程序命令），你需要将你的外接程序的网页文件部署到 Web 服务器或 Web 托管服务，如 [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md)。

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> 在 Visual Studio 中开发和调试加载项时，Visual Studio 使用 IIS Express 在本地部署并运行加载项的网页文件，无需使用其他 Web 服务器。

对于内容和任务窗格外接程序，在受支持的Office客户端应用程序（Excel、PowerPoint、Project或 Word）中，还需要SharePoint上的[应用目录](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)来上传外接程序的 XML 清单文件，或者需要使用[集成应用](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)部署外接程序。

若要测试和运行Outlook外接程序，用户的Outlook电子邮件帐户必须驻留在 2013 或更高版本Exchange，可通过Microsoft 365、Exchange Online或本地安装使用。 用户或管理员在该服务器上安装 Outlook 外接程序的清单文件。

> [!NOTE]
> Outlook 中的 POP 和 IMAP 电子邮件帐户不支持 Office 加载项。

## <a name="client-requirements-windows-desktop-and-tablet"></a>客户端要求：Windows 台式机和平板电脑

为在基于Windows桌面、笔记本电脑或平板电脑设备上运行的支持Office桌面客户端或 Web 客户端开发Office外接程序时，需要以下软件。

- 对于 Windows x86 和 x64 台式机与平板电脑（如 Surface Pro）：
  - 在 Windows 7 或更高版本上运行的 32 位或 64 位版本 Office 2013。
  - Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013 或更高版本的 Office 客户端，（如果您正在专门为这些 Office 桌面客户端测试或运行 Office 外接程序）。Office 桌面客户端可以在本地安装或通过即点即用安装在客户端计算机上。

  如果有有效的Microsoft 365订阅，但无法访问Office客户端，则可以[下载并安装最新版本的Office](https://support.microsoft.com/office/4414eaaf-0478-48be-9c42-23adc4716658)。

- 必须安装 Internet Explorer 11 或 Microsoft Edge（由 Windows 和 Office 版本而定），但它们不能是默认浏览器。 为支持 Office 加载项，充当主机的 Office 客户端使用了 Internet Explorer 11 或 Microsoft Edge 所包含的浏览器组件。 有关更多详细信息，请参阅 [Office加载项使用的浏览器](browsers-used-by-office-web-add-ins.md)。

  > [!NOTE]
  > 必须关闭 Internet Explorer 的增强安全配置 (ESC) 才能使 Office Web 加载项正常工作。 如果在开发加载项时使用 Windows Server 计算机作为客户端，请注意 Windows Server 中会默认打开 ESC。

- 默认浏览器是下述软件之一：Internet Explorer 11，或者 Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) 的最新版。
- HTML 和 JavaScript 编辑器（如记事本）、[Visual Studio 和 Microsoft 开发人员工具](https://www.visualstudio.com/features/office-tools-vs) 或第三方 Web 开发工具。

## <a name="client-requirements-os-x-desktop"></a>客户端要求：OS X 桌面

Mac 上的Outlook作为Microsoft 365的一部分分发，它支持Outlook加载项。在 Mac 上Outlook中运行Outlook外接程序的要求与 Mac 本身Outlook的要求相同：操作系统必须至少为 OS X v10.10“Yosemite”。 由于 Mac 版 Outlook 使用 WebKit 作为布局引擎以呈现加载项页，因此没有其他浏览器依赖项。

以下是支持 Office 加载项的 Mac 版 Office 的最低客户端版本。

- Word 版本 15.18 (160109)
- Excel 版本 15.19 (160206)
- PowerPoint 版本 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>客户端要求：针对 Office Web 客户端和 SharePoint 的浏览器支持

支持 ECMAScript 5.1、HTML5 和 CSS3 的任何浏览器（例如Microsoft Edge、Chrome、Firefox 或 Safari (Mac OS) ）。

## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>客户端要求：非Windows智能手机和平板电脑

专用于在智能手机和非Windows平板电脑设备上运行的Outlook，测试和运行Outlook外接程序需要以下软件。

| Office 应用程序 | 设备 | 操作系统 | Exchange 帐户 | 移动浏览器 |
|:-----|:-----|:-----|:-----|:-----|
|Android 版 Outlook|- Android 平板电脑<br>- Android 智能手机|- Android 4.4 KitKat 或更高版本|最新更新Microsoft 365 商业应用版或Exchange Online|浏览器不适用。 使用适用于 Android 的本机应用。<sup>1</sup>|
|iOS 版 Outlook|- iPad平板电脑<br>- iPhone智能手机|- iOS 11 或更高版本|最新更新Microsoft 365 商业应用版或Exchange Online|浏览器不适用。 使用适用于 iOS 的本机应用。<sup>1</sup>|
|Outlook 网页版 (现代) <sup>2</sup>|- iPad 2 或更高版本<br>- Android 平板电脑 |- iOS 5 或更高版本<br>- Android 4.4 KitKat 或更高版本|Microsoft 365，Exchange Online|- Microsoft Edge<br>- Chrome<br>- Firefox<br>- Safari|
|Outlook 网页版（经典）|- iPhone 4 或更高版本<br>- iPad 2 或更高版本<br>- iPod Touch 4 或更高版本|- iOS 5 或更高版本|本地Exchange Server 2013 或更高版本 <sup>3</sup>|- Safari|

> [!NOTE]
> <sup>1</sup> 个 OWA for Android、OWA for iPad 和 OWA for iPhone 本机应用已[弃用](https://support.microsoft.com/office/076ec122-4576-4900-bc26-937f84d25a4b)。
>
> <sup>2</sup> iPhone和 Android 智能手机上的新式Outlook 网页版不再需要或可用于测试Outlook加载项。
>
> <sup>3</sup> 个加载项在 Android 上的 Outlook、iOS 和具有本地Exchange帐户的新式移动 Web 中不受支持。

> [!TIP]
> 可通过查看邮箱工具栏，在 Web 浏览器中区分经典和新式 Outlook。
>
> **新式**
>
> ![新式 Outlook 工具栏的部分屏幕截图。](../images/outlook-on-the-web-new-toolbar.png)
>
> **经典**
>
> ![经典 Outlook 工具栏的部分屏幕截图。](../images/outlook-on-the-web-classic-toolbar.png)

## <a name="see-also"></a>另请参阅

- [Office 加载项平台概述](../overview/office-add-ins.md)
- [Office 客户端应用程序和平台的 Office 加载项可用性](/javascript/api/requirement-sets)
- [Office 加载项使用的浏览器](browsers-used-by-office-web-add-ins.md)
