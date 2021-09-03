---
title: 打开浏览器窗口要求集
description: 指定哪些Office和版本支持 openBrowserWindow API。
ms.date: 04/09/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8197228f1d428fd48c494825fec0e73cb85609f6
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868755"
---
# <a name="open-browser-window-api-requirement-sets"></a>打开浏览器窗口 API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

OpenBrowserWindow API 集使加载项能够打开浏览器，以完成无法在加载项本身的沙盒 Webview 控件中始终完成的任务;例如，在 webview 控件由 webview 控件提供时下载 PDF Microsoft Edge。

Office 外接程序在多个 Office 版本中运行。 下表列出了 OpenBrowserWindow API 要求集、Office 支持该要求集的主机应用程序，以及 Office 应用程序的版本或版本号。

|  要求集  | Office 2013 Windows或更高版本<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | 不适用 | 版本 1810 (内部版本 16.0.11001.20074) 或更高版本 | 16.0.0.0 或更高版本 | 16.0.0.0 或更高版本 | 不适用 | 不适用|

> [!NOTE]
> OpenBrowserWindowApi 要求集仅按如下方式提供：
>
> - Excel、PowerPoint、Word：Windows、Mac、iPad
> - Outlook：Windows、Mac

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [更新频道发布的版本号和内部版本号Microsoft 365 应用版](/officeupdates/update-history-microsoft365-apps-by-date)
- [使用的是哪一版 Office？](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到客户端应用程序的版本号Office版本号](/officeupdates/update-history-microsoft365-apps-by-date)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

OpenBrowserWindowApi 1.1 是 API 的第一个版本。 有关 API 的详细信息，请参阅[Office.context.ui](/javascript/api/office/office.context#ui)参考主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
