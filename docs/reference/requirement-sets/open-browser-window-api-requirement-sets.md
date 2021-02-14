---
title: 打开浏览器窗口要求集
description: 指定哪些 Office 平台和版本支持 openBrowserWindow API。
ms.date: 02/09/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 759c8265b27fab4589e68fe3f2f90a2a283ef005
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237957"
---
# <a name="open-browser-window-api-requirement-sets"></a>打开浏览器窗口 API 要求集

要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。

OpenBrowserWindow API 集使加载项能够打开浏览器，以完成无法在外接程序本身的沙盒 Web 视图控件中始终完成的任务;例如，在 Microsoft Edge 提供 Web 视图控件时下载 PDF 文件。

Office 外接程序在多个 Office 版本中运行。 下表列出了 OpenBrowserWindow API 要求集、支持该要求集的 Office 主机应用程序以及 Office 应用程序内部版本或版本号。

|  要求集  | Windows 版或更高版本上的 Office 2013<br>（一次性购买） | Windows 版 Office<br>（关联至 Microsoft 365 订阅） |  iPad 版 Office<br>（关联至 Microsoft 365 订阅）  |  Mac 版 Office<br>（关联至 Microsoft 365 订阅）  | Office 网页版  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | 不适用 | 版本 1810 (内部版本 16.0.11001.20074) 或更高版本 | 16.0.0.0 或更高版本 | 16.0.0.0 或更高版本 | 不适用 | 不适用|

若要详细了解版本、内部版本号和 Office Online Server，请参阅：

- [Microsoft 365 应用版的更新频道版本的版本号和内部版本号](/officeupdates/update-history-microsoft365-apps-by-date)
- [使用的是哪一版 Office？](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [在哪里可以找到 Office 客户端应用程序的版本号和内部版本号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概述](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 通用 API 要求集

若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

OpenBrowserWindowApi 1.1 是 API 的第一个版本。 有关 API 的详细信息，请参阅 [Office.context.ui](/javascript/api/office/office.context#ui) 参考主题。

## <a name="see-also"></a>另请参阅

- [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)
- [指定 Office 主机和 API 要求](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office 加载项 XML 清单](../../develop/add-in-manifests.md)
