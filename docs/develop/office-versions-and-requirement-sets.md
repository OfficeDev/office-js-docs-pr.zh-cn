---
title: Office 版本和要求集
description: 使用 JavaScript API 支持的 Office.js 平台。
ms.date: 09/14/2022
ms.localizationpriority: high
ms.openlocfilehash: 669977f87974a1ec5519ddbbe3d38c5a290ec84f
ms.sourcegitcommit: cff5d3450f0c02814c1436f94cd1fc1537094051
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/30/2022
ms.locfileid: "68234905"
---
# <a name="office-versions-and-requirement-sets"></a>Office 版本和要求集

Office 跨多个平台运行且有许多版本，它们并非全都支持 Office JavaScript API (Office.js) 中的所有 API。 Windows 上的 Office 2013 是支持 Office 加载项的最早版本的 Office。你可能并不总是能够控制用户已安装的 Office 版本。 为了处理这种情况，我们提供了一个称为要求集的系统，以帮助你确定 Office 应用程序是否支持 Office 外接程序中所需的功能。

> [!NOTE]
>
> - Office 跨多个平台（包括 Windows、浏览器、Mac 和 iPad）运行。
> - Office 应用程序的示例包括 Office 产品：Excel、Word、PowerPoint、Outlook、OneNote 等。
> - Microsoft 365 订阅或永久许可证提供 Office。 永久版本可通过批量许可协议或零售版获得。
> - 要求集是 API 成员的命名组，例如 `ExcelApi 1.5`， `WordApi 1.3`等等。

## <a name="how-to-check-your-office-version"></a>如何检查 Office 版本

若要确定使用的 Office 版本，请在 Office 应用程序中，依次选择“文件”菜单和“帐户”。 Office 版本显示在 **“产品信息** ”部分。 例如，以下屏幕截图指示 Office 版本 1802 (内部版本 9026.1000) 。

![检查 Office 版本。](../images/office-version.png)

> [!NOTE]
> 如果你的 Office 版本与此版本不同，请参阅 [我拥有哪个版本的 Outlook？](https://support.microsoft.com/office/b3a9568c-edb5-42b9-9825-d48d82b2257c) 或 [关于 Office：我使用的是哪个版本的 Office？](https://support.microsoft.com/topic/932788b8-a3ce-44bf-bb09-e334518b8b19) 了解如何获取版本的此信息。

## <a name="office-requirement-sets-availability"></a>Office 要求集可用性

Office 加载项可以使用 API 要求集来确定 Office 应用程序是否支持它需要使用的 API 成员。 要求集支持因 Office 应用程序而异，Office 应用程序版本 (请参阅前面的部分 [“如何检查 Office 版本](#how-to-check-your-office-version)) 。

Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

此外，通用 API 中还添加了加载项命令（功能区扩展性）和对话框启动功能（对话框 API）等其他功能。 外接程序命令和对话框 API 要求集是各种 Office 应用程序共同共享的 API 集的示例。

An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles.

- [Excel JavaScript API 要求集](/javascript/api/requirement-sets/excel/excel-api-requirement-sets) (ExcelApi)
- [OneNote JavaScript API 要求集](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets) (OneNoteApi)
- [Outlook JavaScript API 要求集](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) (邮箱) 
- [PowerPoint JavaScript API 要求集](/javascript/api/requirement-sets/powerpoint/powerpoint-api-requirement-sets) (PowerPointApi)
- [Word JavaScript API 要求集](/javascript/api/requirement-sets/word/word-api-requirement-sets) (WordApi)

某些要求集包含可由多个 Office 应用程序使用的 API。 有关这些要求集的信息，请参阅以下文章。

- [Office 通用要求集](/javascript/api/requirement-sets/common/office-add-in-requirement-sets)
- [加载项命令要求集](/javascript/api/requirement-sets/common/add-in-commands-requirement-sets)
- [对话框 API 要求集](/javascript/api/requirement-sets/common/dialog-api-requirement-sets)
- [对话框源要求集](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets)
- [标识 API 要求集](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [图像强制要求集](/javascript/api/requirement-sets/common/image-coercion-requirement-sets)
- [键盘快捷方式要求集](/javascript/api/requirement-sets/common/keyboard-shortcuts-requirement-sets)
- [打开浏览器窗口要求集](/javascript/api/requirement-sets/common/open-browser-window-api-requirement-sets)
- [功能区 API 要求集](/javascript/api/requirement-sets/common/ribbon-api-requirement-sets)
- [共享运行时要求集](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

The Office JavaScript API library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).

## <a name="specify-office-applications-and-requirement-sets"></a>指定 Office 应用程序和要求集

There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>另请参阅

- [指定 Office 应用程序和 API 要求集](../develop/specify-office-hosts-and-api-requirements.md)
- [安装最新版 Office](../develop/install-latest-office-version.md)
- [Microsoft 365 应用版更新频道概述](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [利用 Microsoft 365 和 Microsoft Teams 重塑生产力](https://products.office.com/compare-all-microsoft-office-products?tab=2)
