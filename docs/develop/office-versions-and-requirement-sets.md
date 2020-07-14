---
title: Office 版本和要求集
description: 使用 JavaScript API 支持的 Office.js 平台
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 02f3d91256ea05e526ebe2e4e4090b1908d7292a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093579"
---
# <a name="office-versions-and-requirement-sets"></a>Office 版本和要求集

There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in. 

> [!NOTE]
> - Office 跨多个平台（包括 Windows、浏览器、Mac 和 iPad）运行。
> - Office 主机示例包括 Excel、Word、PowerPoint、Outlook、OneNote 等 Office 产品。  
> - 要求集是 API 成员（如 `ExcelApi 1.5`、`WordApi 1.3` 等）的已命名组。  

## <a name="how-to-check-your-office-version"></a>如何检查 Office 版本

To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):

![检查 Office 版本](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Office 要求集可用性

Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).

Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.

此外，通用 API 中还添加了加载项命令（功能区扩展性）和对话框启动功能（对话框 API）等其他功能。 加载项命令和对话框 API 要求集是各种 Office 主机共用的 API 集示例。

An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:

- [Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)
- [Word JavaScript API 要求集](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)
- [OneNote JavaScript API 要求集](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)
- [PowerPoint JavaScript API 要求集](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)
- [了解 Outlook API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md) (MailBox)

Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:

- [Office 通用要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [加载项命令要求集](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [对话框 API 要求集](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [标识 API 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)

The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.

Office JavaScript API 库 (Office.js) 包含当前可用的所有要求集。 虽然有 `ExcelApi 1.3` 和 `WordApi 1.3` 等要求集，但并无 `Office.js 1.3` 要求集。 最新版 Office.js 作为一个通过内容传送网络 (CDN) 提供的 Office 终结点进行维护。 若要详细了解 Office.js CDN（包括如何处理版本控制和向后兼容性），请参阅[了解 Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md)。

## <a name="specify-office-hosts-and-requirement-sets"></a>指定 Office 主机和要求集

There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](../develop/specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>另请参阅

- [指定 Office 主机和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)
- [安装最新版 Office](../develop/install-latest-office-version.md)
- [Microsoft 365 应用版更新频道概述](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [通过 Office 365 充分利用 Office](https://products.office.com/compare-all-microsoft-office-products?tab=2)
