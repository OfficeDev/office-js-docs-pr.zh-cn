---
title: Office 版本和要求集
description: ''
ms.date: 03/29/2018
ms.openlocfilehash: ac3ae4fa3eeca9cfbd56b15168fc39d67139680d
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505991"
---
# <a name="office-versions-and-requirement-sets"></a>Office 版本和要求集

在不同平台上有不同版本的 Office，它们没有全部支持 Office JavaScript API (Office.js) 中的 API 。对用户安装的 Office 版本无法做到完全管控。 这种情况下，我们提供名为要求集的系统，该系统可帮助判断 Office 主机是否支持 Office 加载项中所需的功能。 

> [!NOTE]
> - Office 跨多个平台运行，其中包括 Office for Windows、Office Online、Office for Mac 和 Office for iPad。  
> - Office 主机示例是Office 产品，其中包括 Excel、Word、PowerPoint、Outlook、OneNote 等产品。  
> - 要求集是由 API 成员命名的组，如 `ExcelApi 1.5`、`WordApi 1.3` 等。  


## <a name="how-to-check-your-office-version"></a>如何查看 Office 版本

使用 Office 应用查看正在使用的 Office 版本， 选择 **文件** 目录, 选 **帐户**。Office 版本会出现在 **产品信息** 区。比如, 下面的截图显示了 Office 版本 1802 (生成号 9026.1000):

![查看 Office 版本](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a>Office 要求集有效性

Office 加载项可以使用 API 要求集确定 Office 主机是否支持需要使用的 API 成员。要求集支持 Office 主机和 Office 主机版本的不同 （参阅上一节）。

某些 Office 主机有自己的 API 要求集。例如，第一个 Excel API 的要求集是 `ExcelApi 1.1` 第一个 Word API 的要求集是 `WordApi 1.1`。自此，添加了多个新 ExcelApi 要求集和 WordApi 要求集以提供额外的 API 功能。

此外，其他功能如加载项命令 （功能区扩展性）及弹出对话框 (对话框 API) 的功能已添加到公共 API。加载项命令和对话框 API 要求集是不同的 Office 主机共享的 API 集合。

加载项可以仅在加载项运行的 Office 主机版本支持的要求集中使用 API。若要了解特定 Office 主机版本有哪些要求集可用，参照如下特定主机要求集文章：

- [Excel JavaScript API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)
- [Word JavaScript API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)
- [OneNote JavaScript API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)
- [了解 Outlook API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)

某些要求集包含可由任意 Office 主机使用的 API。欲知这些要求集的信息，请参阅以下文章：

- [Office 通用要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [加载项命令要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets?view=office-js)
- [Dialog API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [标识 API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)

要求集的版本号, 比如 `ExcelApi 1.1` 中的“1.1”， 与 Office 主机有关。 已知的要求集版本号 (比如, `ExcelApi 1.1`) 与 Office.js 版本号或与其他 Office 主机的要求集（如， Word, Outlook, 等等）并不对应。不同 Office主机的要求集于不同速度和时间发布。如， `ExcelApi 1.5` 在 `WordApi 1.3` 要求集之前发布。

用于 Office 的 JavaScript API 库 (Office.js) 包括当前可用的所有要求集。当诸如要求集 `ExcelApi 1.3` 和 `WordApi 1.3`，不存在 `Office.js 1.3` 要求集。最近发布的 Office.js 作为单个 Office 终结点，并通过内容交付网络 (CDN) 传输。欲知 Office.js CDN 的更多详细信息，包括如何处理版本控制和向后兼容性，请参阅 [了解 Office 的 JavaScript API ](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。

## <a name="specify-office-hosts-and-requirement-sets"></a>指定 Office 主机和要求集

有多种方法来明确加载项要求哪些 Office 主机和要求集。 欲知详情，请参阅 [明确 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)


## <a name="see-also"></a>另请参阅

- [指定 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [安装最新版 Office](https://docs.microsoft.com/office/dev/add-ins/develop/install-latest-office-version)
- [Office 365 ProPlus 频道更新概述](https://docs.microsoft.com/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [通过 Office 365 充分使用 Office](https://products.office.com/compare-all-microsoft-office-products?tab=2)
