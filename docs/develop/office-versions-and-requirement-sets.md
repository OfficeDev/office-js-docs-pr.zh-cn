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
# <a name="office-versions-and-requirement-sets"></a><span data-ttu-id="02fa2-102">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="02fa2-102">Office versions and requirement sets</span></span>

<span data-ttu-id="02fa2-p101">在不同平台上有不同版本的 Office，它们没有全部支持 Office JavaScript API (Office.js) 中的 API 。对用户安装的 Office 版本无法做到完全管控。 这种情况下，我们提供名为要求集的系统，该系统可帮助判断 Office 主机是否支持 Office 加载项中所需的功能。</span><span class="sxs-lookup"><span data-stu-id="02fa2-p101">There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in.</span></span> 

> [!NOTE]
> - <span data-ttu-id="02fa2-106">Office 跨多个平台运行，其中包括 Office for Windows、Office Online、Office for Mac 和 Office for iPad。</span><span class="sxs-lookup"><span data-stu-id="02fa2-106">Office runs across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>  
> - <span data-ttu-id="02fa2-107">Office 主机示例是Office 产品，其中包括 Excel、Word、PowerPoint、Outlook、OneNote 等产品。</span><span class="sxs-lookup"><span data-stu-id="02fa2-107">Examples of Office hosts are Office Products: Excel, Word, PowerPoint, Outlook, OneNote, and so forth.</span></span>  
> - <span data-ttu-id="02fa2-108">要求集是由 API 成员命名的组，如 `ExcelApi 1.5`、`WordApi 1.3` 等。</span><span class="sxs-lookup"><span data-stu-id="02fa2-108">A requirement set is a named group of API members e.g., `ExcelApi 1.5`, `WordApi 1.3`, and so on.</span></span>  


## <a name="how-to-check-your-office-version"></a><span data-ttu-id="02fa2-109">如何查看 Office 版本</span><span class="sxs-lookup"><span data-stu-id="02fa2-109">How to check your Office version</span></span>

<span data-ttu-id="02fa2-p102">使用 Office 应用查看正在使用的 Office 版本， 选择 **文件** 目录, 选 **帐户**。Office 版本会出现在 **产品信息** 区。比如, 下面的截图显示了 Office 版本 1802 (生成号 9026.1000):</span><span class="sxs-lookup"><span data-stu-id="02fa2-p102">To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):</span></span>

![查看 Office 版本](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a><span data-ttu-id="02fa2-114">Office 要求集有效性</span><span class="sxs-lookup"><span data-stu-id="02fa2-114">Office requirement sets availability</span></span>

<span data-ttu-id="02fa2-p103">Office 加载项可以使用 API 要求集确定 Office 主机是否支持需要使用的 API 成员。要求集支持 Office 主机和 Office 主机版本的不同 （参阅上一节）。</span><span class="sxs-lookup"><span data-stu-id="02fa2-p103">Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).</span></span>

<span data-ttu-id="02fa2-p104">某些 Office 主机有自己的 API 要求集。例如，第一个 Excel API 的要求集是 `ExcelApi 1.1` 第一个 Word API 的要求集是 `WordApi 1.1`。自此，添加了多个新 ExcelApi 要求集和 WordApi 要求集以提供额外的 API 功能。</span><span class="sxs-lookup"><span data-stu-id="02fa2-p104">Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.</span></span>

<span data-ttu-id="02fa2-p105">此外，其他功能如加载项命令 （功能区扩展性）及弹出对话框 (对话框 API) 的功能已添加到公共 API。加载项命令和对话框 API 要求集是不同的 Office 主机共享的 API 集合。</span><span class="sxs-lookup"><span data-stu-id="02fa2-p105">In addition, other functionality such as add-in commands (ribbon extensibility) and the ability to launch dialog boxes (Dialog API) were added to the common API. Add-in commands and Dialog API requirement sets are examples of API sets that the various Office hosts share in common.</span></span>

<span data-ttu-id="02fa2-p106">加载项可以仅在加载项运行的 Office 主机版本支持的要求集中使用 API。若要了解特定 Office 主机版本有哪些要求集可用，参照如下特定主机要求集文章：</span><span class="sxs-lookup"><span data-stu-id="02fa2-p106">An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:</span></span>

- <span data-ttu-id="02fa2-124">[Excel JavaScript API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)</span><span class="sxs-lookup"><span data-stu-id="02fa2-124">[Excel JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)</span></span>
- <span data-ttu-id="02fa2-125">[Word JavaScript API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)</span><span class="sxs-lookup"><span data-stu-id="02fa2-125">[Word JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)</span></span>
- <span data-ttu-id="02fa2-126">[OneNote JavaScript API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)</span><span class="sxs-lookup"><span data-stu-id="02fa2-126">[OneNote JavaScript API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)</span></span>
- <span data-ttu-id="02fa2-127">[了解 Outlook API 要求集](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)</span><span class="sxs-lookup"><span data-stu-id="02fa2-127">[Understanding Outlook API requirement sets](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)</span></span>

<span data-ttu-id="02fa2-p107">某些要求集包含可由任意 Office 主机使用的 API。欲知这些要求集的信息，请参阅以下文章：</span><span class="sxs-lookup"><span data-stu-id="02fa2-p107">Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:</span></span>

- [<span data-ttu-id="02fa2-130">Office 通用要求集</span><span class="sxs-lookup"><span data-stu-id="02fa2-130">Office common requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [<span data-ttu-id="02fa2-131">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="02fa2-131">Add-in commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets?view=office-js)
- [<span data-ttu-id="02fa2-132">Dialog API 要求集</span><span class="sxs-lookup"><span data-stu-id="02fa2-132">Dialog API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [<span data-ttu-id="02fa2-133">标识 API 要求集</span><span class="sxs-lookup"><span data-stu-id="02fa2-133">Identity API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)

<span data-ttu-id="02fa2-p108">要求集的版本号, 比如 `ExcelApi 1.1` 中的“1.1”， 与 Office 主机有关。 已知的要求集版本号 (比如, `ExcelApi 1.1`) 与 Office.js 版本号或与其他 Office 主机的要求集（如， Word, Outlook, 等等）并不对应。不同 Office主机的要求集于不同速度和时间发布。如， `ExcelApi 1.5` 在 `WordApi 1.3` 要求集之前发布。</span><span class="sxs-lookup"><span data-stu-id="02fa2-p108">The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.</span></span>

<span data-ttu-id="02fa2-p109">用于 Office 的 JavaScript API 库 (Office.js) 包括当前可用的所有要求集。当诸如要求集 `ExcelApi 1.3` 和 `WordApi 1.3`，不存在 `Office.js 1.3` 要求集。最近发布的 Office.js 作为单个 Office 终结点，并通过内容交付网络 (CDN) 传输。欲知 Office.js CDN 的更多详细信息，包括如何处理版本控制和向后兼容性，请参阅 [了解 Office 的 JavaScript API ](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="02fa2-p109">The JavaScript API for Office library (Office.js) includes all requirement sets that are currently available. While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set. The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN). For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the JavaScript API for Office](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

## <a name="specify-office-hosts-and-requirement-sets"></a><span data-ttu-id="02fa2-142">指定 Office 主机和要求集</span><span class="sxs-lookup"><span data-stu-id="02fa2-142">Specify Office hosts and requirement sets</span></span>

<span data-ttu-id="02fa2-p110">有多种方法来明确加载项要求哪些 Office 主机和要求集。 欲知详情，请参阅 [明确 Office 主机和 API 要求](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span><span class="sxs-lookup"><span data-stu-id="02fa2-p110">There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span></span>


## <a name="see-also"></a><span data-ttu-id="02fa2-145">另请参阅</span><span class="sxs-lookup"><span data-stu-id="02fa2-145">See also</span></span>

- [<span data-ttu-id="02fa2-146">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="02fa2-146">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="02fa2-147">安装最新版 Office</span><span class="sxs-lookup"><span data-stu-id="02fa2-147">Install the latest version of Office</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/install-latest-office-version)
- [<span data-ttu-id="02fa2-148">Office 365 ProPlus 频道更新概述</span><span class="sxs-lookup"><span data-stu-id="02fa2-148">Overview of update channels for Office 365 ProPlus</span></span>](https://docs.microsoft.com/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [<span data-ttu-id="02fa2-149">通过 Office 365 充分使用 Office</span><span class="sxs-lookup"><span data-stu-id="02fa2-149">Get the most from Office with Office 365</span></span>](https://products.office.com/compare-all-microsoft-office-products?tab=2)
