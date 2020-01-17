---
title: Office 版本和要求集
description: ''
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: 3b9bcbef47cdc27ad77b6bdce2b5c69f25fa6d62
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217308"
---
# <a name="office-versions-and-requirement-sets"></a><span data-ttu-id="203b0-102">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="203b0-102">Office versions and requirement sets</span></span>

<span data-ttu-id="203b0-p101">Office 跨多个平台运行且有许多版本，它们并非全都支持 Office JavaScript API (Office.js) 中的所有 API。 不一定总能控制用户安装的 Office 版本。  为了应对这种情况，我们提供了名为“要求集”的系统，以帮助确定 Office 主机是否支持 Office 加载项需要的功能。</span><span class="sxs-lookup"><span data-stu-id="203b0-p101">There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office host supports the capabilities you need in your Office Add-in.</span></span> 

> [!NOTE]
> - <span data-ttu-id="203b0-106">Office 跨多个平台（包括 Windows、浏览器、Mac 和 iPad）运行。</span><span class="sxs-lookup"><span data-stu-id="203b0-106">Office runs across multiple platforms, including Windows, in a browser, Mac, and iPad.</span></span>
> - <span data-ttu-id="203b0-107">Office 主机示例包括 Excel、Word、PowerPoint、Outlook、OneNote 等 Office 产品。</span><span class="sxs-lookup"><span data-stu-id="203b0-107">Examples of Office hosts are Office Products: Excel, Word, PowerPoint, Outlook, OneNote, and so forth.</span></span>  
> - <span data-ttu-id="203b0-108">要求集是 API 成员（如 `ExcelApi 1.5`、`WordApi 1.3` 等）的已命名组。</span><span class="sxs-lookup"><span data-stu-id="203b0-108">A requirement set is a named group of API members e.g., `ExcelApi 1.5`, `WordApi 1.3`, and so on.</span></span>  

## <a name="how-to-check-your-office-version"></a><span data-ttu-id="203b0-109">如何检查 Office 版本</span><span class="sxs-lookup"><span data-stu-id="203b0-109">How to check your Office version</span></span>

<span data-ttu-id="203b0-p102">若要确定使用的 Office 版本，请在 Office 应用程序中，依次选择“文件”\*\*\*\* 菜单和“帐户”\*\*\*\*。 Office 版本显示在“产品信息”\*\*\*\* 部分中。 例如，下面的屏幕截图指明 Office 版本 1802（生成号 9026.1000）：</span><span class="sxs-lookup"><span data-stu-id="203b0-p102">To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):</span></span>

![检查 Office 版本](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a><span data-ttu-id="203b0-114">Office 要求集可用性</span><span class="sxs-lookup"><span data-stu-id="203b0-114">Office requirement sets availability</span></span>

<span data-ttu-id="203b0-p103">Office 加载项可使用 API 要求集，以确定 Office 主机是否支持需要使用的 API 成员。 要求集支持因 Office 主机和 Office 主机版本而异（见上一部分）。</span><span class="sxs-lookup"><span data-stu-id="203b0-p103">Office Add-ins can use API requirement sets to determine whether the Office host supports the API members that it need to use. Requirement set support varies by Office host and the Office host version (see previous section).</span></span>

<span data-ttu-id="203b0-p104">一些 Office 主机有自己的 API 要求集。 例如，第一个 Excel API 要求集为 `ExcelApi 1.1`，第一个 Word API 要求集为 `WordApi 1.1`。 从那以后，便新增了多个 ExcelApi 要求集和 WordApi 要求集，以提供其他 API 功能。</span><span class="sxs-lookup"><span data-stu-id="203b0-p104">Some Office hosts have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.</span></span>

<span data-ttu-id="203b0-120">此外，通用 API 中还添加了加载项命令（功能区扩展性）和对话框启动功能（对话框 API）等其他功能。</span><span class="sxs-lookup"><span data-stu-id="203b0-120">In addition, other functionality such as add-in commands (ribbon extensibility) and the ability to launch dialog boxes (Dialog API) were added to the Common API.</span></span> <span data-ttu-id="203b0-121">加载项命令和对话框 API 要求集是各种 Office 主机共用的 API 集示例。</span><span class="sxs-lookup"><span data-stu-id="203b0-121">Add-in commands and Dialog API requirement sets are examples of API sets that the various Office hosts share in common.</span></span>

<span data-ttu-id="203b0-p106">加载项使用的要求集中的 API 只能是受运行加载项的 Office 主机版本支持的 API。 若要确切了解适用于特定 Office 主机版本的要求集，请参阅以下主机专用要求集文章：</span><span class="sxs-lookup"><span data-stu-id="203b0-p106">An add-in can only use APIs in requirement sets that are supported by the version of Office host where the add-in is running. To know exactly which requirement sets are available for a specific Office host version, refer to the following host-specific requirement set articles:</span></span>

- <span data-ttu-id="203b0-124">[Excel JavaScript API 要求集](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)（ExcelApi、[CustomFunctionsRuntime](../excel/custom-functions-architecture.md)）</span><span class="sxs-lookup"><span data-stu-id="203b0-124">[Excel JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) (ExcelApi, [CustomFunctionsRuntime](../excel/custom-functions-architecture.md))</span></span>
- <span data-ttu-id="203b0-125">[Word JavaScript API 要求集](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets) (WordApi)</span><span class="sxs-lookup"><span data-stu-id="203b0-125">[Word JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets) (WordApi)</span></span>
- <span data-ttu-id="203b0-126">[OneNote JavaScript API 要求集](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)</span><span class="sxs-lookup"><span data-stu-id="203b0-126">[OneNote JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)</span></span>
- <span data-ttu-id="203b0-127">[PowerPoint JavaScript API 要求集](/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets) (PowerPointApi)</span><span class="sxs-lookup"><span data-stu-id="203b0-127">[PowerPoint JavaScript API requirement sets](/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets) (PowerPointApi)</span></span>
- <span data-ttu-id="203b0-128">[了解 Outlook API 要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) (MailBox)</span><span class="sxs-lookup"><span data-stu-id="203b0-128">[Understanding Outlook API requirement sets](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) (Mailbox)</span></span>

<span data-ttu-id="203b0-p107">一些要求集包含任何 Office 主机都能使用的 API。 若要了解这些要求集，请参阅以下文章：</span><span class="sxs-lookup"><span data-stu-id="203b0-p107">Some requirement sets contain APIs that can be used by any Office host. For information about these requirement sets, refer to the following articles:</span></span>

- [<span data-ttu-id="203b0-131">Office 通用要求集</span><span class="sxs-lookup"><span data-stu-id="203b0-131">Office common requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="203b0-132">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="203b0-132">Add-in commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="203b0-133">对话框 API 要求集</span><span class="sxs-lookup"><span data-stu-id="203b0-133">Dialog API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [<span data-ttu-id="203b0-134">标识 API 要求集</span><span class="sxs-lookup"><span data-stu-id="203b0-134">Identity API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)

<span data-ttu-id="203b0-p108">要求集的版本号（如 `ExcelApi 1.1` 中的“1.1”）是相对于 Office 主机而言。 给定要求集的版本号（例如，`ExcelApi 1.1`）既不对应于 Office.js 的版本号，也不对应于其他 Office 主机（例如，Word、Outlook 等）的要求集。  各个 Office 主机的要求集的发布速度和时间不同。 例如，`ExcelApi 1.5` 要求集先于 `WordApi 1.3` 要求集发布。</span><span class="sxs-lookup"><span data-stu-id="203b0-p108">The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office host. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office hosts (e.g., Word, Outlook, etc.).  Requirement sets for the different Office hosts are released at different speeds and times. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.</span></span>

<span data-ttu-id="203b0-139">适用于 Office 的 JavaScript API 库 (Office.js) 包含当前可用的所有要求集。</span><span class="sxs-lookup"><span data-stu-id="203b0-139">The JavaScript API for Office library (Office.js) includes all requirement sets that are currently available.</span></span> <span data-ttu-id="203b0-140">虽然有 `ExcelApi 1.3` 和 `WordApi 1.3` 等要求集，但并无 `Office.js 1.3` 要求集。</span><span class="sxs-lookup"><span data-stu-id="203b0-140">While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set.</span></span> <span data-ttu-id="203b0-141">最新版 Office.js 作为一个通过内容传送网络 (CDN) 提供的 Office 终结点进行维护。</span><span class="sxs-lookup"><span data-stu-id="203b0-141">The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN).</span></span> <span data-ttu-id="203b0-142">若要详细了解 Office.js CDN（包括如何处理版本控制和向后兼容性），请参阅[了解适用于 Office 的 JavaScript API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)。</span><span class="sxs-lookup"><span data-stu-id="203b0-142">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the JavaScript API for Office](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

## <a name="specify-office-hosts-and-requirement-sets"></a><span data-ttu-id="203b0-143">指定 Office 主机和要求集</span><span class="sxs-lookup"><span data-stu-id="203b0-143">Specify Office hosts and requirement sets</span></span>

<span data-ttu-id="203b0-p110">可通过多种方法来指定加载项需要的 Office 主机和要求集。  有关详细信息，请参阅[指定 Office 主机和 API 要求](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span><span class="sxs-lookup"><span data-stu-id="203b0-p110">There are various ways to specify which Office hosts and requirement sets are required by an add-in.  For detailed information, see [Specify Office hosts and API requirements](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)</span></span>

## <a name="see-also"></a><span data-ttu-id="203b0-146">另请参阅</span><span class="sxs-lookup"><span data-stu-id="203b0-146">See also</span></span>

- [<span data-ttu-id="203b0-147">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="203b0-147">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="203b0-148">安装最新版 Office</span><span class="sxs-lookup"><span data-stu-id="203b0-148">Install the latest version of Office</span></span>](/office/dev/add-ins/develop/install-latest-office-version)
- [<span data-ttu-id="203b0-149">Office 365 专业增强版的更新通道概述</span><span class="sxs-lookup"><span data-stu-id="203b0-149">Overview of update channels for Office 365 ProPlus</span></span>](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [<span data-ttu-id="203b0-150">通过 Office 365 充分利用 Office</span><span class="sxs-lookup"><span data-stu-id="203b0-150">Get the most from Office with Office 365</span></span>](https://products.office.com/compare-all-microsoft-office-products?tab=2)
