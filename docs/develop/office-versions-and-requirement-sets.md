---
title: Office 版本和要求集
description: 使用 JavaScript API 支持的 Office.js 平台。
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 2fd1393271d50be66dd2bbc2bb8cbb251ae6efbc
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237663"
---
# <a name="office-versions-and-requirement-sets"></a><span data-ttu-id="89229-103">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="89229-103">Office versions and requirement sets</span></span>

<span data-ttu-id="89229-p101">Office 跨多个平台运行且有许多版本，它们并非全都支持 Office JavaScript API (Office.js) 中的所有 API。不一定总能控制用户安装的 Office 版本。为了应对这种情况，我们提供了名为“要求集”的系统，以帮助确定 Office 应用程序是否支持 Office 加载项需要的功能。</span><span class="sxs-lookup"><span data-stu-id="89229-p101">There are many versions of Office on several platforms, and they don't all support every API in Office JavaScript API (Office.js). You may not always have control over the version of Office your users have installed.  To handle this situation, we provide a system called requirement sets to help you determine whether an Office application supports the capabilities you need in your Office Add-in.</span></span> 

> [!NOTE]
> - <span data-ttu-id="89229-107">Office 跨多个平台（包括 Windows、浏览器、Mac 和 iPad）运行。</span><span class="sxs-lookup"><span data-stu-id="89229-107">Office runs across multiple platforms, including Windows, in a browser, Mac, and iPad.</span></span>
> - <span data-ttu-id="89229-108">Office 应用程序示例包括 Excel、Word、PowerPoint、Outlook、OneNote 等 Office 产品。</span><span class="sxs-lookup"><span data-stu-id="89229-108">Examples of Office applications are Office Products: Excel, Word, PowerPoint, Outlook, OneNote, and so forth.</span></span>  
> - <span data-ttu-id="89229-109">要求集是 API 成员（如 `ExcelApi 1.5`、`WordApi 1.3` 等）的已命名组。</span><span class="sxs-lookup"><span data-stu-id="89229-109">A requirement set is a named group of API members e.g., `ExcelApi 1.5`, `WordApi 1.3`, and so on.</span></span>  

## <a name="how-to-check-your-office-version"></a><span data-ttu-id="89229-110">如何检查 Office 版本</span><span class="sxs-lookup"><span data-stu-id="89229-110">How to check your Office version</span></span>

<span data-ttu-id="89229-p102">若要确定使用的 Office 版本，请在 Office 应用程序中，依次选择“文件”\*\*\*\* 菜单和“帐户”\*\*\*\*。 Office 版本显示在“产品信息”\*\*\*\* 部分中。 例如，下面的屏幕截图指明 Office 版本 1802（生成号 9026.1000）：</span><span class="sxs-lookup"><span data-stu-id="89229-p102">To identify the Office version that you're using, from within an Office application, select the **File** menu, and then choose **Account**. The version of Office will appear in the **Product Information** section. For example, the following screenshot indicates Office Version 1802 (Build 9026.1000):</span></span>

![检查 Office 版本](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a><span data-ttu-id="89229-115">Office 要求集可用性</span><span class="sxs-lookup"><span data-stu-id="89229-115">Office requirement sets availability</span></span>

<span data-ttu-id="89229-p103">Office 加载项可使用 API 要求集，以确定 Office 应用程序是否支持需要使用的 API 成员。要求集支持因 Office 应用程序和 Office 应用程序版本而异（见上一部分）。</span><span class="sxs-lookup"><span data-stu-id="89229-p103">Office Add-ins can use API requirement sets to determine whether the Office application supports the API members that it need to use. Requirement set support varies by Office application and the Office application version (see previous section).</span></span>

<span data-ttu-id="89229-p104">一些 Office 应用程序有自己的 API 要求集。 例如，第一个 Excel API 要求集为 `ExcelApi 1.1`，第一个 Word API 要求集为 `WordApi 1.1`。从那以后，便新增了多个 ExcelApi 要求集和 WordApi 要求集，以提供其他 API 功能。</span><span class="sxs-lookup"><span data-stu-id="89229-p104">Some Office applications have their own API requirement sets. For example, the first requirement set for the Excel API was `ExcelApi 1.1` and the first requirement set for the Word API was `WordApi 1.1`. Since then, multiple new ExcelApi requirement sets and WordApi requirement sets have been added to provide additional API functionality.</span></span>

<span data-ttu-id="89229-121">此外，通用 API 中还添加了加载项命令（功能区扩展性）和对话框启动功能（对话框 API）等其他功能。</span><span class="sxs-lookup"><span data-stu-id="89229-121">In addition, other functionality such as add-in commands (ribbon extensibility) and the ability to launch dialog boxes (Dialog API) were added to the Common API.</span></span> <span data-ttu-id="89229-122">加载项命令和对话框 API 要求集是各种 Office 应用程序共用的 API 集示例。</span><span class="sxs-lookup"><span data-stu-id="89229-122">Add-in commands and Dialog API requirement sets are examples of API sets that the various Office applications share in common.</span></span>

<span data-ttu-id="89229-p106">加载项使用的要求集中的 API 只能是受运行加载项的 Office 应用程序版本支持的 API。若要确切了解适用于特定 Office 应用程序版本的要求集，请参阅以下特定于应用程序的要求集文章：</span><span class="sxs-lookup"><span data-stu-id="89229-p106">An add-in can only use APIs in requirement sets that are supported by the version of Office application where the add-in is running. To know exactly which requirement sets are available for a specific Office application version, refer to the following application-specific requirement set articles:</span></span>

- <span data-ttu-id="89229-125">[Excel JavaScript API 要求集](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)</span><span class="sxs-lookup"><span data-stu-id="89229-125">[Excel JavaScript API requirement sets](../reference/requirement-sets/excel-api-requirement-sets.md) (ExcelApi)</span></span>
- <span data-ttu-id="89229-126">[Word JavaScript API 要求集](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)</span><span class="sxs-lookup"><span data-stu-id="89229-126">[Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md) (WordApi)</span></span>
- <span data-ttu-id="89229-127">[OneNote JavaScript API 要求集](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)</span><span class="sxs-lookup"><span data-stu-id="89229-127">[OneNote JavaScript API requirement sets](../reference/requirement-sets/onenote-api-requirement-sets.md) (OneNoteApi)</span></span>
- <span data-ttu-id="89229-128">[PowerPoint JavaScript API 要求集](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)</span><span class="sxs-lookup"><span data-stu-id="89229-128">[PowerPoint JavaScript API requirement sets](../reference/requirement-sets/powerpoint-api-requirement-sets.md) (PowerPointApi)</span></span>
- <span data-ttu-id="89229-129">[了解 Outlook API 要求集](../reference/requirement-sets/outlook-api-requirement-sets.md) (MailBox)</span><span class="sxs-lookup"><span data-stu-id="89229-129">[Understanding Outlook API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md) (Mailbox)</span></span>

<span data-ttu-id="89229-p107">一些要求集包含任何 Office 应用程序都能使用的 API。若要了解这些要求集，请参阅以下文章：</span><span class="sxs-lookup"><span data-stu-id="89229-p107">Some requirement sets contain APIs that can be used by any Office application. For information about these requirement sets, refer to the following articles:</span></span>

- [<span data-ttu-id="89229-132">Office 通用要求集</span><span class="sxs-lookup"><span data-stu-id="89229-132">Office common requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="89229-133">加载项命令要求集</span><span class="sxs-lookup"><span data-stu-id="89229-133">Add-in commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="89229-134">对话框 API 要求集</span><span class="sxs-lookup"><span data-stu-id="89229-134">Dialog API requirement sets</span></span>](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [<span data-ttu-id="89229-135">标识 API 要求集</span><span class="sxs-lookup"><span data-stu-id="89229-135">Identity API requirement sets</span></span>](../reference/requirement-sets/identity-api-requirement-sets.md)

<span data-ttu-id="89229-p108">要求集的版本号（如 `ExcelApi 1.1` 中的“1.1”）是相对于 Office 应用程序而言。给定要求集的版本号（例如，`ExcelApi 1.1`）既不对应于 Office.js 的版本号，也不对应于其他 Office 应用程序（例如，Word、Outlook 等）的要求集。各个 Office 应用程序的要求集的发布速率不同。例如，`ExcelApi 1.5` 要求集先于 `WordApi 1.3` 要求集发布。</span><span class="sxs-lookup"><span data-stu-id="89229-p108">The version number of a requirement set, such as the "1.1" in `ExcelApi 1.1`, is relative to the Office application. The version number of a given requirement set (e.g., `ExcelApi 1.1`) does not correspond to the version number of Office.js or to requirement sets for other Office applications (e.g., Word, Outlook, etc.).  Requirement sets for the different Office applications are released at different rates. For example, `ExcelApi 1.5` was released before the `WordApi 1.3` requirement set.</span></span>


<span data-ttu-id="89229-140">Office JavaScript API 库 (Office.js) 包含当前可用的所有要求集。</span><span class="sxs-lookup"><span data-stu-id="89229-140">The Office JavaScript API library (Office.js) includes all requirement sets that are currently available.</span></span> <span data-ttu-id="89229-141">虽然有 `ExcelApi 1.3` 和 `WordApi 1.3` 等要求集，但并无 `Office.js 1.3` 要求集。</span><span class="sxs-lookup"><span data-stu-id="89229-141">While there is such a thing as requirement sets `ExcelApi 1.3` and `WordApi 1.3`, there is no `Office.js 1.3` requirement set.</span></span> <span data-ttu-id="89229-142">最新版 Office.js 作为一个通过内容传送网络 (CDN) 提供的 Office 终结点进行维护。</span><span class="sxs-lookup"><span data-stu-id="89229-142">The latest release of Office.js is maintained as a single Office endpoint delivered via the content delivery network (CDN).</span></span> <span data-ttu-id="89229-143">若要详细了解 Office.js CDN（包括如何处理版本控制和向后兼容性），请参阅[了解 Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md)。</span><span class="sxs-lookup"><span data-stu-id="89229-143">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Understanding the Office JavaScript API](../develop/understanding-the-javascript-api-for-office.md).</span></span>

## <a name="specify-office-applications-and-requirement-sets"></a><span data-ttu-id="89229-144">指定 Office 应用程序和要求集</span><span class="sxs-lookup"><span data-stu-id="89229-144">Specify Office applications and requirement sets</span></span>

<span data-ttu-id="89229-p110">可通过多种方法来指定加载项需要的 Office 应用程序和要求集。有关详细信息，请参阅[指定 Office 应用程序和 API 要求](../develop/specify-office-hosts-and-api-requirements.md)</span><span class="sxs-lookup"><span data-stu-id="89229-p110">There are various ways to specify which Office applications and requirement sets are required by an add-in.  For detailed information, see [Specify Office applications and API requirements](../develop/specify-office-hosts-and-api-requirements.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="89229-147">另请参阅</span><span class="sxs-lookup"><span data-stu-id="89229-147">See also</span></span>

- [<span data-ttu-id="89229-148">指定 Office 应用程序和 API 要求集</span><span class="sxs-lookup"><span data-stu-id="89229-148">Specify Office applications and API requirements</span></span>](../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="89229-149">安装最新版 Office</span><span class="sxs-lookup"><span data-stu-id="89229-149">Install the latest version of Office</span></span>](../develop/install-latest-office-version.md)
- [<span data-ttu-id="89229-150">Microsoft 365 应用版更新频道概述</span><span class="sxs-lookup"><span data-stu-id="89229-150">Overview of update channels for Microsoft 365 Apps</span></span>](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [<span data-ttu-id="89229-151">利用 Microsoft 365 和 Microsoft Teams 重塑生产力</span><span class="sxs-lookup"><span data-stu-id="89229-151">Reimagine productivity with Microsoft 365 and Microsoft Teams</span></span>](https://products.office.com/compare-all-microsoft-office-products?tab=2)
