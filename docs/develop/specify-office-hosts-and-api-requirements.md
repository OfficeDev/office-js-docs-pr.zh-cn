---
title: 指定 Office 主机和 API 要求
description: 了解如何指定加载项的 Office 应用程序和 API 要求以按预期方式工作。
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 90ee7c3a5ad01252336608c02f995bbcbbe94212
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292627"
---
# <a name="specify-office-applications-and-api-requirements"></a><span data-ttu-id="d0055-103">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="d0055-103">Specify Office applications and API requirements</span></span>

<span data-ttu-id="d0055-104">您的 Office 外接程序可能依赖于特定的 Office 应用程序、要求集、API 成员或 API 版本，以便按预期工作。</span><span class="sxs-lookup"><span data-stu-id="d0055-104">Your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API in order to work as expected.</span></span> <span data-ttu-id="d0055-105">例如，你的外接程序可能：</span><span class="sxs-lookup"><span data-stu-id="d0055-105">For example, your add-in might:</span></span>

- <span data-ttu-id="d0055-106">在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。</span><span class="sxs-lookup"><span data-stu-id="d0055-106">Run in a single Office application (e.g., Word or Excel), or several applications.</span></span>

- <span data-ttu-id="d0055-p102">使用仅在 Office 的某些版本中可用的 JavaScript API。例如，可能会在运行在 Excel 2016 中的外接程序中使用适用于 Excel 的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="d0055-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="d0055-109">只能在 Office 的某些版本中运行，这些版本支持供你的外接程序使用的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="d0055-109">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="d0055-110">本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。</span><span class="sxs-lookup"><span data-stu-id="d0055-110">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="d0055-111">有关当前支持 Office 外接程序的高级别视图，请参阅 office [客户端应用程序和 Office 外接程序的平台可用性](../overview/office-add-in-availability.md) 页面。</span><span class="sxs-lookup"><span data-stu-id="d0055-111">For a high-level view of where Office Add-ins are currently supported, see the [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="d0055-112">下表列出了本文中讨论的核心概念。</span><span class="sxs-lookup"><span data-stu-id="d0055-112">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="d0055-113">**概念**</span><span class="sxs-lookup"><span data-stu-id="d0055-113">**Concept**</span></span>|<span data-ttu-id="d0055-114">**说明**</span><span class="sxs-lookup"><span data-stu-id="d0055-114">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="d0055-115">Office 应用程序、Office 客户端应用程序</span><span class="sxs-lookup"><span data-stu-id="d0055-115">Office application, Office client application</span></span>|<span data-ttu-id="d0055-p103">用于运行加载项的 Office 应用程序。例如 Word、Excel 等。</span><span class="sxs-lookup"><span data-stu-id="d0055-p103">The Office application used to run your add-in. For example, Word, Excel, and so on.</span></span>|
|<span data-ttu-id="d0055-118">平台</span><span class="sxs-lookup"><span data-stu-id="d0055-118">Platform</span></span>|<span data-ttu-id="d0055-119">Office 应用程序的运行位置，例如在浏览器中或在 iPad 上。</span><span class="sxs-lookup"><span data-stu-id="d0055-119">Where the Office application runs, such as in a browser or on an iPad.</span></span>|
|<span data-ttu-id="d0055-120">要求集</span><span class="sxs-lookup"><span data-stu-id="d0055-120">Requirement set</span></span>|<span data-ttu-id="d0055-121">命名的一组相关的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="d0055-121">A named group of related API members.</span></span> <span data-ttu-id="d0055-122">外接程序使用要求集来确定 Office 应用程序是否支持您的外接程序使用的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="d0055-122">Add-ins use requirement sets to determine whether the Office application supports API members used by your add-in.</span></span> <span data-ttu-id="d0055-123">测试对要求集的支持比对单个的 API 成员的支持更为容易。</span><span class="sxs-lookup"><span data-stu-id="d0055-123">It's easier to test for the support of a requirement set than for the support of individual API members.</span></span> <span data-ttu-id="d0055-124">要求集支持因 Office 应用程序和 Office 应用程序的版本而异。</span><span class="sxs-lookup"><span data-stu-id="d0055-124">Requirement set support varies by Office application and the version of the Office application.</span></span> <br ><span data-ttu-id="d0055-125">要求集在清单文件中指定。</span><span class="sxs-lookup"><span data-stu-id="d0055-125">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="d0055-126">当您在清单中指定要求集时，您可以设置 Office 应用程序必须提供的最低级别的 API 支持，以便运行您的外接程序。</span><span class="sxs-lookup"><span data-stu-id="d0055-126">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office application must provide in order to run your add-in.</span></span> <span data-ttu-id="d0055-127">不支持清单中指定的要求集的 Office 应用程序无法运行加载项，并且外接程序不会显示在 <span class="ui">我的外接</span>程序中。这将限制外接程序的可用位置。</span><span class="sxs-lookup"><span data-stu-id="d0055-127">Office applications that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.</span></span> <span data-ttu-id="d0055-128">在使用运行时检查的代码中。</span><span class="sxs-lookup"><span data-stu-id="d0055-128">In code using runtime checks.</span></span> <span data-ttu-id="d0055-129">有关要求集的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="d0055-129">For the complete list of requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>|
|<span data-ttu-id="d0055-130">运行时检查</span><span class="sxs-lookup"><span data-stu-id="d0055-130">Runtime check</span></span>|<span data-ttu-id="d0055-131">在运行时执行的测试，用于确定运行外接程序的 Office 应用程序是否支持您的外接程序使用的要求集或方法。</span><span class="sxs-lookup"><span data-stu-id="d0055-131">A test that is performed at runtime to determine whether the Office application running your add-in supports requirement sets or methods used by your add-in.</span></span> <span data-ttu-id="d0055-132">若要执行运行时检查，请将 **if** 语句与 `isSetSupported` 方法、要求集或不属于要求集的方法名称一起使用。</span><span class="sxs-lookup"><span data-stu-id="d0055-132">To perform a runtime check, you use an **if** statement with the `isSetSupported` method, the requirement sets, or the method names that aren't part of a requirement set.</span></span> <span data-ttu-id="d0055-133">使用运行时检查可确保加载项能够覆盖最大数量的客户。</span><span class="sxs-lookup"><span data-stu-id="d0055-133">Use runtime checks to ensure that your add-in reaches the broadest number of customers.</span></span> <span data-ttu-id="d0055-134">与要求集不同，运行时检查不会指定 Office 应用程序为运行外接程序必须提供的最低级别的 API 支持。</span><span class="sxs-lookup"><span data-stu-id="d0055-134">Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office application must provide for your add-in to run.</span></span> <span data-ttu-id="d0055-135">而是使用 **if** 语句来确定是否支持 API 成员。</span><span class="sxs-lookup"><span data-stu-id="d0055-135">Instead, you use the **if** statement to determine whether an API member is supported.</span></span> <span data-ttu-id="d0055-136">如果支持，则可以在外接程序中提供其他功能。</span><span class="sxs-lookup"><span data-stu-id="d0055-136">If it is, you can provide additional functionality in your add-in.</span></span> <span data-ttu-id="d0055-137">使用运行时检查时，你的外接程序将始终在“**我的外接程序**”中显示。</span><span class="sxs-lookup"><span data-stu-id="d0055-137">Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="d0055-138">开始之前</span><span class="sxs-lookup"><span data-stu-id="d0055-138">Before you begin</span></span>

<span data-ttu-id="d0055-139">您的外接程序必须使用最新版本的外接程序清单架构。</span><span class="sxs-lookup"><span data-stu-id="d0055-139">Your add-in must use the most current version of the add-in manifest schema.</span></span> <span data-ttu-id="d0055-140">如果您在外接程序中使用运行时检查，请确保使用最新的 Office JavaScript API ( # A0) 库。</span><span class="sxs-lookup"><span data-stu-id="d0055-140">If you use runtime checks in your add-in, ensure that you use the latest Office JavaScript API (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="d0055-141">指定最新的外接程序清单架构</span><span class="sxs-lookup"><span data-stu-id="d0055-141">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="d0055-142">外接程序清单必须使用外接程序清单架构版本 1.1。</span><span class="sxs-lookup"><span data-stu-id="d0055-142">Your add-in's manifest must use version 1.1 of the add-in manifest schema.</span></span> <span data-ttu-id="d0055-143">按 `OfficeApp` 如下方式设置外接程序清单中的元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-143">Set the `OfficeApp` element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a><span data-ttu-id="d0055-144">指定最新的 Office JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="d0055-144">Specify the latest Office JavaScript API library</span></span>

<span data-ttu-id="d0055-145">如果您使用运行时检查，请参考内容传送网络 (CDN) 中的 Office JavaScript API 库的最新版本。</span><span class="sxs-lookup"><span data-stu-id="d0055-145">If you use runtime checks, reference the most current version of the Office JavaScript API library from the content delivery network (CDN).</span></span> <span data-ttu-id="d0055-146">若要执行此操作，请将以下 `script` 标记添加到 HTML 中。</span><span class="sxs-lookup"><span data-stu-id="d0055-146">To do this, add the following  `script` tag to your HTML.</span></span> <span data-ttu-id="d0055-147">使用 CDN URL 中的 `/1/` 可以确保引用的是最新版本的 Office.js。</span><span class="sxs-lookup"><span data-stu-id="d0055-147">Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a><span data-ttu-id="d0055-148">用于指定 Office 应用程序或 API 要求的选项</span><span class="sxs-lookup"><span data-stu-id="d0055-148">Options to specify Office applications or API requirements</span></span>

<span data-ttu-id="d0055-149">当您指定 Office 应用程序或 API 要求时，有几个因素需要考虑。</span><span class="sxs-lookup"><span data-stu-id="d0055-149">When you specify Office applications or API requirements, there are several factors to consider.</span></span> <span data-ttu-id="d0055-150">下图显示了如何确定要在外接程序中使用的技术。</span><span class="sxs-lookup"><span data-stu-id="d0055-150">The following diagram shows how to decide which technique to use in your add-in.</span></span>

![在指定 Office 应用程序或 API 要求时，选择适用于你的外接程序的最佳选项](../images/options-for-office-hosts.png)

- <span data-ttu-id="d0055-152">如果你的外接程序在一个 Office 应用程序中运行，请 `Hosts` 在清单中设置该元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-152">If your add-in runs in one Office application, set the `Hosts` element in the manifest.</span></span> <span data-ttu-id="d0055-153">有关详细信息，请参阅 [设置 Hosts 元素](#set-the-hosts-element)。</span><span class="sxs-lookup"><span data-stu-id="d0055-153">For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="d0055-154">若要设置 Office 应用程序必须支持的最低要求集或 API 成员以运行您的外接程序，请 `Requirements` 在清单中设置该元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-154">To set the minimum requirement set or API members that an Office application must support to run your add-in, set the `Requirements` element in the manifest.</span></span> <span data-ttu-id="d0055-155">有关详细信息，请参阅[在清单中设置 Requirements 元素](#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="d0055-155">For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="d0055-156">如果要提供其他功能（如果 Office 应用程序中提供了特定要求集或 API 成员），请在您的外接程序的 JavaScript 代码中执行运行时检查。</span><span class="sxs-lookup"><span data-stu-id="d0055-156">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office application, perform a runtime check in your add-in's JavaScript code.</span></span> <span data-ttu-id="d0055-157">例如，如果加载项在 Excel 2016 中运行，请使用 Excel JavaScript API 中的 API 成员提供附加功能。</span><span class="sxs-lookup"><span data-stu-id="d0055-157">For example, if your add-in runs in Excel 2016, use API members from the Excel JavaScript API to provide additional functionality.</span></span> <span data-ttu-id="d0055-158">有关详细信息，请参阅[在你的 JavaScript 代码中使用运行时检查](#use-runtime-checks-in-your-javascript-code)。</span><span class="sxs-lookup"><span data-stu-id="d0055-158">For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="d0055-159">设置 Hosts 元素</span><span class="sxs-lookup"><span data-stu-id="d0055-159">Set the Hosts element</span></span>

<span data-ttu-id="d0055-160">若要使您的外接程序在一个 Office 客户端应用程序中运行，请使用 `Hosts` `Host` 清单中的和元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-160">To make your add-in run in one Office client application, use the `Hosts` and `Host` elements in the manifest.</span></span> <span data-ttu-id="d0055-161">如果不指定 `Hosts` 元素，则外接程序将在 Office 外接程序支持的所有 office 应用程序中运行。</span><span class="sxs-lookup"><span data-stu-id="d0055-161">If you don't specify the `Hosts` element, your add-in will run in all Office applications supported by Office Add-ins.</span></span>

<span data-ttu-id="d0055-162">例如，以下 `Hosts` 和 `Host` 声明指定外接程序将使用任何版本的 excel，其中包括在 Web、Windows 和 iPad 上的 excel。</span><span class="sxs-lookup"><span data-stu-id="d0055-162">For example, the following `Hosts` and `Host` declaration specifies that the add-in will work with any release of Excel, which includes Excel on the web, Windows, and iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="d0055-163">`Hosts`元素可以包含一个或多个 `Host` 元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-163">The `Hosts` element can contain one or more `Host` elements.</span></span> <span data-ttu-id="d0055-164">`Host`元素指定加载项所需的 Office 应用程序。</span><span class="sxs-lookup"><span data-stu-id="d0055-164">The `Host` element specifies the Office application your add-in requires.</span></span> <span data-ttu-id="d0055-165">`Name`属性是必需的，并且可设置为下列值之一。</span><span class="sxs-lookup"><span data-stu-id="d0055-165">The `Name` attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="d0055-166">名称</span><span class="sxs-lookup"><span data-stu-id="d0055-166">Name</span></span>          | <span data-ttu-id="d0055-167">Office 客户端应用程序</span><span class="sxs-lookup"><span data-stu-id="d0055-167">Office client applications</span></span>                      |
|:--------------|:----------------------------------------------|
| <span data-ttu-id="d0055-168">数据库</span><span class="sxs-lookup"><span data-stu-id="d0055-168">Database</span></span>      | <span data-ttu-id="d0055-169">Access Web App</span><span class="sxs-lookup"><span data-stu-id="d0055-169">Access web apps</span></span>                               |
| <span data-ttu-id="d0055-170">文档</span><span class="sxs-lookup"><span data-stu-id="d0055-170">Document</span></span>      | <span data-ttu-id="d0055-171">Word 网页版、Windows 版、Mac 版、iPad 版</span><span class="sxs-lookup"><span data-stu-id="d0055-171">Word on the web, Windows, Mac, iPad</span></span>           |
| <span data-ttu-id="d0055-172">邮箱</span><span class="sxs-lookup"><span data-stu-id="d0055-172">Mailbox</span></span>       | <span data-ttu-id="d0055-173">Outlook 网页版、Windows 版、Mac 版、Android 版、iOS 版</span><span class="sxs-lookup"><span data-stu-id="d0055-173">Outlook on the web, Windows, Mac, Android, iOS</span></span>|
| <span data-ttu-id="d0055-174">演示文稿</span><span class="sxs-lookup"><span data-stu-id="d0055-174">Presentation</span></span>  | <span data-ttu-id="d0055-175">PowerPoint 网页版、Windows 版、Mac 版、iPad 版</span><span class="sxs-lookup"><span data-stu-id="d0055-175">PowerPoint on the web, Windows, Mac, iPad</span></span>     |
| <span data-ttu-id="d0055-176">项目</span><span class="sxs-lookup"><span data-stu-id="d0055-176">Project</span></span>       | <span data-ttu-id="d0055-177">Windows 版 Project</span><span class="sxs-lookup"><span data-stu-id="d0055-177">Project on Windows</span></span>                            |
| <span data-ttu-id="d0055-178">工作簿</span><span class="sxs-lookup"><span data-stu-id="d0055-178">Workbook</span></span>      | <span data-ttu-id="d0055-179">Excel 网页版、Windows 版、Mac 版、iPad 版</span><span class="sxs-lookup"><span data-stu-id="d0055-179">Excel on the web, Windows, Mac, iPad</span></span>          |

> [!NOTE]
> <span data-ttu-id="d0055-180">`Name`属性指定可以运行外接程序的 Office 客户端应用程序。</span><span class="sxs-lookup"><span data-stu-id="d0055-180">The `Name` attribute specifies the Office client application that can run your add-in.</span></span> <span data-ttu-id="d0055-181">Office 应用程序在不同的平台上受支持，并在桌面、web 浏览器、平板电脑和移动设备上运行。</span><span class="sxs-lookup"><span data-stu-id="d0055-181">Office applications are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices.</span></span> <span data-ttu-id="d0055-182">不能指定用于运行外接程序的平台。</span><span class="sxs-lookup"><span data-stu-id="d0055-182">You can't specify which platform can be used to run your add-in.</span></span> <span data-ttu-id="d0055-183">例如，如果指定 `Mailbox` ，Outlook 网页版和 Windows 都可用于运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="d0055-183">For example, if you specify `Mailbox`, both Outlook on the web and Windows can be used to run your add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d0055-184">我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。</span><span class="sxs-lookup"><span data-stu-id="d0055-184">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="d0055-185">作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。</span><span class="sxs-lookup"><span data-stu-id="d0055-185">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="d0055-186">在清单中设置 Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="d0055-186">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="d0055-187">`Requirements`元素指定 Office 应用程序在运行外接程序时必须支持的最低要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="d0055-187">The `Requirements` element specifies the minimum requirement sets or API members that must be supported by the Office application to run your add-in.</span></span> <span data-ttu-id="d0055-188">`Requirements`元素可以指定要求集和外接程序中使用的各个方法。</span><span class="sxs-lookup"><span data-stu-id="d0055-188">The `Requirements` element can specify both requirement sets and individual methods used in your add-in.</span></span> <span data-ttu-id="d0055-189">在外接程序清单架构的版本1.1 中，除 Outlook 外接程序外接程序外接程序中，该 `Requirements` 元素是可选的。</span><span class="sxs-lookup"><span data-stu-id="d0055-189">In version 1.1 of the add-in manifest schema, the `Requirements` element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="d0055-190">仅使用 `Requirements` 元素指定你的外接程序必须使用的关键要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="d0055-190">Only use the `Requirements` element to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="d0055-191">如果 Office 应用程序或平台不支持元素中指定的要求集或 API 成员 `Requirements` ，则外接程序将不会在该应用程序或平台中运行，并且不会显示在 **我的外接**程序中。相反，我们建议您让外接程序在 Office 应用程序的所有平台上可用，如在 web、Windows 和 iPad 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="d0055-191">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad.</span></span> <span data-ttu-id="d0055-192">若要使你的外接程序在  _所有_ Office 应用程序和平台上可用，请使用运行时检查而不是 `Requirements` 元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-192">To make your add-in available on  _all_ Office applications and platforms, use runtime checks instead of the `Requirements` element.</span></span>

<span data-ttu-id="d0055-193">下面的代码示例演示在支持以下内容的所有 Office 客户端应用程序中加载的外接程序：</span><span class="sxs-lookup"><span data-stu-id="d0055-193">The following code example shows an add-in that loads in all Office client applications that support the following:</span></span>

-  <span data-ttu-id="d0055-194">`TableBindings` 要求集，其最低版本为 "1.1"。</span><span class="sxs-lookup"><span data-stu-id="d0055-194">`TableBindings` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="d0055-195">`OOXML` 要求集，其最低版本为 "1.1"。</span><span class="sxs-lookup"><span data-stu-id="d0055-195">`OOXML` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="d0055-196">`Document.getSelectedDataAsync` 种.</span><span class="sxs-lookup"><span data-stu-id="d0055-196">`Document.getSelectedDataAsync` method.</span></span>

```XML
<Requirements>
   <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" MinVersion="1.1"/>
      <Set Name="OOXML" MinVersion="1.1"/>
   </Sets>
   <Methods>
      <Method Name="Document.getSelectedDataAsync"/>
   </Methods>
</Requirements>
```

- <span data-ttu-id="d0055-197">`Requirements`元素包含 `Sets` 和 `Methods` 子元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-197">The `Requirements` element contains the `Sets` and `Methods` child elements.</span></span>

- <span data-ttu-id="d0055-198">`Sets`元素可以包含一个或多个 `Set` 元素。</span><span class="sxs-lookup"><span data-stu-id="d0055-198">The `Sets` element can contain one or more `Set` elements.</span></span> <span data-ttu-id="d0055-199">`DefaultMinVersion` 指定 `MinVersion` 所有子元素的默认值 `Set` 。</span><span class="sxs-lookup"><span data-stu-id="d0055-199">`DefaultMinVersion` specifies the default `MinVersion` value of all child `Set` elements.</span></span>

- <span data-ttu-id="d0055-200">`Set`元素指定 Office 应用程序必须支持的要求集以运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="d0055-200">The `Set` element specifies requirement sets that the Office application must support to run the add-in.</span></span> <span data-ttu-id="d0055-201">`Name`属性指定要求集的名称。</span><span class="sxs-lookup"><span data-stu-id="d0055-201">The `Name` attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="d0055-202">`MinVersion`指定要求集的最低版本。</span><span class="sxs-lookup"><span data-stu-id="d0055-202">The `MinVersion` specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="d0055-203">`MinVersion` 重写的值 `DefaultMinVersion` 有关您的 API 成员所属的要求集和要求集版本的详细信息，请参阅 [Office 外接程序要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="d0055-203">`MinVersion` overrides the value of `DefaultMinVersion` For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

- <span data-ttu-id="d0055-p122">`Methods`元素可以包含一个或多个 `Method` 元素。不能将 `Methods` 元素与 Outlook 外接程序一起使用。</span><span class="sxs-lookup"><span data-stu-id="d0055-p122">The `Methods` element can contain one or more `Method` elements. You can't use the `Methods` element with Outlook add-ins.</span></span>

- <span data-ttu-id="d0055-p123">`Method`元素指定在运行外接程序的 Office 应用程序中必须支持的单个方法。`Name`属性是必需的，并指定通过其父对象限定的方法的名称。</span><span class="sxs-lookup"><span data-stu-id="d0055-p123">The `Method` element specifies an individual method that must be supported in the Office application where your add-in runs. The `Name` attribute is required and specifies the name of the method qualified with its parent object.</span></span>

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="d0055-208">在你的 JavaScript 代码中使用运行时检查</span><span class="sxs-lookup"><span data-stu-id="d0055-208">Use runtime checks in your JavaScript code</span></span>

<span data-ttu-id="d0055-209">如果 Office 应用程序支持某些要求集，则您可能需要在外接程序中提供其他功能。</span><span class="sxs-lookup"><span data-stu-id="d0055-209">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office application.</span></span> <span data-ttu-id="d0055-210">例如，如果你的加载项在 Word 2016 中运行，则你可能想要在现有的加载项中使用 Word JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="d0055-210">For example, you might want to use the Word JavaScript APIs in your existing add-in if your add-in runs in Word 2016.</span></span> <span data-ttu-id="d0055-211">若要执行此操作，你可以使用含有要求集名称的 [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法。</span><span class="sxs-lookup"><span data-stu-id="d0055-211">To do this, you use the [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method with the name of the requirement set.</span></span> <span data-ttu-id="d0055-212">`isSetSupported` 在运行时确定运行加载项的 Office 应用程序是否支持要求集。</span><span class="sxs-lookup"><span data-stu-id="d0055-212">`isSetSupported` determines, at runtime, whether the Office application running the add-in supports the requirement set.</span></span> <span data-ttu-id="d0055-213">如果支持该要求集，则 `isSetSupported` 返回 **true** ，并运行使用该要求集的 API 成员的其他代码。</span><span class="sxs-lookup"><span data-stu-id="d0055-213">If the requirement set is supported, `isSetSupported` returns **true** and runs the additional code that uses the API members from that requirement set.</span></span> <span data-ttu-id="d0055-214">如果 Office 应用程序不支持要求集，将 `isSetSupported` 返回 **false** ，并且不会运行其他代码。</span><span class="sxs-lookup"><span data-stu-id="d0055-214">If the Office application doesn't support the requirement set, `isSetSupported` returns **false** and the additional code won't run.</span></span> <span data-ttu-id="d0055-215">下面的代码演示与一起使用的语法 `isSetSupported` 。</span><span class="sxs-lookup"><span data-stu-id="d0055-215">The following code shows the syntax to use with `isSetSupported`.</span></span>

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- <span data-ttu-id="d0055-216">_RequirementSetName_（必填）是代表该要求集名称的字符串（例如，“**ExcelApi**”、“**Mailbox**”等）。</span><span class="sxs-lookup"><span data-stu-id="d0055-216">_RequirementSetName_ (required) is a string that represents the name of the requirement set (e.g., "**ExcelApi**", "**Mailbox**", etc.).</span></span> <span data-ttu-id="d0055-217">有关可用要求集的详细信息，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="d0055-217">For more information about available requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>
- <span data-ttu-id="d0055-218">_MinimumVersion_ (optional) 是一个字符串，指定 Office 应用程序必须支持的最低要求集版本，才能运行语句中的代码 `if` (例如，"**1.9**" ) 。</span><span class="sxs-lookup"><span data-stu-id="d0055-218">_MinimumVersion_ (optional) is a string that specifies the minimum requirement set version that the Office application must support in order for the code within the `if` statement to run (e.g., "**1.9**").</span></span>

> [!WARNING]
> <span data-ttu-id="d0055-219">调用方法时 `isSetSupported` ， `MinimumVersion` 如果指定) 的参数 (的值应为字符串。</span><span class="sxs-lookup"><span data-stu-id="d0055-219">When calling the `isSetSupported` method, the value of the `MinimumVersion` parameter (if specified) should be a string.</span></span> <span data-ttu-id="d0055-220">这是因为 JavaScript 分析器无法区分数值，例如 1.1 和 1.10，因为它可以用于字符串值，例如“1.1”和“1.10”。</span><span class="sxs-lookup"><span data-stu-id="d0055-220">This is because the JavaScript parser cannot differentiate between numeric values such as 1.1 and 1.10, where as it can for string values such as "1.1" and "1.10".</span></span>
> <span data-ttu-id="d0055-221">`number` 重载已弃用。</span><span class="sxs-lookup"><span data-stu-id="d0055-221">The `number` overload is deprecated.</span></span>

<span data-ttu-id="d0055-222">`isSetSupported`与 `RequirementSetName` Office 应用程序关联使用，如下所示。</span><span class="sxs-lookup"><span data-stu-id="d0055-222">Use `isSetSupported` with the `RequirementSetName` associated with the Office application as follows.</span></span>

|<span data-ttu-id="d0055-223">Office 应用程序</span><span class="sxs-lookup"><span data-stu-id="d0055-223">Office application</span></span>|<span data-ttu-id="d0055-224">RequirementSetName</span><span class="sxs-lookup"><span data-stu-id="d0055-224">RequirementSetName</span></span>|
|---|---|
|<span data-ttu-id="d0055-225">Excel</span><span class="sxs-lookup"><span data-stu-id="d0055-225">Excel</span></span>|<span data-ttu-id="d0055-226">ExcelApi</span><span class="sxs-lookup"><span data-stu-id="d0055-226">ExcelApi</span></span>|
|<span data-ttu-id="d0055-227">OneNote</span><span class="sxs-lookup"><span data-stu-id="d0055-227">OneNote</span></span>|<span data-ttu-id="d0055-228">OneNoteApi</span><span class="sxs-lookup"><span data-stu-id="d0055-228">OneNoteApi</span></span>|
|<span data-ttu-id="d0055-229">Outlook</span><span class="sxs-lookup"><span data-stu-id="d0055-229">Outlook</span></span>|<span data-ttu-id="d0055-230">Mailbox</span><span class="sxs-lookup"><span data-stu-id="d0055-230">Mailbox</span></span>|
|<span data-ttu-id="d0055-231">Word</span><span class="sxs-lookup"><span data-stu-id="d0055-231">Word</span></span>|<span data-ttu-id="d0055-232">WordApi</span><span class="sxs-lookup"><span data-stu-id="d0055-232">WordApi</span></span>|

<span data-ttu-id="d0055-233">在 `isSetSupported` CDN 上的最新 Office.js 文件中提供了这些应用程序的方法和要求集。</span><span class="sxs-lookup"><span data-stu-id="d0055-233">The `isSetSupported` method and the requirement sets for these applications are available in the latest Office.js file on the CDN.</span></span> <span data-ttu-id="d0055-234">如果不使用 CDN 中的 Office.js，外接程序可能会生成异常，因为这 `isSetSupported` 将是不确定的。</span><span class="sxs-lookup"><span data-stu-id="d0055-234">If you don't use Office.js from the CDN, your add-in might generate exceptions because `isSetSupported` will be undefined.</span></span> <span data-ttu-id="d0055-235">有关详细信息，请参阅 [指定最新的 Office JAVASCRIPT API 库](#specify-the-latest-office-javascript-api-library)。</span><span class="sxs-lookup"><span data-stu-id="d0055-235">For more information, see [Specify the latest Office JavaScript API library](#specify-the-latest-office-javascript-api-library).</span></span>

<span data-ttu-id="d0055-236">下面的代码示例演示外接程序如何为可能支持不同要求集或 API 成员的不同 Office 应用程序提供不同的功能。</span><span class="sxs-lookup"><span data-stu-id="d0055-236">The following code example shows how an add-in can provide different functionality for different Office applications that might support different requirement sets or API members.</span></span>

```js
if (Office.context.requirements.isSetSupported('WordApi', '1.1'))
{
    // Run code that provides additional functionality using the Word JavaScript API when the add-in runs in Word 2016 or later.
}
else if (Office.context.requirements.isSetSupported('CustomXmlParts'))
{
    // Run code that uses API members from the CustomXmlParts requirement set.
}
else
{
    // Run additional code when the Office application is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="d0055-237">使用不属于要求集的方法的运行时检查</span><span class="sxs-lookup"><span data-stu-id="d0055-237">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="d0055-238">部分 API 成员不属于要求集</span><span class="sxs-lookup"><span data-stu-id="d0055-238">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="d0055-239">这仅适用于属于[Office JAVASCRIPT api](../reference/javascript-api-for-office.md)命名空间的 api 成员 (`Office.` 除非[Outlook 邮箱 api](/javascript/api/outlook)) ，而不是属于[Word JavaScript api](../reference/overview/word-add-ins-reference-overview.md)的 api 成员 () 中的任何内容，Excel JavaScript api () 中的任何内容， `Word.` [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) `Excel.` 或者[OneNote JavaScript api](../reference/overview/onenote-add-ins-javascript-reference.md) (命名空间中的任何内容 `OneNote.` 。</span><span class="sxs-lookup"><span data-stu-id="d0055-239">This only applies to API members that are part of the [Office JavaScript API](../reference/javascript-api-for-office.md) namespace (anything under `Office.` except [Outlook Mailbox APIs](/javascript/api/outlook)), but not API members that belong to the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (anything in `Word.`), [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (anything in `Excel.`), or [OneNote JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (anything in `OneNote.`) namespaces.</span></span> <span data-ttu-id="d0055-240">如果外接程序依赖于不属于要求集的方法，则可以使用运行时检查来确定该方法是否受 Office 应用程序支持，如下面的代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="d0055-240">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office application, as shown in the following code example.</span></span> <span data-ttu-id="d0055-241">有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)。</span><span class="sxs-lookup"><span data-stu-id="d0055-241">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).</span></span>

> [!NOTE]
> <span data-ttu-id="d0055-242">建议限制在加载项代码中使用此类型运行时检查。</span><span class="sxs-lookup"><span data-stu-id="d0055-242">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="d0055-243">下面的代码示例检查 Office 应用程序是否支持 `document.setSelectedDataAsync` 。</span><span class="sxs-lookup"><span data-stu-id="d0055-243">The following code example checks whether the Office application supports `document.setSelectedDataAsync`.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a><span data-ttu-id="d0055-244">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d0055-244">See also</span></span>

- [<span data-ttu-id="d0055-245">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="d0055-245">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="d0055-246">Office 加载项要求集</span><span class="sxs-lookup"><span data-stu-id="d0055-246">Office Add-in requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="d0055-247">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="d0055-247">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
