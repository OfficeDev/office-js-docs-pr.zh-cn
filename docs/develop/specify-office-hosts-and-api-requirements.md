---
title: 指定 Office 主机和 API 要求
description: 了解如何指定Office的应用程序和 API 要求，使加载项按预期工作。
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: dad4a562c523291d7af3c975c89111fab410220c
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075962"
---
# <a name="specify-office-applications-and-api-requirements"></a><span data-ttu-id="2bf73-103">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="2bf73-103">Specify Office applications and API requirements</span></span>

<span data-ttu-id="2bf73-104">你的Office加载项可能依赖于特定的 Office 应用程序、要求集、API 成员或 API 版本才能按预期工作。</span><span class="sxs-lookup"><span data-stu-id="2bf73-104">Your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API in order to work as expected.</span></span> <span data-ttu-id="2bf73-105">例如，你的外接程序可能：</span><span class="sxs-lookup"><span data-stu-id="2bf73-105">For example, your add-in might:</span></span>

- <span data-ttu-id="2bf73-106">在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。</span><span class="sxs-lookup"><span data-stu-id="2bf73-106">Run in a single Office application (e.g., Word or Excel), or several applications.</span></span>

- <span data-ttu-id="2bf73-p102">使用仅在 Office 的某些版本中可用的 JavaScript API。例如，可能会在运行在 Excel 2016 中的外接程序中使用适用于 Excel 的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="2bf73-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="2bf73-109">只能在 Office 的某些版本中运行，这些版本支持供你的外接程序使用的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="2bf73-109">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="2bf73-110">本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。</span><span class="sxs-lookup"><span data-stu-id="2bf73-110">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="2bf73-111">有关当前支持Office外接程序的高级别视图，请参阅 Office 外接程序的[Office 客户端](../overview/office-add-in-availability.md)应用程序和平台可用性页面。</span><span class="sxs-lookup"><span data-stu-id="2bf73-111">For a high-level view of where Office Add-ins are currently supported, see the [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="2bf73-112">下表列出了本文中讨论的核心概念。</span><span class="sxs-lookup"><span data-stu-id="2bf73-112">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="2bf73-113">**概念**</span><span class="sxs-lookup"><span data-stu-id="2bf73-113">**Concept**</span></span>|<span data-ttu-id="2bf73-114">**说明**</span><span class="sxs-lookup"><span data-stu-id="2bf73-114">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="2bf73-115">Office应用程序，Office客户端应用程序</span><span class="sxs-lookup"><span data-stu-id="2bf73-115">Office application, Office client application</span></span>|<span data-ttu-id="2bf73-p103">用于运行加载项的 Office 应用程序。例如 Word、Excel 等。</span><span class="sxs-lookup"><span data-stu-id="2bf73-p103">The Office application used to run your add-in. For example, Word, Excel, and so on.</span></span>|
|<span data-ttu-id="2bf73-118">平台</span><span class="sxs-lookup"><span data-stu-id="2bf73-118">Platform</span></span>|<span data-ttu-id="2bf73-119">应用程序Office，例如浏览器或浏览器iPad。</span><span class="sxs-lookup"><span data-stu-id="2bf73-119">Where the Office application runs, such as in a browser or on an iPad.</span></span>|
|<span data-ttu-id="2bf73-120">要求集</span><span class="sxs-lookup"><span data-stu-id="2bf73-120">Requirement set</span></span>|<span data-ttu-id="2bf73-121">命名的一组相关的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="2bf73-121">A named group of related API members.</span></span> <span data-ttu-id="2bf73-122">外接程序使用要求集来确定Office应用程序是否支持外接程序使用的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="2bf73-122">Add-ins use requirement sets to determine whether the Office application supports API members used by your add-in.</span></span> <span data-ttu-id="2bf73-123">测试对要求集的支持比对单个的 API 成员的支持更为容易。</span><span class="sxs-lookup"><span data-stu-id="2bf73-123">It's easier to test for the support of a requirement set than for the support of individual API members.</span></span> <span data-ttu-id="2bf73-124">要求集支持因Office应用程序以及应用程序版本Office而异。</span><span class="sxs-lookup"><span data-stu-id="2bf73-124">Requirement set support varies by Office application and the version of the Office application.</span></span> <br ><span data-ttu-id="2bf73-125">要求集在清单文件中指定。</span><span class="sxs-lookup"><span data-stu-id="2bf73-125">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="2bf73-126">在清单中指定要求集时，应设置应用程序必须提供的最低级别的 API Office应用程序才能运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="2bf73-126">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office application must provide in order to run your add-in.</span></span> <span data-ttu-id="2bf73-127">Office清单中指定的要求集的应用程序无法运行您的外接程序，并且您的外接程序不会显示在"我的外接程序<span class="ui">"中</span>。这将限制外接程序的可用位置。</span><span class="sxs-lookup"><span data-stu-id="2bf73-127">Office applications that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.</span></span> <span data-ttu-id="2bf73-128">在使用运行时检查的代码中。</span><span class="sxs-lookup"><span data-stu-id="2bf73-128">In code using runtime checks.</span></span> <span data-ttu-id="2bf73-129">有关要求集的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-129">For the complete list of requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>|
|<span data-ttu-id="2bf73-130">运行时检查</span><span class="sxs-lookup"><span data-stu-id="2bf73-130">Runtime check</span></span>|<span data-ttu-id="2bf73-131">在运行时执行的一个测试，用于确定Office加载项的加载项应用程序是否支持加载项使用的要求集或方法。</span><span class="sxs-lookup"><span data-stu-id="2bf73-131">A test that is performed at runtime to determine whether the Office application running your add-in supports requirement sets or methods used by your add-in.</span></span> <span data-ttu-id="2bf73-132">若要执行运行时检查，请使用 **if** 语句和方法、要求集或不是要求集 `isSetSupported` 一部分的方法名称。</span><span class="sxs-lookup"><span data-stu-id="2bf73-132">To perform a runtime check, you use an **if** statement with the `isSetSupported` method, the requirement sets, or the method names that aren't part of a requirement set.</span></span> <span data-ttu-id="2bf73-133">使用运行时检查可确保加载项能够覆盖最大数量的客户。</span><span class="sxs-lookup"><span data-stu-id="2bf73-133">Use runtime checks to ensure that your add-in reaches the broadest number of customers.</span></span> <span data-ttu-id="2bf73-134">与要求集不同，运行时检查不指定外接程序运行Office应用程序必须提供的最低级别的 API 支持。</span><span class="sxs-lookup"><span data-stu-id="2bf73-134">Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office application must provide for your add-in to run.</span></span> <span data-ttu-id="2bf73-135">相反，使用 **if** 语句来确定 API 成员是否受支持。</span><span class="sxs-lookup"><span data-stu-id="2bf73-135">Instead, you use the **if** statement to determine whether an API member is supported.</span></span> <span data-ttu-id="2bf73-136">如果支持，则可以在外接程序中提供其他功能。</span><span class="sxs-lookup"><span data-stu-id="2bf73-136">If it is, you can provide additional functionality in your add-in.</span></span> <span data-ttu-id="2bf73-137">使用运行时检查时，你的外接程序将始终在“**我的外接程序**”中显示。</span><span class="sxs-lookup"><span data-stu-id="2bf73-137">Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="2bf73-138">开始之前</span><span class="sxs-lookup"><span data-stu-id="2bf73-138">Before you begin</span></span>

<span data-ttu-id="2bf73-139">您的外接程序必须使用最新版本的外接程序清单架构。</span><span class="sxs-lookup"><span data-stu-id="2bf73-139">Your add-in must use the most current version of the add-in manifest schema.</span></span> <span data-ttu-id="2bf73-140">如果在加载项中使用运行时检查，请确保使用 JavaScript API Office的最新 (office.js) 库。</span><span class="sxs-lookup"><span data-stu-id="2bf73-140">If you use runtime checks in your add-in, ensure that you use the latest Office JavaScript API (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="2bf73-141">指定最新的外接程序清单架构</span><span class="sxs-lookup"><span data-stu-id="2bf73-141">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="2bf73-142">外接程序清单必须使用外接程序清单架构版本 1.1。</span><span class="sxs-lookup"><span data-stu-id="2bf73-142">Your add-in's manifest must use version 1.1 of the add-in manifest schema.</span></span> <span data-ttu-id="2bf73-143">在外接程序清单中设置 [OfficeApp](../reference/manifest/officeapp.md) 元素，如下所示。</span><span class="sxs-lookup"><span data-stu-id="2bf73-143">Set the [OfficeApp](../reference/manifest/officeapp.md) element in your add-in manifest as follows.</span></span> <span data-ttu-id="2bf73-144">本示例显示 `TaskPaneApp` 类型。</span><span class="sxs-lookup"><span data-stu-id="2bf73-144">This example shows the `TaskPaneApp` type.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a><span data-ttu-id="2bf73-145">指定最新 Office JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="2bf73-145">Specify the latest Office JavaScript API library</span></span>

<span data-ttu-id="2bf73-146">如果使用运行时检查，请从内容交付网络Office JavaScript API 库的最新版本 (CDN) 。</span><span class="sxs-lookup"><span data-stu-id="2bf73-146">If you use runtime checks, reference the most current version of the Office JavaScript API library from the content delivery network (CDN).</span></span> <span data-ttu-id="2bf73-147">若要执行此操作，请将以下 `script` 标记添加到 HTML 中。</span><span class="sxs-lookup"><span data-stu-id="2bf73-147">To do this, add the following  `script` tag to your HTML.</span></span> <span data-ttu-id="2bf73-148">使用 CDN URL 中的 `/1/` 可以确保引用的是最新版本的 Office.js。</span><span class="sxs-lookup"><span data-stu-id="2bf73-148">Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-applications-or-api-requirements"></a><span data-ttu-id="2bf73-149">用于指定应用程序Office API 要求的选项</span><span class="sxs-lookup"><span data-stu-id="2bf73-149">Options to specify Office applications or API requirements</span></span>

<span data-ttu-id="2bf73-150">当你指定Office或 API 要求时，有几个因素需要考虑。</span><span class="sxs-lookup"><span data-stu-id="2bf73-150">When you specify Office applications or API requirements, there are several factors to consider.</span></span> <span data-ttu-id="2bf73-151">下图显示了如何确定要在外接程序中使用的技术。</span><span class="sxs-lookup"><span data-stu-id="2bf73-151">The following diagram shows how to decide which technique to use in your add-in.</span></span>

![指定应用程序或 API 要求时，选择Office最佳选项。](../images/options-for-office-hosts.png)

- <span data-ttu-id="2bf73-153">如果外接程序在一个Office中运行，请设置 `Hosts` 清单中的 元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-153">If your add-in runs in one Office application, set the `Hosts` element in the manifest.</span></span> <span data-ttu-id="2bf73-154">有关详细信息，请参阅 [设置 Hosts 元素](#set-the-hosts-element)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-154">For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="2bf73-155">若要设置应用程序运行外接程序Office的最低要求集或 API 成员，请设置清单 `Requirements` 中的 元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-155">To set the minimum requirement set or API members that an Office application must support to run your add-in, set the `Requirements` element in the manifest.</span></span> <span data-ttu-id="2bf73-156">有关详细信息，请参阅[在清单中设置 Requirements 元素](#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-156">For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="2bf73-157">如果要在 Office 应用程序中提供特定要求集或 API 成员时提供其他功能，请在外接程序的 JavaScript 代码中执行运行时检查。</span><span class="sxs-lookup"><span data-stu-id="2bf73-157">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office application, perform a runtime check in your add-in's JavaScript code.</span></span> <span data-ttu-id="2bf73-158">例如，如果加载项在 Excel 2016 中运行，请使用 Excel JavaScript API 中的 API 成员提供附加功能。</span><span class="sxs-lookup"><span data-stu-id="2bf73-158">For example, if your add-in runs in Excel 2016, use API members from the Excel JavaScript API to provide additional functionality.</span></span> <span data-ttu-id="2bf73-159">有关详细信息，请参阅[在你的 JavaScript 代码中使用运行时检查](#use-runtime-checks-in-your-javascript-code)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-159">For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="2bf73-160">设置 Hosts 元素</span><span class="sxs-lookup"><span data-stu-id="2bf73-160">Set the Hosts element</span></span>

<span data-ttu-id="2bf73-161">若要使外接程序在一个客户端Office中运行，请使用清单中的 `Hosts` 和 `Host` 元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-161">To make your add-in run in one Office client application, use the `Hosts` and `Host` elements in the manifest.</span></span> <span data-ttu-id="2bf73-162">如果不指定 元素，外接程序将在指定类型支持的所有 Office 应用程序中运行 (即邮件、任务窗格或内容 `Hosts` `OfficeApp`) 。</span><span class="sxs-lookup"><span data-stu-id="2bf73-162">If you don't specify the `Hosts` element, your add-in will run in all Office applications supported by the specified `OfficeApp` type (that is, Mail, Task pane, or Content).</span></span>

<span data-ttu-id="2bf73-163">例如，以下 和 声明指定外接程序将适用于任何 Excel 版本，其中包括 Excel web 版、Windows 和 `Hosts` `Host` iPad。</span><span class="sxs-lookup"><span data-stu-id="2bf73-163">For example, the following `Hosts` and `Host` declaration specifies that the add-in will work with any release of Excel, which includes Excel on the web, Windows, and iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="2bf73-164">元素 `Hosts` 可以包含一个或多个 `Host` 元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-164">The `Hosts` element can contain one or more `Host` elements.</span></span> <span data-ttu-id="2bf73-165">`Host`元素指定Office所需的应用程序类型。</span><span class="sxs-lookup"><span data-stu-id="2bf73-165">The `Host` element specifies the Office application your add-in requires.</span></span> <span data-ttu-id="2bf73-166">`Name`属性是必需的，可以设置为下列值之一。</span><span class="sxs-lookup"><span data-stu-id="2bf73-166">The `Name` attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="2bf73-167">名称</span><span class="sxs-lookup"><span data-stu-id="2bf73-167">Name</span></span>          | <span data-ttu-id="2bf73-168">Office客户端应用程序</span><span class="sxs-lookup"><span data-stu-id="2bf73-168">Office client applications</span></span>                     | <span data-ttu-id="2bf73-169">可用的外接程序类型</span><span class="sxs-lookup"><span data-stu-id="2bf73-169">Available add-in types</span></span> |
|:--------------|:-----------------------------------------------|:-----------------------|
| <span data-ttu-id="2bf73-170">数据库</span><span class="sxs-lookup"><span data-stu-id="2bf73-170">Database</span></span>      | <span data-ttu-id="2bf73-171">Access Web App</span><span class="sxs-lookup"><span data-stu-id="2bf73-171">Access web apps</span></span>                                | <span data-ttu-id="2bf73-172">任务窗格</span><span class="sxs-lookup"><span data-stu-id="2bf73-172">Task pane</span></span>              |
| <span data-ttu-id="2bf73-173">文档</span><span class="sxs-lookup"><span data-stu-id="2bf73-173">Document</span></span>      | <span data-ttu-id="2bf73-174">Word web 版、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="2bf73-174">Word on the web, Windows, Mac, iPad</span></span>            | <span data-ttu-id="2bf73-175">任务窗格</span><span class="sxs-lookup"><span data-stu-id="2bf73-175">Task pane</span></span>              |
| <span data-ttu-id="2bf73-176">邮箱</span><span class="sxs-lookup"><span data-stu-id="2bf73-176">Mailbox</span></span>       | <span data-ttu-id="2bf73-177">Outlook 网页版、Windows、Mac、Android、iOS</span><span class="sxs-lookup"><span data-stu-id="2bf73-177">Outlook on the web, Windows, Mac, Android, iOS</span></span> | <span data-ttu-id="2bf73-178">邮件</span><span class="sxs-lookup"><span data-stu-id="2bf73-178">Mail</span></span>                   |
| <span data-ttu-id="2bf73-179">笔记本</span><span class="sxs-lookup"><span data-stu-id="2bf73-179">Notebook</span></span>      | <span data-ttu-id="2bf73-180">OneNote 网页版</span><span class="sxs-lookup"><span data-stu-id="2bf73-180">OneNote on the web</span></span>                             | <span data-ttu-id="2bf73-181">任务窗格、内容</span><span class="sxs-lookup"><span data-stu-id="2bf73-181">Task pane, Content</span></span>     |
| <span data-ttu-id="2bf73-182">演示文稿</span><span class="sxs-lookup"><span data-stu-id="2bf73-182">Presentation</span></span>  | <span data-ttu-id="2bf73-183">PowerPoint web 版、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="2bf73-183">PowerPoint on the web, Windows, Mac, iPad</span></span>      | <span data-ttu-id="2bf73-184">任务窗格、内容</span><span class="sxs-lookup"><span data-stu-id="2bf73-184">Task pane, Content</span></span>     |
| <span data-ttu-id="2bf73-185">项目</span><span class="sxs-lookup"><span data-stu-id="2bf73-185">Project</span></span>       | <span data-ttu-id="2bf73-186">Windows 版 Project</span><span class="sxs-lookup"><span data-stu-id="2bf73-186">Project on Windows</span></span>                             | <span data-ttu-id="2bf73-187">任务窗格</span><span class="sxs-lookup"><span data-stu-id="2bf73-187">Task pane</span></span>              |
| <span data-ttu-id="2bf73-188">工作簿</span><span class="sxs-lookup"><span data-stu-id="2bf73-188">Workbook</span></span>      | <span data-ttu-id="2bf73-189">Excel web 版、Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="2bf73-189">Excel on the web, Windows, Mac, iPad</span></span>           | <span data-ttu-id="2bf73-190">任务窗格、内容</span><span class="sxs-lookup"><span data-stu-id="2bf73-190">Task pane, Content</span></span>     |

> [!NOTE]
> <span data-ttu-id="2bf73-191">`Name`属性指定Office加载项的客户端应用程序。</span><span class="sxs-lookup"><span data-stu-id="2bf73-191">The `Name` attribute specifies the Office client application that can run your add-in.</span></span> <span data-ttu-id="2bf73-192">Office应用程序在不同的平台上受支持，并且运行在桌面、Web 浏览器、平板电脑和移动设备上。</span><span class="sxs-lookup"><span data-stu-id="2bf73-192">Office applications are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices.</span></span> <span data-ttu-id="2bf73-193">不能指定用于运行外接程序的平台。</span><span class="sxs-lookup"><span data-stu-id="2bf73-193">You can't specify which platform can be used to run your add-in.</span></span> <span data-ttu-id="2bf73-194">例如，如果指定 ，Outlook 网页版 和 `Mailbox` Windows 都可用于运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="2bf73-194">For example, if you specify `Mailbox`, both Outlook on the web and on Windows can be used to run your add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2bf73-195">我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。</span><span class="sxs-lookup"><span data-stu-id="2bf73-195">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="2bf73-196">作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。</span><span class="sxs-lookup"><span data-stu-id="2bf73-196">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="2bf73-197">在清单中设置 Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="2bf73-197">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="2bf73-198">元素指定应用程序必须支持的最低要求集或 API Office `Requirements` 应用程序运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="2bf73-198">The `Requirements` element specifies the minimum requirement sets or API members that must be supported by the Office application to run your add-in.</span></span> <span data-ttu-id="2bf73-199">`Requirements`元素可以指定要求集和外接程序中使用的单个方法。</span><span class="sxs-lookup"><span data-stu-id="2bf73-199">The `Requirements` element can specify both requirement sets and individual methods used in your add-in.</span></span> <span data-ttu-id="2bf73-200">在外接程序清单架构的版本 1.1 中，元素对于所有外接程序都是可选的，但Outlook `Requirements` 外接程序除外。</span><span class="sxs-lookup"><span data-stu-id="2bf73-200">In version 1.1 of the add-in manifest schema, the `Requirements` element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="2bf73-201">只能使用 `Requirements` 元素指定外接程序必须使用的关键要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="2bf73-201">Only use the `Requirements` element to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="2bf73-202">如果 Office 应用程序或平台不支持 元素中指定的要求集或 API 成员，外接程序将不会在该应用程序或平台中运行，也不会显示在"我的外接程序 `Requirements` **"中**。相反，我们建议你在 Office 应用程序的所有平台上提供外接程序，如 Excel web 版、Windows 和 iPad。</span><span class="sxs-lookup"><span data-stu-id="2bf73-202">If the Office application or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that application or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office application, such as Excel on the web, Windows, and iPad.</span></span> <span data-ttu-id="2bf73-203">若要使外接程序在所有 _应用程序Office都_ 可用，请使用运行时检查而不是 `Requirements` 元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-203">To make your add-in available on  _all_ Office applications and platforms, use runtime checks instead of the `Requirements` element.</span></span>

<span data-ttu-id="2bf73-204">以下代码示例演示在支持以下内容的所有 Office 客户端应用程序中加载的外接程序：</span><span class="sxs-lookup"><span data-stu-id="2bf73-204">The following code example shows an add-in that loads in all Office client applications that support the following:</span></span>

-  <span data-ttu-id="2bf73-205">`TableBindings` 要求集，最低版本为"1.1"。</span><span class="sxs-lookup"><span data-stu-id="2bf73-205">`TableBindings` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="2bf73-206">`OOXML` 要求集，最低版本为"1.1"。</span><span class="sxs-lookup"><span data-stu-id="2bf73-206">`OOXML` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="2bf73-207">`Document.getSelectedDataAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="2bf73-207">`Document.getSelectedDataAsync` method.</span></span>

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

- <span data-ttu-id="2bf73-208">元素 `Requirements` 包含 `Sets` 和 `Methods` 子元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-208">The `Requirements` element contains the `Sets` and `Methods` child elements.</span></span>

- <span data-ttu-id="2bf73-209">元素 `Sets` 可以包含一个或多个 `Set` 元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-209">The `Sets` element can contain one or more `Set` elements.</span></span> <span data-ttu-id="2bf73-210">`DefaultMinVersion` 指定所有 `MinVersion` 子元素的 `Set` 默认值。</span><span class="sxs-lookup"><span data-stu-id="2bf73-210">`DefaultMinVersion` specifies the default `MinVersion` value of all child `Set` elements.</span></span>

- <span data-ttu-id="2bf73-211">`Set`元素指定应用程序必须Office运行外接程序的要求集。</span><span class="sxs-lookup"><span data-stu-id="2bf73-211">The `Set` element specifies requirement sets that the Office application must support to run the add-in.</span></span> <span data-ttu-id="2bf73-212">`Name`属性指定要求集的名称。</span><span class="sxs-lookup"><span data-stu-id="2bf73-212">The `Name` attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="2bf73-213">`MinVersion`指定要求集的最低版本。</span><span class="sxs-lookup"><span data-stu-id="2bf73-213">The `MinVersion` specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="2bf73-214">`MinVersion`替代 的值 有关 API 成员所属的要求集和要求集版本Office外接程序 `DefaultMinVersion` [要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-214">`MinVersion` overrides the value of `DefaultMinVersion` For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

- <span data-ttu-id="2bf73-215">元素 `Methods` 可以包含一个或多个 `Method` 元素。</span><span class="sxs-lookup"><span data-stu-id="2bf73-215">The `Methods` element can contain one or more `Method` elements.</span></span> <span data-ttu-id="2bf73-216">无法将 元素与 `Methods` Outlook一起使用。</span><span class="sxs-lookup"><span data-stu-id="2bf73-216">You can't use the `Methods` element with Outlook add-ins.</span></span>

- <span data-ttu-id="2bf73-217">元素指定在运行加载项的 Office `Method` 应用程序中必须支持的单个方法。</span><span class="sxs-lookup"><span data-stu-id="2bf73-217">The `Method` element specifies an individual method that must be supported in the Office application where your add-in runs.</span></span> <span data-ttu-id="2bf73-218">`Name`属性是必需的，并指定使用其父对象限定的方法的名称。</span><span class="sxs-lookup"><span data-stu-id="2bf73-218">The `Name` attribute is required and specifies the name of the method qualified with its parent object.</span></span>

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="2bf73-219">在你的 JavaScript 代码中使用运行时检查</span><span class="sxs-lookup"><span data-stu-id="2bf73-219">Use runtime checks in your JavaScript code</span></span>

<span data-ttu-id="2bf73-220">如果外接程序应用程序支持某些要求集，您可能希望在外接程序Office功能。</span><span class="sxs-lookup"><span data-stu-id="2bf73-220">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office application.</span></span> <span data-ttu-id="2bf73-221">例如，如果你的加载项在 Word 2016 中运行，则你可能想要在现有的加载项中使用 Word JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="2bf73-221">For example, you might want to use the Word JavaScript APIs in your existing add-in if your add-in runs in Word 2016.</span></span> <span data-ttu-id="2bf73-222">若要执行此操作，你可以使用含有要求集名称的 [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法。</span><span class="sxs-lookup"><span data-stu-id="2bf73-222">To do this, you use the [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method with the name of the requirement set.</span></span> <span data-ttu-id="2bf73-223">`isSetSupported`确定运行外接程序Office的应用程序是否支持要求集。</span><span class="sxs-lookup"><span data-stu-id="2bf73-223">`isSetSupported` determines, at runtime, whether the Office application running the add-in supports the requirement set.</span></span> <span data-ttu-id="2bf73-224">如果要求集受支持，则返回 true 并运行使用该要求集的 API 成员 `isSetSupported` 的其他代码。 </span><span class="sxs-lookup"><span data-stu-id="2bf73-224">If the requirement set is supported, `isSetSupported` returns **true** and runs the additional code that uses the API members from that requirement set.</span></span> <span data-ttu-id="2bf73-225">如果Office应用程序不支持要求集，则返回 `isSetSupported` **false，** 其他代码将不会运行。</span><span class="sxs-lookup"><span data-stu-id="2bf73-225">If the Office application doesn't support the requirement set, `isSetSupported` returns **false** and the additional code won't run.</span></span> <span data-ttu-id="2bf73-226">以下代码显示与 `isSetSupported`结合使用的语法。</span><span class="sxs-lookup"><span data-stu-id="2bf73-226">The following code shows the syntax to use with `isSetSupported`.</span></span>

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- <span data-ttu-id="2bf73-227">_RequirementSetName_（必填）是代表该要求集名称的字符串（例如，“**ExcelApi**”、“**Mailbox**”等）。</span><span class="sxs-lookup"><span data-stu-id="2bf73-227">_RequirementSetName_ (required) is a string that represents the name of the requirement set (e.g., "**ExcelApi**", "**Mailbox**", etc.).</span></span> <span data-ttu-id="2bf73-228">有关可用要求集的详细信息，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-228">For more information about available requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>
- <span data-ttu-id="2bf73-229">_MinimumVersion_ (可选) 是一个字符串，用于指定 Office 应用程序必须支持的最低要求集版本，以便语句中的代码运行 (例如 `if` **"1.9") 。**</span><span class="sxs-lookup"><span data-stu-id="2bf73-229">_MinimumVersion_ (optional) is a string that specifies the minimum requirement set version that the Office application must support in order for the code within the `if` statement to run (e.g., "**1.9**").</span></span>

> [!WARNING]
> <span data-ttu-id="2bf73-230">调用 方法 `isSetSupported` 时，如果指定 (`MinimumVersion` 参数) 应为字符串。</span><span class="sxs-lookup"><span data-stu-id="2bf73-230">When calling the `isSetSupported` method, the value of the `MinimumVersion` parameter (if specified) should be a string.</span></span> <span data-ttu-id="2bf73-231">这是因为 JavaScript 分析器无法区分数值，例如 1.1 和 1.10，因为它可以用于字符串值，例如“1.1”和“1.10”。</span><span class="sxs-lookup"><span data-stu-id="2bf73-231">This is because the JavaScript parser cannot differentiate between numeric values such as 1.1 and 1.10, where as it can for string values such as "1.1" and "1.10".</span></span>
> <span data-ttu-id="2bf73-232">`number` 重载已弃用。</span><span class="sxs-lookup"><span data-stu-id="2bf73-232">The `number` overload is deprecated.</span></span>

<span data-ttu-id="2bf73-233">与 `isSetSupported` 与 `RequirementSetName` 应用程序关联的 Office，如下所示。</span><span class="sxs-lookup"><span data-stu-id="2bf73-233">Use `isSetSupported` with the `RequirementSetName` associated with the Office application as follows.</span></span>

|<span data-ttu-id="2bf73-234">Office 应用程序</span><span class="sxs-lookup"><span data-stu-id="2bf73-234">Office application</span></span>|<span data-ttu-id="2bf73-235">RequirementSetName</span><span class="sxs-lookup"><span data-stu-id="2bf73-235">RequirementSetName</span></span>|
|---|---|
|<span data-ttu-id="2bf73-236">Excel</span><span class="sxs-lookup"><span data-stu-id="2bf73-236">Excel</span></span>|<span data-ttu-id="2bf73-237">ExcelApi</span><span class="sxs-lookup"><span data-stu-id="2bf73-237">ExcelApi</span></span>|
|<span data-ttu-id="2bf73-238">OneNote</span><span class="sxs-lookup"><span data-stu-id="2bf73-238">OneNote</span></span>|<span data-ttu-id="2bf73-239">OneNoteApi</span><span class="sxs-lookup"><span data-stu-id="2bf73-239">OneNoteApi</span></span>|
|<span data-ttu-id="2bf73-240">Outlook</span><span class="sxs-lookup"><span data-stu-id="2bf73-240">Outlook</span></span>|<span data-ttu-id="2bf73-241">Mailbox</span><span class="sxs-lookup"><span data-stu-id="2bf73-241">Mailbox</span></span>|
|<span data-ttu-id="2bf73-242">Word</span><span class="sxs-lookup"><span data-stu-id="2bf73-242">Word</span></span>|<span data-ttu-id="2bf73-243">WordApi</span><span class="sxs-lookup"><span data-stu-id="2bf73-243">WordApi</span></span>|

<span data-ttu-id="2bf73-244">这些 `isSetSupported` 应用程序的 方法和要求集可在 Office.js 上的最新 CDN。</span><span class="sxs-lookup"><span data-stu-id="2bf73-244">The `isSetSupported` method and the requirement sets for these applications are available in the latest Office.js file on the CDN.</span></span> <span data-ttu-id="2bf73-245">如果不从加载项Office.js，CDN可能会生成异常，因为 `isSetSupported` 将是未定义的。</span><span class="sxs-lookup"><span data-stu-id="2bf73-245">If you don't use Office.js from the CDN, your add-in might generate exceptions because `isSetSupported` will be undefined.</span></span> <span data-ttu-id="2bf73-246">有关详细信息，请参阅指定[JavaScript API Office的最新版本](#specify-the-latest-office-javascript-api-library)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-246">For more information, see [Specify the latest Office JavaScript API library](#specify-the-latest-office-javascript-api-library).</span></span>

<span data-ttu-id="2bf73-247">以下代码示例演示外接程序如何为可能支持不同要求集或 API 成员的不同Office应用程序提供不同的功能。</span><span class="sxs-lookup"><span data-stu-id="2bf73-247">The following code example shows how an add-in can provide different functionality for different Office applications that might support different requirement sets or API members.</span></span>

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

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="2bf73-248">使用不属于要求集的方法的运行时检查</span><span class="sxs-lookup"><span data-stu-id="2bf73-248">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="2bf73-249">部分 API 成员不属于要求集</span><span class="sxs-lookup"><span data-stu-id="2bf73-249">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="2bf73-250">这仅适用于[属于 Office JavaScript API](../reference/javascript-api-for-office.md)命名空间 (除 Outlook 邮箱 API) 之外的任何位置的 API 成员，但不包括属于) 中的 Word JavaScript API (任何内容 `Office.` [](/javascript/api/outlook)[](../reference/overview/word-add-ins-reference-overview.md) `Word.` [、Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) () 中的任何内容或 OneNote `Excel.` [JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) `OneNote.` () 命名空间中的任何内容。</span><span class="sxs-lookup"><span data-stu-id="2bf73-250">This only applies to API members that are part of the [Office JavaScript API](../reference/javascript-api-for-office.md) namespace (anything under `Office.` except [Outlook Mailbox APIs](/javascript/api/outlook)), but not API members that belong to the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (anything in `Word.`), [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (anything in `Excel.`), or [OneNote JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (anything in `OneNote.`) namespaces.</span></span> <span data-ttu-id="2bf73-251">当加载项依赖于不是要求集一部分的方法时，可以使用运行时检查来确定 Office 应用程序是否支持该方法，如下面的代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="2bf73-251">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office application, as shown in the following code example.</span></span> <span data-ttu-id="2bf73-252">有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)。</span><span class="sxs-lookup"><span data-stu-id="2bf73-252">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).</span></span>

> [!NOTE]
> <span data-ttu-id="2bf73-253">建议限制在加载项代码中使用此类型运行时检查。</span><span class="sxs-lookup"><span data-stu-id="2bf73-253">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="2bf73-254">下面的代码示例检查应用程序Office是否 `document.setSelectedDataAsync` 支持 。</span><span class="sxs-lookup"><span data-stu-id="2bf73-254">The following code example checks whether the Office application supports `document.setSelectedDataAsync`.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses `document.setSelectedDataAsync`.
}
```


## <a name="see-also"></a><span data-ttu-id="2bf73-255">另请参阅</span><span class="sxs-lookup"><span data-stu-id="2bf73-255">See also</span></span>

- [<span data-ttu-id="2bf73-256">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="2bf73-256">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="2bf73-257">Office 加载项要求集</span><span class="sxs-lookup"><span data-stu-id="2bf73-257">Office Add-in requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="2bf73-258">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="2bf73-258">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
