---
title: 指定 Office 主机和 API 要求
description: 了解如何指定你的外接程序按预期工作的 Office 主机和 API 要求。
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: ab9b97f3d3232339010179097e1fd03dbeb86aa2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718809"
---
# <a name="specify-office-hosts-and-api-requirements"></a><span data-ttu-id="8501a-103">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="8501a-103">Specify Office hosts and API requirements</span></span>

<span data-ttu-id="8501a-p101">你的 Office外接程序可能依赖于特定的 Office 主机、要求集、API 成员或 API 版本才能按预期运行。例如，你的外接程序可能：</span><span class="sxs-lookup"><span data-stu-id="8501a-p101">Your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API in order to work as expected. For example, your add-in might:</span></span>

- <span data-ttu-id="8501a-106">在单个 Office 应用程序（如，Word 或 Excel）或多个应用程序中运行。</span><span class="sxs-lookup"><span data-stu-id="8501a-106">Run in a single Office application (e.g., Word or Excel), or several applications.</span></span>

- <span data-ttu-id="8501a-p102">使用仅在 Office 的某些版本中可用的 JavaScript API。例如，可能会在运行在 Excel 2016 中的外接程序中使用适用于 Excel 的 JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="8501a-p102">Make use of JavaScript APIs that are only available in some versions of Office. For example, you might use the Excel JavaScript APIs in an add-in that runs in Excel 2016.</span></span>

- <span data-ttu-id="8501a-109">只能在 Office 的某些版本中运行，这些版本支持供你的外接程序使用的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="8501a-109">Run only in versions of Office that support API members that your add-in uses.</span></span>

<span data-ttu-id="8501a-110">本文可帮助你了解应选择的选项，以确保你的外接程序按预期运行，并遍及可能的最广泛的访问群体。</span><span class="sxs-lookup"><span data-stu-id="8501a-110">This article helps you understand which options you should choose to ensure that your add-in works as expected and reaches the broadest audience possible.</span></span>

> [!NOTE]
> <span data-ttu-id="8501a-111">若要概览 Office 加载项的当前受支持情况，请参阅 [Office 加载项主机和平台可用性](../overview/office-add-in-availability.md)页面。</span><span class="sxs-lookup"><span data-stu-id="8501a-111">For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>

<span data-ttu-id="8501a-112">下表列出了本文中讨论的核心概念。</span><span class="sxs-lookup"><span data-stu-id="8501a-112">The following table lists core concepts discussed throughout this article.</span></span>

|<span data-ttu-id="8501a-113">**概念**</span><span class="sxs-lookup"><span data-stu-id="8501a-113">**Concept**</span></span>|<span data-ttu-id="8501a-114">**说明**</span><span class="sxs-lookup"><span data-stu-id="8501a-114">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="8501a-115">Office 应用程序、Office 主机应用程序、Office 主机或主机</span><span class="sxs-lookup"><span data-stu-id="8501a-115">Office application, Office host application, Office host, or host</span></span>|<span data-ttu-id="8501a-p103">用于运行加载项的 Office 应用程序。例如 Word、Excel 等。</span><span class="sxs-lookup"><span data-stu-id="8501a-p103">The Office application used to run your add-in. For example, Word, Excel, and so on.</span></span>|
|<span data-ttu-id="8501a-118">平台</span><span class="sxs-lookup"><span data-stu-id="8501a-118">Platform</span></span>|<span data-ttu-id="8501a-119">运行 Office 主机的位置，例如在浏览器或 iPad 中。</span><span class="sxs-lookup"><span data-stu-id="8501a-119">Where the Office host runs, such as in a browser or on an iPad.</span></span>|
|<span data-ttu-id="8501a-120">要求集</span><span class="sxs-lookup"><span data-stu-id="8501a-120">Requirement set</span></span>|<span data-ttu-id="8501a-p104">命名的一组相关的 API 成员。外接程序使用要求集来确定 Office 主机是否支持你的外接程序使用的 API 成员。测试对要求集的支持比对单个的 API 成员的支持更为容易。要求集支持根据 Office 主机和 Office 主机的版本变化。 </span><span class="sxs-lookup"><span data-stu-id="8501a-p104">A named group of related API members. Add-ins use requirement sets to determine whether the Office host supports API members used by your add-in. It's easier to test for the support of a requirement set than for the support of individual API members. Requirement set support varies by Office host and the version of the Office host. </span></span><br ><span data-ttu-id="8501a-125">要求集在清单文件中指定。</span><span class="sxs-lookup"><span data-stu-id="8501a-125">Requirement sets are specified in the manifest file.</span></span> <span data-ttu-id="8501a-126">当你在清单中指定要求集时，你可以设置 Office 主机必须提供的用于运行你的外接程序的最低级别的 API 支持。</span><span class="sxs-lookup"><span data-stu-id="8501a-126">When you specify requirement sets in the manifest, you set the minimum level of API support that the Office host must provide in order to run your add-in.</span></span> <span data-ttu-id="8501a-127">不支持在清单中指定的要求集的 Office 主机不能运行加载项，并且加载项不会显示在“<span class="ui">我的加载项</span>”中。这限制了加载项的使用位置。</span><span class="sxs-lookup"><span data-stu-id="8501a-127">Office hosts that don't support requirement sets specified in the manifest can't run your add-in, and your add-in won't display in <span class="ui">My Add-ins</span>. This restricts where your add-in is available.</span></span> <span data-ttu-id="8501a-128">在使用运行时检查的代码中。</span><span class="sxs-lookup"><span data-stu-id="8501a-128">In code using runtime checks.</span></span> <span data-ttu-id="8501a-129">有关要求集的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="8501a-129">For the complete list of requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>|
|<span data-ttu-id="8501a-130">运行时检查</span><span class="sxs-lookup"><span data-stu-id="8501a-130">Runtime check</span></span>|<span data-ttu-id="8501a-131">在运行时执行的一种测试，用以确定运行加载项的 Office 主机是否支持要求集或加载项使用的方法。</span><span class="sxs-lookup"><span data-stu-id="8501a-131">A test that is performed at runtime to determine whether the Office host running your add-in supports requirement sets or methods used by your add-in.</span></span> <span data-ttu-id="8501a-132">若要执行运行时检查，请将**if**语句与`isSetSupported`方法、要求集或不属于要求集的方法名称一起使用。</span><span class="sxs-lookup"><span data-stu-id="8501a-132">To perform a runtime check, you use an **if** statement with the `isSetSupported` method, the requirement sets, or the method names that aren't part of a requirement set.</span></span> <span data-ttu-id="8501a-133">使用运行时检查可确保加载项能够覆盖最大数量的客户。</span><span class="sxs-lookup"><span data-stu-id="8501a-133">Use runtime checks to ensure that your add-in reaches the broadest number of customers.</span></span> <span data-ttu-id="8501a-134">与要求集不同，运行时检查不指定 Office 主机必须提供的用于运行加载项的最低级别的 API 支持。</span><span class="sxs-lookup"><span data-stu-id="8501a-134">Unlike requirement sets, runtime checks don't specify the minimum level of API support that the Office host must provide for your add-in to run.</span></span> <span data-ttu-id="8501a-135">而是使用**if**语句来确定是否支持 API 成员。</span><span class="sxs-lookup"><span data-stu-id="8501a-135">Instead, you use the **if** statement to determine whether an API member is supported.</span></span> <span data-ttu-id="8501a-136">如果支持，则可以在外接程序中提供其他功能。</span><span class="sxs-lookup"><span data-stu-id="8501a-136">If it is, you can provide additional functionality in your add-in.</span></span> <span data-ttu-id="8501a-137">使用运行时检查时，你的外接程序将始终在“**我的外接程序**”中显示。</span><span class="sxs-lookup"><span data-stu-id="8501a-137">Your add-in will always display in **My Add-ins** when you use runtime checks.</span></span>|

## <a name="before-you-begin"></a><span data-ttu-id="8501a-138">开始之前</span><span class="sxs-lookup"><span data-stu-id="8501a-138">Before you begin</span></span>

<span data-ttu-id="8501a-139">您的外接程序必须使用最新版本的外接程序清单架构。</span><span class="sxs-lookup"><span data-stu-id="8501a-139">Your add-in must use the most current version of the add-in manifest schema.</span></span> <span data-ttu-id="8501a-140">如果您在外接程序中使用运行时检查，请确保使用最新的 Office JavaScript API （node.js）库。</span><span class="sxs-lookup"><span data-stu-id="8501a-140">If you use runtime checks in your add-in, ensure that you use the latest Office JavaScript API (office.js) library.</span></span>

### <a name="specify-the-latest-add-in-manifest-schema"></a><span data-ttu-id="8501a-141">指定最新的外接程序清单架构</span><span class="sxs-lookup"><span data-stu-id="8501a-141">Specify the latest add-in manifest schema</span></span>

<span data-ttu-id="8501a-142">外接程序清单必须使用外接程序清单架构版本 1.1。</span><span class="sxs-lookup"><span data-stu-id="8501a-142">Your add-in's manifest must use version 1.1 of the add-in manifest schema.</span></span> <span data-ttu-id="8501a-143">按如下`OfficeApp`方式设置外接程序清单中的元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-143">Set the `OfficeApp` element in your add-in manifest as follows.</span></span>

```XML
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
```

### <a name="specify-the-latest-office-javascript-api-library"></a><span data-ttu-id="8501a-144">指定最新的 Office JavaScript API 库</span><span class="sxs-lookup"><span data-stu-id="8501a-144">Specify the latest Office JavaScript API library</span></span>

<span data-ttu-id="8501a-145">如果您使用运行时检查，请参考内容传送网络（CDN）中的 Office JavaScript API 库的最新版本。</span><span class="sxs-lookup"><span data-stu-id="8501a-145">If you use runtime checks, reference the most current version of the Office JavaScript API library from the content delivery network (CDN).</span></span> <span data-ttu-id="8501a-146">若要执行此操作，请将以下 `script` 标记添加到 HTML 中。</span><span class="sxs-lookup"><span data-stu-id="8501a-146">To do this, add the following  `script` tag to your HTML.</span></span> <span data-ttu-id="8501a-147">使用 CDN URL 中的 `/1/` 可以确保你引用的是最新版本的 Office.js。</span><span class="sxs-lookup"><span data-stu-id="8501a-147">Using `/1/` in the CDN URL ensures that you reference the most recent version of Office.js.</span></span>

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

## <a name="options-to-specify-office-hosts-or-api-requirements"></a><span data-ttu-id="8501a-148">指定 Office 主机或 API 要求的选项</span><span class="sxs-lookup"><span data-stu-id="8501a-148">Options to specify Office hosts or API requirements</span></span>

<span data-ttu-id="8501a-p110">指定 Office 主机或 API 要求时，有几个决策因素需要考虑。下图显示了如何确定要在外接程序中使用的技术。</span><span class="sxs-lookup"><span data-stu-id="8501a-p110">When you specify Office hosts or API requirements, there are several factors to consider. The following diagram shows how to decide which technique to use in your add-in.</span></span>

![指定 Office 主机或 API 要求时，选择最适用于加载项的选项](../images/options-for-office-hosts.png)

- <span data-ttu-id="8501a-152">如果你的外接程序在一个 Office 主机中运行， `Hosts`请在清单中设置该元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-152">If your add-in runs in one Office host, set the `Hosts` element in the manifest.</span></span> <span data-ttu-id="8501a-153">有关详细信息，请参阅 [设置 Hosts 元素](#set-the-hosts-element)。</span><span class="sxs-lookup"><span data-stu-id="8501a-153">For more information, see [Set the Hosts element](#set-the-hosts-element).</span></span>

- <span data-ttu-id="8501a-154">若要设置 Office 主机必须支持的最低要求集或 API 成员以运行您的外接程序，请在`Requirements`清单中设置该元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-154">To set the minimum requirement set or API members that an Office host must support to run your add-in, set the `Requirements` element in the manifest.</span></span> <span data-ttu-id="8501a-155">有关详细信息，请参阅[在清单中设置 Requirements 元素](#set-the-requirements-element-in-the-manifest)。</span><span class="sxs-lookup"><span data-stu-id="8501a-155">For more information, see [Set the Requirements element in the manifest](#set-the-requirements-element-in-the-manifest).</span></span>

- <span data-ttu-id="8501a-156">如果特定要求集或 API 成员可在 Office 主机中使用，在这种情况下如果你想要提供其他功能，请在外接程序的 JavaScript 代码中执行运行时检查。</span><span class="sxs-lookup"><span data-stu-id="8501a-156">If you would like to provide additional functionality if specific requirement sets or API members are available in the Office host, perform a runtime check in your add-in's JavaScript code.</span></span> <span data-ttu-id="8501a-157">例如，如果加载项在 Excel 2016 中运行，请使用 Excel JavaScript API 中的 API 成员提供附加功能。</span><span class="sxs-lookup"><span data-stu-id="8501a-157">For example, if your add-in runs in Excel 2016, use API members from the Excel JavaScript API to provide additional functionality.</span></span> <span data-ttu-id="8501a-158">有关详细信息，请参阅[在你的 JavaScript 代码中使用运行时检查](#use-runtime-checks-in-your-javascript-code)。</span><span class="sxs-lookup"><span data-stu-id="8501a-158">For more information, see [Use runtime checks in your JavaScript code](#use-runtime-checks-in-your-javascript-code).</span></span>

## <a name="set-the-hosts-element"></a><span data-ttu-id="8501a-159">设置 Hosts 元素</span><span class="sxs-lookup"><span data-stu-id="8501a-159">Set the Hosts element</span></span>

<span data-ttu-id="8501a-160">若要使您的外接程序在一个 Office 主机应用程序中`Hosts`运行`Host` ，请使用清单中的和元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-160">To make your add-in run in one Office host application, use the `Hosts` and `Host` elements in the manifest.</span></span> <span data-ttu-id="8501a-161">如果不指定`Hosts`元素，则外接程序将在所有主机中运行。</span><span class="sxs-lookup"><span data-stu-id="8501a-161">If you don't specify the `Hosts` element, your add-in will run in all hosts.</span></span>

<span data-ttu-id="8501a-162">例如，以下`Hosts`和`Host`声明指定外接程序将使用任何版本的 excel，其中包括在 Web、Windows 和 iPad 上的 excel。</span><span class="sxs-lookup"><span data-stu-id="8501a-162">For example, the following `Hosts` and `Host` declaration specifies that the add-in will work with any release of Excel, which includes Excel on the web, Windows, and iPad.</span></span>

```xml
<Hosts>
  <Host Name="Workbook" />
</Hosts>
```

<span data-ttu-id="8501a-163">`Hosts`元素可以包含一个或多个`Host`元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-163">The `Hosts` element can contain one or more `Host` elements.</span></span> <span data-ttu-id="8501a-164">`Host`元素指定你的外接程序所需的 Office 主机。</span><span class="sxs-lookup"><span data-stu-id="8501a-164">The `Host` element specifies the Office host your add-in requires.</span></span> <span data-ttu-id="8501a-165">`Name`属性是必需的，并且可设置为下列值之一。</span><span class="sxs-lookup"><span data-stu-id="8501a-165">The `Name` attribute is required and can be set to one of the following values.</span></span>

| <span data-ttu-id="8501a-166">名称</span><span class="sxs-lookup"><span data-stu-id="8501a-166">Name</span></span>          | <span data-ttu-id="8501a-167">Office 主机应用程序</span><span class="sxs-lookup"><span data-stu-id="8501a-167">Office host applications</span></span>                                                                  |
|:--------------|:------------------------------------------------------------------------------------------|
| <span data-ttu-id="8501a-168">数据库</span><span class="sxs-lookup"><span data-stu-id="8501a-168">Database</span></span>      | <span data-ttu-id="8501a-169">Access Web App</span><span class="sxs-lookup"><span data-stu-id="8501a-169">Access web apps</span></span>                                                                           |
| <span data-ttu-id="8501a-170">文档</span><span class="sxs-lookup"><span data-stu-id="8501a-170">Document</span></span>      | <span data-ttu-id="8501a-171">Windows 版 Word、Mac 版 Word、iPad 版 Word、Word 网页版</span><span class="sxs-lookup"><span data-stu-id="8501a-171">Word on Windows, Word on Mac, Word on iPad, Word on the web</span></span>                               |
| <span data-ttu-id="8501a-172">邮箱</span><span class="sxs-lookup"><span data-stu-id="8501a-172">Mailbox</span></span>       | <span data-ttu-id="8501a-173">Windows 版 Outlook、Mac 版 Outlook、Outlook 网页版、Android 版 Outlook 和 iOS 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="8501a-173">Outlook on Windows, Outlook on Mac, Outlook on the web, Outlook on Android, Outlook on iOS</span></span>|
| <span data-ttu-id="8501a-174">演示文稿</span><span class="sxs-lookup"><span data-stu-id="8501a-174">Presentation</span></span>  | <span data-ttu-id="8501a-175">Windows 版 PowerPoint、Mac 版 PowerPoint、iPad 版 PowerPoint、PowerPoint 网页版</span><span class="sxs-lookup"><span data-stu-id="8501a-175">PowerPoint on Windows, PowerPoint on Mac, PowerPoint on iPad, PowerPoint on the web</span></span>       |
| <span data-ttu-id="8501a-176">项目</span><span class="sxs-lookup"><span data-stu-id="8501a-176">Project</span></span>       | <span data-ttu-id="8501a-177">Windows 版 Project</span><span class="sxs-lookup"><span data-stu-id="8501a-177">Project on Windows</span></span>                                                                        |
| <span data-ttu-id="8501a-178">工作簿</span><span class="sxs-lookup"><span data-stu-id="8501a-178">Workbook</span></span>      | <span data-ttu-id="8501a-179">Windows 版 Excel、Mac 版 Excel、iPad 版 Excel、Excel 网页版</span><span class="sxs-lookup"><span data-stu-id="8501a-179">Excel on Windows, Excel on Mac, Excel on iPad, Excel on the web</span></span>                           |

> [!NOTE]
> <span data-ttu-id="8501a-180">`Name`属性指定可以运行外接程序的 Office 主机应用程序。</span><span class="sxs-lookup"><span data-stu-id="8501a-180">The `Name` attribute specifies the Office host application that can run your add-in.</span></span> <span data-ttu-id="8501a-181">Office 主机支持不同的平台，且可在台式机、Web 浏览器、平板电脑和移动设备上运行。</span><span class="sxs-lookup"><span data-stu-id="8501a-181">Office hosts are supported on different platforms and run on desktops, web browsers, tablets, and mobile devices.</span></span> <span data-ttu-id="8501a-182">不能指定用于运行外接程序的平台。</span><span class="sxs-lookup"><span data-stu-id="8501a-182">You can't specify which platform can be used to run your add-in.</span></span> <span data-ttu-id="8501a-183">例如，如果你指定 `Mailbox`，则 Windows 版 Outlook 和 Outlook 网页版都可以用来运行你的加载项。</span><span class="sxs-lookup"><span data-stu-id="8501a-183">For example, if you specify `Mailbox`, both Outlook on Windows and on the web can be used to run your add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8501a-184">我们不建议在 SharePoint 中创建和使用 Access Web 应用和数据库。</span><span class="sxs-lookup"><span data-stu-id="8501a-184">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="8501a-185">作为一种替代方法，我们建议你使用 [Microsoft PowerApps](https://powerapps.microsoft.com/) 生成适用于 Web 和移动设备的无代码业务解决方案。</span><span class="sxs-lookup"><span data-stu-id="8501a-185">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>


## <a name="set-the-requirements-element-in-the-manifest"></a><span data-ttu-id="8501a-186">在清单中设置 Requirements 元素</span><span class="sxs-lookup"><span data-stu-id="8501a-186">Set the Requirements element in the manifest</span></span>

<span data-ttu-id="8501a-187">`Requirements`元素指定 Office 主机运行外接程序时必须支持的最低要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="8501a-187">The `Requirements` element specifies the minimum requirement sets or API members that must be supported by the Office host to run your add-in.</span></span> <span data-ttu-id="8501a-188">`Requirements`元素可以指定要求集和外接程序中使用的各个方法。</span><span class="sxs-lookup"><span data-stu-id="8501a-188">The `Requirements` element can specify both requirement sets and individual methods used in your add-in.</span></span> <span data-ttu-id="8501a-189">在外接程序清单架构的版本1.1 中，除`Requirements` Outlook 外接程序外接程序外接程序中，该元素是可选的。</span><span class="sxs-lookup"><span data-stu-id="8501a-189">In version 1.1 of the add-in manifest schema, the `Requirements` element is optional for all add-ins, except for Outlook add-ins.</span></span>

> [!WARNING]
> <span data-ttu-id="8501a-190">仅使用`Requirements`元素指定你的外接程序必须使用的关键要求集或 API 成员。</span><span class="sxs-lookup"><span data-stu-id="8501a-190">Only use the `Requirements` element to specify critical requirement sets or API members that your add-in must use.</span></span> <span data-ttu-id="8501a-191">如果 Office 主机或平台不支持`Requirements`元素中指定的要求集或 API 成员，则外接程序将不会在该主机或平台中运行，并且不会显示在我的**外接程序**中。相反，我们建议您让外接程序在 Office 主机的所有平台上可用，如 web、Windows 和 iPad 上的 Excel。</span><span class="sxs-lookup"><span data-stu-id="8501a-191">If the Office host or platform doesn't support the requirement sets or API members specified in the `Requirements` element, the add-in won't run in that host or platform, and won't display in **My Add-ins**. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel on the web, Windows, and iPad.</span></span> <span data-ttu-id="8501a-192">若要使您的外接程序在_所有_Office 主机和平台上可用，请使用运行`Requirements`时检查而不是元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-192">To make your add-in available on  _all_ Office hosts and platforms, use runtime checks instead of the `Requirements` element.</span></span>

<span data-ttu-id="8501a-193">以下代码示例说明在支持以下内容的所有 Office 主机应用程序中加载的外接程序：</span><span class="sxs-lookup"><span data-stu-id="8501a-193">The following code example shows an add-in that loads in all Office host applications that support the following:</span></span>

-  <span data-ttu-id="8501a-194">`TableBindings`要求集，其最低版本为 "1.1"。</span><span class="sxs-lookup"><span data-stu-id="8501a-194">`TableBindings` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="8501a-195">`OOXML`要求集，其最低版本为 "1.1"。</span><span class="sxs-lookup"><span data-stu-id="8501a-195">`OOXML` requirement set, which has a minimum version of "1.1".</span></span>

-  <span data-ttu-id="8501a-196">`Document.getSelectedDataAsync`种.</span><span class="sxs-lookup"><span data-stu-id="8501a-196">`Document.getSelectedDataAsync` method.</span></span>

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

- <span data-ttu-id="8501a-197">`Requirements`元素包含`Sets`和`Methods`子元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-197">The `Requirements` element contains the `Sets` and `Methods` child elements.</span></span>

- <span data-ttu-id="8501a-198">`Sets`元素可以包含一个或多个`Set`元素。</span><span class="sxs-lookup"><span data-stu-id="8501a-198">The `Sets` element can contain one or more `Set` elements.</span></span> <span data-ttu-id="8501a-199">`DefaultMinVersion`指定所有子`MinVersion` `Set`元素的默认值。</span><span class="sxs-lookup"><span data-stu-id="8501a-199">`DefaultMinVersion` specifies the default `MinVersion` value of all child `Set` elements.</span></span>

- <span data-ttu-id="8501a-200">`Set`元素指定 Office 主机必须支持的要求集以运行外接程序。</span><span class="sxs-lookup"><span data-stu-id="8501a-200">The `Set` element specifies requirement sets that the Office host must support to run the add-in.</span></span> <span data-ttu-id="8501a-201">`Name`属性指定要求集的名称。</span><span class="sxs-lookup"><span data-stu-id="8501a-201">The `Name` attribute specifies the name of the requirement set.</span></span> <span data-ttu-id="8501a-202">`MinVersion`指定要求集的最低版本。</span><span class="sxs-lookup"><span data-stu-id="8501a-202">The `MinVersion` specifies the minimum version of the requirement set.</span></span> <span data-ttu-id="8501a-203">`MinVersion`重写的值`DefaultMinVersion`有关您的 API 成员所属的要求集和要求集版本的详细信息，请参阅[Office 外接程序要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="8501a-203">`MinVersion` overrides the value of `DefaultMinVersion` For more information about requirement sets and requirement set versions that your API members belong to, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>

- <span data-ttu-id="8501a-p122">`Methods`元素可以包含一个或多个`Method`元素。不能将`Methods`元素与 Outlook 外接程序一起使用。</span><span class="sxs-lookup"><span data-stu-id="8501a-p122">The `Methods` element can contain one or more `Method` elements. You can't use the `Methods` element with Outlook add-ins.</span></span>

- <span data-ttu-id="8501a-p123">`Method`元素指定在运行外接程序的 Office 主机中必须支持的单个方法。属性`Name`是必需的，并指定通过其父对象限定的方法的名称。</span><span class="sxs-lookup"><span data-stu-id="8501a-p123">The `Method` element specifies an individual method that must be supported in the Office host where your add-in runs. The `Name` attribute is required and specifies the name of the method qualified with its parent object.</span></span>

## <a name="use-runtime-checks-in-your-javascript-code"></a><span data-ttu-id="8501a-208">在你的 JavaScript 代码中使用运行时检查</span><span class="sxs-lookup"><span data-stu-id="8501a-208">Use runtime checks in your JavaScript code</span></span>

<span data-ttu-id="8501a-209">如果 Office 主机支持某些要求集，你可能想要在你的外接程序中提供其他功能。</span><span class="sxs-lookup"><span data-stu-id="8501a-209">You might want to provide additional functionality in your add-in if certain requirement sets are supported by the Office host.</span></span> <span data-ttu-id="8501a-210">例如，如果你的加载项在 Word 2016 中运行，则你可能想要在现有的加载项中使用 Word JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="8501a-210">For example, you might want to use the Word JavaScript APIs in your existing add-in if your add-in runs in Word 2016.</span></span> <span data-ttu-id="8501a-211">若要执行此操作，你可以使用含有要求集名称的 [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) 方法。</span><span class="sxs-lookup"><span data-stu-id="8501a-211">To do this, you use the [isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) method with the name of the requirement set.</span></span> <span data-ttu-id="8501a-212">`isSetSupported`在运行时确定运行加载项的 Office 主机是否支持要求集。</span><span class="sxs-lookup"><span data-stu-id="8501a-212">`isSetSupported` determines, at runtime, whether the Office host running the add-in supports the requirement set.</span></span> <span data-ttu-id="8501a-213">如果支持该要求集， `isSetSupported`则返回**true** ，并运行使用该要求集的 API 成员的其他代码。</span><span class="sxs-lookup"><span data-stu-id="8501a-213">If the requirement set is supported, `isSetSupported` returns **true** and runs the additional code that uses the API members from that requirement set.</span></span> <span data-ttu-id="8501a-214">如果 Office 主机不支持要求集， `isSetSupported`将返回**false** ，并且不会运行其他代码。</span><span class="sxs-lookup"><span data-stu-id="8501a-214">If the Office host doesn't support the requirement set, `isSetSupported` returns **false** and the additional code won't run.</span></span> <span data-ttu-id="8501a-215">下面的代码演示与一起`isSetSupported`使用的语法。</span><span class="sxs-lookup"><span data-stu-id="8501a-215">The following code shows the syntax to use with `isSetSupported`.</span></span>

```js
if (Office.context.requirements.isSetSupported(RequirementSetName, MinimumVersion))
{
   // Code that uses API members from RequirementSetName.
}

```

- <span data-ttu-id="8501a-216">_RequirementSetName_（必填）是代表该要求集名称的字符串（例如，“**ExcelApi**”、“**Mailbox**”等）。</span><span class="sxs-lookup"><span data-stu-id="8501a-216">_RequirementSetName_ (required) is a string that represents the name of the requirement set (e.g., "**ExcelApi**", "**Mailbox**", etc.).</span></span> <span data-ttu-id="8501a-217">有关可用要求集的详细信息，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="8501a-217">For more information about available requirement sets, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md).</span></span>
- <span data-ttu-id="8501a-218">_MinimumVersion_（可选）是指定要求集的最低版本的字符串，主机必须支持该版本以便运行 `if` 语句中的代码（例如“**1.9**”）。</span><span class="sxs-lookup"><span data-stu-id="8501a-218">_MinimumVersion_ (optional) is a string that specifies the minimum requirement set version that the host must support in order for the code within the `if` statement to run (e.g., "**1.9**").</span></span>

> [!WARNING]
> <span data-ttu-id="8501a-219">调用`isSetSupported`方法时， `MinimumVersion`参数（如果指定）的值应为字符串。</span><span class="sxs-lookup"><span data-stu-id="8501a-219">When calling the `isSetSupported` method, the value of the `MinimumVersion` parameter (if specified) should be a string.</span></span> <span data-ttu-id="8501a-220">这是因为 JavaScript 分析器无法区分数值，例如 1.1 和 1.10，因为它可以用于字符串值，例如“1.1”和“1.10”。</span><span class="sxs-lookup"><span data-stu-id="8501a-220">This is because the JavaScript parser cannot differentiate between numeric values such as 1.1 and 1.10, where as it can for string values such as "1.1" and "1.10".</span></span>
> <span data-ttu-id="8501a-221">`number` 重载已弃用。</span><span class="sxs-lookup"><span data-stu-id="8501a-221">The `number` overload is deprecated.</span></span>

<span data-ttu-id="8501a-222">与`isSetSupported` Office 主机`RequirementSetName`关联使用，如下所示。</span><span class="sxs-lookup"><span data-stu-id="8501a-222">Use `isSetSupported` with the `RequirementSetName` associated with the Office host as follows.</span></span>

|<span data-ttu-id="8501a-223">Office 主机</span><span class="sxs-lookup"><span data-stu-id="8501a-223">Office host</span></span>|<span data-ttu-id="8501a-224">RequirementSetName</span><span class="sxs-lookup"><span data-stu-id="8501a-224">RequirementSetName</span></span>|
|---|---|
|<span data-ttu-id="8501a-225">Excel</span><span class="sxs-lookup"><span data-stu-id="8501a-225">Excel</span></span>|<span data-ttu-id="8501a-226">ExcelApi</span><span class="sxs-lookup"><span data-stu-id="8501a-226">ExcelApi</span></span>|
|<span data-ttu-id="8501a-227">OneNote</span><span class="sxs-lookup"><span data-stu-id="8501a-227">OneNote</span></span>|<span data-ttu-id="8501a-228">OneNoteApi</span><span class="sxs-lookup"><span data-stu-id="8501a-228">OneNoteApi</span></span>|
|<span data-ttu-id="8501a-229">Outlook</span><span class="sxs-lookup"><span data-stu-id="8501a-229">Outlook</span></span>|<span data-ttu-id="8501a-230">Mailbox</span><span class="sxs-lookup"><span data-stu-id="8501a-230">Mailbox</span></span>|
|<span data-ttu-id="8501a-231">Word</span><span class="sxs-lookup"><span data-stu-id="8501a-231">Word</span></span>|<span data-ttu-id="8501a-232">WordApi</span><span class="sxs-lookup"><span data-stu-id="8501a-232">WordApi</span></span>|

<span data-ttu-id="8501a-233">在`isSetSupported` CDN 上的最新的 Office .js 文件中提供了这些主机的方法和要求集。</span><span class="sxs-lookup"><span data-stu-id="8501a-233">The `isSetSupported` method and the requirement sets for these hosts are available in the latest Office.js file on the CDN.</span></span> <span data-ttu-id="8501a-234">如果不从 CDN 中使用 node.js，则外接程序可能会生成异常，因为`isSetSupported`这将是不确定的。</span><span class="sxs-lookup"><span data-stu-id="8501a-234">If you don't use Office.js from the CDN, your add-in might generate exceptions because `isSetSupported` will be undefined.</span></span> <span data-ttu-id="8501a-235">有关详细信息，请参阅[指定最新的 Office JAVASCRIPT API 库](#specify-the-latest-office-javascript-api-library)。</span><span class="sxs-lookup"><span data-stu-id="8501a-235">For more information, see [Specify the latest Office JavaScript API library](#specify-the-latest-office-javascript-api-library).</span></span>

<span data-ttu-id="8501a-236">以下代码示例演示外接程序如何向支持不同要求集或 API 成员的不同 Office 主机提供不同功能。</span><span class="sxs-lookup"><span data-stu-id="8501a-236">The following code example shows how an add-in can provide different functionality for different Office hosts that might support different requirement sets or API members.</span></span>

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
    // Run additional code when the Office host is not Word 2016 or later and does not support the CustomXmlParts requirement set.
}

```

## <a name="runtime-checks-using-methods-not-in-a-requirement-set"></a><span data-ttu-id="8501a-237">使用不属于要求集的方法的运行时检查</span><span class="sxs-lookup"><span data-stu-id="8501a-237">Runtime checks using methods not in a requirement set</span></span>

<span data-ttu-id="8501a-238">部分 API 成员不属于要求集</span><span class="sxs-lookup"><span data-stu-id="8501a-238">Some API members don't belong to requirement sets.</span></span> <span data-ttu-id="8501a-239">这仅适用于属于[Office javascript api](../reference/javascript-api-for-office.md)命名空间的 api 成员（ `Office.` [Outlook 邮箱 api](/javascript/api/outlook)除外），而不是属于[Word JavaScript api](../reference/overview/word-add-ins-reference-overview.md)的 api 成员（in 中`Word.`的任何内容）、 [Excel JavaScript api](../reference/overview/excel-add-ins-reference-overview.md) （任何内容`Excel.`）或[OneNote JavaScript api](../reference/overview/onenote-add-ins-javascript-reference.md) （任何内容都在`OneNote.`）命名空间。</span><span class="sxs-lookup"><span data-stu-id="8501a-239">This only applies to API members that are part of the [Office JavaScript API](../reference/javascript-api-for-office.md) namespace (anything under `Office.` except [Outlook Mailbox APIs](/javascript/api/outlook)), but not API members that belong to the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) (anything in `Word.`), [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) (anything in `Excel.`), or [OneNote JavaScript API](../reference/overview/onenote-add-ins-javascript-reference.md) (anything in `OneNote.`) namespaces.</span></span> <span data-ttu-id="8501a-240">当外接程序依赖于某个不属于要求集的方法时，可以使用运行时检查来确定 Office 主机是否支持此方法，方法如以下代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="8501a-240">When your add-in depends on a method that is not part of a requirement set, you can use the runtime check to determine whether the method is supported by the Office host, as shown in the following code example.</span></span> <span data-ttu-id="8501a-241">有关不属于要求集的方法的完整列表，请参阅 [Office 加载项要求集](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set)。</span><span class="sxs-lookup"><span data-stu-id="8501a-241">For a complete list of methods that don't belong to a requirement set, see [Office Add-in requirement sets](../reference/requirement-sets/office-add-in-requirement-sets.md#methods-that-arent-part-of-a-requirement-set).</span></span>

> [!NOTE]
> <span data-ttu-id="8501a-242">建议限制在加载项代码中使用此类型运行时检查。</span><span class="sxs-lookup"><span data-stu-id="8501a-242">We recommend that you limit the use of this type of runtime check in your add-in's code.</span></span>

<span data-ttu-id="8501a-243">下面的代码示例检查主机是否支持`document.setSelectedDataAsync`。</span><span class="sxs-lookup"><span data-stu-id="8501a-243">The following code example checks whether the host supports `document.setSelectedDataAsync`.</span></span>

```js
if (Office.context.document.setSelectedDataAsync)
{
    // Run code that uses document.setSelectedDataAsync.
}
```


## <a name="see-also"></a><span data-ttu-id="8501a-244">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8501a-244">See also</span></span>

- [<span data-ttu-id="8501a-245">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="8501a-245">Office Add-ins XML manifest</span></span>](add-in-manifests.md)
- [<span data-ttu-id="8501a-246">Office 加载项要求集</span><span class="sxs-lookup"><span data-stu-id="8501a-246">Office Add-in requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="8501a-247">Word-Add-in-Get-Set-EditOpen-XML</span><span class="sxs-lookup"><span data-stu-id="8501a-247">Word-Add-in-Get-Set-EditOpen-XML</span></span>](https://github.com/OfficeDev/Word-Add-in-Get-Set-EditOpen-XML)
