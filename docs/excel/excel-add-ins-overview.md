---
title: Excel 加载项概述
description: ''
ms.date: 07/05/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 645011e7600240e7f4947e8f4495e55383839a42
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596541"
---
# <a name="excel-add-ins-overview"></a><span data-ttu-id="a0d5b-102">Excel 加载项概述</span><span class="sxs-lookup"><span data-stu-id="a0d5b-102">Excel add-ins overview</span></span>

<span data-ttu-id="a0d5b-p101">使用 Excel 加载项，可以跨多个平台（包括 Windows、Mac、iPad 和浏览器）扩展 Excel 应用程序功能。在工作簿内使用 Excel 加载项，可以：</span><span class="sxs-lookup"><span data-stu-id="a0d5b-p101">An Excel add-in allows you to extend Excel application functionality across multiple platforms including Windows, Mac, iPad, and in a browser. Use Excel add-ins within a workbook to:</span></span>

- <span data-ttu-id="a0d5b-105">与 Excel 对象交互、读取和写入 Excel 数据。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-105">Interact with Excel objects, read and write Excel data.</span></span>
- <span data-ttu-id="a0d5b-106">使用基于 Web 的任务窗格或内容窗格扩展功能</span><span class="sxs-lookup"><span data-stu-id="a0d5b-106">Extend functionality using web based task pane or content pane</span></span>
- <span data-ttu-id="a0d5b-107">添加自定义功能区按钮或上下文菜单项</span><span class="sxs-lookup"><span data-stu-id="a0d5b-107">Add custom ribbon buttons or contextual menu items</span></span>
- <span data-ttu-id="a0d5b-108">添加自定义函数</span><span class="sxs-lookup"><span data-stu-id="a0d5b-108">Add custom functions</span></span>
- <span data-ttu-id="a0d5b-109">使用对话框窗口提供更丰富的交互</span><span class="sxs-lookup"><span data-stu-id="a0d5b-109">Provide richer interaction using dialog window</span></span>

<span data-ttu-id="a0d5b-110">Office 加载项平台提供框架和 Office.js JavaScript API，使你能够创建和运行 Excel 加载项。通过使用 Office 加载项平台创建 Excel 加载项，可以获得以下好处：</span><span class="sxs-lookup"><span data-stu-id="a0d5b-110">The Office Add-ins platform provides the framework and Office.js JavaScript APIs that enable you to create and run Excel add-ins. By using the Office Add-ins platform to create your Excel add-in, you'll get the following benefits:</span></span>

* <span data-ttu-id="a0d5b-111">**跨平台支持**：Excel 加载项在 Office 网页版、Windows 版 Office、Mac 版 Office 和 iPad 版 Office中运行。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-111">**Cross-platform support**: Excel add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>
* <span data-ttu-id="a0d5b-112">**集中式部署**：管理员可以在整个组织内为用户快速而轻松地部署 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-112">**Centralized deployment**: Admins can quickly and easily deploy Excel add-ins to users throughout an organization.</span></span>
* <span data-ttu-id="a0d5b-113">**使用标准 Web 技术**：使用熟悉的 Web 技术（如 HTML、CSS 和 JavaScript）创建 Excel 加载项。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-113">**Use of standard web technology**: Create your Excel add-in using familiar web technologies such as HTML, CSS, and JavaScript.</span></span>
* <span data-ttu-id="a0d5b-114">**通过 AppSource 分发**：将 Excel 加载项发布到 [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d)，供广大受众使用。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-114">**Distribution via AppSource**: Share your Excel add-in with a broad audience by publishing it to [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=53245fad-fcbe-41f8-9f97-b0840264f97c&omexanonuid=4a0102fb-b31a-4b9f-9bb0-39d4cc6b789d).</span></span>

> [!NOTE]
> <span data-ttu-id="a0d5b-p102">Excel 加载项不同于 COM 和 VSTO 加载项，后者是旧 Office 集成解决方案，只能在 Windows 版 Office 上运行。与 COM 加载项不同的是，Excel 加载项不需要你在用户设备上，或在 Excel 中安装任何代码。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-p102">Excel add-ins are different from COM and VSTO add-ins, which are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Excel add-ins do not require you to install any code on a user's device, or within Excel.</span></span>

## <a name="components-of-an-excel-add-in"></a><span data-ttu-id="a0d5b-117">Excel 加载项的组件</span><span class="sxs-lookup"><span data-stu-id="a0d5b-117">Components of an Excel add-in</span></span>

<span data-ttu-id="a0d5b-118">Excel 加载项包括两个基本组件：Web 应用程序和称为“清单文件”的配置文件。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-118">An Excel add-in includes two basic components: a web application and a configuration file, called a manifest file.</span></span> 

<span data-ttu-id="a0d5b-p103">Web 应用程序使用 [Office JavaScript API](../reference/javascript-api-for-office.md) 与 Excel 中的对象进行交互，并且还有助于与在线资源进行交互。例如，加载项可以执行下列任意任务：</span><span class="sxs-lookup"><span data-stu-id="a0d5b-p103">The web application uses the [Office JavaScript API](../reference/javascript-api-for-office.md) to interact with objects in Excel, and can also facilitate interaction with online resources. For example, an add-in can perform any of the following tasks:</span></span>

* <span data-ttu-id="a0d5b-121">创建、读取、更新和删除工作簿中的数据（工作表、区域、表、图表、已命名项等）。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-121">Create, read, update, and delete data in the workbook (worksheets, ranges, tables, charts, named items, and more).</span></span>
* <span data-ttu-id="a0d5b-122">使用标准 OAuth 2.0 流通过在线服务执行用户身份验证。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-122">Perform user authorization with an online service by using the standard OAuth 2.0 flow.</span></span>
* <span data-ttu-id="a0d5b-123">向 Microsoft Graph 或任何其他 API 发出 API 请求。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-123">Issue API requests to Microsoft Graph or any other API.</span></span>

<span data-ttu-id="a0d5b-124">Web 应用程序可以托管在任何 Web 服务器上，并且可以使用客户端框架（如 Angular、React、jQuery）或服务器端技术（如 ASP.NET、Node.js、PHP）进行构建。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-124">The web application can be hosted on any web server, and can be built using client-side frameworks (such as Angular, React, jQuery) or server-side technologies (such as ASP.NET, Node.js, PHP).</span></span>

<span data-ttu-id="a0d5b-125">[清单](../develop/add-in-manifests.md)是 XML 配置文件，它定义加载项如何通过指定以下设置和功能与 Office 客户端集成：</span><span class="sxs-lookup"><span data-stu-id="a0d5b-125">The [manifest](../develop/add-in-manifests.md) is an XML configuration file that defines how the add-in integrates with Office clients by specifying settings and capabilities such as:</span></span>

* <span data-ttu-id="a0d5b-126">加载项 Web 应用程序的 URL。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-126">The URL of the add-in's web application.</span></span>
* <span data-ttu-id="a0d5b-127">加载项的显示名称、说明、ID、版本和默认区域设置。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-127">The add-in's display name, description, ID, version, and default locale.</span></span>
* <span data-ttu-id="a0d5b-128">如何将加载项与 Excel 集成，其中包括加载项创建的任何自定义 UI（功能区按钮、上下文菜单等）。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-128">How the add-in integrates with Excel, including any custom UI that the add-in creates (ribbon buttons, context menus, and so on).</span></span>
* <span data-ttu-id="a0d5b-129">加载项所需的权限，如对文档执行读取和写入操作。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-129">Permissions that the add-in requires, such as reading and writing to the document.</span></span>

<span data-ttu-id="a0d5b-130">若要让最终用户能够安装和使用 Excel 加载项，必须将它的清单发布到 AppSource 或加载项目录。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-130">To enable end users to install and use an Excel add-in, you must publish its manifest either to AppSource or to an add-ins catalog.</span></span> <span data-ttu-id="a0d5b-131">要详细了解如何发布到 AppSource，请参阅[将解决方案发布到 AppSource 和 Office 中](/office/dev/store/submit-to-appsource-via-partner-center)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-131">For details about publishing to AppSource, see [Make your solutions available in AppSource and within Office](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="capabilities-of-an-excel-add-in"></a><span data-ttu-id="a0d5b-132">Excel 加载项的功能</span><span class="sxs-lookup"><span data-stu-id="a0d5b-132">Capabilities of an Excel add-in</span></span>

<span data-ttu-id="a0d5b-133">除了能够与工作簿内容进行交互外，Excel 加载项还可以添加自定义功能区按钮或菜单命令、插入任务窗格、添加自定义函数、打开对话框，甚至还能在工作表中嵌入基于 Web 的丰富对象（如图表或交互式可视化效果）。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-133">In addition to interacting with the content in the workbook, Excel add-ins can add custom ribbon buttons or menu commands, insert task panes, add custom functions, open dialog boxes, and even embed rich, web-based objects such as charts or interactive visualizations within a worksheet.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="a0d5b-134">加载项命令</span><span class="sxs-lookup"><span data-stu-id="a0d5b-134">Add-in commands</span></span>

<span data-ttu-id="a0d5b-p105">加载项命令是能够扩展 Excel UI，并在加载项中启动操作的 UI 元素。加载项命令可用于在功能区中添加按钮，也可用于向 Excel 上下文菜单中添加项。选择加载项命令后，用户便启动操作，如运行 JavaScript 代码，或在任务窗格中显示加载项页面。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-p105">Add-in commands are UI elements that extend the Excel UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu in Excel. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane.</span></span> 

<span data-ttu-id="a0d5b-138">**加载项命令**</span><span class="sxs-lookup"><span data-stu-id="a0d5b-138">**Add-in commands**</span></span>

![Excel 中的加载项命令](../images/excel-add-in-commands-script-lab.png)

<span data-ttu-id="a0d5b-140">有关命令功能、受支持的平台和开发加载项命令第最佳做法的详细信息，请参阅[适用于 Excel、Word 和 Powerpoint 的加载项命令](../design/add-in-commands.md)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-140">For more information about command capabilities, supported platforms, and best practices for developing add-in commands, see [Add-in commands for Excel, Word, and PowerPoint](../design/add-in-commands.md).</span></span>

### <a name="task-panes"></a><span data-ttu-id="a0d5b-141">任务窗格</span><span class="sxs-lookup"><span data-stu-id="a0d5b-141">Task panes</span></span>

<span data-ttu-id="a0d5b-p106">任务窗格是接口图面，通常出现在 Excel 中窗口的右侧。使用任务窗格，用户可以访问接口控件，以运行代码来修改 Excel 文档，或显示数据源中的数据。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-p106">Task panes are interface surfaces that typically appear on the right side of the window within Excel. Task panes give users access to interface controls that run code to modify the Excel document or display data from a data source.</span></span> 

<span data-ttu-id="a0d5b-144">**任务窗格**</span><span class="sxs-lookup"><span data-stu-id="a0d5b-144">**Task pane**</span></span>

![Excel 中的任务窗格加载项](../images/excel-add-in-task-pane-insights.png)

<span data-ttu-id="a0d5b-146">有关任务窗格的详细信息，请参阅 [Office 加载项中的任务窗格](../design/task-pane-add-ins.md)。有关在 Excel 中实现任务窗格的示例，请参阅 [Excel 加载项 JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-146">For more information about task panes, see [Task panes in Office Add-ins](../design/task-pane-add-ins.md). For a sample that implements a task pane in Excel, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends).</span></span>

### <a name="custom-functions"></a><span data-ttu-id="a0d5b-147">自定义函数</span><span class="sxs-lookup"><span data-stu-id="a0d5b-147">Custom functions</span></span>

<span data-ttu-id="a0d5b-148">开发人员可以借助自定义函数向 Excel 添加新函数，方法是在 JavaScript 中将这些函数定义为加载项的一部分。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-148">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="a0d5b-149">Excel 中的用户可以访问自定义函数，就像他们访问 Excel 中的任何本机函数一样，比如 `SUM()`。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-149">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> 

<span data-ttu-id="a0d5b-150">**自定义函数**</span><span class="sxs-lookup"><span data-stu-id="a0d5b-150">**Custom function**</span></span>

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

<span data-ttu-id="a0d5b-151">有关自定义函数的详细信息，请参阅[在 Excel 中创建自定义函数](custom-functions-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-151">For more information about custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

### <a name="dialog-boxes"></a><span data-ttu-id="a0d5b-152">对话框</span><span class="sxs-lookup"><span data-stu-id="a0d5b-152">Dialog boxes</span></span>

<span data-ttu-id="a0d5b-p108">对话框是浮动在活动的 Excel 应用程序窗口之上的界面。 可以将对话框用于以下任务，如显示无法直接在任务窗格中打开的登录页、请求用户确认操作，或托管如果局限在任务窗格中可能过小的视频。 若要在 Excel 加载项中打开对话框，请使用[对话框 API](/javascript/api/office/office.ui)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-p108">Dialog boxes are surfaces that float above the active Excel application window. You can use dialog boxes for tasks such as displaying sign-in pages that can't be opened directly in a task pane, requesting that the user confirm an action, or hosting videos that might be too small if confined to a task pane. To open dialog boxes in your Excel add-in, use the [Dialog API](/javascript/api/office/office.ui).</span></span>

<span data-ttu-id="a0d5b-156">**对话框**</span><span class="sxs-lookup"><span data-stu-id="a0d5b-156">**Dialog box**</span></span>

![Excel 中的加载项对话框](../images/excel-add-in-dialog-choose-number.png)

<span data-ttu-id="a0d5b-158">有关对话框和对话框 API 的详细信息，请参阅 [Office 加载项中的对话框](../design/dialog-boxes.md)和[在 Office 加载项中使用对话框 API](../develop/dialog-api-in-office-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-158">For more information about dialog boxes and the Dialog API, see [Dialog boxes in Office Add-ins](../design/dialog-boxes.md) and [Use the Dialog API in your Office Add-ins](../develop/dialog-api-in-office-add-ins.md).</span></span>

### <a name="content-add-ins"></a><span data-ttu-id="a0d5b-159">内容加载项</span><span class="sxs-lookup"><span data-stu-id="a0d5b-159">Content add-ins</span></span>

<span data-ttu-id="a0d5b-p109">内容加载项是可以直接嵌入到 Excel 文档中的图面。 可以使用内容加载项在工作表中嵌入基于 Web 的丰富对象，如图表、数据可视化效果或媒体，或为用户提供对界面控件的访问权限，这些控件运行代码以修改 Excel 文档，或显示来自数据源的数据。 在你要将功能直接嵌入文档时，请使用内容加载项。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-p109">Content add-ins are surfaces that you can embed directly into Excel documents. You can use content add-ins to embed rich, web-based objects such as charts, data visualizations, or media into a worksheet or to give users access to interface controls that run code to modify the Excel document or display data from a data source. Use content add-ins when you want to embed functionality directly into the document.</span></span>

<span data-ttu-id="a0d5b-163">**内容加载项**</span><span class="sxs-lookup"><span data-stu-id="a0d5b-163">**Content add-in**</span></span>

![Excel 中的内容加载项](../images/excel-add-in-content-map.png)

<span data-ttu-id="a0d5b-165">有关内容加载项的详细信息，请参阅 [Office 内容加载项](../design/content-add-ins.md)。有关在 Excel 中实现内容加载项的示例，请参阅 GitHub 中的 [ Excel 内容加载项 Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-165">For more information about content add-ins, see [Content Office Add-ins](../design/content-add-ins.md). For a sample that implements a content add-in in Excel, see [Excel Content Add-in Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance) in GitHub.</span></span>

## <a name="javascript-apis-to-interact-with-workbook-content"></a><span data-ttu-id="a0d5b-166">要与工作簿内容交互的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="a0d5b-166">JavaScript APIs to interact with workbook content</span></span>

<span data-ttu-id="a0d5b-167">Excel 加载项通过使用 [Office JavaScript API](../reference/javascript-api-for-office.md) 与 Excel 中的对象进行交互，JavaScript API 包括两个 JavaScript 对象模型：</span><span class="sxs-lookup"><span data-stu-id="a0d5b-167">An Excel add-in interacts with objects in Excel by using the [Office JavaScript API](../reference/javascript-api-for-office.md), which includes two JavaScript object models:</span></span>

* <span data-ttu-id="a0d5b-168">**Excel JavaScript API**：[Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) 随 Office 2016 引入，提供强类型的 Excel 对象，可用于访问工作表、区域、表、图表等。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-168">**Excel JavaScript API**: Introduced with Office 2016, the [Excel JavaScript API](../reference/overview/excel-add-ins-reference-overview.md) provides strongly-typed Excel objects that you can use to access worksheets, ranges, tables, charts, and more.</span></span> 

* <span data-ttu-id="a0d5b-169">**通用 API**：通用 API 随 Office 2013 引入，可用于访问多种类型的 Office 应用程序中常见的 UI、对话框和客户端设置等功能。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-169">**Common API**: Introduced with Office 2013, the Common API enables you to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span> <span data-ttu-id="a0d5b-170">由于通用 API 确实为 Excel 交互提供了有限的功能，因此，如果加载项需要在 Excel 2013 上运行，则可以使用它。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-170">Because the Common API does provide limited functionality for Excel interaction, you can use it if your add-in needs to run on Excel 2013.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a0d5b-171">后续步骤</span><span class="sxs-lookup"><span data-stu-id="a0d5b-171">Next steps</span></span>

<span data-ttu-id="a0d5b-172">通过[创建第一个 Excel 加载项](../quickstarts/excel-quickstart-jquery.md)开始使用。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-172">Get started by [creating your first Excel add-in](../quickstarts/excel-quickstart-jquery.md).</span></span> <span data-ttu-id="a0d5b-173">接下来，请详细了解与生成 Excel 加载项有关的[核心概念](excel-add-ins-core-concepts.md)。</span><span class="sxs-lookup"><span data-stu-id="a0d5b-173">Then, learn about the [core concepts](excel-add-ins-core-concepts.md) of building Excel add-ins.</span></span>

## <a name="see-also"></a><span data-ttu-id="a0d5b-174">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a0d5b-174">See also</span></span>

- [<span data-ttu-id="a0d5b-175">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="a0d5b-175">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="a0d5b-176">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="a0d5b-176">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="a0d5b-177">Excel JavaScript API 基本编程概念</span><span class="sxs-lookup"><span data-stu-id="a0d5b-177">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="a0d5b-178">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="a0d5b-178">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)