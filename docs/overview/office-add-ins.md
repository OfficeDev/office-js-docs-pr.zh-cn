---
title: Office 加载项平台概述 | Microsoft Docs
description: 使用熟悉的 Web 技术，例如 HTML、CSS 和 JavaScript 来扩展 Word、Excel、PowerPoint、OneNote、Project 和 Outlook，并与其进行交互。
ms.date: 02/13/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 6b162a166bda0c988f5fbbaade3b0bef4b650984
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094069"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="6cef8-103">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="6cef8-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="6cef8-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span><span class="sxs-lookup"><span data-stu-id="6cef8-104">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents.</span></span> <span data-ttu-id="6cef8-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span><span class="sxs-lookup"><span data-stu-id="6cef8-105">With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook.</span></span> <span data-ttu-id="6cef8-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span><span class="sxs-lookup"><span data-stu-id="6cef8-106">Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.</span></span>

![Office 加载项可扩展性图像](../images/addins-overview.png)

<span data-ttu-id="6cef8-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span><span class="sxs-lookup"><span data-stu-id="6cef8-108">Office Add-ins can do almost anything a webpage can do inside a browser.</span></span> <span data-ttu-id="6cef8-109">Use the Office Add-ins platform to:</span><span class="sxs-lookup"><span data-stu-id="6cef8-109">Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="6cef8-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span><span class="sxs-lookup"><span data-stu-id="6cef8-110">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more.</span></span> <span data-ttu-id="6cef8-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span><span class="sxs-lookup"><span data-stu-id="6cef8-111">For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="6cef8-112">**新建可嵌入到 Office 文档的丰富、交互式对象** - 用户可添加到其自己的 Excel 电子表格和 PowerPoint 演示文稿的嵌入式地图、图表和交互式可视化效果。</span><span class="sxs-lookup"><span data-stu-id="6cef8-112">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="6cef8-113">Office 加载项与 COM 和 VSTO 加载项有何不同？</span><span class="sxs-lookup"><span data-stu-id="6cef8-113">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="6cef8-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span><span class="sxs-lookup"><span data-stu-id="6cef8-114">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows.</span></span> <span data-ttu-id="6cef8-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span><span class="sxs-lookup"><span data-stu-id="6cef8-115">Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client.</span></span> <span data-ttu-id="6cef8-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span><span class="sxs-lookup"><span data-stu-id="6cef8-116">For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI.</span></span> <span data-ttu-id="6cef8-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span><span class="sxs-lookup"><span data-stu-id="6cef8-117">When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

![使用 Office 加载项的理由的图像](../images/why.png)

<span data-ttu-id="6cef8-119">相较于使用 VBA、COM 或 VSTO 生成的加载项，Office 加载项提供以下优势：</span><span class="sxs-lookup"><span data-stu-id="6cef8-119">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="6cef8-120">Cross-platform support.</span><span class="sxs-lookup"><span data-stu-id="6cef8-120">Cross-platform support.</span></span> <span data-ttu-id="6cef8-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span><span class="sxs-lookup"><span data-stu-id="6cef8-121">Office Add-ins run in Office on the web, Windows, Mac, and iPad.</span></span>

- <span data-ttu-id="6cef8-122">Centralized deployment and distribution.</span><span class="sxs-lookup"><span data-stu-id="6cef8-122">Centralized deployment and distribution.</span></span> <span data-ttu-id="6cef8-123">Admins can deploy Office Add-ins centrally across an organization.</span><span class="sxs-lookup"><span data-stu-id="6cef8-123">Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="6cef8-124">Easy access via AppSource.</span><span class="sxs-lookup"><span data-stu-id="6cef8-124">Easy access via AppSource.</span></span> <span data-ttu-id="6cef8-125">You can make your solution available to a broad audience by submitting it to AppSource.</span><span class="sxs-lookup"><span data-stu-id="6cef8-125">You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="6cef8-126">Based on standard web technology.</span><span class="sxs-lookup"><span data-stu-id="6cef8-126">Based on standard web technology.</span></span> <span data-ttu-id="6cef8-127">You can use any library you like to build Office Add-ins.</span><span class="sxs-lookup"><span data-stu-id="6cef8-127">You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="6cef8-128">Office 外接程序的组件</span><span class="sxs-lookup"><span data-stu-id="6cef8-128">Components of an Office Add-in</span></span>

<span data-ttu-id="6cef8-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span><span class="sxs-lookup"><span data-stu-id="6cef8-129">An Office Add-in includes two basic components: an XML manifest file, and your own web application.</span></span> <span data-ttu-id="6cef8-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span><span class="sxs-lookup"><span data-stu-id="6cef8-130">The manifest defines various settings, including how your add-in integrates with Office clients.</span></span> <span data-ttu-id="6cef8-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="6cef8-131">Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

### <a name="manifest"></a><span data-ttu-id="6cef8-132">清单</span><span class="sxs-lookup"><span data-stu-id="6cef8-132">Manifest</span></span>

<span data-ttu-id="6cef8-133">清单是一个 XML 文件，它指定外接程序的设置和功能，例如：</span><span class="sxs-lookup"><span data-stu-id="6cef8-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="6cef8-134">外接程序的显示名称、说明、ID、版本和默认区域设置。</span><span class="sxs-lookup"><span data-stu-id="6cef8-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="6cef8-135">如何将外接程序与 Office 集成。</span><span class="sxs-lookup"><span data-stu-id="6cef8-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="6cef8-136">外接程序的权限级别和数据访问要求。</span><span class="sxs-lookup"><span data-stu-id="6cef8-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="6cef8-137">Web 应用</span><span class="sxs-lookup"><span data-stu-id="6cef8-137">Web app</span></span>

<span data-ttu-id="6cef8-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span><span class="sxs-lookup"><span data-stu-id="6cef8-138">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource.</span></span> <span data-ttu-id="6cef8-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span><span class="sxs-lookup"><span data-stu-id="6cef8-139">However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js).</span></span> <span data-ttu-id="6cef8-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span><span class="sxs-lookup"><span data-stu-id="6cef8-140">To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="6cef8-141">*图 2：Hello World Office 加载项的组件*</span><span class="sxs-lookup"><span data-stu-id="6cef8-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Hello World 加载项的组件](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="6cef8-143">扩展并与 Office 客户端交互</span><span class="sxs-lookup"><span data-stu-id="6cef8-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="6cef8-144">Office 外接程序可以在 Office 主机应用程序中执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="6cef8-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="6cef8-145">扩展功能（任何 Office 应用程序）</span><span class="sxs-lookup"><span data-stu-id="6cef8-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="6cef8-146">创建新的对象（Excel 或 PowerPoint）</span><span class="sxs-lookup"><span data-stu-id="6cef8-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="6cef8-147">扩展 Office 功能</span><span class="sxs-lookup"><span data-stu-id="6cef8-147">Extend Office functionality</span></span>

<span data-ttu-id="6cef8-148">可以通过以下方式向 Office 应用程序添加新功能：</span><span class="sxs-lookup"><span data-stu-id="6cef8-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="6cef8-149">自定义功能区按钮和菜单命令（统称为“外接程序命令”）</span><span class="sxs-lookup"><span data-stu-id="6cef8-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="6cef8-150">可插入的任务窗格</span><span class="sxs-lookup"><span data-stu-id="6cef8-150">Insertable task panes</span></span>

<span data-ttu-id="6cef8-151">自定义 UI 和任务窗格在外接程序清单中进行指定。</span><span class="sxs-lookup"><span data-stu-id="6cef8-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="6cef8-152">自定义按钮和菜单命令</span><span class="sxs-lookup"><span data-stu-id="6cef8-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="6cef8-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span><span class="sxs-lookup"><span data-stu-id="6cef8-153">You can add custom ribbon buttons and menu items to the ribbon in Office on the web and Windows.</span></span> <span data-ttu-id="6cef8-154">This makes it easy for users to access your add-in directly from their Office application.</span><span class="sxs-lookup"><span data-stu-id="6cef8-154">This makes it easy for users to access your add-in directly from their Office application.</span></span> <span data-ttu-id="6cef8-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span><span class="sxs-lookup"><span data-stu-id="6cef8-155">Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="6cef8-156">*图 3. 功能区中的加载项命令*</span><span class="sxs-lookup"><span data-stu-id="6cef8-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![自定义按钮和菜单命令](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="6cef8-158">任务窗格</span><span class="sxs-lookup"><span data-stu-id="6cef8-158">Task panes</span></span>  

<span data-ttu-id="6cef8-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span><span class="sxs-lookup"><span data-stu-id="6cef8-159">You can use task panes in addition to add-in commands to enable users to interact with your solution.</span></span> <span data-ttu-id="6cef8-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span><span class="sxs-lookup"><span data-stu-id="6cef8-160">Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane.</span></span> <span data-ttu-id="6cef8-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span><span class="sxs-lookup"><span data-stu-id="6cef8-161">Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="6cef8-162">*图 4：任务窗格*</span><span class="sxs-lookup"><span data-stu-id="6cef8-162">*Figure 4. Task pane*</span></span>

![除加载项命令之外，还可以使用任务窗格](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="6cef8-164">扩展 Outlook 功能</span><span class="sxs-lookup"><span data-stu-id="6cef8-164">Extend Outlook functionality</span></span>

<span data-ttu-id="6cef8-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span><span class="sxs-lookup"><span data-stu-id="6cef8-165">Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it.</span></span> <span data-ttu-id="6cef8-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span><span class="sxs-lookup"><span data-stu-id="6cef8-166">They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="6cef8-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span><span class="sxs-lookup"><span data-stu-id="6cef8-167">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences.</span></span> <span data-ttu-id="6cef8-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span><span class="sxs-lookup"><span data-stu-id="6cef8-168">In most cases, an Outlook add-in runs without modification in the Outlook host application to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="6cef8-169">有关 Outlook 加载项的概述，请参阅 [Outlook 加载项概述](../outlook/outlook-add-ins-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="6cef8-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](../outlook/outlook-add-ins-overview.md).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="6cef8-170">在 Office 文档中新建对象</span><span class="sxs-lookup"><span data-stu-id="6cef8-170">Create new objects in Office documents</span></span>

<span data-ttu-id="6cef8-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span><span class="sxs-lookup"><span data-stu-id="6cef8-171">You can embed web-based objects called content add-ins within Excel and PowerPoint documents.</span></span> <span data-ttu-id="6cef8-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span><span class="sxs-lookup"><span data-stu-id="6cef8-172">With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="6cef8-173">*图 5：内容加载项*</span><span class="sxs-lookup"><span data-stu-id="6cef8-173">*Figure 5. Content add-in*</span></span>

![嵌入称为内容加载项的基于 Web 的对象](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="6cef8-175">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="6cef8-175">Office JavaScript APIs</span></span>

<span data-ttu-id="6cef8-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span><span class="sxs-lookup"><span data-stu-id="6cef8-176">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services.</span></span> <span data-ttu-id="6cef8-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span><span class="sxs-lookup"><span data-stu-id="6cef8-177">There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project.</span></span> <span data-ttu-id="6cef8-178">There are also more extensive host-specific object models for Excel and Word.</span><span class="sxs-lookup"><span data-stu-id="6cef8-178">There are also more extensive host-specific object models for Excel and Word.</span></span> <span data-ttu-id="6cef8-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span><span class="sxs-lookup"><span data-stu-id="6cef8-179">These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="6cef8-180">后续步骤</span><span class="sxs-lookup"><span data-stu-id="6cef8-180">Next steps</span></span>

<span data-ttu-id="6cef8-181">有关开发 Office 加载项的更多详细介绍，请参阅[构建 Office 加载项](../overview/office-add-ins-fundamentals.md)。</span><span class="sxs-lookup"><span data-stu-id="6cef8-181">For a more detailed introduction to developing Office Add-ins, see [Building Office Add-ins](../overview/office-add-ins-fundamentals.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6cef8-182">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6cef8-182">See also</span></span>

- [<span data-ttu-id="6cef8-183">构建 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6cef8-183">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="6cef8-184">Office 加载项的核心概念</span><span class="sxs-lookup"><span data-stu-id="6cef8-184">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="6cef8-185">开发 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6cef8-185">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="6cef8-186">设计 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6cef8-186">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="6cef8-187">测试和调试 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6cef8-187">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="6cef8-188">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="6cef8-188">Publish Office Add-ins</span></span>](../publish/publish.md)
