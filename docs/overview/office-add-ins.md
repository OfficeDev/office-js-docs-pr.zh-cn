---
title: Office 加载项平台概述 | Microsoft Docs
description: 使用熟悉的 Web 技术，例如 HTML、CSS 和 JavaScript 来扩展 Word、Excel、PowerPoint、OneNote、Project 和 Outlook，并与其进行交互。
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: dc0a7755027e1d6a741e97928f3f2bc25f62f6c3
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952346"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="a3d7a-103">Office 加载项平台概述</span><span class="sxs-lookup"><span data-stu-id="a3d7a-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="a3d7a-p101">可以使用 Office 加载项平台来生成解决方案，通过解决方案扩展 Office 应用程序，并与 Office 文档中的内容进行交互。通过 Office 加载项，可以使用熟悉的 Web 技术，例如 HTML、CSS 和 JavaScript 来扩展 Word、Excel、PowerPoint、OneNote，Project 和 Outlook，并与其进行交互。解决方案可以跨多个平台在 Office 中运行，包括 Windows 版 Office、Office Online、Office for Mac 和 Office for iPad。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Office for Windows, Office Online, Office for the Mac, and Office for the iPad.</span></span>

<span data-ttu-id="a3d7a-p102">网页在浏览器中能执行的操作，Office 加载项差不多都能执行。使用 Office 加载项平台可以执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="a3d7a-p103">**将新功能添加到 Office 客户端** - 将外部数据引入 Office、自动处理 Office 文档、在 Office 客户端中公开第三方功能等。例如，使用 Microsoft Graph API，可以连接到提升工作效率的数据。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="a3d7a-111">**新建可嵌入到 Office 文档的丰富、交互式对象** - 用户可添加到其自己的 Excel 电子表格和 PowerPoint 演示文稿的嵌入式地图、图表和交互式可视化效果。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-111">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="a3d7a-112">Office 加载项与 COM 和 VSTO 加载项有何不同？</span><span class="sxs-lookup"><span data-stu-id="a3d7a-112">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="a3d7a-p104">COM 或 VSTO 加载项是旧 Office 集成解决方案，仅在 Windows 版 Office 上运行。与 COM 加载项不同，Office 加载项不涉及在用户设备或 Office 客户端中运行的代码。对于 Office 加载项，主机应用程序（例如 Excel）会读取加载项清单，并挂钩 UI 中的加载项自定义功能区按钮和菜单命令。如果需要，它加载加载项的 JavaScript 和 HTML 代码，此代码在沙盒中的浏览器上下文范围内执行。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p104">COM or VSTO add-ins are earlier Office integration solutions that run only on Office for Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="a3d7a-117">相较于使用 VBA、COM 或 VSTO 生成的加载项，Office 加载项提供以下优势：</span><span class="sxs-lookup"><span data-stu-id="a3d7a-117">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="a3d7a-p105">跨平台支持：Office 加载项在 Windows 版 Office、Mac 版 Office、iOS 版 Office 和 Office Online 中运行。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p105">Cross-platform support. Office Add-ins run in Office for Windows, Mac, iOS, and Office Online.</span></span>

- <span data-ttu-id="a3d7a-p106">集中部署和分发：管理员可以在整个组织内集中部署 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p106">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="a3d7a-p107">可通过 AppSource 轻松使用：可以将解决方案提交到 AppSource，供广大受众使用。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p107">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="a3d7a-p108">以标准 Web 技术为依据：可以使用所需的任何库来生成 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p108">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="a3d7a-126">Office 外接程序的组件</span><span class="sxs-lookup"><span data-stu-id="a3d7a-126">Components of an Office Add-in</span></span>

<span data-ttu-id="a3d7a-p109">Office 外接程序包括两个基本组件：XML 清单文件和你自己的 Web 应用程序。此清单定义各种设置，包括将外接程序与 Office 客户端集成的方式。需要在 Web 服务器或 Web 托管服务上托管 Web 应用程序，例如 Microsoft Azure。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p109">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

<span data-ttu-id="a3d7a-130">*图 1：加载项清单 (XML) + 网页 （HTML、JS）= Office 加载项*</span><span class="sxs-lookup"><span data-stu-id="a3d7a-130">*Figure 1. Add-in manifest (XML) + webpage (HTML, JS) = an Office Add-in*</span></span>

![清单 + 网页 = Office 加载项](../images/about-addins-manifestwebpage.png)

### <a name="manifest"></a><span data-ttu-id="a3d7a-132">清单</span><span class="sxs-lookup"><span data-stu-id="a3d7a-132">Manifest</span></span>

<span data-ttu-id="a3d7a-133">清单是一个 XML 文件，它指定外接程序的设置和功能，例如：</span><span class="sxs-lookup"><span data-stu-id="a3d7a-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="a3d7a-134">外接程序的显示名称、说明、ID、版本和默认区域设置。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="a3d7a-135">如何将外接程序与 Office 集成。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="a3d7a-136">外接程序的权限级别和数据访问要求。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="a3d7a-137">Web 应用</span><span class="sxs-lookup"><span data-stu-id="a3d7a-137">Web app</span></span>

<span data-ttu-id="a3d7a-p110">最基本的 Office 加载项包括在 Office 应用中显示的静态 HTML 页面，但此页面并不与 Office 文档或其他任何 Internet 资源交互。不过，若要创建与 Office 文档交互的体验，或创建允许用户通过 Office 主机应用与在线资源交互的体验，可以使用托管提供程序支持的任何客户端和服务器端技术（如 ASP.NET、PHP 或 Node.js）。若要与 Office 客户端和文档交互，可以使用 Office.js JavaScript API。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p110">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="a3d7a-141">*图 2：Hello World Office 加载项的组件*</span><span class="sxs-lookup"><span data-stu-id="a3d7a-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Hello World 加载项的组件](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="a3d7a-143">扩展并与 Office 客户端交互</span><span class="sxs-lookup"><span data-stu-id="a3d7a-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="a3d7a-144">Office 外接程序可以在 Office 主机应用程序中执行下列操作：</span><span class="sxs-lookup"><span data-stu-id="a3d7a-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="a3d7a-145">扩展功能（任何 Office 应用程序）</span><span class="sxs-lookup"><span data-stu-id="a3d7a-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="a3d7a-146">创建新的对象（Excel 或 PowerPoint）</span><span class="sxs-lookup"><span data-stu-id="a3d7a-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="a3d7a-147">扩展 Office 功能</span><span class="sxs-lookup"><span data-stu-id="a3d7a-147">Extend Office functionality</span></span>

<span data-ttu-id="a3d7a-148">可以通过以下方式向 Office 应用程序添加新功能：</span><span class="sxs-lookup"><span data-stu-id="a3d7a-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="a3d7a-149">自定义功能区按钮和菜单命令（统称为“外接程序命令”）</span><span class="sxs-lookup"><span data-stu-id="a3d7a-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="a3d7a-150">可插入的任务窗格</span><span class="sxs-lookup"><span data-stu-id="a3d7a-150">Insertable task panes</span></span>

<span data-ttu-id="a3d7a-151">自定义 UI 和任务窗格在外接程序清单中进行指定。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="a3d7a-152">自定义按钮和菜单命令</span><span class="sxs-lookup"><span data-stu-id="a3d7a-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="a3d7a-p111">可以向 Windows 桌面版 Office 和 Office Online 中的功能区添加自定义功能区按钮和菜单项。这便于用户直接从他们的 Office 应用程序访问加载项。命令按钮可以启动不同操作，如显示带有自定义 HTML 的任务窗格或执行一个 JavaScript 函数。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p111">You can add custom ribbon buttons and menu items to the ribbon in Office for Windows Desktop and Office Online. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="a3d7a-156">*图 3. 功能区中的加载项命令*</span><span class="sxs-lookup"><span data-stu-id="a3d7a-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![自定义按钮和菜单命令](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="a3d7a-158">任务窗格</span><span class="sxs-lookup"><span data-stu-id="a3d7a-158">Task panes</span></span>  

<span data-ttu-id="a3d7a-p112">除了使用加载项命令以外，还可以使用任务窗格，让用户与解决方案交互。不支持加载项命令的客户端（Office 2013 和 Office for iPad）会以任务窗格的形式运行加载项。用户通过“插入”\*\*\*\* 选项卡上的“我的加载项”\*\*\*\* 按钮，启动任务窗格加载项。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p112">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office for iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span> 

<span data-ttu-id="a3d7a-162">*图 4：任务窗格*</span><span class="sxs-lookup"><span data-stu-id="a3d7a-162">*Figure 4. Task pane*</span></span>

![除加载项命令之外，还可以使用任务窗格](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="a3d7a-164">扩展 Outlook 功能</span><span class="sxs-lookup"><span data-stu-id="a3d7a-164">Extend Outlook functionality</span></span>

<span data-ttu-id="a3d7a-p113">Outlook 外接程序可扩展 Office 功能区，还可以在查看或撰写 Outlook 项目时在其旁边的上下文中显示。当用户查看接收的项目或回复或创建新项目时，它们可以与电子邮件、会议请求、会议响应、会议取消或约会一起使用。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p113">Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="a3d7a-p114">Outlook 加载项可以通过项访问上下文信息（如地址或跟踪 ID），再使用此类数据访问服务器和 Web 服务上的其他信息，以打造有吸引力的用户体验。大多数情况下，Outlook 加载项无需经过修改，即可在各种支持的主机应用（包括 Outlook、Outlook for Mac、Outlook Web App，以及适用于设备的 Outlook Web App）上运行，从而在桌面设备、Web 设备、平板电脑和移动设备上提供无缝体验。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p114">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App, and Outlook Web App for devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span> 

<span data-ttu-id="a3d7a-169">有关 Outlook 加载项的概述，请参阅 [Outlook 加载项概述](/outlook/add-ins/)。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](/outlook/add-ins/).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="a3d7a-170">在 Office 文档中新建对象</span><span class="sxs-lookup"><span data-stu-id="a3d7a-170">Create new objects in Office documents</span></span>

<span data-ttu-id="a3d7a-p115">可以在 Excel 和 PowerPoint 文档中嵌入基于 Web 的对象（称为“内容加载项”）。通过内容加载项，可以集成基于 Web 的丰富数据可视化、媒体（如 YouTube 视频播放器或图片库）和其他外部内容。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p115">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="a3d7a-173">*图 5：内容加载项*</span><span class="sxs-lookup"><span data-stu-id="a3d7a-173">*Figure 5. Content add-in*</span></span>

![嵌入称为内容加载项的基于 Web 的对象](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="a3d7a-175">Office JavaScript API</span><span class="sxs-lookup"><span data-stu-id="a3d7a-175">Office JavaScript APIs</span></span>

<span data-ttu-id="a3d7a-p116">Office JavaScript API 包含的对象和成员适用于生成加载项，并与 Office 内容和 Web 服务交互。Excel、Outlook、Word、PowerPoint、OneNote 和 Project 共用一个常见对象模型。对于 Excel 和 Word，还有更多主机专用对象模型。这些 API 提供对已知对象（如段落和工作簿）的访问权限，以便于能够更轻松地为特定主机创建加载项。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-p116">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="a3d7a-180">后续步骤</span><span class="sxs-lookup"><span data-stu-id="a3d7a-180">Next steps</span></span>

<span data-ttu-id="a3d7a-181">要详细了解如何开始构建 Office 加载项，请尝试 [5 分钟快速入门](/office/dev/add-ins/)。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-181">To learn more about how to start building your Office Add-in, try out our [5-minute Quick Starts](/office/dev/add-ins/).</span></span> <span data-ttu-id="a3d7a-182">可以使用 Visual Studio 或任何其他编辑器立即开始构建加载项。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-182">You can start building add-ins right away using Visual Studio or any other editor.</span></span> 

<span data-ttu-id="a3d7a-183">若要开始计划解决方案并打造有吸引力的有效用户体验，请熟悉 Office 加载项的[设计指南](../design/add-in-design.md)和[最佳做法](../concepts/add-in-development-best-practices.md)。</span><span class="sxs-lookup"><span data-stu-id="a3d7a-183">To start planning solutions that create effective and compelling user experiences, get familiar with the [design guidelines](../design/add-in-design.md) and [best practices](../concepts/add-in-development-best-practices.md) for Office Add-ins.</span></span>    

## <a name="see-also"></a><span data-ttu-id="a3d7a-184">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a3d7a-184">See also</span></span>

- [<span data-ttu-id="a3d7a-185">Office 加载项示例</span><span class="sxs-lookup"><span data-stu-id="a3d7a-185">Office Add-in samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
- [<span data-ttu-id="a3d7a-186">了解适用于 Office 的 JavaScript API</span><span class="sxs-lookup"><span data-stu-id="a3d7a-186">Understanding the JavaScript API for Office</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="a3d7a-187">Office 外接程序主机和平台可用性</span><span class="sxs-lookup"><span data-stu-id="a3d7a-187">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
