---
title: Office 加载项平台概述
description: 使用熟悉的 Web 技术，例如 HTML、CSS 和 JavaScript 来扩展 Word、Excel、PowerPoint、OneNote、Project 和 Outlook，并与其进行交互。
ms.date: 04/14/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 5a780fcc1f863fb6803e2f719fc27338d4a6c366
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810111"
---
# <a name="office-add-ins-platform-overview"></a>Office 加载项平台概述

You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Outlook, Excel, Word, PowerPoint, OneNote, and Project. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.

![Office 应用程序加上嵌入式网站（外接程序）可实现无限扩展性。](../images/addins-overview.png)

Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:

- **Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose functionality from Microsoft and others in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.

- **新建可嵌入到 Office 文档的丰富、交互式对象** - 用户可添加到其自己的 Excel 电子表格和 PowerPoint 演示文稿的嵌入式地图、图表和交互式可视化效果。

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a>Office 加载项与 COM 和 VSTO 加载项有何不同？

COM or VSTO add-ins are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the application (for example, Excel), reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.

![使用 Office 加载项的原因：跨平台、集中部署、通过 AppSource 轻松访问以及基于标准 Web 技术构建。](../images/why.png)

相较于使用 VBA、COM 或 VSTO 生成的加载项，Office 加载项提供以下优势。

- Cross-platform support. Office Add-ins run in Office on the web, Windows, Mac, and iPad.

- Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.

- Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.

- Based on standard web technology. You can use any library you like to build Office Add-ins.

## <a name="components-of-an-office-add-in"></a>Office 加载项的组件

An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.

### <a name="manifest"></a>清单

清单是一个 XML 文件，它指定外接程序的设置和功能，例如：

- 外接程序的显示名称、说明、ID、版本和默认区域设置。

- 如何将外接程序与 Office 集成。  

- 外接程序的权限级别和数据访问要求。

### <a name="web-app"></a>Web 应用

The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office client application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.

![Hello World 加载项的组件。](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a>扩展并与 Office 客户端交互

Office 加载项可以在 Office 客户端应用程序中执行下列操作。

- 扩展功能（任何 Office 应用程序）

- 创建新的对象（Excel 或 PowerPoint）

### <a name="extend-office-functionality"></a>扩展 Office 功能

可以通过以下方式向 Office 应用程序添加新功能：  

- 自定义功能区按钮和菜单命令（统称为“外接程序命令”）

- 可插入的任务窗格

自定义 UI 和任务窗格在外接程序清单中进行指定。  

#### <a name="custom-buttons-and-menu-commands"></a>自定义按钮和菜单命令  

You can add custom ribbon buttons and menu items to the ribbon in Office on the web and on Windows. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.  

![自定义按钮和菜单命令。](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>任务窗格  

You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.

![除加载项命令之外，还可以使用任务窗格。](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>扩展 Outlook 功能

Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.

Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification in the Outlook application to provide a seamless experience on the desktop, web, and tablet and mobile devices.

有关 Outlook 加载项的概述，请参阅 [Outlook 加载项概述](../outlook/outlook-add-ins-overview.md)。

### <a name="create-new-objects-in-office-documents"></a>在 Office 文档中新建对象

You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.

![嵌入称为内容加载项的基于 Web 的对象。](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>Office JavaScript API

Office JavaScript API 包含的对象和成员适用于生成加载项，并与 Office 内容和 Web 服务交互。 Excel、Outlook、Word、PowerPoint、OneNote 和 Project 共享的通用对象模型。 Excel 和 Word 还有更广泛的特定于应用程序的对象模型。 这些 API 提供对已知对象（如段落和工作簿）的访问，从而更轻松地为特定应用程序创建外接程序。

## <a name="next-steps"></a>后续步骤

有关开发 Office 加载项的更多详细介绍，请参阅[开发 Office 加载项](../develop/develop-overview.md)。

## <a name="see-also"></a>另请参阅

- [Office 加载项的核心概念](../overview/core-concepts-office-add-ins.md)
- [开发 Office 加载项](../develop/develop-overview.md)
- [设计 Office 加载项](../design/add-in-design.md)
- [测试和调试 Office 加载项](../testing/test-debug-office-add-ins.md)
- [发布 Office 加载项](../publish/publish.md)
- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
