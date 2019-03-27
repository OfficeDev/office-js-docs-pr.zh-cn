---
title: Office 加载项平台概述 | Microsoft Docs
description: 使用熟悉的 Web 技术，例如 HTML、CSS 和 JavaScript 来扩展 Word、Excel、PowerPoint、OneNote、Project 和 Outlook，并与其进行交互。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 480228c20b20de52a9e1224f6691696b5560986c
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870315"
---
# <a name="office-add-ins-platform-overview"></a>Office 加载项平台概述

可以使用 Office 外接程序平台来生成解决方案，通过解决方案扩展 Office 应用程序，并与 Office 文档中的内容进行交互。通过 Office 外接程序，可以使用熟悉的 Web 技术，例如 HTML、CSS 和 JavaScript 来扩展 Word、Excel、PowerPoint、OneNote，Project 和 Outlook，并与其进行交互。解决方案可以跨多个平台在 Office 中运行，包括 Office for Windows、Office Online、Office for Mac 和 Office for iPad。

网页在浏览器中能执行的操作，Office 加载项差不多都能执行。使用 Office 加载项平台可以执行下列操作：

-  **将新功能添加到 Office 客户端** - 将外部数据引入 Office、自动处理 Office 文档、在 Office 客户端中公开第三方功能等。例如，使用 Microsoft Graph API，可以连接到提升工作效率的数据。

-  **新建可嵌入到 Office 文档的丰富、交互式对象** - 用户可添加到其自己的 Excel 电子表格和 PowerPoint 演示文稿的嵌入式地图、图表和交互式可视化效果。

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a>Office 加载项与 COM 和 VSTO 加载项有何不同？

COM 或 VSTO 加载项是旧 Office 集成解决方案，仅在 Office for Windows 上运行。与 COM 加载项不同，Office 加载项不涉及在用户设备或 Office 客户端中运行的代码。对于 Office 加载项，主机应用（例如 Excel）会读取加载项清单，并挂钩 UI 中的加载项自定义功能区按钮和菜单命令。如果需要，它加载加载项的 JavaScript 和 HTML 代码，此代码在沙盒中的浏览器上下文范围内执行。

相较于使用 VBA、COM 或 VSTO 生成的加载项，Office 加载项提供以下优势：

- 跨平台支持：Office 加载项在 Office for Windows、Mac、iOS 和 Office Online 中运行。

- 集中部署和分发：管理员可以在整个组织内集中部署 Office 加载项。

- 可通过 AppSource 轻松使用：可以将解决方案提交到 AppSource，供广大受众使用。

- 以标准 Web 技术为依据：可以使用所需的任何库来生成 Office 加载项。

## <a name="components-of-an-office-add-in"></a>Office 外接程序的组件

Office 外接程序包括两个基本组件：XML 清单文件和你自己的 Web 应用程序。此清单定义各种设置，包括将外接程序与 Office 客户端集成的方式。需要在 Web 服务器或 Web 托管服务上托管 Web 应用程序，例如 Microsoft Azure。

*图 1：加载项清单 (XML) + 网页 （HTML、JS）= Office 加载项*

![清单 + 网页 = Office 加载项](../images/about-addins-manifestwebpage.png)

### <a name="manifest"></a>清单

清单是一个 XML 文件，它指定外接程序的设置和功能，例如：

- 外接程序的显示名称、说明、ID、版本和默认区域设置。

- 如何将外接程序与 Office 集成。  

- 外接程序的权限级别和数据访问要求。

### <a name="web-app"></a>Web 应用

最基本的 Office 加载项包括在 Office 应用中显示的静态 HTML 页面，但此页面并不与 Office 文档或其他任何 Internet 资源交互。不过，若要创建与 Office 文档交互的体验，或创建允许用户通过 Office 主机应用与在线资源交互的体验，可以使用托管提供程序支持的任何客户端和服务器端技术（如 ASP.NET、PHP 或 Node.js）。若要与 Office 客户端和文档交互，可以使用 Office.js JavaScript API。

*图 2：Hello World Office 加载项的组件*

![Hello World 加载项的组件](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a>扩展并与 Office 客户端交互

Office 外接程序可以在 Office 主机应用程序中执行下列操作：

-  扩展功能（任何 Office 应用程序）

-  创建新的对象（Excel 或 PowerPoint）
 
### <a name="extend-office-functionality"></a>扩展 Office 功能

可以通过以下方式向 Office 应用程序添加新功能：  

-  自定义功能区按钮和菜单命令（统称为“外接程序命令”）

-  可插入的任务窗格

自定义 UI 和任务窗格在外接程序清单中进行指定。  

#### <a name="custom-buttons-and-menu-commands"></a>自定义按钮和菜单命令  

可以向 Office for Windows Desktop 和 Office Online 中的功能区添加自定义功能区按钮和菜单项。这便于用户直接从他们的 Office 应用程序访问外接程序。命令按钮可以启动不同操作，如显示带有自定义 HTML 的任务窗格或执行一个 JavaScript 函数。  

*图 3. 功能区中的加载项命令*

![自定义按钮和菜单命令](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>任务窗格  

除了使用加载项命令以外，还可以使用任务窗格，让用户与解决方案交互。不支持加载项命令的客户端（Office 2013 和 Office for iPad）会以任务窗格的形式运行加载项。用户通过“插入”**** 选项卡上的“我的加载项”**** 按钮，启动任务窗格加载项。 

*图 4：任务窗格*

![除加载项命令之外，还可以使用任务窗格](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>扩展 Outlook 功能

Outlook 外接程序可扩展 Office 功能区，还可以在查看或撰写 Outlook 项目时在其旁边的上下文中显示。当用户查看接收的项目或回复或创建新项目时，它们可以与电子邮件、会议请求、会议响应、会议取消或约会一起使用。 

Outlook 加载项可以通过项访问上下文信息（如地址或跟踪 ID），再使用此类数据访问服务器和 Web 服务上的其他信息，以打造有吸引力的用户体验。大多数情况下，Outlook 加载项无需经过修改，即可在各种支持的主机应用（包括 Outlook、Outlook for Mac、Outlook Web App，以及适用于设备的 Outlook Web App）上运行，从而在桌面设备、Web 设备、平板电脑和移动设备上提供无缝体验。 

有关 Outlook 加载项的概述，请参阅 [Outlook 加载项概述](/outlook/add-ins/)。

### <a name="create-new-objects-in-office-documents"></a>在 Office 文档中新建对象

可以在 Excel 和 PowerPoint 文档中嵌入基于 Web 的对象（称为“内容加载项”）。通过内容加载项，可以集成基于 Web 的丰富数据可视化、媒体（如 YouTube 视频播放器或图片库）和其他外部内容。

*图 5：内容加载项*

![嵌入称为内容加载项的基于 Web 的对象](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>Office JavaScript API

Office JavaScript API 包含的对象和成员适用于生成加载项，并与 Office 内容和 Web 服务交互。Excel、Outlook、Word、PowerPoint、OneNote 和 Project 共用一个常见对象模型。对于 Excel 和 Word，还有更多主机专用对象模型。这些 API 提供对已知对象（如段落和工作簿）的访问权限，以便于能够更轻松地为特定主机创建加载项。  

## <a name="next-steps"></a>后续步骤

要详细了解如何开始构建 Office 加载项，请尝试 [5 分钟快速入门](/office/dev/add-ins/)。 可以使用 Visual Studio 或任何其他编辑器立即开始构建加载项。 

若要开始计划解决方案并打造有吸引力的有效用户体验，请熟悉 Office 加载项的[设计指南](../design/add-in-design.md)和[最佳做法](../concepts/add-in-development-best-practices.md)。    

## <a name="see-also"></a>另请参阅

- [Office 加载项示例](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
- [了解适用于 Office 的 JavaScript API](../develop/understanding-the-javascript-api-for-office.md)
- [Office 外接程序主机和平台可用性](../overview/office-add-in-availability.md)
