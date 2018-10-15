---
title: Office 外接程序开发生命周期
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 5b056527deaf03beb51d755b582be715fbd14233
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505893"
---
# <a name="office-add-ins-development-lifecycle"></a>Office 外接程序开发生命周期

> [!NOTE]
> 如果计划将外接程序[发布](../publish/publish.md)到 AppSource 并适用于 Office 体验，请务必遵循 [AppSource 验证策略](https://docs.microsoft.com/office/dev/store/validation-policies)。例如，外接程序必须适用于支持已定义方法的所有平台，才能通过验证（有关详细信息，请参阅[第 4.12 部分](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably)以及 [Office 外接程序主机和可用性页面](../overview/office-add-in-availability.md)）。 

Office 外接程序的典型开发生命周期包括下列步骤：


## <a name="1-decide-on-the-purpose-of-the-add-in"></a>1. 确定外接程序的用途
    
提出以下问题：
    
- 外接程序有何作用？ 
        
- 它如何帮助您的客户提高工作效率？
        
- 您的外接程序功能支持哪些方案？
    
确定最重要的功能和方案，并围绕它们进行集中设计。 

    
## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a>2. 确定外接程序的数据和数据源
    
- 是文档、工作簿、演示文稿、项目中的数据，还是基于 Access 浏览器数据库中的数据？ 
    
- 数据是否关于 Exchange Server 或 Exchange Online 邮箱中的一个或多个项？ 
    
- 数据是否来自外部源（如 Web 服务）？

    
## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a>3. 确定外接程序类型和最能支持其用途的 Office 主机应用
    
为确定方案，请考虑以下几点：
    
- 客户是否要使用外接程序来丰富文档或基于 Access 浏览器数据库的内容？如果是，建议考虑创建**内容外接程序**。 
    
- 客户是否要在查看、撰写电子邮件或约会时使用该外接程序？能够根据当前上下文公开外接程序是否很重要？是否优先考虑使外接程序不仅在台式机上可用，而且在平板电脑或智能手机上也可用？
    
    如果上述任一问题的答案是肯定的，请考虑创建 **Outlook 外接程序**。确定外接程序的触发上下文（例如，撰写表单中的用户、特定消息类型、是否有附件、地址、任务建议、会议建议，或者电子邮件或约会内容中的特定字符串模式）。 
        
    若要了解如何根据上下文激活 Outlook 外接程序，请参阅 [Outlook 外接程序的激活规则](https://docs.microsoft.com/outlook/add-ins/activation-rules)。 
    
- 客户是否要使用外接程序来增强文档的查看或创作体验？如果是，建议考虑创建**任务窗格外接程序**。 

某些外接程序 API 的支持可能因 Office 应用及其运行平台（Windows、Mac、Web、移动平台）而异。若要了解客户端和平台的当前 API 覆盖范围，请参阅 [Office 外接程序主机和平台可用性](../overview/office-add-in-availability.md)页面。  

    
## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a>4. 设计并实现外接程序的用户体验和用户接口
    
设计快速、流畅、连贯和易于学习的用户体验，主要方案只需几个步骤即可完成。根据外接程序的用途，可使用第三方 API 或 Web 服务。
    
可从各种 Web 开发工具中进行选择，并使用 HTML 和 JavaScript 实现用户接口。

    
## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a>5. 根据 Office 外接程序清单架构创建 XML 清单文件
    
创建 XML 清单，以确定外接程序及其要求，指定外接程序使用的 HTML 以及任何 JavaScript 和 CSS 文件的位置，并根据外接程序的类型指定默认大小和权限。
    
对于 Outlook 外接程序，可以根据当前邮件或约会指定上下文，外接程序在其中不仅相关，还可供 Outlook 在 UI 中使用。您还可以确定希望外接程序支持的设备。在清单中，将上下文指定为激活规则和受支持的设备。
    

## <a name="6-install-and-test-the-add-in"></a>6. 安装和测试外接程序
    
将 HTML 文件以及任何 JavaScript 和 CSS 文件放在外接程序清单文件中指定的 Web 服务器上。安装外接程序的过程取决于外接程序的类型。有关详细信息，请参阅[旁加载 Office 外接程序进行测试](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)。
    
对于 Outlook 外接程序，将其安装在 Exchange 邮箱中，并指定 Exchange 管理中心 (EAC) 中外接程序清单文件的位置。有关详细信息，请参阅[部署和安装 Outlook 外接程序以供测试](https://docs.microsoft.com/outlook/add-ins/testing-and-tips)。

    
## <a name="7-publish-the-add-in"></a>7. 发布外接程序
    
可以将外接程序提交到 AppSource，客户从中能够安装外接程序。此外，还可以向 SharePoint 上的专有文件夹外接程序目录或共享网络文件夹发布任务窗格和内容外接程序，并在组织的 Exchange 服务器上直接部署 Outlook 外接程序。有关详细信息，请参阅[发布 Office 外接程序](../publish/publish.md)。
    
    
## <a name="8-maintain-the-add-in"></a>8. 维护外接程序
    
如果外接程序调用 web 服务，且在发布外接程序后对 web 服务进行更新，则无需重新发布外接程序。但是，如果您对提交的外接程序的任何项目或数据进行了更改（如外接程序清单、屏幕截图、图标、HTML 或 JavaScript 文件），则需重新发布外接程序。 
    
特别是，如果已将外接程序发布到 AppSource，您需要重新提交外接程序，以便 AppSource 能够实现这些更改。重新提交外接程序时，必须附带包含新版本号的更新后外接程序清单。还必须确保更新提交表单中的外接程序版本号，需与新清单版本号一致。对于 Outlook 外接程序，应确保 [Id](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id?view=office-js) 元素包含外接程序清单中的不同 UUID。
    
