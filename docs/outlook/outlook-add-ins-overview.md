---
title: Outlook 加载项概述
description: Outlook 加载项由第三方使用基于 Web 的平台集成到 Outlook 中。
ms.date: 10/09/2019
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 82778f7118166f7ed566fc175599efd7049b9d3a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609025"
---
# <a name="outlook-add-ins-overview"></a>Outlook 加载项概述

Outlook 加载项由第三方使用基于 Web 的平台集成到 Outlook 中。 Outlook 加载项有三个主要方面：

- 相同的加载项和业务逻辑可跨桌面（Windows 版 Outlook 和 Outlook for Mac）、Web（Office 365 和 Outlook.com）和移动平台使用。
- Outlook 外接程序包括一个清单，其中介绍了如何将外接程序集成到 Outlook（例如，按钮或任务窗格）中，以及构成外接程序 UI 和业务逻辑的 JavaScript/HTML 代码。
- 最终用户或管理员可以从 [AppSource](https://appsource.microsoft.com) 获取 Outlook 加载项，也可以进行[旁加载](sideload-outlook-add-ins-for-testing.md)。

Outlook 加载项不同于 COM 或 VSTO 的加载项，后者为特定于 Windows 版 Outlook的较旧集成。 Outlook 加载项与 COM 加载项不同，它在用户的设备或 Outlook 客户端上没有通过物理方式安装任何代码。 对于 Outlook 加载项，Outlook 读取清单，挂钩 UI 中的指定控件，然后加载 JavaScript 和 HTML。 Web 组件都在沙盒浏览器的上下文中运行。

支持加载项的 Outlook 项目包括电子邮件、会议请求、响应和取消及约会。 每个 Outlook 加载项均定义其可用的上下文，包括项目类型以及用户是在阅读还是撰写项目。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>扩展点

扩展点是加载项与 Outlook 集成的方式。以下是执行此操作的方法：

- 加载项可以声明出现在所有邮件和约会的命令界面中的按钮。有关详细信息，请参阅 [用于 Outlook 的加载项命令](add-in-commands-for-outlook.md)。

    **功能区上具有命令按钮的加载项**

    ![加载项命令无 UI 形状](../images/uiless-command-shape.png)

- 加载项可以在邮件和约会中中断与正则表达式匹配项或检测实体的链接。 有关详细信息，请参阅 [上下文 Outlook 加载项](contextual-outlook-add-ins.md)。

    **用于突出显示的实体（地址）的上下文相关加载项**

    ![在卡片中显示上下文相关应用程序](../images/outlook-detected-entity-card.png)


> [!NOTE]
> [已弃用自定义窗格](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/)，因此，请确保使用的是受支持的扩展点。

## <a name="mailbox-items-available-to-add-ins"></a>外接程序可用的邮箱项目

在撰写或阅读时，Outlook 外接程序对邮件或约会可用，但对其他项目类型不可用。如果撰写或阅读窗体中的当前邮件项目为以下项之一，则 Outlook 不会激活邮件外接程序：

- 使用信息权限管理 (IRM) 进行保护，或使用其他保护方式进行加密。数字签名邮件便是其中一个例子，因为数字签名依赖于这些机制之一。

- 具有邮件类别 IPM.Report.* 的送达报告或通知，包括送达和未送达报告 (NDR)，以及已读、未读和延迟通知。

- 草稿（没有为其分配发件人），或位于 Outlook 草稿文件夹中。

- 属于其他邮件的附件的 .msg 或 .eml 文件。

- 从文件系统打开的 .msg 或 .eml 文件。

- 在共享邮箱、其他用户的邮箱、存档邮箱或公用文件夹中。

- 使用自定义窗体。

通常，Outlook 可以为"已发送邮件"文件夹中的项目在阅读窗体中激活加载项，基于已知实体字符串匹配激活的加载项除外。 欲了解其背后原因的详细信息，请参阅[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。

## <a name="supported-hosts"></a>支持的主机

Windows 版 Outlook 2013 或更高版本、Mac 版 Outlook 2016 或更高版本、Outlook 网页版（本地 Exchange 2013 和更高版本）、iOS 版 Outlook、Android 版 Outlook 及 Outlook 网页版（Office 365 和 Outlook.com）支持 Outlook 加载项。 并非所有[客户端](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)都同时支持全部最新功能。 有关这些功能的信息，请参阅文章和 API 参考，了解它们可能在哪些主机中受支持或不受支持。


## <a name="get-started-building-outlook-add-ins"></a>开始构建 Outlook 外接程序

若要开始生成 Outlook 加载项，请尝试执行以下操作。

- [快速入门](../quickstarts/outlook-quickstart.md) - 生成简单的任务窗格。
- [教程](../tutorials/outlook-tutorial.md) - 了解如何创建将 GitHub Gist 插入新邮件的加载项。


## <a name="see-also"></a>另请参阅

- [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- [Office 加载项的设计准则](../design/add-in-design.md)
- [许可 Office 和 SharePoint 加载项](/office/dev/store/license-your-add-ins)
- [发布 Office 加载项](../publish/publish.md)
- [将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-the-office-store)
