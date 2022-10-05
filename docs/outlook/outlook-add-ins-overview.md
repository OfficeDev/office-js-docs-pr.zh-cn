---
title: Outlook 加载项概述
description: Outlook 加载项由第三方使用基于 Web 的平台集成到 Outlook 中。
ms.date: 08/09/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: fd17728f840188fbedfdeba7d3ee8f97852d702a
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467256"
---
# <a name="outlook-add-ins-overview"></a>Outlook 加载项概述

Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:

- 相同的加载项和业务逻辑可跨桌面（Windows 版和 Mac 版 Outlook）、Web（Microsoft 365 和 Outlook.com）和移动平台使用。
- Outlook 外接程序包括一个清单，其中介绍了如何将外接程序集成到 Outlook（例如，按钮或任务窗格）中，以及构成外接程序 UI 和业务逻辑的 JavaScript/HTML 代码。
- 最终用户或管理员可以从 [AppSource](https://appsource.microsoft.com) 获取 Outlook 加载项，也可以进行[旁加载](sideload-outlook-add-ins-for-testing.md)。

Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.

The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>扩展点

Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done.

- Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

    **功能区上具有命令按钮的加载项**

    ![加载项函数命令。](../images/uiless-command-shape.png)

- Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).

    **用于突出显示的实体（地址）的上下文相关加载项**

    ![在卡片中显示上下文相关应用。](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>外接程序可用的邮箱项目

当用户正在撰写或阅读邮件或约会，而不是其他项目类型时，Outlook 加载项会激活。 但是，如果撰写或阅读窗体中的当前邮件项目为以下项之一，则 Outlook *不会* 激活邮件加载项：

- 受信息权限管理 (IRM 的保护) 或以其他方式加密，以便在非 Windows 客户端上通过 Outlook 进行保护和访问。 由于数字签名依赖于这些机制之一，数字签名邮件就是一个示例。

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- 具有邮件类别 IPM.Report.* 的送达报告或通知，包括送达和未送达报告 (NDR)，以及已读、未读和延迟通知。

- 属于其他邮件的附件的 .msg 或 .eml 文件。

- 从文件系统打开的 .msg 或 .eml 文件。

- 在 [组邮箱](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes)、共享邮箱\*、另一用户邮箱\*、 [存档邮箱](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-client-and-compliance-&-security-feature-details?tabs=Archive-features#archive-mailbox)或公用文件夹中。

  > [!IMPORTANT]
  > [要求集 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8)中引入了 \* 对委托访问方案的支持（例如，从其他用户的邮箱共享的文件夹）。 现在，共享邮箱支持在 Windows 版和 Mac 版 Outlook 中进行预览。 若要了解详细信息，请参阅 [“启用共享文件夹”和“共享邮箱”方案](delegate-access.md)。

- 使用自定义窗体。

- 通过简单 MAPI 创建。 如果 Outlook 关闭时，Office 用户从 Windows 上的 Office 应用程序创建或发送电子邮件，则将使用简单 MAPI。 例如，用户在 Word 中工作时可以创建 Outlook 电子邮件，这会触发 Outlook 撰写窗口，而无需启动完整的 Outlook 应用程序。 但是，如果用户从 Word 创建电子邮件时 Outlook 已在运行，则这不属于简单 MAPI 方案，因此只要满足其他激活要求，Outlook 加载项就会在撰写窗体中工作。

通常，Outlook 可以为"已发送邮件"文件夹中的项目在阅读窗体中激活加载项，基于已知实体字符串匹配激活的加载项除外。 有关其背后的具体原因，请参阅 [支持已知实体](match-strings-in-an-item-as-well-known-entities.md#support-for-well-known-entities)。

目前，设计和实现移动客户端的加载项时还有其他注意事项。 若要了解详细信息，请参阅 [将移动支持添加到 Outlook 加载项](add-mobile-support.md#compose-mode-and-appointments)。

## <a name="supported-clients"></a>支持的客户端

Windows 版 Outlook 2013 或更高版本、Mac 版 Outlook 2016 或更高版本、适用于本地 Exchange 2013 和更高版本的 Outlook 网页版、iOS 版 Outlook、Android 版 Outlook 及 Outlook 网页版和 Outlook.com 支持 Outlook 加载项。 并非所有[客户端](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)都同时支持全部最新功能。 请参阅有关这些功能的文章和 API 参考，了解它们可能在哪些应用程序中受支持或不受支持。

## <a name="get-started-building-outlook-add-ins"></a>开始构建 Outlook 外接程序

要开始生成 Outlook 加载项，请尝试执行以下操作：

- [快速入门](../quickstarts/outlook-quickstart.md) - 生成简单的任务窗格。
- [教程](../tutorials/outlook-tutorial.md) - 了解如何创建将 GitHub Gist 插入新邮件的加载项。

## <a name="see-also"></a>另请参阅

- [了解 Microsoft 365 开发人员计划](https://developer.microsoft.com/microsoft-365/dev-program)
- [Office 加载项开发最佳做法](../concepts/add-in-development-best-practices.md)
- [Office 加载项的设计准则](../design/add-in-design.md)
- [许可 Office 和 SharePoint 加载项](/office/dev/store/license-your-add-ins)
- [发布 Office 加载项](../publish/publish.md)
- [将解决方案提交到 AppSource 和 Office 应用商店](/office/dev/store/submit-to-the-office-store)
