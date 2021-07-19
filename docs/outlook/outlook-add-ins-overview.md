---
title: Outlook 加载项概述
description: Outlook 加载项由第三方使用基于 Web 的平台集成到 Outlook 中。
ms.date: 07/14/2021
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 0d9dd51627cd797351e4e43957375b7a493b2b57
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455486"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="bc7ab-103">Outlook 加载项概述</span><span class="sxs-lookup"><span data-stu-id="bc7ab-103">Outlook add-ins overview</span></span>

<span data-ttu-id="bc7ab-p101">Outlook 加载项是由第三方通过使用基于 Web 的平台构建到 Outlook 中的集成。Outlook 加载项具有三个关键方面：</span><span class="sxs-lookup"><span data-stu-id="bc7ab-p101">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="bc7ab-106">相同的加载项和业务逻辑可跨桌面（Windows 版和 Mac 版 Outlook）、Web（Microsoft 365 和 Outlook.com）和移动平台使用。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Microsoft 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="bc7ab-107">Outlook 外接程序包括一个清单，其中介绍了如何将外接程序集成到 Outlook（例如，按钮或任务窗格）中，以及构成外接程序 UI 和业务逻辑的 JavaScript/HTML 代码。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="bc7ab-108">最终用户或管理员可以从 [AppSource](https://appsource.microsoft.com) 获取 Outlook 加载项，也可以进行[旁加载](sideload-outlook-add-ins-for-testing.md)。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="bc7ab-p102">Outlook 外接程序与 COM 或 VSTO 外接程序（特定于在 Windows 上运行的 Outlook 的较早集成项）不同。与 COM 外接程序不同的是，Outlook 外接程序不具有任何实际安装到用户设备或 Outlook 客户端的代码。对于 Outlook 外接程序，Outlook 读取清单并挂钩在 UI 中指定的控件，然后加载 JavaScript 和 HTML。web 部件全部在沙盒的浏览器的上下文中执行。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-p102">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="bc7ab-p103">支持加载项的 Outlook 项目包括电子邮件、会议请求、响应和取消及约会。每个 Outlook 加载项均定义其可用的上下文，包括项目类型以及用户是在阅读还是撰写项目。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-p103">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a><span data-ttu-id="bc7ab-115">扩展点</span><span class="sxs-lookup"><span data-stu-id="bc7ab-115">Extension points</span></span>

<span data-ttu-id="bc7ab-p104">扩展点是加载项与 Outlook 集成的方式。以下是执行此操作的方法。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-p104">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done.</span></span>

- <span data-ttu-id="bc7ab-p105">加载项可以声明出现在所有邮件和约会的命令界面中的按钮。有关详细信息，请参阅 [用于 Outlook 的加载项命令](add-in-commands-for-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-p105">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="bc7ab-120">**功能区上具有命令按钮的加载项**</span><span class="sxs-lookup"><span data-stu-id="bc7ab-120">**An add-in with command buttons on the ribbon**</span></span>

    ![加载项命令无 UI 形状。](../images/uiless-command-shape.png)

- <span data-ttu-id="bc7ab-p106">加载项可以在邮件和约会中中断与正则表达式匹配项或检测实体的链接。 有关详细信息，请参阅 [上下文 Outlook 加载项](contextual-outlook-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-p106">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="bc7ab-124">**用于突出显示的实体（地址）的上下文相关加载项**</span><span class="sxs-lookup"><span data-stu-id="bc7ab-124">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![在卡片中显示上下文相关应用。](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="bc7ab-126">外接程序可用的邮箱项目</span><span class="sxs-lookup"><span data-stu-id="bc7ab-126">Mailbox items available to add-ins</span></span>

<span data-ttu-id="bc7ab-127">当用户正在撰写或阅读邮件或约会，而不是其他项目类型时，Outlook 加载项会激活。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-127">Outlook add-ins activate when the user is composing or reading a message or appointment, but not other item types.</span></span> <span data-ttu-id="bc7ab-128">但是，如果撰写或阅读窗体中的当前邮件项目为以下项之一，则 Outlook *不会* 激活邮件加载项：</span><span class="sxs-lookup"><span data-stu-id="bc7ab-128">However, add-ins are *not* activated if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="bc7ab-p108">使用信息权限管理 (IRM) 进行保护，或使用其他保护方式进行加密。数字签名邮件便是其中一个例子，因为数字签名依赖于这些机制之一。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-p108">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

  > [!IMPORTANT]
  >
  > - <span data-ttu-id="bc7ab-131">加载项在与 Microsoft 365 订阅相关联的 Outlook 电子签名邮件上激活。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-131">Add-ins activate on digitally signed messages in Outlook associated with a Microsoft 365 subscription.</span></span> <span data-ttu-id="bc7ab-132">在Windows上，这个支持是通过8711.1000版本中引入的。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-132">On Windows, this support was introduced with build 8711.1000.</span></span>
  >
  > - <span data-ttu-id="bc7ab-133">现在，Windows 版 Outlook 从内部版本 13229.10000 开始可以在受 IRM 保护的项目上激活加载项。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-133">Starting with Outlook build 13229.10000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="bc7ab-134">有关处于预览阶段的此功能的详细信息，请参阅 [在受信息权限管理 (IRM) 保护的项目上激活加载项](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-134">For more information about this feature in preview, refer to [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="bc7ab-135">具有邮件类别 IPM.Report.\* 的送达报告或通知，包括送达和未送达报告 (NDR)，以及已读、未读和延迟通知。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-135">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="bc7ab-136">属于其他邮件的附件的 .msg 或 .eml 文件。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="bc7ab-137">从文件系统打开的 .msg 或 .eml 文件。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="bc7ab-138">在 [组邮箱](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes)、共享邮箱\*、另一用户邮箱\*、 [存档邮箱](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-features#archive-mailbox)或公用文件夹中。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-138">In a [group mailbox](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), in a shared mailbox\*, in another user's mailbox\*, in an [archive mailbox](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-features#archive-mailbox), or in a public folder.</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="bc7ab-139">[要求集 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)中引入了 \* 对委托访问方案的支持（例如，从其他用户的邮箱共享的文件夹）。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-139">\* Support for delegate access scenarios (for example, folders shared from another user's mailbox) was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="bc7ab-140">共享邮箱支持现已提供预览版。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-140">Shared mailbox support is now in preview.</span></span> <span data-ttu-id="bc7ab-141">要了解详细信息，请参阅 [启用共享文件夹和共享邮箱方案](delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-141">To learn more, refer to [Enable shared folders and shared mailbox scenarios](delegate-access.md).</span></span>

- <span data-ttu-id="bc7ab-142">使用自定义窗体。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-142">Using a custom form.</span></span>

- <span data-ttu-id="bc7ab-143">通过[简单 MAPI](https://support.microsoft.com/topic/a3d3f856-eaf6-b6d8-3617-186c0a1123c5) 创建。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-143">Created through [Simple MAPI](https://support.microsoft.com/topic/a3d3f856-eaf6-b6d8-3617-186c0a1123c5).</span></span> <span data-ttu-id="bc7ab-144">如果 Outlook 关闭时，Office 用户从 Windows 上的 Office 应用程序创建或发送电子邮件，则将使用简单 MAPI。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-144">Simple MAPI is used when an Office user creates or sends an email from an Office application on Windows while Outlook is closed.</span></span> <span data-ttu-id="bc7ab-145">例如，用户在 Word 中工作时可以创建 Outlook 电子邮件，这会触发 Outlook 撰写窗口，而无需启动完整的 Outlook 应用程序。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-145">For example, a user can create an Outlook email while working in Word which triggers an Outlook compose window without launching the full Outlook application.</span></span> <span data-ttu-id="bc7ab-146">但是，如果用户从 Word 创建电子邮件时 Outlook 已在运行，则这不属于简单 MAPI 方案，因此只要满足其他激活要求，Outlook 加载项就会在撰写窗体中工作。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-146">If, however, Outlook is already running when the user creates the email from Word, that isn't a Simple MAPI scenario so Outlook add-ins work in the compose form as long as other activation requirements are met.</span></span>

<span data-ttu-id="bc7ab-147">通常，Outlook 可以为"已发送邮件"文件夹中的项目在阅读窗体中激活加载项，基于已知实体字符串匹配激活的加载项除外。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-147">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="bc7ab-148">欲了解其背后原因的详细信息，请参阅[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-148">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-clients"></a><span data-ttu-id="bc7ab-149">支持的客户端</span><span class="sxs-lookup"><span data-stu-id="bc7ab-149">Supported clients</span></span>

<span data-ttu-id="bc7ab-150">Windows 版 Outlook 2013 或更高版本、Mac 版 Outlook 2016 或更高版本、适用于本地 Exchange 2013 和更高版本的 Outlook 网页版、iOS 版 Outlook、Android 版 Outlook 及 Outlook 网页版和 Outlook.com 支持 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-150">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web and Outlook.com.</span></span> <span data-ttu-id="bc7ab-151">并非所有[客户端](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)都同时支持全部最新功能。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-151">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="bc7ab-152">请参阅有关这些功能的文章和 API 参考，了解它们可能在哪些应用程序中受支持或不受支持。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-152">Please refer to articles and API references for those features to see which applications they may or may not be supported in.</span></span>

## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="bc7ab-153">开始构建 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="bc7ab-153">Get started building Outlook add-ins</span></span>

<span data-ttu-id="bc7ab-154">要开始生成 Outlook 加载项，请尝试执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="bc7ab-154">To get started building Outlook add-ins, try the following:</span></span>

- <span data-ttu-id="bc7ab-155">[快速入门](../quickstarts/outlook-quickstart.md) - 生成简单的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-155">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="bc7ab-156">[教程](../tutorials/outlook-tutorial.md) - 了解如何创建将 GitHub Gist 插入新邮件的加载项。</span><span class="sxs-lookup"><span data-stu-id="bc7ab-156">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>

## <a name="see-also"></a><span data-ttu-id="bc7ab-157">另请参阅</span><span class="sxs-lookup"><span data-stu-id="bc7ab-157">See also</span></span>

- [<span data-ttu-id="bc7ab-158">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="bc7ab-158">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="bc7ab-159">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="bc7ab-159">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="bc7ab-160">Office 加载项的设计准则</span><span class="sxs-lookup"><span data-stu-id="bc7ab-160">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="bc7ab-161">许可 Office 和 SharePoint 加载项</span><span class="sxs-lookup"><span data-stu-id="bc7ab-161">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="bc7ab-162">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="bc7ab-162">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="bc7ab-163">将解决方案提交到 AppSource 和 Office 应用商店</span><span class="sxs-lookup"><span data-stu-id="bc7ab-163">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
