---
title: Outlook 加载项概述
description: Outlook 加载项由第三方使用基于 Web 的平台集成到 Outlook 中。
ms.date: 06/15/2021
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 3fb6c47d0dc2b41ecf657ea4d453c2ffcb8a8902
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076754"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="c7a95-103">Outlook 加载项概述</span><span class="sxs-lookup"><span data-stu-id="c7a95-103">Outlook add-ins overview</span></span>

<span data-ttu-id="c7a95-p101">Outlook 加载项是由第三方通过使用基于 Web 的平台构建到 Outlook 中的集成。Outlook 加载项具有三个关键方面：</span><span class="sxs-lookup"><span data-stu-id="c7a95-p101">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="c7a95-106">相同的加载项和业务逻辑可跨桌面（Windows 版和 Mac 版 Outlook）、Web（Microsoft 365 和 Outlook.com）和移动平台使用。</span><span class="sxs-lookup"><span data-stu-id="c7a95-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Microsoft 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="c7a95-107">Outlook 外接程序包括一个清单，其中介绍了如何将外接程序集成到 Outlook（例如，按钮或任务窗格）中，以及构成外接程序 UI 和业务逻辑的 JavaScript/HTML 代码。</span><span class="sxs-lookup"><span data-stu-id="c7a95-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="c7a95-108">最终用户或管理员可以从 [AppSource](https://appsource.microsoft.com) 获取 Outlook 加载项，也可以进行[旁加载](sideload-outlook-add-ins-for-testing.md)。</span><span class="sxs-lookup"><span data-stu-id="c7a95-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="c7a95-p102">Outlook 外接程序与 COM 或 VSTO 外接程序（特定于在 Windows 上运行的 Outlook 的较早集成项）不同。与 COM 外接程序不同的是，Outlook 外接程序不具有任何实际安装到用户设备或 Outlook 客户端的代码。对于 Outlook 外接程序，Outlook 读取清单并挂钩在 UI 中指定的控件，然后加载 JavaScript 和 HTML。web 部件全部在沙盒的浏览器的上下文中执行。</span><span class="sxs-lookup"><span data-stu-id="c7a95-p102">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="c7a95-p103">支持加载项的 Outlook 项目包括电子邮件、会议请求、响应和取消及约会。每个 Outlook 加载项均定义其可用的上下文，包括项目类型以及用户是在阅读还是撰写项目。</span><span class="sxs-lookup"><span data-stu-id="c7a95-p103">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a><span data-ttu-id="c7a95-115">扩展点</span><span class="sxs-lookup"><span data-stu-id="c7a95-115">Extension points</span></span>

<span data-ttu-id="c7a95-p104">扩展点是加载项与 Outlook 集成的方式。以下是执行此操作的方法：</span><span class="sxs-lookup"><span data-stu-id="c7a95-p104">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="c7a95-p105">加载项可以声明出现在所有邮件和约会的命令界面中的按钮。有关详细信息，请参阅 [用于 Outlook 的加载项命令](add-in-commands-for-outlook.md)。</span><span class="sxs-lookup"><span data-stu-id="c7a95-p105">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="c7a95-120">**功能区上具有命令按钮的加载项**</span><span class="sxs-lookup"><span data-stu-id="c7a95-120">**An add-in with command buttons on the ribbon**</span></span>

    ![加载项命令无 UI 形状。](../images/uiless-command-shape.png)

- <span data-ttu-id="c7a95-p106">加载项可以在邮件和约会中中断与正则表达式匹配项或检测实体的链接。 有关详细信息，请参阅 [上下文 Outlook 加载项](contextual-outlook-add-ins.md)。</span><span class="sxs-lookup"><span data-stu-id="c7a95-p106">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="c7a95-124">**用于突出显示的实体（地址）的上下文相关加载项**</span><span class="sxs-lookup"><span data-stu-id="c7a95-124">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![在卡片中显示上下文相关应用。](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="c7a95-126">外接程序可用的邮箱项目</span><span class="sxs-lookup"><span data-stu-id="c7a95-126">Mailbox items available to add-ins</span></span>

<span data-ttu-id="c7a95-127">当用户正在撰写或阅读邮件或约会，而不是其他项目类型时，Outlook 加载项会激活。</span><span class="sxs-lookup"><span data-stu-id="c7a95-127">Outlook add-ins activate when the user is composing or reading a message or appointment, but not other item types.</span></span> <span data-ttu-id="c7a95-128">但是，如果撰写或阅读窗体中的当前邮件项目为以下项之一，则 Outlook *不会* 激活邮件加载项：</span><span class="sxs-lookup"><span data-stu-id="c7a95-128">However, add-ins are *not* activated if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="c7a95-p108">使用信息权限管理 (IRM) 进行保护，或使用其他保护方式进行加密。数字签名邮件便是其中一个例子，因为数字签名依赖于这些机制之一。</span><span class="sxs-lookup"><span data-stu-id="c7a95-p108">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

  > [!IMPORTANT]
  >
  > - <span data-ttu-id="c7a95-131">加载项在与 Microsoft 365 订阅相关联的 Outlook 电子签名邮件上激活。</span><span class="sxs-lookup"><span data-stu-id="c7a95-131">Add-ins activate on digitally signed messages in Outlook associated with a Microsoft 365 subscription.</span></span> <span data-ttu-id="c7a95-132">在Windows上，这个支持是通过8711.1000版本中引入的。</span><span class="sxs-lookup"><span data-stu-id="c7a95-132">On Windows, this support was introduced with build 8711.1000.</span></span>
  >
  > - <span data-ttu-id="c7a95-133">现在，Windows 版 Outlook 从内部版本 13229.10000 开始可以在受 IRM 保护的项目上激活加载项。</span><span class="sxs-lookup"><span data-stu-id="c7a95-133">Starting with Outlook build 13229.10000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="c7a95-134">有关处于预览阶段的此功能的详细信息，请参阅 [在受信息权限管理 (IRM) 保护的项目上激活加载项](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm)。</span><span class="sxs-lookup"><span data-stu-id="c7a95-134">For more information about this feature in preview, refer to [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="c7a95-135">具有邮件类别 IPM.Report.\* 的送达报告或通知，包括送达和未送达报告 (NDR)，以及已读、未读和延迟通知。</span><span class="sxs-lookup"><span data-stu-id="c7a95-135">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="c7a95-136">属于其他邮件的附件的 .msg 或 .eml 文件。</span><span class="sxs-lookup"><span data-stu-id="c7a95-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="c7a95-137">从文件系统打开的 .msg 或 .eml 文件。</span><span class="sxs-lookup"><span data-stu-id="c7a95-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="c7a95-138">在 [组邮箱](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) 中、在共享邮箱 \* 中、在其他用户的邮箱 \* 中、在存档邮箱中或在公用文件夹中。</span><span class="sxs-lookup"><span data-stu-id="c7a95-138">In a [group mailbox](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), in a shared mailbox\*, in another user's mailbox\*, in an archive mailbox, or in a public folder.</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="c7a95-139">[要求集 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)中引入了 \* 对委托访问方案的支持（例如，从其他用户的邮箱共享的文件夹）。</span><span class="sxs-lookup"><span data-stu-id="c7a95-139">\* Support for delegate access scenarios (for example, folders shared from another user's mailbox) was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="c7a95-140">共享邮箱支持现已提供预览版。</span><span class="sxs-lookup"><span data-stu-id="c7a95-140">Shared mailbox support is now in preview.</span></span> <span data-ttu-id="c7a95-141">要了解详细信息，请参阅 [启用共享文件夹和共享邮箱方案](delegate-access.md)。</span><span class="sxs-lookup"><span data-stu-id="c7a95-141">To learn more, refer to [Enable shared folders and shared mailbox scenarios](delegate-access.md).</span></span>

- <span data-ttu-id="c7a95-142">使用自定义窗体。</span><span class="sxs-lookup"><span data-stu-id="c7a95-142">Using a custom form.</span></span>

<span data-ttu-id="c7a95-143">通常，Outlook 可以为"已发送邮件"文件夹中的项目在阅读窗体中激活加载项，基于已知实体字符串匹配激活的加载项除外。</span><span class="sxs-lookup"><span data-stu-id="c7a95-143">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="c7a95-144">欲了解其背后原因的详细信息，请参阅[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。</span><span class="sxs-lookup"><span data-stu-id="c7a95-144">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-clients"></a><span data-ttu-id="c7a95-145">支持的客户端</span><span class="sxs-lookup"><span data-stu-id="c7a95-145">Supported clients</span></span>

<span data-ttu-id="c7a95-146">Windows 版 Outlook 2013 或更高版本、Mac 版 Outlook 2016 或更高版本、适用于本地 Exchange 2013 和更高版本的 Outlook 网页版、iOS 版 Outlook、Android 版 Outlook 及 Outlook 网页版和 Outlook.com 支持 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="c7a95-146">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web and Outlook.com.</span></span> <span data-ttu-id="c7a95-147">并非所有[客户端](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)都同时支持全部最新功能。</span><span class="sxs-lookup"><span data-stu-id="c7a95-147">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="c7a95-148">请参阅有关这些功能的文章和 API 参考，了解它们可能在哪些应用程序中受支持或不受支持。</span><span class="sxs-lookup"><span data-stu-id="c7a95-148">Please refer to articles and API references for those features to see which applications they may or may not be supported in.</span></span>

## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="c7a95-149">开始构建 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="c7a95-149">Get started building Outlook add-ins</span></span>

<span data-ttu-id="c7a95-150">要开始生成 Outlook 加载项，请尝试执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="c7a95-150">To get started building Outlook add-ins, try the following:</span></span>

- <span data-ttu-id="c7a95-151">[快速入门](../quickstarts/outlook-quickstart.md) - 生成简单的任务窗格。</span><span class="sxs-lookup"><span data-stu-id="c7a95-151">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="c7a95-152">[教程](../tutorials/outlook-tutorial.md) - 了解如何创建将 GitHub Gist 插入新邮件的加载项。</span><span class="sxs-lookup"><span data-stu-id="c7a95-152">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>

## <a name="see-also"></a><span data-ttu-id="c7a95-153">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c7a95-153">See also</span></span>

- [<span data-ttu-id="c7a95-154">了解 Microsoft 365 开发人员计划</span><span class="sxs-lookup"><span data-stu-id="c7a95-154">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="c7a95-155">Office 加载项开发最佳做法</span><span class="sxs-lookup"><span data-stu-id="c7a95-155">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="c7a95-156">Office 加载项的设计准则</span><span class="sxs-lookup"><span data-stu-id="c7a95-156">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="c7a95-157">许可 Office 和 SharePoint 加载项</span><span class="sxs-lookup"><span data-stu-id="c7a95-157">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="c7a95-158">发布 Office 加载项</span><span class="sxs-lookup"><span data-stu-id="c7a95-158">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="c7a95-159">将解决方案提交到 AppSource 和 Office 应用商店</span><span class="sxs-lookup"><span data-stu-id="c7a95-159">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
