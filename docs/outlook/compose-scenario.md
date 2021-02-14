---
title: 创建适用于撰写窗体的 Outlook 加载项
description: 了解有关适用于撰写窗体的 Outlook 加载项的方案和功能。
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 9156f2e1393c27eea359a6b63da47bc24a8a6828
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234252"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a><span data-ttu-id="3cd81-103">创建适用于撰写窗体的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="3cd81-103">Create Outlook add-ins for compose forms</span></span>

<span data-ttu-id="3cd81-104">从用于 Office 加载项清单的架构版本 1.1 和 Office.js 的版本 1.1 开始，可以创建撰写加载项，这是在撰写窗体中激活的 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="3cd81-104">Starting with version 1.1 of the schema for Office Add-ins manifests and v1.1 of Office.js, you can create compose add-ins, which are Outlook add-ins activated in compose forms.</span></span> <span data-ttu-id="3cd81-105">与读取加载项（用户查看邮件或约会时，以阅读模式激活的 Outlook 加载项）相反，在以下用户应用场景中可以使用撰写加载项：</span><span class="sxs-lookup"><span data-stu-id="3cd81-105">In contrast with read add-ins (Outlook add-ins that are activated in read mode when a user is viewing a message or appointment), compose add-ins are available in the following user scenarios:</span></span>

- <span data-ttu-id="3cd81-106">在撰写窗体中撰写新的邮件、会议请求或约会。</span><span class="sxs-lookup"><span data-stu-id="3cd81-106">Composing a new message, meeting request, or appointment in a compose form.</span></span>

- <span data-ttu-id="3cd81-107">查看或编辑现有约会或用户是组织者的会议项目。</span><span class="sxs-lookup"><span data-stu-id="3cd81-107">Viewing or editing an existing appointment, or meeting item in which the user is the organizer.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="3cd81-108">如果用户使用的是 Outlook 2013 和 Exchange 2013 的 RTM 版本，在查看由用户组织的会议项目时，用户可以发现读取加载项是可用的。</span><span class="sxs-lookup"><span data-stu-id="3cd81-108">If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available.</span></span> <span data-ttu-id="3cd81-109">从 Office 2013 SP1 版本开始进行了更改，在同一方案中，只有撰写外接程序能够激活并可用。</span><span class="sxs-lookup"><span data-stu-id="3cd81-109">Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.</span></span>

- <span data-ttu-id="3cd81-110">在单独的撰写窗体中撰写内嵌响应邮件或答复邮件。</span><span class="sxs-lookup"><span data-stu-id="3cd81-110">Composing an inline response message or replying to a message in a separate compose form.</span></span>

- <span data-ttu-id="3cd81-111">编辑会议请求或会议项目答复（“接受”、“暂定”或“拒绝”）。</span><span class="sxs-lookup"><span data-stu-id="3cd81-111">Editing a response (**Accept**, **Tentative**, or **Decline**) to a meeting request or meeting item.</span></span>

- <span data-ttu-id="3cd81-112">建议新的会议项目时间。</span><span class="sxs-lookup"><span data-stu-id="3cd81-112">Proposing a new time for a meeting item.</span></span>

- <span data-ttu-id="3cd81-113">转发或答复会议请求或会议项目。</span><span class="sxs-lookup"><span data-stu-id="3cd81-113">Forwarding or replying to a meeting request or meeting item.</span></span>

<span data-ttu-id="3cd81-114">在每个撰写方案中，显示由加载项定义的任何加载项命令按钮。</span><span class="sxs-lookup"><span data-stu-id="3cd81-114">In each of these compose scenarios, any add-in command buttons defined by the add-in are shown.</span></span> <span data-ttu-id="3cd81-115">对于未执行加载项命令的较旧加载项，用户可以选择功能区中的“Office 加载项”打开加载项选择窗格，然后选择并启动撰写加载项。</span><span class="sxs-lookup"><span data-stu-id="3cd81-115">For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in.</span></span> <span data-ttu-id="3cd81-116">下图显示了撰写窗体中的加载项命令。</span><span class="sxs-lookup"><span data-stu-id="3cd81-116">The following figure shows add-in commands in a compose form.</span></span>

![显示 Outlook 撰写窗体，其中包含外接程序命令。](../images/compose-form-commands.png)

<span data-ttu-id="3cd81-118">下图显示了外接程序选择窗格，该窗格由两个不实施外接程序命令的撰写外接程序组成，当用户在 Outlook 中撰写内嵌答复时将激活这两个撰写外接程序。</span><span class="sxs-lookup"><span data-stu-id="3cd81-118">The following figure shows the add-in selection pane consisting of two compose add-ins that do not implement add-in commands, activated when the user is composing an inline reply in Outlook.</span></span>

![为编写项目激活的模板邮件应用程序](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a><span data-ttu-id="3cd81-120">撰写模式下可用的外接程序的类型</span><span class="sxs-lookup"><span data-stu-id="3cd81-120">Types of add-ins available in compose mode</span></span>

<span data-ttu-id="3cd81-121">撰写加载项作为[用于 Outlook 的加载项命令](add-in-commands-for-outlook.md)实现。</span><span class="sxs-lookup"><span data-stu-id="3cd81-121">Compose add-ins are implemented as [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span> <span data-ttu-id="3cd81-122">若要激活用于撰写电子邮件或会议答复的加载项，则加载项在清单中包括 [MessageComposeCommandSurface 扩展点元素](../reference/manifest/extensionpoint.md#messagecomposecommandsurface)。</span><span class="sxs-lookup"><span data-stu-id="3cd81-122">To activate add-ins for composing email or meeting responses, add-ins include a [MessageComposeCommandSurface extension point element](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) in the manifest.</span></span> <span data-ttu-id="3cd81-123">若要激活用于撰写或编辑用户是组织者的约会或会议的加载项，则加载项包括 [AppointmentOrganizerCommandSurface 扩展点元素](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface)。</span><span class="sxs-lookup"><span data-stu-id="3cd81-123">To activate add-ins for composing or editing appointments or meetings where the user is the organizer, add-ins include a [AppointmentOrganizerCommandSurface extension point element](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span></span>

> [!NOTE]
> <span data-ttu-id="3cd81-124">为不支持加载项命令在包含在 [OfficeApp](../reference/manifest/officeapp.md) 元素中的 [Rule](../reference/manifest/rule.md) 元素使用[激活规则](activation-rules.md)的服务器或客户端开发的加载项。</span><span class="sxs-lookup"><span data-stu-id="3cd81-124">Add-ins developed for servers or clients that do not support add-in commands use [activation rules](activation-rules.md) in a [Rule](../reference/manifest/rule.md) element contained in the [OfficeApp](../reference/manifest/officeapp.md) element.</span></span> <span data-ttu-id="3cd81-125">除非加载项是为较早的客户端和服务器专门开发的，否则新加载项应使用加载项命令。</span><span class="sxs-lookup"><span data-stu-id="3cd81-125">Unless the add-in is being specifically developed for older clients and servers, new add-ins should use add-in commands.</span></span>

## <a name="api-features-available-to-compose-add-ins"></a><span data-ttu-id="3cd81-126">撰写加载项可用的 API 功能</span><span class="sxs-lookup"><span data-stu-id="3cd81-126">API features available to compose add-ins</span></span>

- [<span data-ttu-id="3cd81-127">在 Outlook 的撰写窗体中添加和删除项目附件</span><span class="sxs-lookup"><span data-stu-id="3cd81-127">Add and remove attachments to an item in a compose form in Outlook</span></span>](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [<span data-ttu-id="3cd81-128">在 Outlook 的撰写窗体中获取和设置项目数据</span><span class="sxs-lookup"><span data-stu-id="3cd81-128">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="3cd81-129">在 Outlook 中撰写约会或邮件时获取、设置或添加收件人</span><span class="sxs-lookup"><span data-stu-id="3cd81-129">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)
- [<span data-ttu-id="3cd81-130">在 Outlook 中撰写约会或邮件时获取或设置主题</span><span class="sxs-lookup"><span data-stu-id="3cd81-130">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="3cd81-131">在 Outlook 中撰写约会或邮件时将数据插入到正文中</span><span class="sxs-lookup"><span data-stu-id="3cd81-131">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="3cd81-132">在 Outlook 中撰写约会时获取或设置位置</span><span class="sxs-lookup"><span data-stu-id="3cd81-132">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="3cd81-133">在 Outlook 中撰写约会时获取或设置时间</span><span class="sxs-lookup"><span data-stu-id="3cd81-133">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a><span data-ttu-id="3cd81-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3cd81-134">See also</span></span>

- [<span data-ttu-id="3cd81-135">适用于 Office 的 Outlook 加载项入门</span><span class="sxs-lookup"><span data-stu-id="3cd81-135">Get Started with Outlook add-ins for Office</span></span>](../quickstarts/outlook-quickstart.md)
