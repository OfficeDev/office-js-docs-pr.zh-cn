---
title: 创建适用于阅读窗体的 Outlook 加载项
description: 阅读加载项是在 Outlook 中的阅读窗格或阅读检查器中激活的 Outlook 加载项。
ms.date: 04/12/2018
localization_priority: Priority
ms.openlocfilehash: 815234ed046b4c00b91f5acd6cd2c4dcd226dba2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605302"
---
# <a name="create-outlook-add-ins-for-read-forms"></a><span data-ttu-id="381e9-103">创建适用于阅读窗体的 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="381e9-103">Create Outlook add-ins for read forms</span></span>

<span data-ttu-id="381e9-p101">阅读外接程序是在 Outlook 中的阅读窗格或阅读检查器中激活的 Outlook 外接程序。与撰写外接程序（用户创建邮件或约会时激活的 Outlook 外接程序）不同，阅读外接程序在以下用户方案中可用：</span><span class="sxs-lookup"><span data-stu-id="381e9-p101">Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user is creating a message or appointment), read add-ins are available when users:</span></span> 

- <span data-ttu-id="381e9-106">查看电子邮件、会议请求、会议响应或会议取消。</span><span class="sxs-lookup"><span data-stu-id="381e9-106">View an email message, meeting request, meeting response, or meeting cancellation.</span></span>

   > [!NOTE]
   > <span data-ttu-id="381e9-107">Outlook 不会在阅读窗体中针对特定邮件类型激活外接程序，这些类型包括另一封邮件附加的项目、Outlook“草稿”文件夹中的项目，或以其他方式加密或保护的项目。</span><span class="sxs-lookup"><span data-stu-id="381e9-107">Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts folder, or items that are encrypted or protected in other ways.</span></span>
    
- <span data-ttu-id="381e9-108">查看用户参与的会议项。</span><span class="sxs-lookup"><span data-stu-id="381e9-108">View a meeting item in which the user is an attendee.</span></span>
    
- <span data-ttu-id="381e9-109">查看用户作为组织者的会议项目（仅限 Outlook 2013 和 Exchange 2013 的 RTM 版本）。</span><span class="sxs-lookup"><span data-stu-id="381e9-109">View a meeting item in which the user is the organizer (RTM release of Outlook 2013 and Exchange 2013 only).</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="381e9-p102">从 Office 2013 SP1 版本开始，如果用户查看由用户组织的会议项目，则只有撰写外接程序才能够激活并可用。这种情况下不再提供读取外接程序。</span><span class="sxs-lookup"><span data-stu-id="381e9-p102">Starting in the Office 2013 SP1 release, if the user is viewing a meeting item that the user has organized, only compose add-ins can activate and be available. Read add-ins are no longer available in this scenario.</span></span>


<span data-ttu-id="381e9-p103">在每个阅读应用场景中，当激活条件满足时，Outlook 便会激活加载项，用户可以在加载项栏中选择并打开在阅读窗格或阅读检查器中激活的加载项。下图展示了当用户在阅读包含地理位置地址的邮件时激活和打开的**必应地图**加载项。</span><span class="sxs-lookup"><span data-stu-id="381e9-p103">In each of these read scenarios, Outlook activates add-ins when their activation conditions are fulfilled, and users can choose and open activated add-ins in the add-in bar in the Reading Pane or read inspector. The following figure shows the **Bing Maps** add-in activated and opened as the user is reading a message that contains a geographic address.</span></span>


<span data-ttu-id="381e9-114">**加载项窗格，展示了包含地址的选定 Outlook 邮件的必应地图加载项的实际效果**</span><span class="sxs-lookup"><span data-stu-id="381e9-114">**The add-in pane showing the Bing Maps add-in in action for the selected Outlook message that contains an address**</span></span>

![Outlook 中的 Bing 地图邮件应用程序](../images/bing-maps-add-in.jpg)


## <a name="types-of-add-ins-available-in-read-mode"></a><span data-ttu-id="381e9-116">阅读模式下可用的外接程序的类型</span><span class="sxs-lookup"><span data-stu-id="381e9-116">Types of add-ins available in read mode</span></span>

<span data-ttu-id="381e9-117">阅读外接程序可以为下列类型的任意组合。</span><span class="sxs-lookup"><span data-stu-id="381e9-117">Read add-ins can be any combination of the following types.</span></span>

- [<span data-ttu-id="381e9-118">适用于 Outlook 的外接程序命令</span><span class="sxs-lookup"><span data-stu-id="381e9-118">Add-in commands for Outlook</span></span>](add-in-commands-for-outlook.md)   
- [<span data-ttu-id="381e9-119">上下文 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="381e9-119">Contextual Outlook add-ins</span></span>](contextual-outlook-add-ins.md)
    

## <a name="api-features-available-to-read-add-ins"></a><span data-ttu-id="381e9-120">阅读外接程序可用的 API 功能</span><span class="sxs-lookup"><span data-stu-id="381e9-120">API features available to read add-ins</span></span>

- <span data-ttu-id="381e9-121">要激活阅读窗体中的外接程序：请参阅[在清单中指定激活规则](activation-rules.md#specify-activation-rules-in-a-manifest)中的表 1。</span><span class="sxs-lookup"><span data-stu-id="381e9-121">For activating add-ins in read forms: see Table 1 in [Specify activation rules in a manifest](activation-rules.md#specify-activation-rules-in-a-manifest).</span></span>    
- [<span data-ttu-id="381e9-122">使用正则表达式激活规则显示 Outlook 外接程序</span><span class="sxs-lookup"><span data-stu-id="381e9-122">Use regular expression activation rules to show an Outlook add-in</span></span>](use-regular-expressions-to-show-an-outlook-add-in.md)    
- [<span data-ttu-id="381e9-123">将 Outlook 项中的字符串作为已知实体进行匹配</span><span class="sxs-lookup"><span data-stu-id="381e9-123">Match strings in an Outlook item as well-known entities</span></span>](match-strings-in-an-item-as-well-known-entities.md)    
- [<span data-ttu-id="381e9-124">从 Outlook 项中提取实体字符串</span><span class="sxs-lookup"><span data-stu-id="381e9-124">Extract entity strings from an Outlook item</span></span>](extract-entity-strings-from-an-item.md)   
- [<span data-ttu-id="381e9-125">从服务器获取 Outlook 项的附件</span><span class="sxs-lookup"><span data-stu-id="381e9-125">Get attachments of an Outlook item from the server</span></span>](get-attachments-of-an-outlook-item.md)
    

## <a name="see-also"></a><span data-ttu-id="381e9-126">另请参阅</span><span class="sxs-lookup"><span data-stu-id="381e9-126">See also</span></span>

- [<span data-ttu-id="381e9-127">编写第一个 Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="381e9-127">Write your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
