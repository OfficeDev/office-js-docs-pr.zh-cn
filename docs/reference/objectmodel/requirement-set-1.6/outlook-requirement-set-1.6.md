---
title: Outlook 外接程序 API 要求集 1.6
description: 作为邮箱 API 1.6 的一部分引入的 Outlook 外接程序和 Office JavaScript Api 的功能和 Api。
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: adcfcb49a76fd3f0df2c2c3acfc6e1861a02f3b1
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431449"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="0924f-103">Outlook 外接程序 API 要求集 1.6</span><span class="sxs-lookup"><span data-stu-id="0924f-103">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="0924f-104">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="0924f-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="0924f-105">本文档适用于最新要求集之外的[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="0924f-105">This documentation is for a [requirement set](../../requirement-sets/outlook-api-requirement-sets.md) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="0924f-106">1.6 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="0924f-106">What's new in 1.6?</span></span>

<span data-ttu-id="0924f-107">要求集 1.6 包括[要求集 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="0924f-107">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="0924f-108">它还添加了下列功能。</span><span class="sxs-lookup"><span data-stu-id="0924f-108">It added the following features.</span></span>

- <span data-ttu-id="0924f-109">为上下文外接程序添加了新 API，以获取用户选择用于激活外接程序的实体或 RegEx 匹配项。</span><span class="sxs-lookup"><span data-stu-id="0924f-109">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="0924f-110">添加了新 API，用于打开新邮件窗体。</span><span class="sxs-lookup"><span data-stu-id="0924f-110">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="0924f-111">添加了通过外接程序来确定用户邮箱的帐户类型的功能。</span><span class="sxs-lookup"><span data-stu-id="0924f-111">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="0924f-112">更改日志</span><span class="sxs-lookup"><span data-stu-id="0924f-112">Change log</span></span>

- <span data-ttu-id="0924f-113">添加了 [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods)：添加了一个新函数，该函数可用于获取用户选择的突出显示匹配项中的实体。</span><span class="sxs-lookup"><span data-stu-id="0924f-113">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="0924f-114">突出显示的匹配项适用于上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="0924f-114">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="0924f-115">添加了 [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods)：添加了一个新函数，该函数可用于返回突出显示匹配项中与清单 XML 文件中定义的正则表达式匹配的字符串值。</span><span class="sxs-lookup"><span data-stu-id="0924f-115">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="0924f-116">突出显示的匹配项适用于上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="0924f-116">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="0924f-117">添加了 [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods)：添加了一个新函数，该函数将打开新邮件窗体。</span><span class="sxs-lookup"><span data-stu-id="0924f-117">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="0924f-118">添加了 [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#accounttype)：向指示用户帐户类型的用户配置文件添加了一个新成员。</span><span class="sxs-lookup"><span data-stu-id="0924f-118">Added [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#accounttype): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="0924f-119">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0924f-119">See also</span></span>

- [<span data-ttu-id="0924f-120">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="0924f-120">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="0924f-121">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="0924f-121">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="0924f-122">入门</span><span class="sxs-lookup"><span data-stu-id="0924f-122">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="0924f-123">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="0924f-123">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
