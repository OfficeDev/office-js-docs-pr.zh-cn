---
title: Outlook 外接程序 API 要求集 1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e1f920c259ca1ef8a137bab07132b015d9c75d2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451729"
---
# <a name="outlook-add-in-api-requirement-set-16"></a><span data-ttu-id="eec45-102">Outlook 外接程序 API 要求集 1.6</span><span class="sxs-lookup"><span data-stu-id="eec45-102">Outlook add-in API requirement set 1.6</span></span>

<span data-ttu-id="eec45-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="eec45-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="eec45-104">本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="eec45-104">This documentation is for a [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) other than the latest requirement set.</span></span>

## <a name="whats-new-in-16"></a><span data-ttu-id="eec45-105">1.6 中的新增功能有哪些？</span><span class="sxs-lookup"><span data-stu-id="eec45-105">What's new in 1.6?</span></span>

<span data-ttu-id="eec45-106">要求集 1.6 包括[要求集 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="eec45-106">Requirement set 1.6 includes all of the features of [Requirement set 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md).</span></span> <span data-ttu-id="eec45-107">它还添加了下列功能。</span><span class="sxs-lookup"><span data-stu-id="eec45-107">It added the following features.</span></span>

- <span data-ttu-id="eec45-108">为上下文外接程序添加了新 API，以获取用户选择用于激活外接程序的实体或 RegEx 匹配项。</span><span class="sxs-lookup"><span data-stu-id="eec45-108">Added new APIs for contextual add-ins to get the entity or RegEx match that the user selected to activate the add-in.</span></span>
- <span data-ttu-id="eec45-109">添加了新 API，用于打开新邮件窗体。</span><span class="sxs-lookup"><span data-stu-id="eec45-109">Added a new API to open a new message form.</span></span>
- <span data-ttu-id="eec45-110">添加了通过外接程序来确定用户邮箱的帐户类型的功能。</span><span class="sxs-lookup"><span data-stu-id="eec45-110">Added the ability for the add-in to determine the account type of the user's mailbox.</span></span>

### <a name="change-log"></a><span data-ttu-id="eec45-111">更改日志</span><span class="sxs-lookup"><span data-stu-id="eec45-111">Change log</span></span>

- <span data-ttu-id="eec45-112">添加了 [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities)：添加了一个新函数，该函数可用于获取用户选择的突出显示匹配项中的实体。</span><span class="sxs-lookup"><span data-stu-id="eec45-112">Added [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities): Adds a new function that gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="eec45-113">突出显示的匹配项适用于上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="eec45-113">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="eec45-114">添加了 [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object)：添加了一个新函数，该函数可用于返回突出显示匹配项中与清单 XML 文件中定义的正则表达式匹配的字符串值。</span><span class="sxs-lookup"><span data-stu-id="eec45-114">Added [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object): Adds a new function that returns string values in a highlighted match that match the regular expressions defined in the manifest XML file.</span></span> <span data-ttu-id="eec45-115">突出显示的匹配项适用于上下文外接程序。</span><span class="sxs-lookup"><span data-stu-id="eec45-115">Highlighted matches apply to contextual add-ins.</span></span>
- <span data-ttu-id="eec45-116">添加了 [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)：添加了一个新函数，该函数将打开新邮件窗体。</span><span class="sxs-lookup"><span data-stu-id="eec45-116">Added [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters): Adds a new function that opens a new message form.</span></span>
- <span data-ttu-id="eec45-117">添加了 [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string)：向指示用户帐户类型的用户配置文件添加了一个新成员。</span><span class="sxs-lookup"><span data-stu-id="eec45-117">Added [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string): Adds a new member to the user profile that indicates the type of the user's account.</span></span>

## <a name="see-also"></a><span data-ttu-id="eec45-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="eec45-118">See also</span></span>

- [<span data-ttu-id="eec45-119">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="eec45-119">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="eec45-120">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="eec45-120">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="eec45-121">入门</span><span class="sxs-lookup"><span data-stu-id="eec45-121">Get started</span></span>](/outlook/add-ins/quick-start)
