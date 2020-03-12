---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 03/04/2020
localization_priority: Normal
ms.openlocfilehash: 4365dab3d8dd1ddb876536b3030926d68a89ac49
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605671"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="0c243-102">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="0c243-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="0c243-103">Office JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="0c243-103">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0c243-104">本文档适用于**预览**[要求集](../../requirement-sets/outlook-api-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="0c243-104">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="0c243-105">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="0c243-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="0c243-106">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="0c243-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="0c243-107">预览要求集包括[要求集 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="0c243-107">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="0c243-108">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="0c243-108">Features in preview</span></span>

<span data-ttu-id="0c243-109">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="0c243-109">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="0c243-110">发送时追加</span><span class="sxs-lookup"><span data-stu-id="0c243-110">Append on send</span></span>

#### <a name="officebodyappendonsendasync"></a>[<span data-ttu-id="0c243-111">AppendOnSendAsync</span><span class="sxs-lookup"><span data-stu-id="0c243-111">Office.Body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="0c243-112">向`Body`对象添加了一个新函数，该函数在撰写模式下将数据追加到项正文的末尾。</span><span class="sxs-lookup"><span data-stu-id="0c243-112">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="0c243-113">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="0c243-113">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="0c243-114">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="0c243-114">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="0c243-115">向清单添加了一个新元素，其中`AppendOnSend`扩展权限必须包含在扩展权限的集合中。</span><span class="sxs-lookup"><span data-stu-id="0c243-115">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="0c243-116">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="0c243-116">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="0c243-117">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="0c243-117">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="0c243-118">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="0c243-118">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="0c243-119">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="0c243-119">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="0c243-120">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="0c243-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="0c243-121">Office 主题</span><span class="sxs-lookup"><span data-stu-id="0c243-121">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="0c243-122">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="0c243-122">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="0c243-123">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="0c243-123">Added ability to get Office theme.</span></span>

<span data-ttu-id="0c243-124">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="0c243-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="0c243-125">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="0c243-125">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="0c243-126">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="0c243-126">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="0c243-127">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="0c243-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="0c243-128">SSO</span><span class="sxs-lookup"><span data-stu-id="0c243-128">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="0c243-129">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="0c243-129">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="0c243-130">添加了对 `getAccessToken` 的访问，使外接程序[能够访问](../../../outlook/authenticate-a-user-with-an-sso-token.md) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="0c243-130">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="0c243-131">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="0c243-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="0c243-132">另请参阅</span><span class="sxs-lookup"><span data-stu-id="0c243-132">See also</span></span>

- [<span data-ttu-id="0c243-133">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="0c243-133">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="0c243-134">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="0c243-134">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="0c243-135">入门</span><span class="sxs-lookup"><span data-stu-id="0c243-135">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="0c243-136">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="0c243-136">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
