---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: b563d6cfc279a18a6a61f39c33a5ab42e1bd6984
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395706"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="20817-102">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="20817-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="20817-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="20817-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="20817-104">本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="20817-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="20817-105">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="20817-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="20817-106">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="20817-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="20817-107">在此要求集中引入的方法和属性应在使用前单独测试其可用性。</span><span class="sxs-lookup"><span data-stu-id="20817-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span>
>
> <span data-ttu-id="20817-108">要使用预览 API：</span><span class="sxs-lookup"><span data-stu-id="20817-108">To use preview APIs:</span></span>
>
> - <span data-ttu-id="20817-109">必须参考 CDN 上的 **beta** 库 (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)。</span><span class="sxs-lookup"><span data-stu-id="20817-109">You must reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span>
> - <span data-ttu-id="20817-110">还可能需要加入 [Office 预览体验计划](https://products.office.com/office-insider)才能访问更新的 Office 版本。</span><span class="sxs-lookup"><span data-stu-id="20817-110">You may also need to join the [Office Insider program](https://products.office.com/office-insider) for access to more recent Office builds.</span></span>

<span data-ttu-id="20817-111">预览要求集包括[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="20817-111">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="20817-112">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="20817-112">Features in preview</span></span>

<span data-ttu-id="20817-113">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="20817-113">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="20817-114">附件</span><span class="sxs-lookup"><span data-stu-id="20817-114">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="20817-115">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="20817-115">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="20817-116">新增了表示附件内容的对象。</span><span class="sxs-lookup"><span data-stu-id="20817-116">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="20817-117">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-117">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="20817-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="20817-118">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="20817-119">新增了一个方法，可将 base64 编码字符串形式的文件附加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="20817-119">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="20817-120">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-120">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="20817-121">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="20817-121">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="20817-122">新增了一个方法，可获取特定附件的内容。</span><span class="sxs-lookup"><span data-stu-id="20817-122">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="20817-123">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-123">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="20817-124">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="20817-124">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="20817-125">新增了一个方法，可在撰写模式下获取项目附件。</span><span class="sxs-lookup"><span data-stu-id="20817-125">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="20817-126">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-126">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="20817-127">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="20817-127">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="20817-128">新增了一个枚举，可指定应用于附件内容的格式设置。</span><span class="sxs-lookup"><span data-stu-id="20817-128">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="20817-129">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-129">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="20817-130">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="20817-130">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="20817-131">新增了一个枚举，可指定将附件添加至项目还是从项目中删除附件。</span><span class="sxs-lookup"><span data-stu-id="20817-131">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="20817-132">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-132">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="20817-133">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="20817-133">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="20817-134">向 `Item` 中添加了 `AttachmentsChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="20817-134">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="20817-135">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="block-on-send"></a><span data-ttu-id="20817-136">阻止发送</span><span class="sxs-lookup"><span data-stu-id="20817-136">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="20817-137">Event.completed</span><span class="sxs-lookup"><span data-stu-id="20817-137">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="20817-138">新增了可选参数 `options`，它是有效值为 `allowEvent` 的字典。</span><span class="sxs-lookup"><span data-stu-id="20817-138">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="20817-139">此值可用于取消执行事件。</span><span class="sxs-lookup"><span data-stu-id="20817-139">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="20817-140">**适用于**：Outlook 网页版（经典）、Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-140">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="categories"></a><span data-ttu-id="20817-141">类别</span><span class="sxs-lookup"><span data-stu-id="20817-141">Categories</span></span>

<span data-ttu-id="20817-142">在 Outlook 中，用户可以使用类别对邮件和约会进行颜色编码。</span><span class="sxs-lookup"><span data-stu-id="20817-142">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="20817-143">用户在其邮箱的主列表中定义类别。</span><span class="sxs-lookup"><span data-stu-id="20817-143">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="20817-144">然后，他们可以将一个或多个类别应用于项目。</span><span class="sxs-lookup"><span data-stu-id="20817-144">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="20817-145">iOS 版 Outlook 或 Android 版 Outlook 不支持此功能。</span><span class="sxs-lookup"><span data-stu-id="20817-145">This feature is not supported in Outlook for iOS or Outlook for Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="20817-146">Categories</span><span class="sxs-lookup"><span data-stu-id="20817-146">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="20817-147">新增了一个表示项目类别的对象。</span><span class="sxs-lookup"><span data-stu-id="20817-147">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="20817-148">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="20817-149">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="20817-149">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="20817-150">新增了一个表示类别详细信息（其名称以及对应的颜色）的对象。</span><span class="sxs-lookup"><span data-stu-id="20817-150">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="20817-151">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-151">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="20817-152">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="20817-152">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="20817-153">新增了一个表示邮箱上类别主列表的对象。</span><span class="sxs-lookup"><span data-stu-id="20817-153">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="20817-154">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-154">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="20817-155">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="20817-155">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="20817-156">新增了一个表示邮箱上类别主列表的属性。</span><span class="sxs-lookup"><span data-stu-id="20817-156">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="20817-157">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-157">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="20817-158">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="20817-158">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="20817-159">新增了一个表示项目上类别集的属性。</span><span class="sxs-lookup"><span data-stu-id="20817-159">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="20817-160">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-160">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="20817-161">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="20817-161">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="20817-162">新增了一个指定可用于与类别关联的颜色的枚举。</span><span class="sxs-lookup"><span data-stu-id="20817-162">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="20817-163">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="delegate-access"></a><span data-ttu-id="20817-164">委托访问</span><span class="sxs-lookup"><span data-stu-id="20817-164">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="20817-165">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="20817-165">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="20817-166">新增了一个对象，表示共享文件夹、日历或邮箱中的约会或邮件项目的属性。</span><span class="sxs-lookup"><span data-stu-id="20817-166">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="20817-167">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-167">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="20817-168">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="20817-168">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="20817-169">添加了用于获取已保存约会或邮件项目的 ID 的新方法。</span><span class="sxs-lookup"><span data-stu-id="20817-169">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="20817-170">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="20817-171">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="20817-171">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="20817-172">新增了一个对象，用于获取表示约会或邮件项目的 sharedProperties 的对象。</span><span class="sxs-lookup"><span data-stu-id="20817-172">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="20817-173">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-173">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="20817-174">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="20817-174">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="20817-175">新增了一个位标志枚举，可指定委派权限。</span><span class="sxs-lookup"><span data-stu-id="20817-175">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="20817-176">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="20817-177">SupportsSharedFolders manifest element</span><span class="sxs-lookup"><span data-stu-id="20817-177">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="20817-178">向 [DesktopFormFactor](../../manifest/desktopformfactor.md) 清单元素中添加了子元素。</span><span class="sxs-lookup"><span data-stu-id="20817-178">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="20817-179">它定义外接程序是否在代理应用场景中可用。</span><span class="sxs-lookup"><span data-stu-id="20817-179">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="20817-180">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="enhanced-location"></a><span data-ttu-id="20817-181">增强位置</span><span class="sxs-lookup"><span data-stu-id="20817-181">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="20817-182">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="20817-182">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="20817-183">新增了一个对象，显示约会的位置。</span><span class="sxs-lookup"><span data-stu-id="20817-183">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="20817-184">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="20817-185">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="20817-185">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="20817-186">新增了一个表示位置的对象。</span><span class="sxs-lookup"><span data-stu-id="20817-186">Added a new object that represents a location.</span></span> <span data-ttu-id="20817-187">只读。</span><span class="sxs-lookup"><span data-stu-id="20817-187">Read only.</span></span>

<span data-ttu-id="20817-188">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-188">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="20817-189">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="20817-189">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="20817-190">新增了一个表示位置 ID 的对象。</span><span class="sxs-lookup"><span data-stu-id="20817-190">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="20817-191">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-191">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="20817-192">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="20817-192">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="20817-193">新增了一个表示约会位置的属性。</span><span class="sxs-lookup"><span data-stu-id="20817-193">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="20817-194">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-194">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="20817-195">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="20817-195">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="20817-196">新增了一个用于指定约会位置类型的枚举。</span><span class="sxs-lookup"><span data-stu-id="20817-196">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="20817-197">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-197">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="20817-198">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="20817-198">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="20817-199">向 `Item` 中添加了 `EnhancedLocationsChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="20817-199">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="20817-200">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-200">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="20817-201">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="20817-201">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="20817-202">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="20817-202">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="20817-203">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="20817-203">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="20817-204">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="20817-204">**Available in**: Outlook on Windows (connected to Office 365), Outlook on the web (Classic)</span></span>

---

### <a name="internet-headers"></a><span data-ttu-id="20817-205">Internet 标头：</span><span class="sxs-lookup"><span data-stu-id="20817-205">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="20817-206">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="20817-206">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="20817-207">添加了一个表示邮件项目的自定义 Internet 标头的新对象。</span><span class="sxs-lookup"><span data-stu-id="20817-207">Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="20817-208">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-208">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[<span data-ttu-id="20817-209">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="20817-209">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheaders)

<span data-ttu-id="20817-210">添加了一个表示邮件项目的自定义 Internet 标头的新属性。</span><span class="sxs-lookup"><span data-stu-id="20817-210">Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="20817-211">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-211">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

---

### <a name="office-theme"></a><span data-ttu-id="20817-212">Office 主题</span><span class="sxs-lookup"><span data-stu-id="20817-212">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="20817-213">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="20817-213">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="20817-214">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="20817-214">Added ability to get Office theme.</span></span>

<span data-ttu-id="20817-215">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-215">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="20817-216">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="20817-216">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="20817-217">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="20817-217">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="20817-218">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="20817-218">**Available in**: Outlook on Windows (connected to Office 365)</span></span>

---

### <a name="sso"></a><span data-ttu-id="20817-219">SSO</span><span class="sxs-lookup"><span data-stu-id="20817-219">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="20817-220">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="20817-220">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="20817-221">添加了对 `getAccessTokenAsync` 的访问，使外接程序[能够访问](/outlook/add-ins/authenticate-a-user-with-an-sso-token) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="20817-221">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="20817-222">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="20817-222">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="20817-223">另请参阅</span><span class="sxs-lookup"><span data-stu-id="20817-223">See also</span></span>

- [<span data-ttu-id="20817-224">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="20817-224">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="20817-225">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="20817-225">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="20817-226">入门</span><span class="sxs-lookup"><span data-stu-id="20817-226">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="20817-227">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="20817-227">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
