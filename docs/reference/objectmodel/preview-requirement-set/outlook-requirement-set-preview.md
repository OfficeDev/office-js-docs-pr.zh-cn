---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 10/18/2019
localization_priority: Priority
ms.openlocfilehash: 40bf17a6bfcc429b3de013a1b232a7c054b22768
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682527"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="3d444-102">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="3d444-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="3d444-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="3d444-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3d444-104">本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="3d444-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="3d444-105">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="3d444-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="3d444-106">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="3d444-106">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="3d444-107">预览要求集包括[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="3d444-107">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="3d444-108">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="3d444-108">Features in preview</span></span>

<span data-ttu-id="3d444-109">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="3d444-109">The following features are in preview.</span></span>

### <a name="attachments"></a><span data-ttu-id="3d444-110">附件</span><span class="sxs-lookup"><span data-stu-id="3d444-110">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="3d444-111">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="3d444-111">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="3d444-112">新增了表示附件内容的对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-112">Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="3d444-113">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-113">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="3d444-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="3d444-114">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="3d444-115">新增了一个方法，可将 base64 编码字符串形式的文件附加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3d444-115">Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="3d444-116">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-116">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[<span data-ttu-id="3d444-117">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="3d444-117">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

<span data-ttu-id="3d444-118">新增了一个方法，可获取特定附件的内容。</span><span class="sxs-lookup"><span data-stu-id="3d444-118">Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="3d444-119">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-119">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[<span data-ttu-id="3d444-120">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="3d444-120">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

<span data-ttu-id="3d444-121">新增了一个方法，可在撰写模式下获取项目附件。</span><span class="sxs-lookup"><span data-stu-id="3d444-121">Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="3d444-122">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-122">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="3d444-123">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="3d444-123">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="3d444-124">新增了一个枚举，可指定应用于附件内容的格式设置。</span><span class="sxs-lookup"><span data-stu-id="3d444-124">Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="3d444-125">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="3d444-126">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="3d444-126">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="3d444-127">新增了一个枚举，可指定将附件添加至项目还是从项目中删除附件。</span><span class="sxs-lookup"><span data-stu-id="3d444-127">Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="3d444-128">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="3d444-129">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="3d444-129">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="3d444-130">向 `Item` 中添加了 `AttachmentsChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="3d444-130">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="3d444-131">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-131">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="block-on-send"></a><span data-ttu-id="3d444-132">阻止发送</span><span class="sxs-lookup"><span data-stu-id="3d444-132">Block on send</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="3d444-133">Event.completed</span><span class="sxs-lookup"><span data-stu-id="3d444-133">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="3d444-134">新增了可选参数 `options`，它是有效值为 `allowEvent` 的字典。</span><span class="sxs-lookup"><span data-stu-id="3d444-134">Added a new optional parameter `options`, which is a dictionary with one valid value `allowEvent`.</span></span> <span data-ttu-id="3d444-135">此值可用于取消执行事件。</span><span class="sxs-lookup"><span data-stu-id="3d444-135">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="3d444-136">**适用于**：Outlook 网页版（经典）、Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-136">**Available in**: Outlook on the web (classic), Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="categories"></a><span data-ttu-id="3d444-137">类别</span><span class="sxs-lookup"><span data-stu-id="3d444-137">Categories</span></span>

<span data-ttu-id="3d444-138">在 Outlook 中，用户可以使用类别对邮件和约会进行颜色编码。</span><span class="sxs-lookup"><span data-stu-id="3d444-138">In Outlook, a user can group messages and appointments by using a category to color-code them.</span></span> <span data-ttu-id="3d444-139">用户在其邮箱的主列表中定义类别。</span><span class="sxs-lookup"><span data-stu-id="3d444-139">The user defines categories in a master list on their mailbox.</span></span> <span data-ttu-id="3d444-140">然后，他们可以将一个或多个类别应用于项目。</span><span class="sxs-lookup"><span data-stu-id="3d444-140">They can then apply one or more categories to an item.</span></span>

> [!NOTE]
> <span data-ttu-id="3d444-141">iOS 版 Outlook 或 Android 版 Outlook 不支持此功能。</span><span class="sxs-lookup"><span data-stu-id="3d444-141">This feature is not supported in Outlook on iOS or Android.</span></span>

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[<span data-ttu-id="3d444-142">Categories</span><span class="sxs-lookup"><span data-stu-id="3d444-142">Categories</span></span>](/javascript/api/outlook/office.categories)

<span data-ttu-id="3d444-143">新增了一个表示项目类别的对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-143">Added a new object that represents an item's categories.</span></span>

<span data-ttu-id="3d444-144">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-144">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[<span data-ttu-id="3d444-145">CategoryDetails</span><span class="sxs-lookup"><span data-stu-id="3d444-145">CategoryDetails</span></span>](/javascript/api/outlook/office.categorydetails)

<span data-ttu-id="3d444-146">新增了一个表示类别详细信息（其名称以及对应的颜色）的对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-146">Added a new object that represents a category's details (its name and associated color).</span></span>

<span data-ttu-id="3d444-147">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-147">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[<span data-ttu-id="3d444-148">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="3d444-148">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

<span data-ttu-id="3d444-149">新增了一个表示邮箱上类别主列表的对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-149">Added a new object that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="3d444-150">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-150">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[<span data-ttu-id="3d444-151">Office.context.mailbox.masterCategories</span><span class="sxs-lookup"><span data-stu-id="3d444-151">Office.context.mailbox.masterCategories</span></span>](/javascript/api/outlook/office.mailbox#mastercategories)

<span data-ttu-id="3d444-152">新增了一个表示邮箱上类别主列表的属性。</span><span class="sxs-lookup"><span data-stu-id="3d444-152">Added a new property that represents the categories master list on a mailbox.</span></span>

<span data-ttu-id="3d444-153">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-153">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[<span data-ttu-id="3d444-154">Office.context.mailbox.item.categories</span><span class="sxs-lookup"><span data-stu-id="3d444-154">Office.context.mailbox.item.categories</span></span>](/javascript/api/outlook/office.item#categories)

<span data-ttu-id="3d444-155">新增了一个表示项目上类别集的属性。</span><span class="sxs-lookup"><span data-stu-id="3d444-155">Added a new property that represents the set of categories on an item.</span></span>

<span data-ttu-id="3d444-156">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[<span data-ttu-id="3d444-157">Office.MailboxEnums.CategoryColor</span><span class="sxs-lookup"><span data-stu-id="3d444-157">Office.MailboxEnums.CategoryColor</span></span>](/javascript/api/outlook/office.mailboxenums.categorycolor)

<span data-ttu-id="3d444-158">新增了一个指定可用于与类别关联的颜色的枚举。</span><span class="sxs-lookup"><span data-stu-id="3d444-158">Added a new enum that specifies the colors available to be associated with categories.</span></span>

<span data-ttu-id="3d444-159">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-159">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="delegate-access"></a><span data-ttu-id="3d444-160">委托访问</span><span class="sxs-lookup"><span data-stu-id="3d444-160">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="3d444-161">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="3d444-161">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="3d444-162">新增了一个对象，表示共享文件夹、日历或邮箱中的约会或邮件项目的属性。</span><span class="sxs-lookup"><span data-stu-id="3d444-162">Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="3d444-163">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-163">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[<span data-ttu-id="3d444-164">Office.context.mailbox.item.getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="3d444-164">Office.context.mailbox.item.getItemIdAsync</span></span>](office.context.mailbox.item.md#getitemidasyncoptions-callback)

<span data-ttu-id="3d444-165">添加了用于获取已保存约会或邮件项目的 ID 的新方法。</span><span class="sxs-lookup"><span data-stu-id="3d444-165">Added a new method that gets the ID of a saved appointment or message item.</span></span>

<span data-ttu-id="3d444-166">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-166">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="3d444-167">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="3d444-167">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="3d444-168">新增了一个对象，用于获取表示约会或邮件项目的 sharedProperties 的对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-168">Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="3d444-169">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-169">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="3d444-170">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="3d444-170">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="3d444-171">新增了一个位标志枚举，可指定委派权限。</span><span class="sxs-lookup"><span data-stu-id="3d444-171">Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="3d444-172">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-172">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="3d444-173">SupportsSharedFolders manifest element</span><span class="sxs-lookup"><span data-stu-id="3d444-173">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="3d444-174">向 [DesktopFormFactor](../../manifest/desktopformfactor.md) 清单元素中添加了子元素。</span><span class="sxs-lookup"><span data-stu-id="3d444-174">Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="3d444-175">它定义外接程序是否在代理应用场景中可用。</span><span class="sxs-lookup"><span data-stu-id="3d444-175">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="3d444-176">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="enhanced-location"></a><span data-ttu-id="3d444-177">增强位置</span><span class="sxs-lookup"><span data-stu-id="3d444-177">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="3d444-178">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="3d444-178">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="3d444-179">新增了一个对象，显示约会的位置。</span><span class="sxs-lookup"><span data-stu-id="3d444-179">Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="3d444-180">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-180">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="3d444-181">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="3d444-181">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="3d444-182">新增了一个表示位置的对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-182">Added a new object that represents a location.</span></span> <span data-ttu-id="3d444-183">只读。</span><span class="sxs-lookup"><span data-stu-id="3d444-183">Read only.</span></span>

<span data-ttu-id="3d444-184">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-184">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="3d444-185">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="3d444-185">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="3d444-186">新增了一个表示位置 ID 的对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-186">Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="3d444-187">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[<span data-ttu-id="3d444-188">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="3d444-188">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

<span data-ttu-id="3d444-189">新增了一个表示约会位置的属性。</span><span class="sxs-lookup"><span data-stu-id="3d444-189">Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="3d444-190">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-190">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="3d444-191">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="3d444-191">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="3d444-192">新增了一个用于指定约会位置类型的枚举。</span><span class="sxs-lookup"><span data-stu-id="3d444-192">Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="3d444-193">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-193">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="3d444-194">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="3d444-194">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="3d444-195">向 `Item` 中添加了 `EnhancedLocationsChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="3d444-195">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="3d444-196">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-196">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="3d444-197">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="3d444-197">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="3d444-198">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="3d444-198">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="3d444-199">新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="3d444-199">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="3d444-200">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="3d444-200">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

### <a name="internet-headers"></a><span data-ttu-id="3d444-201">Internet 标头：</span><span class="sxs-lookup"><span data-stu-id="3d444-201">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="3d444-202">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="3d444-202">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="3d444-203">添加了一个表示邮件项目的自定义 Internet 标头的新对象。</span><span class="sxs-lookup"><span data-stu-id="3d444-203">Added a new object that represents the custom internet headers of a message item.</span></span> <span data-ttu-id="3d444-204">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="3d444-204">Compose mode only.</span></span>

<span data-ttu-id="3d444-205">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-205">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersjavascriptapioutlookofficemessagecomposeinternetheaders"></a>[<span data-ttu-id="3d444-206">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="3d444-206">Office.context.mailbox.item.internetHeaders</span></span>](/javascript/api/outlook/office.messagecompose#internetheaders)

<span data-ttu-id="3d444-207">添加了一个表示邮件项目的自定义 Internet 标头的新属性。</span><span class="sxs-lookup"><span data-stu-id="3d444-207">Added a new property that represents the custom internet headers on a message item.</span></span> <span data-ttu-id="3d444-208">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="3d444-208">Compose mode only.</span></span>

<span data-ttu-id="3d444-209">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-209">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetallinternetheadersasyncjavascriptapioutlookofficemessagereadgetallinternetheadersasync-options--callback-"></a>[<span data-ttu-id="3d444-210">Office.context.mailbox.item.getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="3d444-210">Office.context.mailbox.item.getAllInternetHeadersAsync</span></span>](/javascript/api/outlook/office.messageread#getallinternetheadersasync-options--callback-)

<span data-ttu-id="3d444-211">添加了一种新方法，可获取邮件项目的所有 Internet 标头。</span><span class="sxs-lookup"><span data-stu-id="3d444-211">Added a new method that gets all the internet headers for a message item.</span></span> <span data-ttu-id="3d444-212">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d444-212">Read mode only.</span></span>

<span data-ttu-id="3d444-213">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-213">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="office-theme"></a><span data-ttu-id="3d444-214">Office 主题</span><span class="sxs-lookup"><span data-stu-id="3d444-214">Office theme</span></span>

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[<span data-ttu-id="3d444-215">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="3d444-215">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="3d444-216">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="3d444-216">Added ability to get Office theme.</span></span>

<span data-ttu-id="3d444-217">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-217">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="3d444-218">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="3d444-218">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="3d444-219">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="3d444-219">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="3d444-220">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d444-220">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

### <a name="sso"></a><span data-ttu-id="3d444-221">SSO</span><span class="sxs-lookup"><span data-stu-id="3d444-221">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="3d444-222">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3d444-222">Office.context.auth.getAccessTokenAsync</span></span>](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="3d444-223">添加了对 `getAccessTokenAsync` 的访问，使外接程序[能够访问](/outlook/add-ins/authenticate-a-user-with-an-sso-token) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="3d444-223">Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="3d444-224">**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="3d444-224">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="3d444-225">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3d444-225">See also</span></span>

- [<span data-ttu-id="3d444-226">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="3d444-226">Outlook add-ins</span></span>](/outlook/add-ins/)
- [<span data-ttu-id="3d444-227">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="3d444-227">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="3d444-228">入门</span><span class="sxs-lookup"><span data-stu-id="3d444-228">Get started</span></span>](/outlook/add-ins/quick-start)
- [<span data-ttu-id="3d444-229">要求集和支持的客户端</span><span class="sxs-lookup"><span data-stu-id="3d444-229">Requirement sets and supported clients</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
