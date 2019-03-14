---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 03/07/2019
localization_priority: Priority
ms.openlocfilehash: b1a3f5c675b2bcb43003ad15b3358e3febd80260
ms.sourcegitcommit: 8e7b7b0cfb68b91a3a95585d094cf5f5ffd00178
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/09/2019
ms.locfileid: "30512858"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="8e624-102">Outlook 外接程序 API 预览要求集</span><span class="sxs-lookup"><span data-stu-id="8e624-102">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="8e624-103">适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。</span><span class="sxs-lookup"><span data-stu-id="8e624-103">The Outlook add-in API subset of the JavaScript API for Office includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="8e624-104">本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="8e624-104">This documentation is for a **preview** [requirement set](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).</span></span> <span data-ttu-id="8e624-105">此要求集尚未完全实现，客户端不会准确报告对它的支持。</span><span class="sxs-lookup"><span data-stu-id="8e624-105">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="8e624-106">不应在外接程序清单中指定此要求集。</span><span class="sxs-lookup"><span data-stu-id="8e624-106">You should not specify this requirement set in your add-in manifest.</span></span> <span data-ttu-id="8e624-107">在此要求集中引入的方法和属性应在使用前单独测试其可用性。</span><span class="sxs-lookup"><span data-stu-id="8e624-107">Methods and properties that are introduced in this requirement set should be individually tested for availability before using them.</span></span> <span data-ttu-id="8e624-108">此外，你还需要加入 [Office 预览体验成员计划](https://products.office.com/office-insider)。</span><span class="sxs-lookup"><span data-stu-id="8e624-108">You may also need to join the [Office Insider program](https://products.office.com/office-insider).</span></span>

<span data-ttu-id="8e624-109">预览要求集包括[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) 的所有功能。</span><span class="sxs-lookup"><span data-stu-id="8e624-109">The Preview Requirement set includes all of the features of [Requirement set 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="8e624-110">预览阶段的功能</span><span class="sxs-lookup"><span data-stu-id="8e624-110">Features in preview</span></span>

<span data-ttu-id="8e624-111">以下是预览版中的功能。</span><span class="sxs-lookup"><span data-stu-id="8e624-111">The following features are in preview.</span></span>

### <a name="add-in-commands"></a><span data-ttu-id="8e624-112">加载项命令</span><span class="sxs-lookup"><span data-stu-id="8e624-112">Add-in commands</span></span>

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[<span data-ttu-id="8e624-113">Event.completed</span><span class="sxs-lookup"><span data-stu-id="8e624-113">Event.completed</span></span>](/javascript/api/office/office.addincommands.event#completed-options-)

<span data-ttu-id="8e624-114">新增了可选参数 `options`，它是有效值为 `allowEvent` 的字典。</span><span class="sxs-lookup"><span data-stu-id="8e624-114">Event.completed - A new optional parameter , which is a dictionary with one valid value .</span></span> <span data-ttu-id="8e624-115">此值可用于取消执行事件。</span><span class="sxs-lookup"><span data-stu-id="8e624-115">This value is used to cancel execution of an event.</span></span>

<span data-ttu-id="8e624-116">**适用对象**：Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="8e624-116">**Available in**: Outlook on the web (Classic)</span></span>

### <a name="attachments"></a><span data-ttu-id="8e624-117">附件</span><span class="sxs-lookup"><span data-stu-id="8e624-117">Attachments</span></span>

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="8e624-118">AttachmentContent</span><span class="sxs-lookup"><span data-stu-id="8e624-118">AttachmentContent</span></span>](/javascript/api/outlook/office.attachmentcontent)

<span data-ttu-id="8e624-119">新增了表示附件内容的对象。</span><span class="sxs-lookup"><span data-stu-id="8e624-119">AttachmentContent - Added a new object that represents the content of an attachment.</span></span>

<span data-ttu-id="8e624-120">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-120">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[<span data-ttu-id="8e624-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="8e624-121">Office.context.mailbox.item.addFileAttachmentFromBase64Async</span></span>](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

<span data-ttu-id="8e624-122">新增了一个方法，可将 base64 编码字符串形式的文件附加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="8e624-122">Office.context.mailbox.item.addFileAttachmentFromBase64Async - Added a new method that allows you to attach a file represented as a base64 encoded string to a message or appointment.</span></span>

<span data-ttu-id="8e624-123">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-123">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[<span data-ttu-id="8e624-124">Office.context.mailbox.item.getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="8e624-124">Office.context.mailbox.item.getAttachmentContentAsync</span></span>](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent)

<span data-ttu-id="8e624-125">新增了一个方法，可获取特定附件的内容。</span><span class="sxs-lookup"><span data-stu-id="8e624-125">Office.context.mailbox.item.getAttachmentContentAsync - Added a new method to get the content of a specific attachment.</span></span>

<span data-ttu-id="8e624-126">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-126">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>[<span data-ttu-id="8e624-127">Office.context.mailbox.item.getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="8e624-127">Office.context.mailbox.item.getAttachmentsAsync</span></span>](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails)

<span data-ttu-id="8e624-128">新增了一个方法，可在撰写模式下获取项目附件。</span><span class="sxs-lookup"><span data-stu-id="8e624-128">Office.context.mailbox.item.getAttachmentsAsync - Added a new method that gets an item's attachments in compose mode.</span></span>

<span data-ttu-id="8e624-129">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-129">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[<span data-ttu-id="8e624-130">Office.MailboxEnums.AttachmentContentFormat</span><span class="sxs-lookup"><span data-stu-id="8e624-130">Office.MailboxEnums.AttachmentContentFormat</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

<span data-ttu-id="8e624-131">新增了一个枚举，可指定应用于附件内容的格式设置。</span><span class="sxs-lookup"><span data-stu-id="8e624-131">Office.MailboxEnums.AttachmentContentFormat - Added a new enum that specifies the formatting that applies to an attachment's content.</span></span>

<span data-ttu-id="8e624-132">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-132">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[<span data-ttu-id="8e624-133">Office.MailboxEnums.AttachmentStatus</span><span class="sxs-lookup"><span data-stu-id="8e624-133">Office.MailboxEnums.AttachmentStatus</span></span>](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

<span data-ttu-id="8e624-134">新增了一个枚举，可指定将附件添加至项目还是从项目中删除附件。</span><span class="sxs-lookup"><span data-stu-id="8e624-134">Office.MailboxEnums.AttachmentStatus - Added a new enum that specifies whether an attachment was added to or removed from an item.</span></span>

<span data-ttu-id="8e624-135">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-135">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="8e624-136">Office.EventType.AttachmentsChanged</span><span class="sxs-lookup"><span data-stu-id="8e624-136">Office.EventType.AttachmentsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="8e624-137">向 `Item` 中添加了 `AttachmentsChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="8e624-137">Added `AttachmentsChanged` event to `Item`.</span></span>

<span data-ttu-id="8e624-138">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-138">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="delegate-access"></a><span data-ttu-id="8e624-139">委托访问</span><span class="sxs-lookup"><span data-stu-id="8e624-139">Delegate access</span></span>

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[<span data-ttu-id="8e624-140">SharedProperties</span><span class="sxs-lookup"><span data-stu-id="8e624-140">SharedProperties</span></span>](/javascript/api/outlook/office.sharedproperties)

<span data-ttu-id="8e624-141">新增了一个对象，表示共享文件夹、日历或邮箱中的约会或邮件项目的属性。</span><span class="sxs-lookup"><span data-stu-id="8e624-141">SharedProperties - Added a new object that represents the properties of an appointment or message item in a shared folder, calendar, or mailbox.</span></span>

<span data-ttu-id="8e624-142">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-142">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[<span data-ttu-id="8e624-143">Office.context.mailbox.item.getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="8e624-143">Office.context.mailbox.item.getSharedPropertiesAsync</span></span>](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

<span data-ttu-id="8e624-144">新增了一个对象，用于获取表示约会或邮件项目的 sharedProperties 的对象。</span><span class="sxs-lookup"><span data-stu-id="8e624-144">Office.context.mailbox.item.getSharedPropertiesAsync - Added a new method that gets an object which represents the sharedProperties of an appointment or message item.</span></span>

<span data-ttu-id="8e624-145">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-145">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[<span data-ttu-id="8e624-146">Office.MailboxEnums.DelegatePermissions</span><span class="sxs-lookup"><span data-stu-id="8e624-146">Office.MailboxEnums.DelegatePermissions</span></span>](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

<span data-ttu-id="8e624-147">新增了一个位标志枚举，可指定委派权限。</span><span class="sxs-lookup"><span data-stu-id="8e624-147">Office.MailboxEnums.DelegatePermissions - Added a new bit flag enum that specifies the delegate permissions.</span></span>

<span data-ttu-id="8e624-148">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-148">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[<span data-ttu-id="8e624-149">SupportsSharedFolders manifest element</span><span class="sxs-lookup"><span data-stu-id="8e624-149">SupportsSharedFolders manifest element</span></span>](../../manifest/supportssharedfolders.md)

<span data-ttu-id="8e624-150">向 [DesktopFormFactor](../../manifest/desktopformfactor.md) 清单元素中添加了子元素。</span><span class="sxs-lookup"><span data-stu-id="8e624-150">SupportsSharedFolders manifest element - Added a child element to the [DesktopFormFactor](../../manifest/desktopformfactor.md) manifest element.</span></span> <span data-ttu-id="8e624-151">它定义外接程序是否在代理应用场景中可用。</span><span class="sxs-lookup"><span data-stu-id="8e624-151">It defines whether the add-in is available in delegate scenarios.</span></span>

<span data-ttu-id="8e624-152">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-152">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="enhanced-location"></a><span data-ttu-id="8e624-153">增强位置</span><span class="sxs-lookup"><span data-stu-id="8e624-153">Enhanced location</span></span>

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="8e624-154">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="8e624-154">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

<span data-ttu-id="8e624-155">新增了一个对象，显示约会的位置。</span><span class="sxs-lookup"><span data-stu-id="8e624-155">EnhancedLocation - Added a new object that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="8e624-156">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-156">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[<span data-ttu-id="8e624-157">LocationDetails</span><span class="sxs-lookup"><span data-stu-id="8e624-157">LocationDetails</span></span>](/javascript/api/outlook/office.locationdetails)

<span data-ttu-id="8e624-158">新增了一个表示位置的对象。</span><span class="sxs-lookup"><span data-stu-id="8e624-158">LocationDetails - Added a new object that represents a location.</span></span> <span data-ttu-id="8e624-159">只读。</span><span class="sxs-lookup"><span data-stu-id="8e624-159">Read only.</span></span>

<span data-ttu-id="8e624-160">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-160">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[<span data-ttu-id="8e624-161">LocationIdentifier</span><span class="sxs-lookup"><span data-stu-id="8e624-161">LocationIdentifier</span></span>](/javascript/api/outlook/office.locationidentifier)

<span data-ttu-id="8e624-162">新增了一个表示位置 ID 的对象。</span><span class="sxs-lookup"><span data-stu-id="8e624-162">LocationIdentifier - Added a new object that represents the id of a location.</span></span>

<span data-ttu-id="8e624-163">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-163">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[<span data-ttu-id="8e624-164">Office.context.mailbox.item.enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="8e624-164">Office.context.mailbox.item.enhancedLocation</span></span>](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation)

<span data-ttu-id="8e624-165">新增了一个表示约会位置的属性。</span><span class="sxs-lookup"><span data-stu-id="8e624-165">Office.context.mailbox.item.enhancedLocation - Added a new property that represents the set of locations on an appointment.</span></span>

<span data-ttu-id="8e624-166">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-166">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[<span data-ttu-id="8e624-167">Office.MailboxEnums.LocationType</span><span class="sxs-lookup"><span data-stu-id="8e624-167">Office.MailboxEnums.LocationType</span></span>](/javascript/api/outlook/office.mailboxenums.locationtype)

<span data-ttu-id="8e624-168">新增了一个用于指定约会位置类型的枚举。</span><span class="sxs-lookup"><span data-stu-id="8e624-168">Office.MailboxEnums.LocationType - Added a new enum that specifies an appointment location's type.</span></span>

<span data-ttu-id="8e624-169">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-169">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="8e624-170">Office.EventType.EnhancedLocationsChanged</span><span class="sxs-lookup"><span data-stu-id="8e624-170">Office.EventType.EnhancedLocationsChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="8e624-171">向 `Item` 中添加了 `EnhancedLocationsChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="8e624-171">Added `EnhancedLocationsChanged` event to `Item`.</span></span>

<span data-ttu-id="8e624-172">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-172">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="8e624-173">与可操作邮件集成</span><span class="sxs-lookup"><span data-stu-id="8e624-173">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[<span data-ttu-id="8e624-174">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="8e624-174">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

<span data-ttu-id="8e624-175">新增了一个函数，当外接程序[由可操作邮件激活时](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="8e624-175">Office.context.mailbox.item.getInitializationContextAsync - Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="8e624-176">**适用对象**：Office 2019 for Windows（Office 365 订阅）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="8e624-176">**Available in**: Office 2019 for Windows (Office 365 subscription), Outlook on the web (Classic)</span></span>

### <a name="internet-headers"></a><span data-ttu-id="8e624-177">Internet 标头：</span><span class="sxs-lookup"><span data-stu-id="8e624-177">Internet headers</span></span>

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="8e624-178">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="8e624-178">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

<span data-ttu-id="8e624-179">新增了一个对象，显示邮件项目的 Internet 标头。</span><span class="sxs-lookup"><span data-stu-id="8e624-179">InternetHeaders - Added a new object that represents the internet headers of a message item.</span></span>

<span data-ttu-id="8e624-180">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-180">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>[<span data-ttu-id="8e624-181">Office.context.mailbox.item.internetHeaders</span><span class="sxs-lookup"><span data-stu-id="8e624-181">Office.context.mailbox.item.internetHeaders</span></span>](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders)

<span data-ttu-id="8e624-182">新增了一个属性，显示邮件项目的 Internet 标头。</span><span class="sxs-lookup"><span data-stu-id="8e624-182">Office.context.mailbox.item.internetHeaders - Added a new property that represents the internet headers on a message item.</span></span>

<span data-ttu-id="8e624-183">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-183">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="office-theme"></a><span data-ttu-id="8e624-184">Office 主题</span><span class="sxs-lookup"><span data-stu-id="8e624-184">Office Theme</span></span>

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[<span data-ttu-id="8e624-185">Office.context.mailbox.officeTheme</span><span class="sxs-lookup"><span data-stu-id="8e624-185">Office.context.mailbox.officeTheme</span></span>](/javascript/api/office/office.officetheme)

<span data-ttu-id="8e624-186">增加了获取 Office 主题的功能。</span><span class="sxs-lookup"><span data-stu-id="8e624-186">Added ability to get Office theme.</span></span>

<span data-ttu-id="8e624-187">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-187">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[<span data-ttu-id="8e624-188">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="8e624-188">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="8e624-189">向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。</span><span class="sxs-lookup"><span data-stu-id="8e624-189">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="8e624-190">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="8e624-190">**Available in**: Outlook 2019 for Windows (Office 365 subscription)</span></span>

### <a name="sso"></a><span data-ttu-id="8e624-191">SSO</span><span class="sxs-lookup"><span data-stu-id="8e624-191">SSO</span></span>

#### <a name="officecontextauthgetaccesstokenasynchttpsdocsmicrosoftcomofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[<span data-ttu-id="8e624-192">Office.context.auth.getAccessTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8e624-192">Office.context.auth.getAccessTokenAsync</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

<span data-ttu-id="8e624-193">添加了对 `getAccessTokenAsync` 的访问，使外接程序[能够访问](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) Microsoft Graph API 的访问令牌。</span><span class="sxs-lookup"><span data-stu-id="8e624-193">Office.context.auth.getAccessTokenAsync - Added access to `getAccessTokenAsync`, which allows add-ins to [get an access token](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token) for the Microsoft Graph API.</span></span>

<span data-ttu-id="8e624-194">**适用对象**：Outlook 2019 for Windows（Office 365 订阅）、Outlook 2019 for Mac、Outlook 网页版（Office 365 和 Outlook.com）、Outlook 网页版（经典）</span><span class="sxs-lookup"><span data-stu-id="8e624-194">**Available in**: Outlook 2019 for Windows (Office 365 subscription), Outlook 2019 for Mac, Outlook on the web (Office 365 and Outlook.com), Outlook on the web (Classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="8e624-195">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8e624-195">See also</span></span>

- [<span data-ttu-id="8e624-196">Outlook 加载项</span><span class="sxs-lookup"><span data-stu-id="8e624-196">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="8e624-197">Outlook 外接程序代码示例</span><span class="sxs-lookup"><span data-stu-id="8e624-197">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="8e624-198">入门</span><span class="sxs-lookup"><span data-stu-id="8e624-198">Get started</span></span>](https://docs.microsoft.com/outlook/add-ins/quick-start)
