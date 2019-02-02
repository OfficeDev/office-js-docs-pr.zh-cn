---
title: Office.context.mailbox.item - 要求集 1.5
description: ''
ms.date: 01/30/2019
localization_priority: Priority
ms.openlocfilehash: cca0bb4baa15d72a58909ca1417eb52a9bf70a8f
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701839"
---
# <a name="item"></a><span data-ttu-id="524d9-102">item</span><span class="sxs-lookup"><span data-stu-id="524d9-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="524d9-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="524d9-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="524d9-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="524d9-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="524d9-106">Requirements</span></span>

|<span data-ttu-id="524d9-107">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-107">Requirement</span></span>| <span data-ttu-id="524d9-108">值</span><span class="sxs-lookup"><span data-stu-id="524d9-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-110">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-110">1.0</span></span>|
|[<span data-ttu-id="524d9-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-112">受限</span><span class="sxs-lookup"><span data-stu-id="524d9-112">Restricted</span></span>|
|[<span data-ttu-id="524d9-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="524d9-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="524d9-115">Members and methods</span></span>

| <span data-ttu-id="524d9-116">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-116">Member</span></span> | <span data-ttu-id="524d9-117">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="524d9-118">attachments</span><span class="sxs-lookup"><span data-stu-id="524d9-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="524d9-119">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-119">Member</span></span> |
| [<span data-ttu-id="524d9-120">bcc</span><span class="sxs-lookup"><span data-stu-id="524d9-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="524d9-121">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-121">Member</span></span> |
| [<span data-ttu-id="524d9-122">body</span><span class="sxs-lookup"><span data-stu-id="524d9-122">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="524d9-123">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-123">Member</span></span> |
| [<span data-ttu-id="524d9-124">cc</span><span class="sxs-lookup"><span data-stu-id="524d9-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="524d9-125">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-125">Member</span></span> |
| [<span data-ttu-id="524d9-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="524d9-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="524d9-127">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-127">Member</span></span> |
| [<span data-ttu-id="524d9-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="524d9-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="524d9-129">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-129">Member</span></span> |
| [<span data-ttu-id="524d9-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="524d9-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="524d9-131">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-131">Member</span></span> |
| [<span data-ttu-id="524d9-132">end</span><span class="sxs-lookup"><span data-stu-id="524d9-132">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="524d9-133">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-133">Member</span></span> |
| [<span data-ttu-id="524d9-134">from</span><span class="sxs-lookup"><span data-stu-id="524d9-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="524d9-135">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-135">Member</span></span> |
| [<span data-ttu-id="524d9-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="524d9-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="524d9-137">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-137">Member</span></span> |
| [<span data-ttu-id="524d9-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="524d9-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="524d9-139">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-139">Member</span></span> |
| [<span data-ttu-id="524d9-140">itemId</span><span class="sxs-lookup"><span data-stu-id="524d9-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="524d9-141">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-141">Member</span></span> |
| [<span data-ttu-id="524d9-142">itemType</span><span class="sxs-lookup"><span data-stu-id="524d9-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="524d9-143">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-143">Member</span></span> |
| [<span data-ttu-id="524d9-144">location</span><span class="sxs-lookup"><span data-stu-id="524d9-144">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="524d9-145">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-145">Member</span></span> |
| [<span data-ttu-id="524d9-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="524d9-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="524d9-147">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-147">Member</span></span> |
| [<span data-ttu-id="524d9-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="524d9-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="524d9-149">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-149">Member</span></span> |
| [<span data-ttu-id="524d9-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="524d9-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="524d9-151">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-151">Member</span></span> |
| [<span data-ttu-id="524d9-152">organizer</span><span class="sxs-lookup"><span data-stu-id="524d9-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="524d9-153">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-153">Member</span></span> |
| [<span data-ttu-id="524d9-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="524d9-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="524d9-155">Member</span><span class="sxs-lookup"><span data-stu-id="524d9-155">Member</span></span> |
| [<span data-ttu-id="524d9-156">sender</span><span class="sxs-lookup"><span data-stu-id="524d9-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="524d9-157">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-157">Member</span></span> |
| [<span data-ttu-id="524d9-158">start</span><span class="sxs-lookup"><span data-stu-id="524d9-158">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="524d9-159">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-159">Member</span></span> |
| [<span data-ttu-id="524d9-160">subject</span><span class="sxs-lookup"><span data-stu-id="524d9-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="524d9-161">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-161">Member</span></span> |
| [<span data-ttu-id="524d9-162">to</span><span class="sxs-lookup"><span data-stu-id="524d9-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="524d9-163">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-163">Member</span></span> |
| [<span data-ttu-id="524d9-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="524d9-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="524d9-165">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-165">Method</span></span> |
| [<span data-ttu-id="524d9-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="524d9-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="524d9-167">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-167">Method</span></span> |
| [<span data-ttu-id="524d9-168">close</span><span class="sxs-lookup"><span data-stu-id="524d9-168">close</span></span>](#close) | <span data-ttu-id="524d9-169">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-169">Method</span></span> |
| [<span data-ttu-id="524d9-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="524d9-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="524d9-171">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-171">Method</span></span> |
| [<span data-ttu-id="524d9-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="524d9-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="524d9-173">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-173">Method</span></span> |
| [<span data-ttu-id="524d9-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="524d9-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="524d9-175">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-175">Method</span></span> |
| [<span data-ttu-id="524d9-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="524d9-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="524d9-177">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-177">Method</span></span> |
| [<span data-ttu-id="524d9-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="524d9-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="524d9-179">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-179">Method</span></span> |
| [<span data-ttu-id="524d9-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="524d9-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="524d9-181">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-181">Method</span></span> |
| [<span data-ttu-id="524d9-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="524d9-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="524d9-183">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-183">Method</span></span> |
| [<span data-ttu-id="524d9-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="524d9-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="524d9-185">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-185">Method</span></span> |
| [<span data-ttu-id="524d9-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="524d9-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="524d9-187">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-187">Method</span></span> |
| [<span data-ttu-id="524d9-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="524d9-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="524d9-189">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-189">Method</span></span> |
| [<span data-ttu-id="524d9-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="524d9-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="524d9-191">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-191">Method</span></span> |
| [<span data-ttu-id="524d9-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="524d9-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="524d9-193">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="524d9-194">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-194">Example</span></span>

<span data-ttu-id="524d9-195">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="524d9-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
    });
}
```

### <a name="members"></a><span data-ttu-id="524d9-196">成员</span><span class="sxs-lookup"><span data-stu-id="524d9-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="524d9-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="524d9-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="524d9-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-200">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="524d9-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="524d9-201">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="524d9-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-202">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-202">Type:</span></span>

*   <span data-ttu-id="524d9-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="524d9-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-204">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-204">Requirements</span></span>

|<span data-ttu-id="524d9-205">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-205">Requirement</span></span>| <span data-ttu-id="524d9-206">值</span><span class="sxs-lookup"><span data-stu-id="524d9-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-207">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-208">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-208">1.0</span></span>|
|[<span data-ttu-id="524d9-209">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-209">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-210">ReadItem</span></span>|
|[<span data-ttu-id="524d9-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-211">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-212">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-213">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-213">Example</span></span>

<span data-ttu-id="524d9-214">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="524d9-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="524d9-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="524d9-216">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="524d9-217">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-218">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-218">Type:</span></span>

*   [<span data-ttu-id="524d9-219">收件人</span><span class="sxs-lookup"><span data-stu-id="524d9-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="524d9-220">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-220">Requirements</span></span>

|<span data-ttu-id="524d9-221">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-221">Requirement</span></span>| <span data-ttu-id="524d9-222">值</span><span class="sxs-lookup"><span data-stu-id="524d9-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-224">1.1</span><span class="sxs-lookup"><span data-stu-id="524d9-224">1.1</span></span>|
|[<span data-ttu-id="524d9-225">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-225">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-226">ReadItem</span></span>|
|[<span data-ttu-id="524d9-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-227">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-228">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-229">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-229">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="524d9-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="524d9-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="524d9-231">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-232">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-232">Type:</span></span>

*   [<span data-ttu-id="524d9-233">Body</span><span class="sxs-lookup"><span data-stu-id="524d9-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="524d9-234">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-234">Requirements</span></span>

|<span data-ttu-id="524d9-235">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-235">Requirement</span></span>| <span data-ttu-id="524d9-236">值</span><span class="sxs-lookup"><span data-stu-id="524d9-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-238">1.1</span><span class="sxs-lookup"><span data-stu-id="524d9-238">1.1</span></span>|
|[<span data-ttu-id="524d9-239">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-240">ReadItem</span></span>|
|[<span data-ttu-id="524d9-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-242">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-242">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="524d9-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-243">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="524d9-244">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="524d9-244">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="524d9-245">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-245">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-246">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-246">Read mode</span></span>

<span data-ttu-id="524d9-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="524d9-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="524d9-249">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-249">Compose mode</span></span>

<span data-ttu-id="524d9-250">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-250">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-251">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-251">Type:</span></span>

*   <span data-ttu-id="524d9-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-252">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-253">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-253">Requirements</span></span>

|<span data-ttu-id="524d9-254">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-254">Requirement</span></span>| <span data-ttu-id="524d9-255">值</span><span class="sxs-lookup"><span data-stu-id="524d9-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-256">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-257">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-257">1.0</span></span>|
|[<span data-ttu-id="524d9-258">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-258">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-259">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-259">ReadItem</span></span>|
|[<span data-ttu-id="524d9-260">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-260">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-261">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-261">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-262">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-262">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="524d9-263">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="524d9-263">(nullable) conversationId :String</span></span>

<span data-ttu-id="524d9-264">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="524d9-264">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="524d9-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="524d9-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="524d9-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="524d9-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-269">类型:</span><span class="sxs-lookup"><span data-stu-id="524d9-269">Type:</span></span>

*   <span data-ttu-id="524d9-270">String</span><span class="sxs-lookup"><span data-stu-id="524d9-270">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-271">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-271">Requirements</span></span>

|<span data-ttu-id="524d9-272">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-272">Requirement</span></span>| <span data-ttu-id="524d9-273">值</span><span class="sxs-lookup"><span data-stu-id="524d9-273">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-274">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-275">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-275">1.0</span></span>|
|[<span data-ttu-id="524d9-276">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-277">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-277">ReadItem</span></span>|
|[<span data-ttu-id="524d9-278">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-279">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-279">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="524d9-280">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="524d9-280">dateTimeCreated :Date</span></span>

<span data-ttu-id="524d9-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-283">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-283">Type:</span></span>

*   <span data-ttu-id="524d9-284">日期</span><span class="sxs-lookup"><span data-stu-id="524d9-284">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-285">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-285">Requirements</span></span>

|<span data-ttu-id="524d9-286">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-286">Requirement</span></span>| <span data-ttu-id="524d9-287">值</span><span class="sxs-lookup"><span data-stu-id="524d9-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-288">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-289">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-289">1.0</span></span>|
|[<span data-ttu-id="524d9-290">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-291">ReadItem</span></span>|
|[<span data-ttu-id="524d9-292">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-293">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-293">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-294">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-294">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="524d9-295">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="524d9-295">dateTimeModified :Date</span></span>

<span data-ttu-id="524d9-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-298">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="524d9-298">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-299">类型:</span><span class="sxs-lookup"><span data-stu-id="524d9-299">Type:</span></span>

*   <span data-ttu-id="524d9-300">日期</span><span class="sxs-lookup"><span data-stu-id="524d9-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-301">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-301">Requirements</span></span>

|<span data-ttu-id="524d9-302">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-302">Requirement</span></span>| <span data-ttu-id="524d9-303">值</span><span class="sxs-lookup"><span data-stu-id="524d9-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-304">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-305">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-305">1.0</span></span>|
|[<span data-ttu-id="524d9-306">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-307">ReadItem</span></span>|
|[<span data-ttu-id="524d9-308">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-309">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-310">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-310">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="524d9-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="524d9-311">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="524d9-312">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="524d9-312">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="524d9-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="524d9-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-315">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-315">Read mode</span></span>

<span data-ttu-id="524d9-316">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-316">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="524d9-317">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-317">Compose mode</span></span>

<span data-ttu-id="524d9-318">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-318">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="524d9-319">使用 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="524d9-319">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-320">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-320">Type:</span></span>

*   <span data-ttu-id="524d9-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="524d9-321">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-322">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-322">Requirements</span></span>

|<span data-ttu-id="524d9-323">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-323">Requirement</span></span>| <span data-ttu-id="524d9-324">值</span><span class="sxs-lookup"><span data-stu-id="524d9-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-325">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-326">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-326">1.0</span></span>|
|[<span data-ttu-id="524d9-327">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-328">ReadItem</span></span>|
|[<span data-ttu-id="524d9-329">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-330">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-330">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-331">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-331">Example</span></span>

<span data-ttu-id="524d9-332">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="524d9-332">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="524d9-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="524d9-333">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="524d9-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="524d9-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="524d9-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-338">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="524d9-338">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-339">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-339">Type:</span></span>

*   [<span data-ttu-id="524d9-340">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="524d9-340">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="524d9-341">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-341">Requirements</span></span>

|<span data-ttu-id="524d9-342">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-342">Requirement</span></span>| <span data-ttu-id="524d9-343">值</span><span class="sxs-lookup"><span data-stu-id="524d9-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-344">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-345">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-345">1.0</span></span>|
|[<span data-ttu-id="524d9-346">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-347">ReadItem</span></span>|
|[<span data-ttu-id="524d9-348">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-349">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-349">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="524d9-350">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="524d9-350">internetMessageId :String</span></span>

<span data-ttu-id="524d9-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-353">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-353">Type:</span></span>

*   <span data-ttu-id="524d9-354">String</span><span class="sxs-lookup"><span data-stu-id="524d9-354">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-355">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-355">Requirements</span></span>

|<span data-ttu-id="524d9-356">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-356">Requirement</span></span>| <span data-ttu-id="524d9-357">值</span><span class="sxs-lookup"><span data-stu-id="524d9-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-358">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-359">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-359">1.0</span></span>|
|[<span data-ttu-id="524d9-360">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-361">ReadItem</span></span>|
|[<span data-ttu-id="524d9-362">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-363">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-363">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-364">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-364">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="524d9-365">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="524d9-365">itemClass :String</span></span>

<span data-ttu-id="524d9-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="524d9-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="524d9-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="524d9-370">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-370">Type</span></span> | <span data-ttu-id="524d9-371">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-371">Description</span></span> | <span data-ttu-id="524d9-372">项目类</span><span class="sxs-lookup"><span data-stu-id="524d9-372">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="524d9-373">约会项目</span><span class="sxs-lookup"><span data-stu-id="524d9-373">Appointment items</span></span> | <span data-ttu-id="524d9-374">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="524d9-374">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="524d9-375">邮件项目</span><span class="sxs-lookup"><span data-stu-id="524d9-375">Message items</span></span> | <span data-ttu-id="524d9-376">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="524d9-376">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="524d9-377">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="524d9-377">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-378">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-378">Type:</span></span>

*   <span data-ttu-id="524d9-379">String</span><span class="sxs-lookup"><span data-stu-id="524d9-379">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-380">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-380">Requirements</span></span>

|<span data-ttu-id="524d9-381">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-381">Requirement</span></span>| <span data-ttu-id="524d9-382">值</span><span class="sxs-lookup"><span data-stu-id="524d9-382">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-383">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-383">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-384">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-384">1.0</span></span>|
|[<span data-ttu-id="524d9-385">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-385">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-386">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-386">ReadItem</span></span>|
|[<span data-ttu-id="524d9-387">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-387">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-388">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-388">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-389">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-389">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="524d9-390">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="524d9-390">(nullable) itemId :String</span></span>

<span data-ttu-id="524d9-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-393">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="524d9-393">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="524d9-394">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="524d9-394">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="524d9-395">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="524d9-395">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="524d9-396">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="524d9-396">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="524d9-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-399">类型:</span><span class="sxs-lookup"><span data-stu-id="524d9-399">Type:</span></span>

*   <span data-ttu-id="524d9-400">String</span><span class="sxs-lookup"><span data-stu-id="524d9-400">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-401">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-401">Requirements</span></span>

|<span data-ttu-id="524d9-402">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-402">Requirement</span></span>| <span data-ttu-id="524d9-403">值</span><span class="sxs-lookup"><span data-stu-id="524d9-403">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-404">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-404">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-405">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-405">1.0</span></span>|
|[<span data-ttu-id="524d9-406">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-406">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-407">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-407">ReadItem</span></span>|
|[<span data-ttu-id="524d9-408">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-408">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-409">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-409">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-410">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-410">Example</span></span>

<span data-ttu-id="524d9-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="524d9-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="524d9-413">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="524d9-414">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="524d9-414">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="524d9-415">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="524d9-415">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-416">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-416">Type:</span></span>

*   [<span data-ttu-id="524d9-417">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="524d9-417">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="524d9-418">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-418">Requirements</span></span>

|<span data-ttu-id="524d9-419">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-419">Requirement</span></span>| <span data-ttu-id="524d9-420">值</span><span class="sxs-lookup"><span data-stu-id="524d9-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-421">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-422">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-422">1.0</span></span>|
|[<span data-ttu-id="524d9-423">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-424">ReadItem</span></span>|
|[<span data-ttu-id="524d9-425">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-426">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-426">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-427">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-427">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="524d9-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="524d9-428">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="524d9-429">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="524d9-429">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-430">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-430">Read mode</span></span>

<span data-ttu-id="524d9-431">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="524d9-431">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="524d9-432">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-432">Compose mode</span></span>

<span data-ttu-id="524d9-433">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-433">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-434">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-434">Type:</span></span>

*   <span data-ttu-id="524d9-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="524d9-435">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-436">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-436">Requirements</span></span>

|<span data-ttu-id="524d9-437">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-437">Requirement</span></span>| <span data-ttu-id="524d9-438">值</span><span class="sxs-lookup"><span data-stu-id="524d9-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-439">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-440">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-440">1.0</span></span>|
|[<span data-ttu-id="524d9-441">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-441">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-442">ReadItem</span></span>|
|[<span data-ttu-id="524d9-443">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-443">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-444">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-444">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-445">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-445">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="524d9-446">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="524d9-446">normalizedSubject :String</span></span>

<span data-ttu-id="524d9-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="524d9-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="524d9-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-451">类型:</span><span class="sxs-lookup"><span data-stu-id="524d9-451">Type:</span></span>

*   <span data-ttu-id="524d9-452">String</span><span class="sxs-lookup"><span data-stu-id="524d9-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-453">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-453">Requirements</span></span>

|<span data-ttu-id="524d9-454">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-454">Requirement</span></span>| <span data-ttu-id="524d9-455">值</span><span class="sxs-lookup"><span data-stu-id="524d9-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-456">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-457">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-457">1.0</span></span>|
|[<span data-ttu-id="524d9-458">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-459">ReadItem</span></span>|
|[<span data-ttu-id="524d9-460">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-461">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-462">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="524d9-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="524d9-463">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="524d9-464">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="524d9-464">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-465">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-465">Type:</span></span>

*   [<span data-ttu-id="524d9-466">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="524d9-466">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="524d9-467">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-467">Requirements</span></span>

|<span data-ttu-id="524d9-468">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-468">Requirement</span></span>| <span data-ttu-id="524d9-469">值</span><span class="sxs-lookup"><span data-stu-id="524d9-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-470">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-471">1.3</span><span class="sxs-lookup"><span data-stu-id="524d9-471">1.3</span></span>|
|[<span data-ttu-id="524d9-472">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-473">ReadItem</span></span>|
|[<span data-ttu-id="524d9-474">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-475">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-475">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="524d9-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-476">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="524d9-477">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="524d9-477">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="524d9-478">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-478">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-479">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-479">Read mode</span></span>

<span data-ttu-id="524d9-480">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-480">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="524d9-481">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-481">Compose mode</span></span>

<span data-ttu-id="524d9-482">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-482">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-483">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-483">Type:</span></span>

*   <span data-ttu-id="524d9-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-484">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-485">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-485">Requirements</span></span>

|<span data-ttu-id="524d9-486">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-486">Requirement</span></span>| <span data-ttu-id="524d9-487">值</span><span class="sxs-lookup"><span data-stu-id="524d9-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-489">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-489">1.0</span></span>|
|[<span data-ttu-id="524d9-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-491">ReadItem</span></span>|
|[<span data-ttu-id="524d9-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-493">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-494">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-494">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="524d9-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="524d9-495">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="524d9-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-498">类型:</span><span class="sxs-lookup"><span data-stu-id="524d9-498">Type:</span></span>

*   [<span data-ttu-id="524d9-499">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="524d9-499">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="524d9-500">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-500">Requirements</span></span>

|<span data-ttu-id="524d9-501">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-501">Requirement</span></span>| <span data-ttu-id="524d9-502">值</span><span class="sxs-lookup"><span data-stu-id="524d9-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-504">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-504">1.0</span></span>|
|[<span data-ttu-id="524d9-505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-506">ReadItem</span></span>|
|[<span data-ttu-id="524d9-507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-508">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-508">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-509">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-509">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="524d9-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-510">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="524d9-511">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="524d9-511">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="524d9-512">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-512">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-513">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-513">Read mode</span></span>

<span data-ttu-id="524d9-514">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-514">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="524d9-515">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-515">Compose mode</span></span>

<span data-ttu-id="524d9-516">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-516">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-517">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-517">Type:</span></span>

*   <span data-ttu-id="524d9-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-519">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-519">Requirements</span></span>

|<span data-ttu-id="524d9-520">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-520">Requirement</span></span>| <span data-ttu-id="524d9-521">值</span><span class="sxs-lookup"><span data-stu-id="524d9-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-522">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-523">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-523">1.0</span></span>|
|[<span data-ttu-id="524d9-524">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-525">ReadItem</span></span>|
|[<span data-ttu-id="524d9-526">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-527">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-528">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-528">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="524d9-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="524d9-529">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="524d9-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="524d9-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="524d9-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-534">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="524d9-534">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-535">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-535">Type:</span></span>

*   [<span data-ttu-id="524d9-536">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="524d9-536">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="524d9-537">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-537">Requirements</span></span>

|<span data-ttu-id="524d9-538">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-538">Requirement</span></span>| <span data-ttu-id="524d9-539">值</span><span class="sxs-lookup"><span data-stu-id="524d9-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-540">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-541">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-541">1.0</span></span>|
|[<span data-ttu-id="524d9-542">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-543">ReadItem</span></span>|
|[<span data-ttu-id="524d9-544">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-545">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-545">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-546">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-546">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="524d9-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="524d9-547">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="524d9-548">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="524d9-548">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="524d9-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="524d9-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-551">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-551">Read mode</span></span>

<span data-ttu-id="524d9-552">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-552">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="524d9-553">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-553">Compose mode</span></span>

<span data-ttu-id="524d9-554">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-554">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="524d9-555">使用 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="524d9-555">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-556">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-556">Type:</span></span>

*   <span data-ttu-id="524d9-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="524d9-557">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-558">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-558">Requirements</span></span>

|<span data-ttu-id="524d9-559">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-559">Requirement</span></span>| <span data-ttu-id="524d9-560">值</span><span class="sxs-lookup"><span data-stu-id="524d9-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-561">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-562">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-562">1.0</span></span>|
|[<span data-ttu-id="524d9-563">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-563">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-564">ReadItem</span></span>|
|[<span data-ttu-id="524d9-565">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-565">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-566">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-566">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-567">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-567">Example</span></span>

<span data-ttu-id="524d9-568">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="524d9-568">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="524d9-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="524d9-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="524d9-570">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="524d9-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="524d9-571">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="524d9-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-572">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-572">Read mode</span></span>

<span data-ttu-id="524d9-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="524d9-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="524d9-575">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-575">Compose mode</span></span>

<span data-ttu-id="524d9-576">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="524d9-577">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-577">Type:</span></span>

*   <span data-ttu-id="524d9-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="524d9-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-579">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-579">Requirements</span></span>

|<span data-ttu-id="524d9-580">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-580">Requirement</span></span>| <span data-ttu-id="524d9-581">值</span><span class="sxs-lookup"><span data-stu-id="524d9-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-582">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-583">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-583">1.0</span></span>|
|[<span data-ttu-id="524d9-584">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-585">ReadItem</span></span>|
|[<span data-ttu-id="524d9-586">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-587">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-587">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="524d9-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="524d9-589">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="524d9-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="524d9-590">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="524d9-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="524d9-591">阅读模式</span><span class="sxs-lookup"><span data-stu-id="524d9-591">Read mode</span></span>

<span data-ttu-id="524d9-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="524d9-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="524d9-594">撰写模式</span><span class="sxs-lookup"><span data-stu-id="524d9-594">Compose mode</span></span>

<span data-ttu-id="524d9-595">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="524d9-596">类型：</span><span class="sxs-lookup"><span data-stu-id="524d9-596">Type:</span></span>

*   <span data-ttu-id="524d9-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="524d9-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-598">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-598">Requirements</span></span>

|<span data-ttu-id="524d9-599">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-599">Requirement</span></span>| <span data-ttu-id="524d9-600">值</span><span class="sxs-lookup"><span data-stu-id="524d9-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-601">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-602">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-602">1.0</span></span>|
|[<span data-ttu-id="524d9-603">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-603">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-604">ReadItem</span></span>|
|[<span data-ttu-id="524d9-605">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-605">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-606">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-606">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-607">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-607">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="524d9-608">方法</span><span class="sxs-lookup"><span data-stu-id="524d9-608">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="524d9-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="524d9-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="524d9-610">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="524d9-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="524d9-611">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="524d9-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="524d9-612">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="524d9-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-613">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-613">Parameters:</span></span>

|<span data-ttu-id="524d9-614">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-614">Name</span></span>| <span data-ttu-id="524d9-615">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-615">Type</span></span>| <span data-ttu-id="524d9-616">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-616">Attributes</span></span>| <span data-ttu-id="524d9-617">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="524d9-618">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-618">String</span></span>||<span data-ttu-id="524d9-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="524d9-621">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-621">String</span></span>||<span data-ttu-id="524d9-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="524d9-624">Object</span><span class="sxs-lookup"><span data-stu-id="524d9-624">Object</span></span>| <span data-ttu-id="524d9-625">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-625">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-626">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="524d9-626">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="524d9-627">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-627">Object</span></span> | <span data-ttu-id="524d9-628">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-628">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-629">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="524d9-630">布尔值</span><span class="sxs-lookup"><span data-stu-id="524d9-630">Boolean</span></span> | <span data-ttu-id="524d9-631">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-631">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-632">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="524d9-632">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="524d9-633">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-633">function</span></span>| <span data-ttu-id="524d9-634">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-634">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-635">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="524d9-636">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="524d9-636">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="524d9-637">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-637">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="524d9-638">错误</span><span class="sxs-lookup"><span data-stu-id="524d9-638">Errors</span></span>

| <span data-ttu-id="524d9-639">错误代码</span><span class="sxs-lookup"><span data-stu-id="524d9-639">Error code</span></span> | <span data-ttu-id="524d9-640">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-640">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="524d9-641">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="524d9-641">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="524d9-642">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="524d9-642">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="524d9-643">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="524d9-643">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="524d9-644">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-644">Requirements</span></span>

|<span data-ttu-id="524d9-645">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-645">Requirement</span></span>| <span data-ttu-id="524d9-646">值</span><span class="sxs-lookup"><span data-stu-id="524d9-646">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-647">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-648">1.1</span><span class="sxs-lookup"><span data-stu-id="524d9-648">1.1</span></span>|
|[<span data-ttu-id="524d9-649">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-649">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="524d9-650">ReadWriteItem</span></span>|
|[<span data-ttu-id="524d9-651">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-651">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-652">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-652">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="524d9-653">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-653">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="524d9-654">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="524d9-654">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        
      }
    );
  }
);
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="524d9-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="524d9-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="524d9-656">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="524d9-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="524d9-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="524d9-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="524d9-660">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="524d9-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="524d9-661">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="524d9-661">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-662">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-662">Parameters:</span></span>

|<span data-ttu-id="524d9-663">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-663">Name</span></span>| <span data-ttu-id="524d9-664">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-664">Type</span></span>| <span data-ttu-id="524d9-665">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-665">Attributes</span></span>| <span data-ttu-id="524d9-666">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="524d9-667">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-667">String</span></span>||<span data-ttu-id="524d9-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="524d9-670">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-670">String</span></span>||<span data-ttu-id="524d9-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="524d9-673">Object</span><span class="sxs-lookup"><span data-stu-id="524d9-673">Object</span></span>| <span data-ttu-id="524d9-674">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-674">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-675">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="524d9-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="524d9-676">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-676">Object</span></span>| <span data-ttu-id="524d9-677">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-677">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-678">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="524d9-679">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-679">function</span></span>| <span data-ttu-id="524d9-680">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-680">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-681">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="524d9-682">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="524d9-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="524d9-683">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="524d9-684">错误</span><span class="sxs-lookup"><span data-stu-id="524d9-684">Errors</span></span>

| <span data-ttu-id="524d9-685">错误代码</span><span class="sxs-lookup"><span data-stu-id="524d9-685">Error code</span></span> | <span data-ttu-id="524d9-686">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="524d9-687">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="524d9-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="524d9-688">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-688">Requirements</span></span>

|<span data-ttu-id="524d9-689">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-689">Requirement</span></span>| <span data-ttu-id="524d9-690">值</span><span class="sxs-lookup"><span data-stu-id="524d9-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-691">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-692">1.1</span><span class="sxs-lookup"><span data-stu-id="524d9-692">1.1</span></span>|
|[<span data-ttu-id="524d9-693">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="524d9-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="524d9-695">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-696">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-697">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-697">Example</span></span>

<span data-ttu-id="524d9-698">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="524d9-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="524d9-699">close()</span><span class="sxs-lookup"><span data-stu-id="524d9-699">close()</span></span>

<span data-ttu-id="524d9-700">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="524d9-700">Closes the current item that is being composed.</span></span>

<span data-ttu-id="524d9-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="524d9-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-703">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="524d9-703">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="524d9-704">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="524d9-704">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-705">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-705">Requirements</span></span>

|<span data-ttu-id="524d9-706">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-706">Requirement</span></span>| <span data-ttu-id="524d9-707">值</span><span class="sxs-lookup"><span data-stu-id="524d9-707">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-708">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-708">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-709">1.3</span><span class="sxs-lookup"><span data-stu-id="524d9-709">1.3</span></span>|
|[<span data-ttu-id="524d9-710">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-710">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-711">受限</span><span class="sxs-lookup"><span data-stu-id="524d9-711">Restricted</span></span>|
|[<span data-ttu-id="524d9-712">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-712">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-713">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-713">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="524d9-714">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="524d9-714">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="524d9-715">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="524d9-715">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-716">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="524d9-717">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="524d9-717">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="524d9-718">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="524d9-718">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="524d9-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="524d9-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-722">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-722">Parameters:</span></span>

| <span data-ttu-id="524d9-723">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-723">Name</span></span> | <span data-ttu-id="524d9-724">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-724">Type</span></span> | <span data-ttu-id="524d9-725">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-725">Attributes</span></span> | <span data-ttu-id="524d9-726">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-726">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="524d9-727">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="524d9-727">String &#124; Object</span></span>| |<span data-ttu-id="524d9-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="524d9-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="524d9-730">**或**</span><span class="sxs-lookup"><span data-stu-id="524d9-730">**OR**</span></span><br/><span data-ttu-id="524d9-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="524d9-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="524d9-733">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-733">String</span></span> | <span data-ttu-id="524d9-734">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-734">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="524d9-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="524d9-737">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-737">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="524d9-738">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-738">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-739">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="524d9-739">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="524d9-740">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-740">String</span></span> | | <span data-ttu-id="524d9-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="524d9-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="524d9-743">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-743">String</span></span> | | <span data-ttu-id="524d9-744">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-744">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="524d9-745">String</span><span class="sxs-lookup"><span data-stu-id="524d9-745">String</span></span> | | <span data-ttu-id="524d9-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="524d9-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="524d9-748">布尔</span><span class="sxs-lookup"><span data-stu-id="524d9-748">Boolean</span></span> | | <span data-ttu-id="524d9-p144">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="524d9-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="524d9-751">String</span><span class="sxs-lookup"><span data-stu-id="524d9-751">String</span></span> | | <span data-ttu-id="524d9-p145">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="524d9-755">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-755">function</span></span> | <span data-ttu-id="524d9-756">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-756">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-757">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-757">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="524d9-758">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-758">Requirements</span></span>

|<span data-ttu-id="524d9-759">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-759">Requirement</span></span>| <span data-ttu-id="524d9-760">值</span><span class="sxs-lookup"><span data-stu-id="524d9-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-761">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-762">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-762">1.0</span></span>|
|[<span data-ttu-id="524d9-763">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-764">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-764">ReadItem</span></span>|
|[<span data-ttu-id="524d9-765">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-766">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-766">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="524d9-767">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-767">Examples</span></span>

<span data-ttu-id="524d9-768">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-768">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="524d9-769">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-769">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="524d9-770">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-770">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="524d9-771">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-771">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="524d9-772">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-772">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="524d9-773">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-773">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="524d9-774">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="524d9-774">displayReplyForm(formData)</span></span>

<span data-ttu-id="524d9-775">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="524d9-775">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-776">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-776">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="524d9-777">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="524d9-777">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="524d9-778">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="524d9-778">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="524d9-p146">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="524d9-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-782">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-782">Parameters:</span></span>

| <span data-ttu-id="524d9-783">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-783">Name</span></span> | <span data-ttu-id="524d9-784">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-784">Type</span></span> | <span data-ttu-id="524d9-785">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-785">Attributes</span></span> | <span data-ttu-id="524d9-786">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-786">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="524d9-787">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="524d9-787">String &#124; Object</span></span>| | <span data-ttu-id="524d9-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="524d9-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="524d9-790">**或**</span><span class="sxs-lookup"><span data-stu-id="524d9-790">**OR**</span></span><br/><span data-ttu-id="524d9-p148">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="524d9-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="524d9-793">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-793">String</span></span> | <span data-ttu-id="524d9-794">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-794">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-p149">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="524d9-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="524d9-797">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-797">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="524d9-798">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-798">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-799">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="524d9-799">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="524d9-800">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-800">String</span></span> | | <span data-ttu-id="524d9-p150">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="524d9-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="524d9-803">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-803">String</span></span> | | <span data-ttu-id="524d9-804">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-804">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="524d9-805">String</span><span class="sxs-lookup"><span data-stu-id="524d9-805">String</span></span> | | <span data-ttu-id="524d9-p151">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="524d9-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="524d9-808">布尔</span><span class="sxs-lookup"><span data-stu-id="524d9-808">Boolean</span></span> | | <span data-ttu-id="524d9-p152">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="524d9-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="524d9-811">String</span><span class="sxs-lookup"><span data-stu-id="524d9-811">String</span></span> | | <span data-ttu-id="524d9-p153">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="524d9-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="524d9-815">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-815">function</span></span> | <span data-ttu-id="524d9-816">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-816">&lt;optional&gt;</span></span> | <span data-ttu-id="524d9-817">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-817">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="524d9-818">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-818">Requirements</span></span>

|<span data-ttu-id="524d9-819">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-819">Requirement</span></span>| <span data-ttu-id="524d9-820">值</span><span class="sxs-lookup"><span data-stu-id="524d9-820">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-821">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-821">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-822">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-822">1.0</span></span>|
|[<span data-ttu-id="524d9-823">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-823">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-824">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-824">ReadItem</span></span>|
|[<span data-ttu-id="524d9-825">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-825">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-826">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-826">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="524d9-827">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-827">Examples</span></span>

<span data-ttu-id="524d9-828">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-828">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="524d9-829">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-829">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="524d9-830">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-830">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="524d9-831">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-831">Reply with a body and a file attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

<span data-ttu-id="524d9-832">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-832">Reply with a body and an item attachment.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

<span data-ttu-id="524d9-833">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="524d9-833">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="524d9-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="524d9-834">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="524d9-835">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="524d9-835">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-836">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-836">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-837">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-837">Requirements</span></span>

|<span data-ttu-id="524d9-838">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-838">Requirement</span></span>| <span data-ttu-id="524d9-839">值</span><span class="sxs-lookup"><span data-stu-id="524d9-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-840">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-841">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-841">1.0</span></span>|
|[<span data-ttu-id="524d9-842">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-843">ReadItem</span></span>|
|[<span data-ttu-id="524d9-844">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-845">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="524d9-846">返回：</span><span class="sxs-lookup"><span data-stu-id="524d9-846">Returns:</span></span>

<span data-ttu-id="524d9-847">类型：[Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="524d9-847">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="524d9-848">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-848">Example</span></span>

<span data-ttu-id="524d9-849">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="524d9-849">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="524d9-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="524d9-850">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="524d9-851">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="524d9-851">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-852">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-852">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-853">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-853">Parameters:</span></span>

|<span data-ttu-id="524d9-854">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-854">Name</span></span>| <span data-ttu-id="524d9-855">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-855">Type</span></span>| <span data-ttu-id="524d9-856">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-856">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="524d9-857">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="524d9-857">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="524d9-858">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="524d9-858">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="524d9-859">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-859">Requirements</span></span>

|<span data-ttu-id="524d9-860">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-860">Requirement</span></span>| <span data-ttu-id="524d9-861">值</span><span class="sxs-lookup"><span data-stu-id="524d9-861">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-862">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-862">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-863">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-863">1.0</span></span>|
|[<span data-ttu-id="524d9-864">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-864">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-865">受限</span><span class="sxs-lookup"><span data-stu-id="524d9-865">Restricted</span></span>|
|[<span data-ttu-id="524d9-866">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-866">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-867">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-867">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="524d9-868">返回：</span><span class="sxs-lookup"><span data-stu-id="524d9-868">Returns:</span></span>

<span data-ttu-id="524d9-869">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="524d9-869">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="524d9-870">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="524d9-870">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="524d9-871">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="524d9-871">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="524d9-872">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="524d9-872">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="524d9-873">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="524d9-873">Value of `entityType`</span></span> | <span data-ttu-id="524d9-874">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="524d9-874">Type of objects in returned array</span></span> | <span data-ttu-id="524d9-875">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-875">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="524d9-876">String</span><span class="sxs-lookup"><span data-stu-id="524d9-876">String</span></span> | <span data-ttu-id="524d9-877">**受限**</span><span class="sxs-lookup"><span data-stu-id="524d9-877">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="524d9-878">Contact</span><span class="sxs-lookup"><span data-stu-id="524d9-878">Contact</span></span> | <span data-ttu-id="524d9-879">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="524d9-879">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="524d9-880">String</span><span class="sxs-lookup"><span data-stu-id="524d9-880">String</span></span> | <span data-ttu-id="524d9-881">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="524d9-881">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="524d9-882">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="524d9-882">MeetingSuggestion</span></span> | <span data-ttu-id="524d9-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="524d9-883">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="524d9-884">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="524d9-884">PhoneNumber</span></span> | <span data-ttu-id="524d9-885">**受限**</span><span class="sxs-lookup"><span data-stu-id="524d9-885">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="524d9-886">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="524d9-886">TaskSuggestion</span></span> | <span data-ttu-id="524d9-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="524d9-887">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="524d9-888">String</span><span class="sxs-lookup"><span data-stu-id="524d9-888">String</span></span> | <span data-ttu-id="524d9-889">**受限**</span><span class="sxs-lookup"><span data-stu-id="524d9-889">**Restricted**</span></span> |

<span data-ttu-id="524d9-890">类型：Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="524d9-890">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="524d9-891">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-891">Example</span></span>

<span data-ttu-id="524d9-892">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="524d9-892">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
}
```

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="524d9-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="524d9-893">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="524d9-894">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="524d9-894">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-895">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-895">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="524d9-896">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="524d9-896">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-897">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-897">Parameters:</span></span>

|<span data-ttu-id="524d9-898">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-898">Name</span></span>| <span data-ttu-id="524d9-899">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-899">Type</span></span>| <span data-ttu-id="524d9-900">描述</span><span class="sxs-lookup"><span data-stu-id="524d9-900">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="524d9-901">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-901">String</span></span>|<span data-ttu-id="524d9-902">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="524d9-902">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="524d9-903">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-903">Requirements</span></span>

|<span data-ttu-id="524d9-904">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-904">Requirement</span></span>| <span data-ttu-id="524d9-905">值</span><span class="sxs-lookup"><span data-stu-id="524d9-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-906">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-907">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-907">1.0</span></span>|
|[<span data-ttu-id="524d9-908">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-909">ReadItem</span></span>|
|[<span data-ttu-id="524d9-910">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-911">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="524d9-912">返回：</span><span class="sxs-lookup"><span data-stu-id="524d9-912">Returns:</span></span>

<span data-ttu-id="524d9-p155">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="524d9-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="524d9-915">类型：Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="524d9-915">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="524d9-916">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="524d9-916">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="524d9-917">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="524d9-917">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-918">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-918">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="524d9-p156">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="524d9-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="524d9-922">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="524d9-922">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="524d9-923">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="524d9-923">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="524d9-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="524d9-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="524d9-927">Requirements</span><span class="sxs-lookup"><span data-stu-id="524d9-927">Requirements</span></span>

|<span data-ttu-id="524d9-928">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-928">Requirement</span></span>| <span data-ttu-id="524d9-929">值</span><span class="sxs-lookup"><span data-stu-id="524d9-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-930">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-931">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-931">1.0</span></span>|
|[<span data-ttu-id="524d9-932">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-932">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-933">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-933">ReadItem</span></span>|
|[<span data-ttu-id="524d9-934">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-934">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-935">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-935">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="524d9-936">返回：</span><span class="sxs-lookup"><span data-stu-id="524d9-936">Returns:</span></span>

<span data-ttu-id="524d9-p158">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="524d9-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="524d9-939">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="524d9-939">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="524d9-940">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-940">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="524d9-941">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-941">Example</span></span>

<span data-ttu-id="524d9-942">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="524d9-942">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="524d9-943">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="524d9-943">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="524d9-944">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="524d9-944">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-945">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="524d9-945">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="524d9-946">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="524d9-946">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="524d9-p159">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="524d9-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-949">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-949">Parameters:</span></span>

|<span data-ttu-id="524d9-950">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-950">Name</span></span>| <span data-ttu-id="524d9-951">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-951">Type</span></span>| <span data-ttu-id="524d9-952">描述</span><span class="sxs-lookup"><span data-stu-id="524d9-952">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="524d9-953">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-953">String</span></span>|<span data-ttu-id="524d9-954">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="524d9-954">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="524d9-955">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-955">Requirements</span></span>

|<span data-ttu-id="524d9-956">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-956">Requirement</span></span>| <span data-ttu-id="524d9-957">值</span><span class="sxs-lookup"><span data-stu-id="524d9-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-958">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-959">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-959">1.0</span></span>|
|[<span data-ttu-id="524d9-960">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-961">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-961">ReadItem</span></span>|
|[<span data-ttu-id="524d9-962">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-963">阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="524d9-964">返回：</span><span class="sxs-lookup"><span data-stu-id="524d9-964">Returns:</span></span>

<span data-ttu-id="524d9-965">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="524d9-965">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="524d9-966">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="524d9-966">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="524d9-967">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="524d9-967">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="524d9-968">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-968">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="524d9-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="524d9-969">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="524d9-970">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="524d9-970">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="524d9-p160">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="524d9-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-973">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-973">Parameters:</span></span>

|<span data-ttu-id="524d9-974">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-974">Name</span></span>| <span data-ttu-id="524d9-975">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-975">Type</span></span>| <span data-ttu-id="524d9-976">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-976">Attributes</span></span>| <span data-ttu-id="524d9-977">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-977">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="524d9-978">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="524d9-978">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="524d9-p161">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="524d9-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="524d9-982">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-982">Object</span></span>| <span data-ttu-id="524d9-983">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-983">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-984">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="524d9-984">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="524d9-985">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-985">Object</span></span>| <span data-ttu-id="524d9-986">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-986">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-987">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-987">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="524d9-988">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-988">function</span></span>||<span data-ttu-id="524d9-989">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-989">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="524d9-990">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="524d9-990">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="524d9-991">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="524d9-991">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="524d9-992">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-992">Requirements</span></span>

|<span data-ttu-id="524d9-993">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-993">Requirement</span></span>| <span data-ttu-id="524d9-994">值</span><span class="sxs-lookup"><span data-stu-id="524d9-994">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-995">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-995">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-996">1.2</span><span class="sxs-lookup"><span data-stu-id="524d9-996">1.2</span></span>|
|[<span data-ttu-id="524d9-997">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-997">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-998">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="524d9-998">ReadWriteItem</span></span>|
|[<span data-ttu-id="524d9-999">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-999">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-1000">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-1000">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="524d9-1001">返回：</span><span class="sxs-lookup"><span data-stu-id="524d9-1001">Returns:</span></span>

<span data-ttu-id="524d9-1002">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="524d9-1002">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="524d9-1003">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="524d9-1003">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="524d9-1004">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-1004">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="524d9-1005">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-1005">Example</span></span>

```js
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="524d9-1006">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="524d9-1006">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="524d9-1007">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="524d9-1007">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="524d9-p163">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="524d9-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-1011">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-1011">Parameters:</span></span>

|<span data-ttu-id="524d9-1012">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-1012">Name</span></span>| <span data-ttu-id="524d9-1013">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-1013">Type</span></span>| <span data-ttu-id="524d9-1014">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-1014">Attributes</span></span>| <span data-ttu-id="524d9-1015">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-1015">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="524d9-1016">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-1016">function</span></span>||<span data-ttu-id="524d9-1017">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-1017">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="524d9-1018">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="524d9-1018">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="524d9-1019">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="524d9-1019">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="524d9-1020">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-1020">Object</span></span>| <span data-ttu-id="524d9-1021">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1021">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1022">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-1022">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="524d9-1023">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="524d9-1023">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="524d9-1024">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-1024">Requirements</span></span>

|<span data-ttu-id="524d9-1025">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-1025">Requirement</span></span>| <span data-ttu-id="524d9-1026">值</span><span class="sxs-lookup"><span data-stu-id="524d9-1026">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-1027">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-1027">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-1028">1.0</span><span class="sxs-lookup"><span data-stu-id="524d9-1028">1.0</span></span>|
|[<span data-ttu-id="524d9-1029">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-1029">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-1030">ReadItem</span><span class="sxs-lookup"><span data-stu-id="524d9-1030">ReadItem</span></span>|
|[<span data-ttu-id="524d9-1031">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-1031">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-1032">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="524d9-1032">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-1033">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-1033">Example</span></span>

<span data-ttu-id="524d9-p166">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="524d9-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
  // After the DOM is loaded, add-in-specific code can run.
  var item = Office.context.mailbox.item;
  item.loadCustomPropertiesAsync(customPropsCallback);
  });
}

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="524d9-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="524d9-1037">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="524d9-1038">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="524d9-1038">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="524d9-p167">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="524d9-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-1043">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-1043">Parameters:</span></span>

|<span data-ttu-id="524d9-1044">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-1044">Name</span></span>| <span data-ttu-id="524d9-1045">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-1045">Type</span></span>| <span data-ttu-id="524d9-1046">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-1046">Attributes</span></span>| <span data-ttu-id="524d9-1047">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-1047">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="524d9-1048">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-1048">String</span></span>||<span data-ttu-id="524d9-1049">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="524d9-1049">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="524d9-1050">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-1050">Object</span></span>| <span data-ttu-id="524d9-1051">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1052">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="524d9-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="524d9-1053">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-1053">Object</span></span>| <span data-ttu-id="524d9-1054">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1055">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="524d9-1056">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-1056">function</span></span>| <span data-ttu-id="524d9-1057">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1058">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="524d9-1059">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="524d9-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="524d9-1060">错误</span><span class="sxs-lookup"><span data-stu-id="524d9-1060">Errors</span></span>

| <span data-ttu-id="524d9-1061">错误代码</span><span class="sxs-lookup"><span data-stu-id="524d9-1061">Error code</span></span> | <span data-ttu-id="524d9-1062">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="524d9-1063">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="524d9-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="524d9-1064">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-1064">Requirements</span></span>

|<span data-ttu-id="524d9-1065">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-1065">Requirement</span></span>| <span data-ttu-id="524d9-1066">值</span><span class="sxs-lookup"><span data-stu-id="524d9-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-1067">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="524d9-1068">1.1</span></span>|
|[<span data-ttu-id="524d9-1069">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="524d9-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="524d9-1071">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-1072">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-1073">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-1073">Example</span></span>

<span data-ttu-id="524d9-1074">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="524d9-1074">The following code removes an attachment with an identifier of '0'.</span></span>

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="524d9-1075">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="524d9-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="524d9-1076">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="524d9-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="524d9-p168">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="524d9-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-1080">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="524d9-1080">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="524d9-1081">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="524d9-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="524d9-p170">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="524d9-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="524d9-1085">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="524d9-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="524d9-1086">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="524d9-1086">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="524d9-1087">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="524d9-1087">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="524d9-1088">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="524d9-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-1089">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-1089">Parameters:</span></span>

|<span data-ttu-id="524d9-1090">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-1090">Name</span></span>| <span data-ttu-id="524d9-1091">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-1091">Type</span></span>| <span data-ttu-id="524d9-1092">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-1092">Attributes</span></span>| <span data-ttu-id="524d9-1093">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="524d9-1094">Object</span><span class="sxs-lookup"><span data-stu-id="524d9-1094">Object</span></span>| <span data-ttu-id="524d9-1095">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1096">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="524d9-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="524d9-1097">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-1097">Object</span></span>| <span data-ttu-id="524d9-1098">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1099">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="524d9-1100">函数</span><span class="sxs-lookup"><span data-stu-id="524d9-1100">function</span></span>||<span data-ttu-id="524d9-1101">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="524d9-1102">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="524d9-1102">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="524d9-1103">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-1103">Requirements</span></span>

|<span data-ttu-id="524d9-1104">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-1104">Requirement</span></span>| <span data-ttu-id="524d9-1105">值</span><span class="sxs-lookup"><span data-stu-id="524d9-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-1106">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="524d9-1107">1.3</span></span>|
|[<span data-ttu-id="524d9-1108">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="524d9-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="524d9-1110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-1111">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="524d9-1112">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-1112">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="524d9-p172">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="524d9-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="524d9-1115">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="524d9-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="524d9-1116">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="524d9-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="524d9-p173">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="524d9-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="524d9-1120">参数：</span><span class="sxs-lookup"><span data-stu-id="524d9-1120">Parameters:</span></span>

|<span data-ttu-id="524d9-1121">名称</span><span class="sxs-lookup"><span data-stu-id="524d9-1121">Name</span></span>| <span data-ttu-id="524d9-1122">类型</span><span class="sxs-lookup"><span data-stu-id="524d9-1122">Type</span></span>| <span data-ttu-id="524d9-1123">属性</span><span class="sxs-lookup"><span data-stu-id="524d9-1123">Attributes</span></span>| <span data-ttu-id="524d9-1124">说明</span><span class="sxs-lookup"><span data-stu-id="524d9-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="524d9-1125">字符串</span><span class="sxs-lookup"><span data-stu-id="524d9-1125">String</span></span>||<span data-ttu-id="524d9-p174">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="524d9-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="524d9-1129">Object</span><span class="sxs-lookup"><span data-stu-id="524d9-1129">Object</span></span>| <span data-ttu-id="524d9-1130">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1131">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="524d9-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="524d9-1132">对象</span><span class="sxs-lookup"><span data-stu-id="524d9-1132">Object</span></span>| <span data-ttu-id="524d9-1133">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-1134">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="524d9-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="524d9-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="524d9-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="524d9-1136">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="524d9-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="524d9-p175">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="524d9-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="524d9-p176">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="524d9-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="524d9-1141">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="524d9-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="524d9-1142">function</span><span class="sxs-lookup"><span data-stu-id="524d9-1142">function</span></span>||<span data-ttu-id="524d9-1143">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="524d9-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="524d9-1144">Requirements</span><span class="sxs-lookup"><span data-stu-id="524d9-1144">Requirements</span></span>

|<span data-ttu-id="524d9-1145">要求</span><span class="sxs-lookup"><span data-stu-id="524d9-1145">Requirement</span></span>| <span data-ttu-id="524d9-1146">值</span><span class="sxs-lookup"><span data-stu-id="524d9-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="524d9-1147">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="524d9-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="524d9-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="524d9-1148">1.2</span></span>|
|[<span data-ttu-id="524d9-1149">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="524d9-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="524d9-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="524d9-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="524d9-1151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="524d9-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="524d9-1152">撰写</span><span class="sxs-lookup"><span data-stu-id="524d9-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="524d9-1153">示例</span><span class="sxs-lookup"><span data-stu-id="524d9-1153">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
