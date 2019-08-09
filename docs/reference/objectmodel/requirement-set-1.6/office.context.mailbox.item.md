---
title: "\"Context\"-\"邮箱\"。项目-要求集1。6"
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: e67664200baea89f14360465b34cdfededa3aa7d
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268339"
---
# <a name="item"></a><span data-ttu-id="02d3c-102">item</span><span class="sxs-lookup"><span data-stu-id="02d3c-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="02d3c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="02d3c-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="02d3c-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="02d3c-106">Requirements</span></span>

|<span data-ttu-id="02d3c-107">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-107">Requirement</span></span>| <span data-ttu-id="02d3c-108">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-110">1.0</span></span>|
|[<span data-ttu-id="02d3c-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-112">受限</span><span class="sxs-lookup"><span data-stu-id="02d3c-112">Restricted</span></span>|
|[<span data-ttu-id="02d3c-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="02d3c-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-115">Members and methods</span></span>

| <span data-ttu-id="02d3c-116">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-116">Member</span></span> | <span data-ttu-id="02d3c-117">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="02d3c-118">attachments</span><span class="sxs-lookup"><span data-stu-id="02d3c-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="02d3c-119">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-119">Member</span></span> |
| [<span data-ttu-id="02d3c-120">bcc</span><span class="sxs-lookup"><span data-stu-id="02d3c-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="02d3c-121">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-121">Member</span></span> |
| [<span data-ttu-id="02d3c-122">body</span><span class="sxs-lookup"><span data-stu-id="02d3c-122">body</span></span>](#body-body) | <span data-ttu-id="02d3c-123">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-123">Member</span></span> |
| [<span data-ttu-id="02d3c-124">cc</span><span class="sxs-lookup"><span data-stu-id="02d3c-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="02d3c-125">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-125">Member</span></span> |
| [<span data-ttu-id="02d3c-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="02d3c-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="02d3c-127">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-127">Member</span></span> |
| [<span data-ttu-id="02d3c-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="02d3c-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="02d3c-129">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-129">Member</span></span> |
| [<span data-ttu-id="02d3c-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="02d3c-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="02d3c-131">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-131">Member</span></span> |
| [<span data-ttu-id="02d3c-132">end</span><span class="sxs-lookup"><span data-stu-id="02d3c-132">end</span></span>](#end-datetime) | <span data-ttu-id="02d3c-133">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-133">Member</span></span> |
| [<span data-ttu-id="02d3c-134">from</span><span class="sxs-lookup"><span data-stu-id="02d3c-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="02d3c-135">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-135">Member</span></span> |
| [<span data-ttu-id="02d3c-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="02d3c-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="02d3c-137">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-137">Member</span></span> |
| [<span data-ttu-id="02d3c-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="02d3c-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="02d3c-139">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-139">Member</span></span> |
| [<span data-ttu-id="02d3c-140">itemId</span><span class="sxs-lookup"><span data-stu-id="02d3c-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="02d3c-141">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-141">Member</span></span> |
| [<span data-ttu-id="02d3c-142">itemType</span><span class="sxs-lookup"><span data-stu-id="02d3c-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="02d3c-143">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-143">Member</span></span> |
| [<span data-ttu-id="02d3c-144">location</span><span class="sxs-lookup"><span data-stu-id="02d3c-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="02d3c-145">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-145">Member</span></span> |
| [<span data-ttu-id="02d3c-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="02d3c-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="02d3c-147">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-147">Member</span></span> |
| [<span data-ttu-id="02d3c-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="02d3c-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="02d3c-149">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-149">Member</span></span> |
| [<span data-ttu-id="02d3c-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="02d3c-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="02d3c-151">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-151">Member</span></span> |
| [<span data-ttu-id="02d3c-152">organizer</span><span class="sxs-lookup"><span data-stu-id="02d3c-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="02d3c-153">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-153">Member</span></span> |
| [<span data-ttu-id="02d3c-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="02d3c-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="02d3c-155">Member</span><span class="sxs-lookup"><span data-stu-id="02d3c-155">Member</span></span> |
| [<span data-ttu-id="02d3c-156">sender</span><span class="sxs-lookup"><span data-stu-id="02d3c-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="02d3c-157">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-157">Member</span></span> |
| [<span data-ttu-id="02d3c-158">start</span><span class="sxs-lookup"><span data-stu-id="02d3c-158">start</span></span>](#start-datetime) | <span data-ttu-id="02d3c-159">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-159">Member</span></span> |
| [<span data-ttu-id="02d3c-160">subject</span><span class="sxs-lookup"><span data-stu-id="02d3c-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="02d3c-161">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-161">Member</span></span> |
| [<span data-ttu-id="02d3c-162">to</span><span class="sxs-lookup"><span data-stu-id="02d3c-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="02d3c-163">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-163">Member</span></span> |
| [<span data-ttu-id="02d3c-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="02d3c-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="02d3c-165">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-165">Method</span></span> |
| [<span data-ttu-id="02d3c-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="02d3c-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="02d3c-167">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-167">Method</span></span> |
| [<span data-ttu-id="02d3c-168">close</span><span class="sxs-lookup"><span data-stu-id="02d3c-168">close</span></span>](#close) | <span data-ttu-id="02d3c-169">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-169">Method</span></span> |
| [<span data-ttu-id="02d3c-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="02d3c-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="02d3c-171">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-171">Method</span></span> |
| [<span data-ttu-id="02d3c-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="02d3c-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="02d3c-173">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-173">Method</span></span> |
| [<span data-ttu-id="02d3c-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="02d3c-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="02d3c-175">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-175">Method</span></span> |
| [<span data-ttu-id="02d3c-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="02d3c-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="02d3c-177">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-177">Method</span></span> |
| [<span data-ttu-id="02d3c-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="02d3c-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="02d3c-179">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-179">Method</span></span> |
| [<span data-ttu-id="02d3c-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="02d3c-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="02d3c-181">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-181">Method</span></span> |
| [<span data-ttu-id="02d3c-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="02d3c-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="02d3c-183">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-183">Method</span></span> |
| [<span data-ttu-id="02d3c-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="02d3c-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="02d3c-185">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-185">Method</span></span> |
| [<span data-ttu-id="02d3c-186">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="02d3c-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="02d3c-187">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-187">Method</span></span> |
| [<span data-ttu-id="02d3c-188">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="02d3c-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="02d3c-189">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-189">Method</span></span> |
| [<span data-ttu-id="02d3c-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="02d3c-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="02d3c-191">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-191">Method</span></span> |
| [<span data-ttu-id="02d3c-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="02d3c-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="02d3c-193">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-193">Method</span></span> |
| [<span data-ttu-id="02d3c-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="02d3c-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="02d3c-195">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-195">Method</span></span> |
| [<span data-ttu-id="02d3c-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="02d3c-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="02d3c-197">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="02d3c-198">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-198">Example</span></span>

<span data-ttu-id="02d3c-199">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="02d3c-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
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
};
```

### <a name="members"></a><span data-ttu-id="02d3c-200">成员</span><span class="sxs-lookup"><span data-stu-id="02d3c-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="02d3c-201">附件: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="02d3c-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="02d3c-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-204">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="02d3c-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="02d3c-205">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="02d3c-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-206">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-206">Type</span></span>

*   <span data-ttu-id="02d3c-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="02d3c-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-208">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-208">Requirements</span></span>

|<span data-ttu-id="02d3c-209">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-209">Requirement</span></span>| <span data-ttu-id="02d3c-210">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-212">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-212">1.0</span></span>|
|[<span data-ttu-id="02d3c-213">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-214">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-215">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-216">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-217">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-217">Example</span></span>

<span data-ttu-id="02d3c-218">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="02d3c-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="02d3c-219">密件抄送:[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-220">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="02d3c-221">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-222">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-222">Type</span></span>

*   [<span data-ttu-id="02d3c-223">收件人</span><span class="sxs-lookup"><span data-stu-id="02d3c-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="02d3c-224">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-224">Requirements</span></span>

|<span data-ttu-id="02d3c-225">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-225">Requirement</span></span>| <span data-ttu-id="02d3c-226">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-228">1.1</span><span class="sxs-lookup"><span data-stu-id="02d3c-228">1.1</span></span>|
|[<span data-ttu-id="02d3c-229">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-230">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-231">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-232">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-233">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-233">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="02d3c-234">正文:[正文](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-235">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-236">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-236">Type</span></span>

*   [<span data-ttu-id="02d3c-237">Body</span><span class="sxs-lookup"><span data-stu-id="02d3c-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="02d3c-238">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-238">Requirements</span></span>

|<span data-ttu-id="02d3c-239">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-239">Requirement</span></span>| <span data-ttu-id="02d3c-240">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-242">1.1</span><span class="sxs-lookup"><span data-stu-id="02d3c-242">1.1</span></span>|
|[<span data-ttu-id="02d3c-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-244">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-247">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-247">Example</span></span>

<span data-ttu-id="02d3c-248">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="02d3c-248">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="02d3c-249">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="02d3c-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="02d3c-250"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="02d3c-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-251">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02d3c-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="02d3c-252">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-253">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-253">Read mode</span></span>

<span data-ttu-id="02d3c-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-256">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-256">Compose mode</span></span>

<span data-ttu-id="02d3c-257">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-257">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02d3c-258">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-258">Type</span></span>

*   <span data-ttu-id="02d3c-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-260">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-260">Requirements</span></span>

|<span data-ttu-id="02d3c-261">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-261">Requirement</span></span>| <span data-ttu-id="02d3c-262">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-264">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-264">1.0</span></span>|
|[<span data-ttu-id="02d3c-265">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-266">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-267">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-268">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-268">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="02d3c-269">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="02d3c-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="02d3c-270">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="02d3c-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="02d3c-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-275">Type</span><span class="sxs-lookup"><span data-stu-id="02d3c-275">Type</span></span>

*   <span data-ttu-id="02d3c-276">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-277">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-277">Requirements</span></span>

|<span data-ttu-id="02d3c-278">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-278">Requirement</span></span>| <span data-ttu-id="02d3c-279">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-281">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-281">1.0</span></span>|
|[<span data-ttu-id="02d3c-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-283">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-285">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-286">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-286">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="02d3c-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="02d3c-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="02d3c-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-290">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-290">Type</span></span>

*   <span data-ttu-id="02d3c-291">日期</span><span class="sxs-lookup"><span data-stu-id="02d3c-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-292">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-292">Requirements</span></span>

|<span data-ttu-id="02d3c-293">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-293">Requirement</span></span>| <span data-ttu-id="02d3c-294">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-295">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-296">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-296">1.0</span></span>|
|[<span data-ttu-id="02d3c-297">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-298">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-299">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-300">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-301">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-301">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="02d3c-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="02d3c-302">dateTimeModified: Date</span></span>

<span data-ttu-id="02d3c-303">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02d3c-303">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="02d3c-304">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-304">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-305">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="02d3c-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-306">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-306">Type</span></span>

*   <span data-ttu-id="02d3c-307">日期</span><span class="sxs-lookup"><span data-stu-id="02d3c-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-308">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-308">Requirements</span></span>

|<span data-ttu-id="02d3c-309">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-309">Requirement</span></span>| <span data-ttu-id="02d3c-310">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-312">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-312">1.0</span></span>|
|[<span data-ttu-id="02d3c-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-314">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-316">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-317">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-317">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="02d3c-318">结束: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-319">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02d3c-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="02d3c-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-322">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-322">Read mode</span></span>

<span data-ttu-id="02d3c-323">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-323">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-324">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-324">Compose mode</span></span>

<span data-ttu-id="02d3c-325">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="02d3c-326">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="02d3c-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="02d3c-327">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="02d3c-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="02d3c-328">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-328">Type</span></span>

*   <span data-ttu-id="02d3c-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-330">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-330">Requirements</span></span>

|<span data-ttu-id="02d3c-331">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-331">Requirement</span></span>| <span data-ttu-id="02d3c-332">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-333">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-334">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-334">1.0</span></span>|
|[<span data-ttu-id="02d3c-335">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-336">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-337">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-338">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-338">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="02d3c-339">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="02d3c-p113">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-344">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-345">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-345">Type</span></span>

*   [<span data-ttu-id="02d3c-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="02d3c-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="02d3c-347">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="02d3c-348">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-348">Requirements</span></span>

|<span data-ttu-id="02d3c-349">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-349">Requirement</span></span>| <span data-ttu-id="02d3c-350">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-351">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-352">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-352">1.0</span></span>|
|[<span data-ttu-id="02d3c-353">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-354">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-355">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-356">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-356">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="02d3c-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="02d3c-357">internetMessageId: String</span></span>

<span data-ttu-id="02d3c-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-360">Type</span><span class="sxs-lookup"><span data-stu-id="02d3c-360">Type</span></span>

*   <span data-ttu-id="02d3c-361">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-362">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-362">Requirements</span></span>

|<span data-ttu-id="02d3c-363">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-363">Requirement</span></span>| <span data-ttu-id="02d3c-364">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-365">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-366">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-366">1.0</span></span>|
|[<span data-ttu-id="02d3c-367">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-368">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-369">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-370">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-371">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-371">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="02d3c-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="02d3c-372">itemClass: String</span></span>

<span data-ttu-id="02d3c-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="02d3c-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="02d3c-377">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-377">Type</span></span> | <span data-ttu-id="02d3c-378">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-378">Description</span></span> | <span data-ttu-id="02d3c-379">项目类</span><span class="sxs-lookup"><span data-stu-id="02d3c-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="02d3c-380">约会项目</span><span class="sxs-lookup"><span data-stu-id="02d3c-380">Appointment items</span></span> | <span data-ttu-id="02d3c-381">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="02d3c-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="02d3c-382">邮件项目</span><span class="sxs-lookup"><span data-stu-id="02d3c-382">Message items</span></span> | <span data-ttu-id="02d3c-383">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="02d3c-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="02d3c-384">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-385">Type</span><span class="sxs-lookup"><span data-stu-id="02d3c-385">Type</span></span>

*   <span data-ttu-id="02d3c-386">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-387">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-387">Requirements</span></span>

|<span data-ttu-id="02d3c-388">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-388">Requirement</span></span>| <span data-ttu-id="02d3c-389">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-390">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-391">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-391">1.0</span></span>|
|[<span data-ttu-id="02d3c-392">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-393">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-394">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-395">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-396">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-396">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="02d3c-397">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="02d3c-397">(nullable) itemId: String</span></span>

<span data-ttu-id="02d3c-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-400">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="02d3c-400">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="02d3c-401">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="02d3c-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="02d3c-402">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="02d3c-402">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="02d3c-403">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="02d3c-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="02d3c-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-406">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-406">Type</span></span>

*   <span data-ttu-id="02d3c-407">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-408">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-408">Requirements</span></span>

|<span data-ttu-id="02d3c-409">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-409">Requirement</span></span>| <span data-ttu-id="02d3c-410">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-412">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-412">1.0</span></span>|
|[<span data-ttu-id="02d3c-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-414">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-416">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-417">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-417">Example</span></span>

<span data-ttu-id="02d3c-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="02d3c-420">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-420">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-421">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="02d3c-421">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="02d3c-422">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="02d3c-422">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-423">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-423">Type</span></span>

*   [<span data-ttu-id="02d3c-424">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="02d3c-424">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="02d3c-425">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-425">Requirements</span></span>

|<span data-ttu-id="02d3c-426">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-426">Requirement</span></span>| <span data-ttu-id="02d3c-427">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-428">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-429">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-429">1.0</span></span>|
|[<span data-ttu-id="02d3c-430">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-431">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-432">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-433">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-434">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-434">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="02d3c-435">位置: 字符串 |[位置](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-435">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-436">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="02d3c-436">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-437">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-437">Read mode</span></span>

<span data-ttu-id="02d3c-438">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="02d3c-438">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-439">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-439">Compose mode</span></span>

<span data-ttu-id="02d3c-440">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-440">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02d3c-441">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-441">Type</span></span>

*   <span data-ttu-id="02d3c-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-442">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-443">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-443">Requirements</span></span>

|<span data-ttu-id="02d3c-444">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-444">Requirement</span></span>| <span data-ttu-id="02d3c-445">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-446">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-447">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-447">1.0</span></span>|
|[<span data-ttu-id="02d3c-448">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-448">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-449">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-450">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-450">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-451">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-451">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="02d3c-452">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="02d3c-452">normalizedSubject: String</span></span>

<span data-ttu-id="02d3c-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="02d3c-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-457">Type</span><span class="sxs-lookup"><span data-stu-id="02d3c-457">Type</span></span>

*   <span data-ttu-id="02d3c-458">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-458">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-459">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-459">Requirements</span></span>

|<span data-ttu-id="02d3c-460">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-460">Requirement</span></span>| <span data-ttu-id="02d3c-461">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-462">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-463">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-463">1.0</span></span>|
|[<span data-ttu-id="02d3c-464">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-465">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-466">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-467">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-468">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-468">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="02d3c-469">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-469">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-470">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-470">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-471">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-471">Type</span></span>

*   [<span data-ttu-id="02d3c-472">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="02d3c-472">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="02d3c-473">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-473">Requirements</span></span>

|<span data-ttu-id="02d3c-474">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-474">Requirement</span></span>| <span data-ttu-id="02d3c-475">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-475">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-476">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-477">1.3</span><span class="sxs-lookup"><span data-stu-id="02d3c-477">1.3</span></span>|
|[<span data-ttu-id="02d3c-478">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-479">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-479">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-480">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-481">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-481">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-482">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-482">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="02d3c-483">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)的数组</span><span class="sxs-lookup"><span data-stu-id="02d3c-483">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-484">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02d3c-484">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="02d3c-485">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-485">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-486">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-486">Read mode</span></span>

<span data-ttu-id="02d3c-487">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-487">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-488">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-488">Compose mode</span></span>

<span data-ttu-id="02d3c-489">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-489">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02d3c-490">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-490">Type</span></span>

*   <span data-ttu-id="02d3c-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-491">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-492">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-492">Requirements</span></span>

|<span data-ttu-id="02d3c-493">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-493">Requirement</span></span>| <span data-ttu-id="02d3c-494">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-495">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-496">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-496">1.0</span></span>|
|[<span data-ttu-id="02d3c-497">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-498">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-499">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-500">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-500">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="02d3c-501">组织者: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-501">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-504">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-504">Type</span></span>

*   [<span data-ttu-id="02d3c-505">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="02d3c-505">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="02d3c-506">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-506">Requirements</span></span>

|<span data-ttu-id="02d3c-507">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-507">Requirement</span></span>| <span data-ttu-id="02d3c-508">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-509">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-510">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-510">1.0</span></span>|
|[<span data-ttu-id="02d3c-511">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-512">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-513">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-514">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-514">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-515">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-515">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="02d3c-516">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)的数组</span><span class="sxs-lookup"><span data-stu-id="02d3c-516">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-517">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02d3c-517">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="02d3c-518">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-518">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-519">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-519">Read mode</span></span>

<span data-ttu-id="02d3c-520">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-520">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-521">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-521">Compose mode</span></span>

<span data-ttu-id="02d3c-522">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-522">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="02d3c-523">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-523">Type</span></span>

*   <span data-ttu-id="02d3c-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-524">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-525">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-525">Requirements</span></span>

|<span data-ttu-id="02d3c-526">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-526">Requirement</span></span>| <span data-ttu-id="02d3c-527">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-527">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-528">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-528">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-529">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-529">1.0</span></span>|
|[<span data-ttu-id="02d3c-530">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-530">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-531">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-531">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-532">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-532">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-533">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-533">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="02d3c-534">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-534">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="02d3c-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-539">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-539">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="02d3c-540">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-540">Type</span></span>

*   [<span data-ttu-id="02d3c-541">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="02d3c-541">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="02d3c-542">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-542">Requirements</span></span>

|<span data-ttu-id="02d3c-543">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-543">Requirement</span></span>| <span data-ttu-id="02d3c-544">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-545">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-546">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-546">1.0</span></span>|
|[<span data-ttu-id="02d3c-547">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-548">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-549">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-550">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-550">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-551">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-551">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="02d3c-552">开始日期: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-552">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-553">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02d3c-553">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="02d3c-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-556">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-556">Read mode</span></span>

<span data-ttu-id="02d3c-557">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-557">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-558">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-558">Compose mode</span></span>

<span data-ttu-id="02d3c-559">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-559">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="02d3c-560">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="02d3c-560">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="02d3c-561">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="02d3c-561">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a><span data-ttu-id="02d3c-562">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-562">Type</span></span>

*   <span data-ttu-id="02d3c-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-563">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-564">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-564">Requirements</span></span>

|<span data-ttu-id="02d3c-565">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-565">Requirement</span></span>| <span data-ttu-id="02d3c-566">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-567">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-568">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-568">1.0</span></span>|
|[<span data-ttu-id="02d3c-569">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-570">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-571">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-572">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-572">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="02d3c-573">subject: String |[主题](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-573">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-574">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="02d3c-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="02d3c-575">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="02d3c-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-576">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-576">Read mode</span></span>

<span data-ttu-id="02d3c-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-579">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-579">Compose mode</span></span>

<span data-ttu-id="02d3c-580">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="02d3c-581">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-581">Type</span></span>

*   <span data-ttu-id="02d3c-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-582">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-583">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-583">Requirements</span></span>

|<span data-ttu-id="02d3c-584">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-584">Requirement</span></span>| <span data-ttu-id="02d3c-585">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-586">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-587">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-587">1.0</span></span>|
|[<span data-ttu-id="02d3c-588">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-588">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-589">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-590">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-590">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-591">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-591">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="02d3c-592">to: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)的数组</span><span class="sxs-lookup"><span data-stu-id="02d3c-592">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="02d3c-593">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02d3c-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="02d3c-594">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02d3c-595">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-595">Read mode</span></span>

<span data-ttu-id="02d3c-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="02d3c-598">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-598">Compose mode</span></span>

<span data-ttu-id="02d3c-599">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02d3c-600">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-600">Type</span></span>

*   <span data-ttu-id="02d3c-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-601">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-602">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-602">Requirements</span></span>

|<span data-ttu-id="02d3c-603">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-603">Requirement</span></span>| <span data-ttu-id="02d3c-604">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-605">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-606">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-606">1.0</span></span>|
|[<span data-ttu-id="02d3c-607">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-608">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-609">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-610">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-610">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="02d3c-611">方法</span><span class="sxs-lookup"><span data-stu-id="02d3c-611">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="02d3c-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="02d3c-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="02d3c-613">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="02d3c-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="02d3c-614">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="02d3c-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="02d3c-615">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-616">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-616">Parameters</span></span>

|<span data-ttu-id="02d3c-617">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-617">Name</span></span>| <span data-ttu-id="02d3c-618">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-618">Type</span></span>| <span data-ttu-id="02d3c-619">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-619">Attributes</span></span>| <span data-ttu-id="02d3c-620">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="02d3c-621">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-621">String</span></span>||<span data-ttu-id="02d3c-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="02d3c-624">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-624">String</span></span>||<span data-ttu-id="02d3c-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="02d3c-627">Object</span><span class="sxs-lookup"><span data-stu-id="02d3c-627">Object</span></span>| <span data-ttu-id="02d3c-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-628">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-629">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02d3c-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="02d3c-630">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-630">Object</span></span> | <span data-ttu-id="02d3c-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-631">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-632">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="02d3c-633">布尔值</span><span class="sxs-lookup"><span data-stu-id="02d3c-633">Boolean</span></span> | <span data-ttu-id="02d3c-634">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-634">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-635">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="02d3c-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="02d3c-636">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-636">function</span></span>| <span data-ttu-id="02d3c-637">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-637">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-638">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="02d3c-639">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="02d3c-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="02d3c-640">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="02d3c-641">错误</span><span class="sxs-lookup"><span data-stu-id="02d3c-641">Errors</span></span>

| <span data-ttu-id="02d3c-642">错误代码</span><span class="sxs-lookup"><span data-stu-id="02d3c-642">Error code</span></span> | <span data-ttu-id="02d3c-643">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="02d3c-644">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="02d3c-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="02d3c-645">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="02d3c-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="02d3c-646">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="02d3c-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02d3c-647">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-647">Requirements</span></span>

|<span data-ttu-id="02d3c-648">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-648">Requirement</span></span>| <span data-ttu-id="02d3c-649">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-650">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-651">1.1</span><span class="sxs-lookup"><span data-stu-id="02d3c-651">1.1</span></span>|
|[<span data-ttu-id="02d3c-652">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-652">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="02d3c-654">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-654">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-655">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="02d3c-656">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-656">Examples</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="02d3c-657">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```javascript
Office.context.mailbox.item.addFileAttachmentAsync(
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
        // Do something here.
      });
  });
```

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="02d3c-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="02d3c-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="02d3c-659">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="02d3c-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="02d3c-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="02d3c-663">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="02d3c-664">如果 Office 外接程序在 web 上的 Outlook 中运行, 则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是, 不支持这种情况, 建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="02d3c-664">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-665">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-665">Parameters</span></span>

|<span data-ttu-id="02d3c-666">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-666">Name</span></span>| <span data-ttu-id="02d3c-667">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-667">Type</span></span>| <span data-ttu-id="02d3c-668">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-668">Attributes</span></span>| <span data-ttu-id="02d3c-669">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="02d3c-670">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-670">String</span></span>||<span data-ttu-id="02d3c-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="02d3c-673">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-673">String</span></span>||<span data-ttu-id="02d3c-674">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="02d3c-674">The subject of the item to be attached.</span></span> <span data-ttu-id="02d3c-675">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-675">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="02d3c-676">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-676">Object</span></span>| <span data-ttu-id="02d3c-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-677">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-678">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02d3c-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02d3c-679">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-679">Object</span></span>| <span data-ttu-id="02d3c-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-680">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-681">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="02d3c-682">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-682">function</span></span>| <span data-ttu-id="02d3c-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-683">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-684">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="02d3c-685">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="02d3c-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="02d3c-686">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="02d3c-687">错误</span><span class="sxs-lookup"><span data-stu-id="02d3c-687">Errors</span></span>

| <span data-ttu-id="02d3c-688">错误代码</span><span class="sxs-lookup"><span data-stu-id="02d3c-688">Error code</span></span> | <span data-ttu-id="02d3c-689">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="02d3c-690">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="02d3c-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02d3c-691">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-691">Requirements</span></span>

|<span data-ttu-id="02d3c-692">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-692">Requirement</span></span>| <span data-ttu-id="02d3c-693">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-694">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-695">1.1</span><span class="sxs-lookup"><span data-stu-id="02d3c-695">1.1</span></span>|
|[<span data-ttu-id="02d3c-696">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-696">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="02d3c-698">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-698">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-699">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-700">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-700">Example</span></span>

<span data-ttu-id="02d3c-701">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

#### <a name="close"></a><span data-ttu-id="02d3c-702">close()</span><span class="sxs-lookup"><span data-stu-id="02d3c-702">close()</span></span>

<span data-ttu-id="02d3c-703">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="02d3c-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="02d3c-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-706">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="02d3c-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="02d3c-707">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="02d3c-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-708">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-708">Requirements</span></span>

|<span data-ttu-id="02d3c-709">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-709">Requirement</span></span>| <span data-ttu-id="02d3c-710">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-711">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-712">1.3</span><span class="sxs-lookup"><span data-stu-id="02d3c-712">1.3</span></span>|
|[<span data-ttu-id="02d3c-713">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-713">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-714">受限</span><span class="sxs-lookup"><span data-stu-id="02d3c-714">Restricted</span></span>|
|[<span data-ttu-id="02d3c-715">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-715">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-716">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-716">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="02d3c-717">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="02d3c-717">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="02d3c-718">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="02d3c-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-719">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-719">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02d3c-720">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-720">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="02d3c-721">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="02d3c-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="02d3c-722">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-722">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="02d3c-723">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="02d3c-723">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="02d3c-724">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="02d3c-724">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-725">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-725">Parameters</span></span>

| <span data-ttu-id="02d3c-726">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-726">Name</span></span> | <span data-ttu-id="02d3c-727">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-727">Type</span></span> | <span data-ttu-id="02d3c-728">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-728">Attributes</span></span> | <span data-ttu-id="02d3c-729">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="02d3c-730">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-730">String &#124; Object</span></span>| |<span data-ttu-id="02d3c-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="02d3c-733">**或**</span><span class="sxs-lookup"><span data-stu-id="02d3c-733">**OR**</span></span><br/><span data-ttu-id="02d3c-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="02d3c-736">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-736">String</span></span> | <span data-ttu-id="02d3c-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-737">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="02d3c-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="02d3c-741">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-741">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-742">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="02d3c-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="02d3c-743">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-743">String</span></span> | | <span data-ttu-id="02d3c-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="02d3c-746">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-746">String</span></span> | | <span data-ttu-id="02d3c-747">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="02d3c-748">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-748">String</span></span> | | <span data-ttu-id="02d3c-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="02d3c-751">布尔</span><span class="sxs-lookup"><span data-stu-id="02d3c-751">Boolean</span></span> | | <span data-ttu-id="02d3c-p144">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="02d3c-754">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-754">String</span></span> | | <span data-ttu-id="02d3c-p145">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="02d3c-758">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-758">function</span></span> | <span data-ttu-id="02d3c-759">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-759">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-760">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02d3c-761">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-761">Requirements</span></span>

|<span data-ttu-id="02d3c-762">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-762">Requirement</span></span>| <span data-ttu-id="02d3c-763">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-764">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-765">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-765">1.0</span></span>|
|[<span data-ttu-id="02d3c-766">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-766">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-767">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-768">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-768">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-769">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="02d3c-770">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-770">Examples</span></span>

<span data-ttu-id="02d3c-771">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="02d3c-772">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-772">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="02d3c-773">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-773">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="02d3c-774">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-774">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="02d3c-775">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-775">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="02d3c-776">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="02d3c-777">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="02d3c-777">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="02d3c-778">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="02d3c-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-779">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-779">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02d3c-780">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-780">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="02d3c-781">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="02d3c-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="02d3c-782">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-782">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="02d3c-783">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="02d3c-783">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="02d3c-784">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="02d3c-784">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-785">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-785">Parameters</span></span>

| <span data-ttu-id="02d3c-786">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-786">Name</span></span> | <span data-ttu-id="02d3c-787">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-787">Type</span></span> | <span data-ttu-id="02d3c-788">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-788">Attributes</span></span> | <span data-ttu-id="02d3c-789">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="02d3c-790">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-790">String &#124; Object</span></span>| | <span data-ttu-id="02d3c-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="02d3c-793">**或**</span><span class="sxs-lookup"><span data-stu-id="02d3c-793">**OR**</span></span><br/><span data-ttu-id="02d3c-p148">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="02d3c-796">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-796">String</span></span> | <span data-ttu-id="02d3c-797">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-797">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-p149">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="02d3c-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="02d3c-801">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-801">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-802">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="02d3c-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="02d3c-803">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-803">String</span></span> | | <span data-ttu-id="02d3c-p150">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="02d3c-806">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-806">String</span></span> | | <span data-ttu-id="02d3c-807">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="02d3c-808">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-808">String</span></span> | | <span data-ttu-id="02d3c-p151">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="02d3c-811">布尔</span><span class="sxs-lookup"><span data-stu-id="02d3c-811">Boolean</span></span> | | <span data-ttu-id="02d3c-p152">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="02d3c-814">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-814">String</span></span> | | <span data-ttu-id="02d3c-p153">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="02d3c-818">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-818">function</span></span> | <span data-ttu-id="02d3c-819">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-819">&lt;optional&gt;</span></span> | <span data-ttu-id="02d3c-820">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02d3c-821">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-821">Requirements</span></span>

|<span data-ttu-id="02d3c-822">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-822">Requirement</span></span>| <span data-ttu-id="02d3c-823">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-824">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-825">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-825">1.0</span></span>|
|[<span data-ttu-id="02d3c-826">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-826">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-827">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-828">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-828">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-829">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="02d3c-830">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-830">Examples</span></span>

<span data-ttu-id="02d3c-831">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="02d3c-832">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-832">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="02d3c-833">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-833">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="02d3c-834">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-834">Reply with a body and a file attachment.</span></span>

```javascript
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

<span data-ttu-id="02d3c-835">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-835">Reply with a body and an item attachment.</span></span>

```javascript
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

<span data-ttu-id="02d3c-836">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="02d3c-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```javascript
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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="02d3c-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="02d3c-837">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="02d3c-838">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-838">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-839">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-840">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-840">Requirements</span></span>

|<span data-ttu-id="02d3c-841">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-841">Requirement</span></span>| <span data-ttu-id="02d3c-842">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-843">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-844">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-844">1.0</span></span>|
|[<span data-ttu-id="02d3c-845">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-846">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-847">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-848">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-849">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-849">Returns:</span></span>

<span data-ttu-id="02d3c-850">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-850">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="02d3c-851">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-851">Example</span></span>

<span data-ttu-id="02d3c-852">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-852">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="02d3c-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="02d3c-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="02d3c-854">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="02d3c-854">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-855">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-855">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-856">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-856">Parameters</span></span>

|<span data-ttu-id="02d3c-857">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-857">Name</span></span>| <span data-ttu-id="02d3c-858">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-858">Type</span></span>| <span data-ttu-id="02d3c-859">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="02d3c-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="02d3c-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="02d3c-861">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="02d3c-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02d3c-862">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-862">Requirements</span></span>

|<span data-ttu-id="02d3c-863">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-863">Requirement</span></span>| <span data-ttu-id="02d3c-864">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-865">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-866">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-866">1.0</span></span>|
|[<span data-ttu-id="02d3c-867">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-868">受限</span><span class="sxs-lookup"><span data-stu-id="02d3c-868">Restricted</span></span>|
|[<span data-ttu-id="02d3c-869">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-870">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-871">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-871">Returns:</span></span>

<span data-ttu-id="02d3c-872">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="02d3c-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="02d3c-873">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="02d3c-873">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="02d3c-874">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="02d3c-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="02d3c-875">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="02d3c-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="02d3c-876">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="02d3c-876">Value of `entityType`</span></span> | <span data-ttu-id="02d3c-877">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-877">Type of objects in returned array</span></span> | <span data-ttu-id="02d3c-878">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="02d3c-879">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-879">String</span></span> | <span data-ttu-id="02d3c-880">**受限**</span><span class="sxs-lookup"><span data-stu-id="02d3c-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="02d3c-881">Contact</span><span class="sxs-lookup"><span data-stu-id="02d3c-881">Contact</span></span> | <span data-ttu-id="02d3c-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02d3c-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="02d3c-883">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-883">String</span></span> | <span data-ttu-id="02d3c-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02d3c-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="02d3c-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="02d3c-885">MeetingSuggestion</span></span> | <span data-ttu-id="02d3c-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02d3c-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="02d3c-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="02d3c-887">PhoneNumber</span></span> | <span data-ttu-id="02d3c-888">**受限**</span><span class="sxs-lookup"><span data-stu-id="02d3c-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="02d3c-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="02d3c-889">TaskSuggestion</span></span> | <span data-ttu-id="02d3c-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02d3c-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="02d3c-891">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-891">String</span></span> | <span data-ttu-id="02d3c-892">**受限**</span><span class="sxs-lookup"><span data-stu-id="02d3c-892">**Restricted**</span></span> |

<span data-ttu-id="02d3c-893">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="02d3c-893">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="02d3c-894">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-894">Example</span></span>

<span data-ttu-id="02d3c-895">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="02d3c-895">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

```javascript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="02d3c-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="02d3c-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="02d3c-897">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-898">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-898">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02d3c-899">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-900">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-900">Parameters</span></span>

|<span data-ttu-id="02d3c-901">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-901">Name</span></span>| <span data-ttu-id="02d3c-902">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-902">Type</span></span>| <span data-ttu-id="02d3c-903">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="02d3c-904">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-904">String</span></span>|<span data-ttu-id="02d3c-905">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="02d3c-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02d3c-906">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-906">Requirements</span></span>

|<span data-ttu-id="02d3c-907">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-907">Requirement</span></span>| <span data-ttu-id="02d3c-908">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-909">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-910">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-910">1.0</span></span>|
|[<span data-ttu-id="02d3c-911">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-911">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-912">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-913">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-913">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-914">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-915">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-915">Returns:</span></span>

<span data-ttu-id="02d3c-p155">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="02d3c-918">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="02d3c-918">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="02d3c-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="02d3c-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="02d3c-920">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="02d3c-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-921">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-921">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02d3c-p156">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="02d3c-925">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="02d3c-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="02d3c-926">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="02d3c-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="02d3c-930">Requirements</span></span>

|<span data-ttu-id="02d3c-931">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-931">Requirement</span></span>| <span data-ttu-id="02d3c-932">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-933">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-934">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-934">1.0</span></span>|
|[<span data-ttu-id="02d3c-935">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-936">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-937">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-938">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-939">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-939">Returns:</span></span>

<span data-ttu-id="02d3c-p158">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="02d3c-942">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="02d3c-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="02d3c-943">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="02d3c-944">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-944">Example</span></span>

<span data-ttu-id="02d3c-945">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="02d3c-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="02d3c-946">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="02d3c-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="02d3c-947">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="02d3c-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-948">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-948">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02d3c-949">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="02d3c-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="02d3c-p159">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-952">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-952">Parameters</span></span>

|<span data-ttu-id="02d3c-953">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-953">Name</span></span>| <span data-ttu-id="02d3c-954">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-954">Type</span></span>| <span data-ttu-id="02d3c-955">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="02d3c-956">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-956">String</span></span>|<span data-ttu-id="02d3c-957">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="02d3c-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02d3c-958">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-958">Requirements</span></span>

|<span data-ttu-id="02d3c-959">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-959">Requirement</span></span>| <span data-ttu-id="02d3c-960">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-961">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-962">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-962">1.0</span></span>|
|[<span data-ttu-id="02d3c-963">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-963">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-964">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-965">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-965">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-966">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-967">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-967">Returns:</span></span>

<span data-ttu-id="02d3c-968">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="02d3c-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="02d3c-969">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="02d3c-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="02d3c-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="02d3c-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="02d3c-971">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-971">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="02d3c-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="02d3c-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="02d3c-973">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="02d3c-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="02d3c-p160">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-976">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-976">Parameters</span></span>

|<span data-ttu-id="02d3c-977">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-977">Name</span></span>| <span data-ttu-id="02d3c-978">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-978">Type</span></span>| <span data-ttu-id="02d3c-979">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-979">Attributes</span></span>| <span data-ttu-id="02d3c-980">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="02d3c-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="02d3c-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="02d3c-p161">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="02d3c-985">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-985">Object</span></span>| <span data-ttu-id="02d3c-986">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-986">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-987">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02d3c-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02d3c-988">Object</span><span class="sxs-lookup"><span data-stu-id="02d3c-988">Object</span></span>| <span data-ttu-id="02d3c-989">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-989">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-990">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="02d3c-991">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-991">function</span></span>||<span data-ttu-id="02d3c-992">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="02d3c-993">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="02d3c-994">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-994">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02d3c-995">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-995">Requirements</span></span>

|<span data-ttu-id="02d3c-996">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-996">Requirement</span></span>| <span data-ttu-id="02d3c-997">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-998">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-999">1.2</span><span class="sxs-lookup"><span data-stu-id="02d3c-999">1.2</span></span>|
|[<span data-ttu-id="02d3c-1000">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-1000">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="02d3c-1002">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-1002">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-1003">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-1004">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-1004">Returns:</span></span>

<span data-ttu-id="02d3c-1005">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="02d3c-1006">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="02d3c-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="02d3c-1007">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="02d3c-1008">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-1008">Example</span></span>

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="02d3c-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="02d3c-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="02d3c-1010">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1010">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="02d3c-1011">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1011">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-1012">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1012">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-1013">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1013">Requirements</span></span>

|<span data-ttu-id="02d3c-1014">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1014">Requirement</span></span>| <span data-ttu-id="02d3c-1015">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-1016">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="02d3c-1017">1.6</span></span> |
|[<span data-ttu-id="02d3c-1018">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-1018">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-1019">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-1020">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-1020">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-1021">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-1022">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-1022">Returns:</span></span>

<span data-ttu-id="02d3c-1023">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="02d3c-1023">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="02d3c-1024">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-1024">Example</span></span>

<span data-ttu-id="02d3c-1025">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="02d3c-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="02d3c-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="02d3c-p164">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-1029">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1029">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02d3c-p165">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="02d3c-1033">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="02d3c-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="02d3c-1034">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="02d3c-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02d3c-1038">Requirements</span><span class="sxs-lookup"><span data-stu-id="02d3c-1038">Requirements</span></span>

|<span data-ttu-id="02d3c-1039">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1039">Requirement</span></span>| <span data-ttu-id="02d3c-1040">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-1041">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="02d3c-1042">1.6</span></span> |
|[<span data-ttu-id="02d3c-1043">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-1044">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-1045">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-1046">阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02d3c-1047">返回：</span><span class="sxs-lookup"><span data-stu-id="02d3c-1047">Returns:</span></span>

<span data-ttu-id="02d3c-p167">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="02d3c-1050">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-1050">Example</span></span>

<span data-ttu-id="02d3c-1051">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="02d3c-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="02d3c-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="02d3c-1053">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="02d3c-p168">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-1057">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-1057">Parameters</span></span>

|<span data-ttu-id="02d3c-1058">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-1058">Name</span></span>| <span data-ttu-id="02d3c-1059">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-1059">Type</span></span>| <span data-ttu-id="02d3c-1060">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-1060">Attributes</span></span>| <span data-ttu-id="02d3c-1061">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="02d3c-1062">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-1062">function</span></span>||<span data-ttu-id="02d3c-1063">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="02d3c-1064">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="02d3c-1065">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1065">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="02d3c-1066">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-1066">Object</span></span>| <span data-ttu-id="02d3c-1067">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1068">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1068">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="02d3c-1069">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02d3c-1070">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1070">Requirements</span></span>

|<span data-ttu-id="02d3c-1071">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1071">Requirement</span></span>| <span data-ttu-id="02d3c-1072">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-1073">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="02d3c-1074">1.0</span></span>|
|[<span data-ttu-id="02d3c-1075">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-1075">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-1076">ReadItem</span></span>|
|[<span data-ttu-id="02d3c-1077">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-1077">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-1078">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02d3c-1078">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-1079">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-1079">Example</span></span>

<span data-ttu-id="02d3c-p171">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="02d3c-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="02d3c-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="02d3c-1084">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="02d3c-1085">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1085">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="02d3c-1086">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1086">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="02d3c-1087">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1087">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="02d3c-1088">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1088">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-1089">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-1089">Parameters</span></span>

|<span data-ttu-id="02d3c-1090">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-1090">Name</span></span>| <span data-ttu-id="02d3c-1091">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-1091">Type</span></span>| <span data-ttu-id="02d3c-1092">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-1092">Attributes</span></span>| <span data-ttu-id="02d3c-1093">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="02d3c-1094">String</span><span class="sxs-lookup"><span data-stu-id="02d3c-1094">String</span></span>||<span data-ttu-id="02d3c-1095">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1095">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="02d3c-1096">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-1096">Object</span></span>| <span data-ttu-id="02d3c-1097">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1098">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1098">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02d3c-1099">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-1099">Object</span></span>| <span data-ttu-id="02d3c-1100">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1100">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1101">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1101">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="02d3c-1102">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-1102">function</span></span>| <span data-ttu-id="02d3c-1103">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1103">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1104">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1104">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="02d3c-1105">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1105">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="02d3c-1106">错误</span><span class="sxs-lookup"><span data-stu-id="02d3c-1106">Errors</span></span>

| <span data-ttu-id="02d3c-1107">错误代码</span><span class="sxs-lookup"><span data-stu-id="02d3c-1107">Error code</span></span> | <span data-ttu-id="02d3c-1108">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-1108">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="02d3c-1109">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1109">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02d3c-1110">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1110">Requirements</span></span>

|<span data-ttu-id="02d3c-1111">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1111">Requirement</span></span>| <span data-ttu-id="02d3c-1112">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-1113">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-1114">1.1</span><span class="sxs-lookup"><span data-stu-id="02d3c-1114">1.1</span></span>|
|[<span data-ttu-id="02d3c-1115">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-1115">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-1116">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-1116">ReadWriteItem</span></span>|
|[<span data-ttu-id="02d3c-1117">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-1117">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-1118">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-1118">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-1119">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-1119">Example</span></span>

<span data-ttu-id="02d3c-1120">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1120">The following code removes an attachment with an identifier of '0'.</span></span>

```javascript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="02d3c-1121">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="02d3c-1121">saveAsync([options], callback)</span></span>

<span data-ttu-id="02d3c-1122">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1122">Asynchronously saves an item.</span></span>

<span data-ttu-id="02d3c-1123">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1123">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="02d3c-1124">在 Outlook 网页或 Outlook 的联机模式中, 将项目保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1124">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="02d3c-1125">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1125">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-1126">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1126">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="02d3c-1127">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1127">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="02d3c-p175">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="02d3c-1131">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="02d3c-1131">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="02d3c-1132">Mac 上的 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1132">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="02d3c-1133">在`saveAsync`撰写模式下从会议中调用时, 此方法将失败。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1133">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="02d3c-1134">若要解决此问题, 请参阅[使用 OFFICE JS API 将会议保存为 Outlook For Mac 中的草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1134">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="02d3c-1135">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-1136">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-1136">Parameters</span></span>

|<span data-ttu-id="02d3c-1137">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-1137">Name</span></span>| <span data-ttu-id="02d3c-1138">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-1138">Type</span></span>| <span data-ttu-id="02d3c-1139">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-1139">Attributes</span></span>| <span data-ttu-id="02d3c-1140">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="02d3c-1141">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-1141">Object</span></span>| <span data-ttu-id="02d3c-1142">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1143">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02d3c-1144">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-1144">Object</span></span>| <span data-ttu-id="02d3c-1145">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1146">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="02d3c-1147">函数</span><span class="sxs-lookup"><span data-stu-id="02d3c-1147">function</span></span>||<span data-ttu-id="02d3c-1148">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="02d3c-1149">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02d3c-1150">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1150">Requirements</span></span>

|<span data-ttu-id="02d3c-1151">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1151">Requirement</span></span>| <span data-ttu-id="02d3c-1152">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-1153">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="02d3c-1154">1.3</span></span>|
|[<span data-ttu-id="02d3c-1155">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-1155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="02d3c-1157">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-1157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-1158">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="02d3c-1159">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-1159">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="02d3c-p177">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="02d3c-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="02d3c-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="02d3c-1163">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="02d3c-p178">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02d3c-1167">参数</span><span class="sxs-lookup"><span data-stu-id="02d3c-1167">Parameters</span></span>

|<span data-ttu-id="02d3c-1168">名称</span><span class="sxs-lookup"><span data-stu-id="02d3c-1168">Name</span></span>| <span data-ttu-id="02d3c-1169">类型</span><span class="sxs-lookup"><span data-stu-id="02d3c-1169">Type</span></span>| <span data-ttu-id="02d3c-1170">属性</span><span class="sxs-lookup"><span data-stu-id="02d3c-1170">Attributes</span></span>| <span data-ttu-id="02d3c-1171">说明</span><span class="sxs-lookup"><span data-stu-id="02d3c-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="02d3c-1172">字符串</span><span class="sxs-lookup"><span data-stu-id="02d3c-1172">String</span></span>||<span data-ttu-id="02d3c-p179">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="02d3c-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="02d3c-1176">Object</span><span class="sxs-lookup"><span data-stu-id="02d3c-1176">Object</span></span>| <span data-ttu-id="02d3c-1177">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1178">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02d3c-1179">对象</span><span class="sxs-lookup"><span data-stu-id="02d3c-1179">Object</span></span>| <span data-ttu-id="02d3c-1180">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1181">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="02d3c-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="02d3c-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="02d3c-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02d3c-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="02d3c-1184">如果`text`为, 则当前样式应用于 web 上的 Outlook 和桌面客户端。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1184">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="02d3c-1185">如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1185">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="02d3c-1186">如果`html`和字段支持 HTML (主题不), 则当前样式应用于 web 上的 outlook, 并且在 outlook 桌面客户端中应用了默认样式。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="02d3c-1187">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1187">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="02d3c-1188">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="02d3c-1189">function</span><span class="sxs-lookup"><span data-stu-id="02d3c-1189">function</span></span>||<span data-ttu-id="02d3c-1190">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02d3c-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02d3c-1191">Requirements</span><span class="sxs-lookup"><span data-stu-id="02d3c-1191">Requirements</span></span>

|<span data-ttu-id="02d3c-1192">要求</span><span class="sxs-lookup"><span data-stu-id="02d3c-1192">Requirement</span></span>| <span data-ttu-id="02d3c-1193">值</span><span class="sxs-lookup"><span data-stu-id="02d3c-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="02d3c-1194">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02d3c-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02d3c-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="02d3c-1195">1.2</span></span>|
|[<span data-ttu-id="02d3c-1196">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02d3c-1196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02d3c-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02d3c-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="02d3c-1198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02d3c-1198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02d3c-1199">撰写</span><span class="sxs-lookup"><span data-stu-id="02d3c-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02d3c-1200">示例</span><span class="sxs-lookup"><span data-stu-id="02d3c-1200">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
