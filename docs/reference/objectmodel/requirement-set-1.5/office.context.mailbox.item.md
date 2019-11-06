---
title: Office.context.mailbox.item - 要求集 1.5
description: ''
ms.date: 11/05/2019
localization_priority: Priority
ms.openlocfilehash: 7cb755ecb7bcc836e93cf11e0caa5db55a6ddc29
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001577"
---
# <a name="item"></a><span data-ttu-id="daa6d-102">item</span><span class="sxs-lookup"><span data-stu-id="daa6d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="daa6d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="daa6d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="daa6d-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="daa6d-106">Requirements</span></span>

|<span data-ttu-id="daa6d-107">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-107">Requirement</span></span>| <span data-ttu-id="daa6d-108">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-110">1.0</span></span>|
|[<span data-ttu-id="daa6d-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-112">受限</span><span class="sxs-lookup"><span data-stu-id="daa6d-112">Restricted</span></span>|
|[<span data-ttu-id="daa6d-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="daa6d-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-115">Members and methods</span></span>

| <span data-ttu-id="daa6d-116">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-116">Member</span></span> | <span data-ttu-id="daa6d-117">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="daa6d-118">attachments</span><span class="sxs-lookup"><span data-stu-id="daa6d-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="daa6d-119">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-119">Member</span></span> |
| [<span data-ttu-id="daa6d-120">bcc</span><span class="sxs-lookup"><span data-stu-id="daa6d-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="daa6d-121">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-121">Member</span></span> |
| [<span data-ttu-id="daa6d-122">body</span><span class="sxs-lookup"><span data-stu-id="daa6d-122">body</span></span>](#body-body) | <span data-ttu-id="daa6d-123">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-123">Member</span></span> |
| [<span data-ttu-id="daa6d-124">cc</span><span class="sxs-lookup"><span data-stu-id="daa6d-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="daa6d-125">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-125">Member</span></span> |
| [<span data-ttu-id="daa6d-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="daa6d-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="daa6d-127">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-127">Member</span></span> |
| [<span data-ttu-id="daa6d-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="daa6d-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="daa6d-129">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-129">Member</span></span> |
| [<span data-ttu-id="daa6d-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="daa6d-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="daa6d-131">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-131">Member</span></span> |
| [<span data-ttu-id="daa6d-132">end</span><span class="sxs-lookup"><span data-stu-id="daa6d-132">end</span></span>](#end-datetime) | <span data-ttu-id="daa6d-133">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-133">Member</span></span> |
| [<span data-ttu-id="daa6d-134">from</span><span class="sxs-lookup"><span data-stu-id="daa6d-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="daa6d-135">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-135">Member</span></span> |
| [<span data-ttu-id="daa6d-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="daa6d-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="daa6d-137">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-137">Member</span></span> |
| [<span data-ttu-id="daa6d-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="daa6d-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="daa6d-139">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-139">Member</span></span> |
| [<span data-ttu-id="daa6d-140">itemId</span><span class="sxs-lookup"><span data-stu-id="daa6d-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="daa6d-141">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-141">Member</span></span> |
| [<span data-ttu-id="daa6d-142">itemType</span><span class="sxs-lookup"><span data-stu-id="daa6d-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="daa6d-143">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-143">Member</span></span> |
| [<span data-ttu-id="daa6d-144">location</span><span class="sxs-lookup"><span data-stu-id="daa6d-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="daa6d-145">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-145">Member</span></span> |
| [<span data-ttu-id="daa6d-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="daa6d-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="daa6d-147">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-147">Member</span></span> |
| [<span data-ttu-id="daa6d-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="daa6d-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="daa6d-149">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-149">Member</span></span> |
| [<span data-ttu-id="daa6d-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="daa6d-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="daa6d-151">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-151">Member</span></span> |
| [<span data-ttu-id="daa6d-152">organizer</span><span class="sxs-lookup"><span data-stu-id="daa6d-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="daa6d-153">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-153">Member</span></span> |
| [<span data-ttu-id="daa6d-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="daa6d-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="daa6d-155">Member</span><span class="sxs-lookup"><span data-stu-id="daa6d-155">Member</span></span> |
| [<span data-ttu-id="daa6d-156">sender</span><span class="sxs-lookup"><span data-stu-id="daa6d-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="daa6d-157">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-157">Member</span></span> |
| [<span data-ttu-id="daa6d-158">start</span><span class="sxs-lookup"><span data-stu-id="daa6d-158">start</span></span>](#start-datetime) | <span data-ttu-id="daa6d-159">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-159">Member</span></span> |
| [<span data-ttu-id="daa6d-160">subject</span><span class="sxs-lookup"><span data-stu-id="daa6d-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="daa6d-161">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-161">Member</span></span> |
| [<span data-ttu-id="daa6d-162">to</span><span class="sxs-lookup"><span data-stu-id="daa6d-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="daa6d-163">成员</span><span class="sxs-lookup"><span data-stu-id="daa6d-163">Member</span></span> |
| [<span data-ttu-id="daa6d-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="daa6d-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="daa6d-165">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-165">Method</span></span> |
| [<span data-ttu-id="daa6d-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="daa6d-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="daa6d-167">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-167">Method</span></span> |
| [<span data-ttu-id="daa6d-168">close</span><span class="sxs-lookup"><span data-stu-id="daa6d-168">close</span></span>](#close) | <span data-ttu-id="daa6d-169">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-169">Method</span></span> |
| [<span data-ttu-id="daa6d-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="daa6d-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="daa6d-171">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-171">Method</span></span> |
| [<span data-ttu-id="daa6d-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="daa6d-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="daa6d-173">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-173">Method</span></span> |
| [<span data-ttu-id="daa6d-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="daa6d-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="daa6d-175">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-175">Method</span></span> |
| [<span data-ttu-id="daa6d-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="daa6d-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="daa6d-177">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-177">Method</span></span> |
| [<span data-ttu-id="daa6d-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="daa6d-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="daa6d-179">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-179">Method</span></span> |
| [<span data-ttu-id="daa6d-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="daa6d-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="daa6d-181">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-181">Method</span></span> |
| [<span data-ttu-id="daa6d-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="daa6d-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="daa6d-183">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-183">Method</span></span> |
| [<span data-ttu-id="daa6d-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="daa6d-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="daa6d-185">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-185">Method</span></span> |
| [<span data-ttu-id="daa6d-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="daa6d-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="daa6d-187">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-187">Method</span></span> |
| [<span data-ttu-id="daa6d-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="daa6d-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="daa6d-189">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-189">Method</span></span> |
| [<span data-ttu-id="daa6d-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="daa6d-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="daa6d-191">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-191">Method</span></span> |
| [<span data-ttu-id="daa6d-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="daa6d-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="daa6d-193">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="daa6d-194">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-194">Example</span></span>

<span data-ttu-id="daa6d-195">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="daa6d-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="daa6d-196">Members</span><span class="sxs-lookup"><span data-stu-id="daa6d-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-15"></a><span data-ttu-id="daa6d-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="daa6d-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

<span data-ttu-id="daa6d-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-200">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="daa6d-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="daa6d-201">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="daa6d-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-202">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-202">Type</span></span>

*   <span data-ttu-id="daa6d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span><span class="sxs-lookup"><span data-stu-id="daa6d-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.5)></span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-204">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-204">Requirements</span></span>

|<span data-ttu-id="daa6d-205">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-205">Requirement</span></span>| <span data-ttu-id="daa6d-206">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-207">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-208">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-208">1.0</span></span>|
|[<span data-ttu-id="daa6d-209">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-210">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-212">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-213">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-213">Example</span></span>

<span data-ttu-id="daa6d-214">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="daa6d-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="daa6d-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-216">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="daa6d-217">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-217">Compose mode only.</span></span>

<span data-ttu-id="daa6d-218">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-218">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-219">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="daa6d-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="daa6d-220">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="daa6d-221">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-222">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-222">Type</span></span>

*   [<span data-ttu-id="daa6d-223">收件人</span><span class="sxs-lookup"><span data-stu-id="daa6d-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="daa6d-224">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-224">Requirements</span></span>

|<span data-ttu-id="daa6d-225">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-225">Requirement</span></span>| <span data-ttu-id="daa6d-226">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-228">1.1</span><span class="sxs-lookup"><span data-stu-id="daa6d-228">1.1</span></span>|
|[<span data-ttu-id="daa6d-229">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-230">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-231">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-232">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-233">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-15"></a><span data-ttu-id="daa6d-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-235">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-236">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-236">Type</span></span>

*   [<span data-ttu-id="daa6d-237">Body</span><span class="sxs-lookup"><span data-stu-id="daa6d-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="daa6d-238">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-238">Requirements</span></span>

|<span data-ttu-id="daa6d-239">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-239">Requirement</span></span>| <span data-ttu-id="daa6d-240">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-242">1.1</span><span class="sxs-lookup"><span data-stu-id="daa6d-242">1.1</span></span>|
|[<span data-ttu-id="daa6d-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-244">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-247">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-247">Example</span></span>

<span data-ttu-id="daa6d-248">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="daa6d-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="daa6d-249">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="daa6d-249">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="daa6d-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-251">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="daa6d-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="daa6d-252">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-253">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-253">Read mode</span></span>

<span data-ttu-id="daa6d-254">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="daa6d-255">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-255">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-256">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-257">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-257">Compose mode</span></span>

<span data-ttu-id="daa6d-258">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="daa6d-259">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-260">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="daa6d-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="daa6d-261">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="daa6d-262">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="daa6d-263">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-263">Type</span></span>

*   <span data-ttu-id="daa6d-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-265">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-265">Requirements</span></span>

|<span data-ttu-id="daa6d-266">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-266">Requirement</span></span>| <span data-ttu-id="daa6d-267">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-269">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-269">1.0</span></span>|
|[<span data-ttu-id="daa6d-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-271">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-273">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="daa6d-274">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="daa6d-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="daa6d-275">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="daa6d-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="daa6d-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-280">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-280">Type</span></span>

*   <span data-ttu-id="daa6d-281">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-282">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-282">Requirements</span></span>

|<span data-ttu-id="daa6d-283">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-283">Requirement</span></span>| <span data-ttu-id="daa6d-284">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-285">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-286">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-286">1.0</span></span>|
|[<span data-ttu-id="daa6d-287">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-288">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-289">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-290">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-291">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="daa6d-292">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="daa6d-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="daa6d-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-295">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-295">Type</span></span>

*   <span data-ttu-id="daa6d-296">日期</span><span class="sxs-lookup"><span data-stu-id="daa6d-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-297">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-297">Requirements</span></span>

|<span data-ttu-id="daa6d-298">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-298">Requirement</span></span>| <span data-ttu-id="daa6d-299">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-300">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-301">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-301">1.0</span></span>|
|[<span data-ttu-id="daa6d-302">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-303">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-304">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-305">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-306">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="daa6d-307">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="daa6d-307">dateTimeModified: Date</span></span>

<span data-ttu-id="daa6d-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-310">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-311">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-311">Type</span></span>

*   <span data-ttu-id="daa6d-312">日期</span><span class="sxs-lookup"><span data-stu-id="daa6d-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-313">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-313">Requirements</span></span>

|<span data-ttu-id="daa6d-314">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-314">Requirement</span></span>| <span data-ttu-id="daa6d-315">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-316">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-317">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-317">1.0</span></span>|
|[<span data-ttu-id="daa6d-318">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-319">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-320">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-321">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-322">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="daa6d-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-324">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="daa6d-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="daa6d-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-327">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-327">Read mode</span></span>

<span data-ttu-id="daa6d-328">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-329">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-329">Compose mode</span></span>

<span data-ttu-id="daa6d-330">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="daa6d-331">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="daa6d-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="daa6d-332">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="daa6d-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="daa6d-333">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-333">Type</span></span>

*   <span data-ttu-id="daa6d-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-335">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-335">Requirements</span></span>

|<span data-ttu-id="daa6d-336">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-336">Requirement</span></span>| <span data-ttu-id="daa6d-337">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-338">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-339">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-339">1.0</span></span>|
|[<span data-ttu-id="daa6d-340">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-341">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-342">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-343">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="daa6d-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-p114">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="daa6d-p115">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-349">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-350">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-350">Type</span></span>

*   [<span data-ttu-id="daa6d-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="daa6d-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="daa6d-352">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-352">Requirements</span></span>

|<span data-ttu-id="daa6d-353">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-353">Requirement</span></span>| <span data-ttu-id="daa6d-354">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-355">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-356">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-356">1.0</span></span>|
|[<span data-ttu-id="daa6d-357">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-358">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-359">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-360">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-361">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="daa6d-362">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="daa6d-362">internetMessageId: String</span></span>

<span data-ttu-id="daa6d-p116">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-365">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-365">Type</span></span>

*   <span data-ttu-id="daa6d-366">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-367">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-367">Requirements</span></span>

|<span data-ttu-id="daa6d-368">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-368">Requirement</span></span>| <span data-ttu-id="daa6d-369">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-370">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-371">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-371">1.0</span></span>|
|[<span data-ttu-id="daa6d-372">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-373">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-374">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-375">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-376">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="daa6d-377">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="daa6d-377">itemClass: String</span></span>

<span data-ttu-id="daa6d-p117">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="daa6d-p118">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="daa6d-382">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-382">Type</span></span> | <span data-ttu-id="daa6d-383">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-383">Description</span></span> | <span data-ttu-id="daa6d-384">项目类</span><span class="sxs-lookup"><span data-stu-id="daa6d-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="daa6d-385">约会项目</span><span class="sxs-lookup"><span data-stu-id="daa6d-385">Appointment items</span></span> | <span data-ttu-id="daa6d-386">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="daa6d-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="daa6d-387">邮件项目</span><span class="sxs-lookup"><span data-stu-id="daa6d-387">Message items</span></span> | <span data-ttu-id="daa6d-388">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="daa6d-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="daa6d-389">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-390">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-390">Type</span></span>

*   <span data-ttu-id="daa6d-391">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-392">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-392">Requirements</span></span>

|<span data-ttu-id="daa6d-393">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-393">Requirement</span></span>| <span data-ttu-id="daa6d-394">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-395">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-396">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-396">1.0</span></span>|
|[<span data-ttu-id="daa6d-397">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-398">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-399">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-400">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-401">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="daa6d-402">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="daa6d-402">(nullable) itemId: String</span></span>

<span data-ttu-id="daa6d-p119">获取当前项目的 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-405">`itemId` 属性返回的标识符与 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)相同。</span><span class="sxs-lookup"><span data-stu-id="daa6d-405">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="daa6d-406">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="daa6d-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="daa6d-407">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="daa6d-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="daa6d-408">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="daa6d-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="daa6d-p121">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-411">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-411">Type</span></span>

*   <span data-ttu-id="daa6d-412">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-413">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-413">Requirements</span></span>

|<span data-ttu-id="daa6d-414">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-414">Requirement</span></span>| <span data-ttu-id="daa6d-415">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-416">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-417">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-417">1.0</span></span>|
|[<span data-ttu-id="daa6d-418">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-419">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-420">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-421">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-422">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-422">Example</span></span>

<span data-ttu-id="daa6d-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-15"></a><span data-ttu-id="daa6d-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-426">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="daa6d-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="daa6d-427">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="daa6d-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-428">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-428">Type</span></span>

*   [<span data-ttu-id="daa6d-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="daa6d-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="daa6d-430">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-430">Requirements</span></span>

|<span data-ttu-id="daa6d-431">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-431">Requirement</span></span>| <span data-ttu-id="daa6d-432">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-433">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-434">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-434">1.0</span></span>|
|[<span data-ttu-id="daa6d-435">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-436">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-437">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-438">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-439">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-439">Example</span></span>

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-15"></a><span data-ttu-id="daa6d-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-441">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="daa6d-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-442">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-442">Read mode</span></span>

<span data-ttu-id="daa6d-443">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="daa6d-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-444">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-444">Compose mode</span></span>

<span data-ttu-id="daa6d-445">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="daa6d-446">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-446">Type</span></span>

*   <span data-ttu-id="daa6d-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-448">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-448">Requirements</span></span>

|<span data-ttu-id="daa6d-449">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-449">Requirement</span></span>| <span data-ttu-id="daa6d-450">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-451">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-452">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-452">1.0</span></span>|
|[<span data-ttu-id="daa6d-453">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-454">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-455">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-456">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="daa6d-457">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="daa6d-457">normalizedSubject: String</span></span>

<span data-ttu-id="daa6d-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="daa6d-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-462">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-462">Type</span></span>

*   <span data-ttu-id="daa6d-463">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-464">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-464">Requirements</span></span>

|<span data-ttu-id="daa6d-465">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-465">Requirement</span></span>| <span data-ttu-id="daa6d-466">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-468">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-468">1.0</span></span>|
|[<span data-ttu-id="daa6d-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-470">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-472">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-473">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-15"></a><span data-ttu-id="daa6d-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-475">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-476">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-476">Type</span></span>

*   [<span data-ttu-id="daa6d-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="daa6d-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="daa6d-478">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-478">Requirements</span></span>

|<span data-ttu-id="daa6d-479">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-479">Requirement</span></span>| <span data-ttu-id="daa6d-480">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-481">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-482">1.3</span><span class="sxs-lookup"><span data-stu-id="daa6d-482">1.3</span></span>|
|[<span data-ttu-id="daa6d-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-484">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-486">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-487">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-487">Example</span></span>

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="daa6d-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-489">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="daa6d-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="daa6d-490">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-491">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-491">Read mode</span></span>

<span data-ttu-id="daa6d-492">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="daa6d-493">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-493">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-494">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-495">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-495">Compose mode</span></span>

<span data-ttu-id="daa6d-496">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="daa6d-497">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-498">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="daa6d-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="daa6d-499">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="daa6d-500">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="daa6d-501">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-501">Type</span></span>

*   <span data-ttu-id="daa6d-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-503">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-503">Requirements</span></span>

|<span data-ttu-id="daa6d-504">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-504">Requirement</span></span>| <span data-ttu-id="daa6d-505">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-506">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-507">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-507">1.0</span></span>|
|[<span data-ttu-id="daa6d-508">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-509">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-510">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-511">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="daa6d-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-p128">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-515">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-515">Type</span></span>

*   [<span data-ttu-id="daa6d-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="daa6d-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="daa6d-517">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-517">Requirements</span></span>

|<span data-ttu-id="daa6d-518">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-518">Requirement</span></span>| <span data-ttu-id="daa6d-519">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-520">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-521">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-521">1.0</span></span>|
|[<span data-ttu-id="daa6d-522">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-523">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-524">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-525">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-526">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="daa6d-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-528">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="daa6d-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="daa6d-529">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-530">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-530">Read mode</span></span>

<span data-ttu-id="daa6d-531">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="daa6d-532">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-532">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-533">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-534">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-534">Compose mode</span></span>

<span data-ttu-id="daa6d-535">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="daa6d-536">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-537">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="daa6d-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="daa6d-538">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="daa6d-539">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="daa6d-540">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-540">Type</span></span>

*   <span data-ttu-id="daa6d-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-542">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-542">Requirements</span></span>

|<span data-ttu-id="daa6d-543">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-543">Requirement</span></span>| <span data-ttu-id="daa6d-544">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-545">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-546">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-546">1.0</span></span>|
|[<span data-ttu-id="daa6d-547">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-548">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-549">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-550">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15"></a><span data-ttu-id="daa6d-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-p132">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="daa6d-p133">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-556">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="daa6d-557">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-557">Type</span></span>

*   [<span data-ttu-id="daa6d-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="daa6d-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="daa6d-559">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-559">Requirements</span></span>

|<span data-ttu-id="daa6d-560">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-560">Requirement</span></span>| <span data-ttu-id="daa6d-561">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-562">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-563">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-563">1.0</span></span>|
|[<span data-ttu-id="daa6d-564">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-565">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-566">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-567">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-568">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-15"></a><span data-ttu-id="daa6d-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-570">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="daa6d-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="daa6d-p134">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-573">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-573">Read mode</span></span>

<span data-ttu-id="daa6d-574">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-575">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-575">Compose mode</span></span>

<span data-ttu-id="daa6d-576">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="daa6d-577">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="daa6d-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="daa6d-578">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="daa6d-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.5#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

##### <a name="type"></a><span data-ttu-id="daa6d-579">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-579">Type</span></span>

*   <span data-ttu-id="daa6d-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-581">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-581">Requirements</span></span>

|<span data-ttu-id="daa6d-582">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-582">Requirement</span></span>| <span data-ttu-id="daa6d-583">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-584">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-585">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-585">1.0</span></span>|
|[<span data-ttu-id="daa6d-586">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-587">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-588">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-589">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-15"></a><span data-ttu-id="daa6d-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-591">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="daa6d-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="daa6d-592">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="daa6d-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-593">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-593">Read mode</span></span>

<span data-ttu-id="daa6d-p135">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-596">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-596">Compose mode</span></span>

<span data-ttu-id="daa6d-597">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="daa6d-598">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-598">Type</span></span>

*   <span data-ttu-id="daa6d-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-600">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-600">Requirements</span></span>

|<span data-ttu-id="daa6d-601">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-601">Requirement</span></span>| <span data-ttu-id="daa6d-602">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-603">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-604">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-604">1.0</span></span>|
|[<span data-ttu-id="daa6d-605">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-606">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-607">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-608">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-15recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-15"></a><span data-ttu-id="daa6d-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

<span data-ttu-id="daa6d-610">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="daa6d-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="daa6d-611">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="daa6d-612">阅读模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-612">Read mode</span></span>

<span data-ttu-id="daa6d-613">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="daa6d-614">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-614">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-615">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="daa6d-616">撰写模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-616">Compose mode</span></span>

<span data-ttu-id="daa6d-617">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="daa6d-618">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="daa6d-619">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="daa6d-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="daa6d-620">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="daa6d-621">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="daa6d-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="daa6d-622">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-622">Type</span></span>

*   <span data-ttu-id="daa6d-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.5)</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-624">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-624">Requirements</span></span>

|<span data-ttu-id="daa6d-625">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-625">Requirement</span></span>| <span data-ttu-id="daa6d-626">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-627">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-628">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-628">1.0</span></span>|
|[<span data-ttu-id="daa6d-629">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-630">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-631">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-632">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="daa6d-633">方法</span><span class="sxs-lookup"><span data-stu-id="daa6d-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="daa6d-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="daa6d-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="daa6d-635">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="daa6d-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="daa6d-636">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="daa6d-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="daa6d-637">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-638">参数</span><span class="sxs-lookup"><span data-stu-id="daa6d-638">Parameters</span></span>

|<span data-ttu-id="daa6d-639">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-639">Name</span></span>| <span data-ttu-id="daa6d-640">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-640">Type</span></span>| <span data-ttu-id="daa6d-641">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-641">Attributes</span></span>| <span data-ttu-id="daa6d-642">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="daa6d-643">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-643">String</span></span>||<span data-ttu-id="daa6d-p139">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="daa6d-646">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-646">String</span></span>||<span data-ttu-id="daa6d-p140">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="daa6d-649">Object</span><span class="sxs-lookup"><span data-stu-id="daa6d-649">Object</span></span>| <span data-ttu-id="daa6d-650">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-650">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-651">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="daa6d-651">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="daa6d-652">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-652">Object</span></span> | <span data-ttu-id="daa6d-653">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-653">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-654">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-654">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="daa6d-655">布尔值</span><span class="sxs-lookup"><span data-stu-id="daa6d-655">Boolean</span></span> | <span data-ttu-id="daa6d-656">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-656">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-657">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="daa6d-657">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="daa6d-658">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-658">function</span></span>| <span data-ttu-id="daa6d-659">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-659">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-660">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-660">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="daa6d-661">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="daa6d-661">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="daa6d-662">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-662">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="daa6d-663">错误</span><span class="sxs-lookup"><span data-stu-id="daa6d-663">Errors</span></span>

| <span data-ttu-id="daa6d-664">错误代码</span><span class="sxs-lookup"><span data-stu-id="daa6d-664">Error code</span></span> | <span data-ttu-id="daa6d-665">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-665">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="daa6d-666">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="daa6d-666">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="daa6d-667">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="daa6d-667">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="daa6d-668">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="daa6d-668">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="daa6d-669">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-669">Requirements</span></span>

|<span data-ttu-id="daa6d-670">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-670">Requirement</span></span>| <span data-ttu-id="daa6d-671">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-671">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-672">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-672">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-673">1.1</span><span class="sxs-lookup"><span data-stu-id="daa6d-673">1.1</span></span>|
|[<span data-ttu-id="daa6d-674">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-674">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-675">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-675">ReadWriteItem</span></span>|
|[<span data-ttu-id="daa6d-676">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-676">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-677">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-677">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="daa6d-678">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-678">Examples</span></span>

```js
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

<span data-ttu-id="daa6d-679">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-679">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="daa6d-680">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="daa6d-680">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="daa6d-681">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="daa6d-681">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="daa6d-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="daa6d-685">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-685">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="daa6d-686">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="daa6d-686">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-687">Parameters</span><span class="sxs-lookup"><span data-stu-id="daa6d-687">Parameters</span></span>

|<span data-ttu-id="daa6d-688">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-688">Name</span></span>| <span data-ttu-id="daa6d-689">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-689">Type</span></span>| <span data-ttu-id="daa6d-690">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-690">Attributes</span></span>| <span data-ttu-id="daa6d-691">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-691">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="daa6d-692">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-692">String</span></span>||<span data-ttu-id="daa6d-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="daa6d-695">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-695">String</span></span>||<span data-ttu-id="daa6d-696">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="daa6d-696">The subject of the item to be attached.</span></span> <span data-ttu-id="daa6d-697">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-697">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="daa6d-698">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-698">Object</span></span>| <span data-ttu-id="daa6d-699">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-699">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-700">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="daa6d-700">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="daa6d-701">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-701">Object</span></span>| <span data-ttu-id="daa6d-702">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-702">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-703">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-703">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="daa6d-704">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-704">function</span></span>| <span data-ttu-id="daa6d-705">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-705">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-706">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-706">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="daa6d-707">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="daa6d-707">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="daa6d-708">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-708">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="daa6d-709">错误</span><span class="sxs-lookup"><span data-stu-id="daa6d-709">Errors</span></span>

| <span data-ttu-id="daa6d-710">错误代码</span><span class="sxs-lookup"><span data-stu-id="daa6d-710">Error code</span></span> | <span data-ttu-id="daa6d-711">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-711">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="daa6d-712">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="daa6d-712">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="daa6d-713">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-713">Requirements</span></span>

|<span data-ttu-id="daa6d-714">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-714">Requirement</span></span>| <span data-ttu-id="daa6d-715">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-716">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-717">1.1</span><span class="sxs-lookup"><span data-stu-id="daa6d-717">1.1</span></span>|
|[<span data-ttu-id="daa6d-718">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-719">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-719">ReadWriteItem</span></span>|
|[<span data-ttu-id="daa6d-720">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-721">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-721">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-722">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-722">Example</span></span>

<span data-ttu-id="daa6d-723">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-723">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

<br>

---
---

#### <a name="close"></a><span data-ttu-id="daa6d-724">close()</span><span class="sxs-lookup"><span data-stu-id="daa6d-724">close()</span></span>

<span data-ttu-id="daa6d-725">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="daa6d-725">Closes the current item that is being composed.</span></span>

<span data-ttu-id="daa6d-p144">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-728">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="daa6d-728">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="daa6d-729">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="daa6d-729">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-730">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-730">Requirements</span></span>

|<span data-ttu-id="daa6d-731">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-731">Requirement</span></span>| <span data-ttu-id="daa6d-732">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-732">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-733">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-733">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-734">1.3</span><span class="sxs-lookup"><span data-stu-id="daa6d-734">1.3</span></span>|
|[<span data-ttu-id="daa6d-735">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-735">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-736">受限</span><span class="sxs-lookup"><span data-stu-id="daa6d-736">Restricted</span></span>|
|[<span data-ttu-id="daa6d-737">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-737">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-738">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-738">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="daa6d-739">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="daa6d-739">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="daa6d-740">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="daa6d-740">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-741">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-741">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="daa6d-742">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="daa6d-742">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="daa6d-743">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="daa6d-743">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="daa6d-p145">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-747">Parameters</span><span class="sxs-lookup"><span data-stu-id="daa6d-747">Parameters</span></span>

| <span data-ttu-id="daa6d-748">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-748">Name</span></span> | <span data-ttu-id="daa6d-749">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-749">Type</span></span> | <span data-ttu-id="daa6d-750">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-750">Attributes</span></span> | <span data-ttu-id="daa6d-751">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-751">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="daa6d-752">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-752">String &#124; Object</span></span>| |<span data-ttu-id="daa6d-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="daa6d-755">**或**</span><span class="sxs-lookup"><span data-stu-id="daa6d-755">**OR**</span></span><br/><span data-ttu-id="daa6d-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="daa6d-758">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-758">String</span></span> | <span data-ttu-id="daa6d-759">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-759">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="daa6d-762">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-762">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="daa6d-763">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-763">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-764">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-764">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="daa6d-765">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-765">String</span></span> | | <span data-ttu-id="daa6d-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="daa6d-768">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-768">String</span></span> | | <span data-ttu-id="daa6d-769">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-769">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="daa6d-770">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-770">String</span></span> | | <span data-ttu-id="daa6d-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="daa6d-773">布尔</span><span class="sxs-lookup"><span data-stu-id="daa6d-773">Boolean</span></span> | | <span data-ttu-id="daa6d-p151">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="daa6d-776">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-776">String</span></span> | | <span data-ttu-id="daa6d-p152">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="daa6d-780">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-780">function</span></span> | <span data-ttu-id="daa6d-781">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-781">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-782">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-782">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="daa6d-783">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-783">Requirements</span></span>

|<span data-ttu-id="daa6d-784">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-784">Requirement</span></span>| <span data-ttu-id="daa6d-785">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-786">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-787">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-787">1.0</span></span>|
|[<span data-ttu-id="daa6d-788">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-788">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-789">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-790">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-790">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-791">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-791">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="daa6d-792">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-792">Examples</span></span>

<span data-ttu-id="daa6d-793">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-793">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="daa6d-794">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-794">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="daa6d-795">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-795">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="daa6d-796">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-796">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="daa6d-797">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-797">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="daa6d-798">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-798">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="daa6d-799">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="daa6d-799">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="daa6d-800">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="daa6d-800">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-801">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-801">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="daa6d-802">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="daa6d-802">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="daa6d-803">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="daa6d-803">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="daa6d-p153">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-807">Parameters</span><span class="sxs-lookup"><span data-stu-id="daa6d-807">Parameters</span></span>

| <span data-ttu-id="daa6d-808">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-808">Name</span></span> | <span data-ttu-id="daa6d-809">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-809">Type</span></span> | <span data-ttu-id="daa6d-810">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-810">Attributes</span></span> | <span data-ttu-id="daa6d-811">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-811">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="daa6d-812">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-812">String &#124; Object</span></span>| | <span data-ttu-id="daa6d-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="daa6d-815">**或**</span><span class="sxs-lookup"><span data-stu-id="daa6d-815">**OR**</span></span><br/><span data-ttu-id="daa6d-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="daa6d-818">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-818">String</span></span> | <span data-ttu-id="daa6d-819">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-819">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="daa6d-822">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-822">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="daa6d-823">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-823">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-824">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-824">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="daa6d-825">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-825">String</span></span> | | <span data-ttu-id="daa6d-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="daa6d-828">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-828">String</span></span> | | <span data-ttu-id="daa6d-829">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-829">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="daa6d-830">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-830">String</span></span> | | <span data-ttu-id="daa6d-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="daa6d-833">布尔</span><span class="sxs-lookup"><span data-stu-id="daa6d-833">Boolean</span></span> | | <span data-ttu-id="daa6d-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="daa6d-836">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-836">String</span></span> | | <span data-ttu-id="daa6d-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="daa6d-840">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-840">function</span></span> | <span data-ttu-id="daa6d-841">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-841">&lt;optional&gt;</span></span> | <span data-ttu-id="daa6d-842">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="daa6d-843">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-843">Requirements</span></span>

|<span data-ttu-id="daa6d-844">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-844">Requirement</span></span>| <span data-ttu-id="daa6d-845">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-845">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-846">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-846">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-847">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-847">1.0</span></span>|
|[<span data-ttu-id="daa6d-848">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-848">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-849">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-849">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-850">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-850">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-851">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-851">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="daa6d-852">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-852">Examples</span></span>

<span data-ttu-id="daa6d-853">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-853">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="daa6d-854">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-854">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="daa6d-855">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-855">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="daa6d-856">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-856">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="daa6d-857">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-857">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="daa6d-858">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="daa6d-858">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-15"></a><span data-ttu-id="daa6d-859">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="daa6d-859">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="daa6d-860">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="daa6d-860">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-861">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-861">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-862">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-862">Requirements</span></span>

|<span data-ttu-id="daa6d-863">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-863">Requirement</span></span>| <span data-ttu-id="daa6d-864">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-865">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-866">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-866">1.0</span></span>|
|[<span data-ttu-id="daa6d-867">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-867">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-868">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-868">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-869">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-869">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-870">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="daa6d-871">返回：</span><span class="sxs-lookup"><span data-stu-id="daa6d-871">Returns:</span></span>

<span data-ttu-id="daa6d-872">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="daa6d-872">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.5)</span></span>

##### <a name="example"></a><span data-ttu-id="daa6d-873">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-873">Example</span></span>

<span data-ttu-id="daa6d-874">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="daa6d-874">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="daa6d-875">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="daa6d-875">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="daa6d-876">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-876">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-877">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-877">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-878">Parameters</span><span class="sxs-lookup"><span data-stu-id="daa6d-878">Parameters</span></span>

|<span data-ttu-id="daa6d-879">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-879">Name</span></span>| <span data-ttu-id="daa6d-880">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-880">Type</span></span>| <span data-ttu-id="daa6d-881">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-881">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="daa6d-882">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="daa6d-882">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.5)|<span data-ttu-id="daa6d-883">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="daa6d-883">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="daa6d-884">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-884">Requirements</span></span>

|<span data-ttu-id="daa6d-885">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-885">Requirement</span></span>| <span data-ttu-id="daa6d-886">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-887">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-888">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-888">1.0</span></span>|
|[<span data-ttu-id="daa6d-889">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-890">受限</span><span class="sxs-lookup"><span data-stu-id="daa6d-890">Restricted</span></span>|
|[<span data-ttu-id="daa6d-891">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-892">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-892">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="daa6d-893">返回：</span><span class="sxs-lookup"><span data-stu-id="daa6d-893">Returns:</span></span>

<span data-ttu-id="daa6d-894">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="daa6d-894">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="daa6d-895">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-895">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="daa6d-896">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="daa6d-896">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="daa6d-897">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="daa6d-897">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="daa6d-898">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="daa6d-898">Value of `entityType`</span></span> | <span data-ttu-id="daa6d-899">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-899">Type of objects in returned array</span></span> | <span data-ttu-id="daa6d-900">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-900">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="daa6d-901">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-901">String</span></span> | <span data-ttu-id="daa6d-902">**受限**</span><span class="sxs-lookup"><span data-stu-id="daa6d-902">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="daa6d-903">Contact</span><span class="sxs-lookup"><span data-stu-id="daa6d-903">Contact</span></span> | <span data-ttu-id="daa6d-904">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="daa6d-904">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="daa6d-905">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-905">String</span></span> | <span data-ttu-id="daa6d-906">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="daa6d-906">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="daa6d-907">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="daa6d-907">MeetingSuggestion</span></span> | <span data-ttu-id="daa6d-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="daa6d-908">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="daa6d-909">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="daa6d-909">PhoneNumber</span></span> | <span data-ttu-id="daa6d-910">**受限**</span><span class="sxs-lookup"><span data-stu-id="daa6d-910">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="daa6d-911">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="daa6d-911">TaskSuggestion</span></span> | <span data-ttu-id="daa6d-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="daa6d-912">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="daa6d-913">String</span><span class="sxs-lookup"><span data-stu-id="daa6d-913">String</span></span> | <span data-ttu-id="daa6d-914">**受限**</span><span class="sxs-lookup"><span data-stu-id="daa6d-914">**Restricted**</span></span> |

<span data-ttu-id="daa6d-915">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="daa6d-915">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

##### <a name="example"></a><span data-ttu-id="daa6d-916">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-916">Example</span></span>

<span data-ttu-id="daa6d-917">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-917">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-15meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-15phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-15tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-15"></a><span data-ttu-id="daa6d-918">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span><span class="sxs-lookup"><span data-stu-id="daa6d-918">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))>}</span></span>

<span data-ttu-id="daa6d-919">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="daa6d-919">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-920">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="daa6d-921">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="daa6d-921">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-922">参数</span><span class="sxs-lookup"><span data-stu-id="daa6d-922">Parameters</span></span>

|<span data-ttu-id="daa6d-923">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-923">Name</span></span>| <span data-ttu-id="daa6d-924">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-924">Type</span></span>| <span data-ttu-id="daa6d-925">描述</span><span class="sxs-lookup"><span data-stu-id="daa6d-925">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="daa6d-926">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-926">String</span></span>|<span data-ttu-id="daa6d-927">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="daa6d-927">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="daa6d-928">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-928">Requirements</span></span>

|<span data-ttu-id="daa6d-929">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-929">Requirement</span></span>| <span data-ttu-id="daa6d-930">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-930">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-931">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-931">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-932">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-932">1.0</span></span>|
|[<span data-ttu-id="daa6d-933">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-933">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-934">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-934">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-935">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-935">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-936">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-936">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="daa6d-937">返回：</span><span class="sxs-lookup"><span data-stu-id="daa6d-937">Returns:</span></span>

<span data-ttu-id="daa6d-p162">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="daa6d-940">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span><span class="sxs-lookup"><span data-stu-id="daa6d-940">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.5)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.5)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.5)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.5))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="daa6d-941">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="daa6d-941">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="daa6d-942">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="daa6d-942">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-943">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-943">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="daa6d-p163">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="daa6d-947">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="daa6d-947">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="daa6d-948">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-948">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="daa6d-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.5#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="daa6d-952">Requirements</span><span class="sxs-lookup"><span data-stu-id="daa6d-952">Requirements</span></span>

|<span data-ttu-id="daa6d-953">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-953">Requirement</span></span>| <span data-ttu-id="daa6d-954">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-954">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-955">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-955">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-956">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-956">1.0</span></span>|
|[<span data-ttu-id="daa6d-957">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-957">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-958">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-958">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-959">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-959">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-960">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-960">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="daa6d-961">返回：</span><span class="sxs-lookup"><span data-stu-id="daa6d-961">Returns:</span></span>

<span data-ttu-id="daa6d-p165">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="daa6d-964">类型：对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-964">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="daa6d-965">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-965">Example</span></span>

<span data-ttu-id="daa6d-966">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="daa6d-966">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="daa6d-967">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="daa6d-967">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="daa6d-968">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="daa6d-968">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-969">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-969">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="daa6d-970">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="daa6d-970">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="daa6d-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-973">参数</span><span class="sxs-lookup"><span data-stu-id="daa6d-973">Parameters</span></span>

|<span data-ttu-id="daa6d-974">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-974">Name</span></span>| <span data-ttu-id="daa6d-975">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-975">Type</span></span>| <span data-ttu-id="daa6d-976">描述</span><span class="sxs-lookup"><span data-stu-id="daa6d-976">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="daa6d-977">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-977">String</span></span>|<span data-ttu-id="daa6d-978">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="daa6d-978">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="daa6d-979">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-979">Requirements</span></span>

|<span data-ttu-id="daa6d-980">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-980">Requirement</span></span>| <span data-ttu-id="daa6d-981">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-981">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-982">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-982">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-983">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-983">1.0</span></span>|
|[<span data-ttu-id="daa6d-984">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-984">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-985">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-985">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-986">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-986">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-987">阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-987">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="daa6d-988">返回：</span><span class="sxs-lookup"><span data-stu-id="daa6d-988">Returns:</span></span>

<span data-ttu-id="daa6d-989">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="daa6d-989">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="daa6d-990">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="daa6d-990">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="daa6d-991">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-991">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="daa6d-992">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="daa6d-992">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="daa6d-993">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="daa6d-993">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="daa6d-p167">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-996">在 Outlook 网页版中，如果未选中任何文本，但光标位于正文中，则该方法返回字符串“null”。</span><span class="sxs-lookup"><span data-stu-id="daa6d-996">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="daa6d-997">要检查这种情况，请包含类似于以下内容的代码：</span><span class="sxs-lookup"><span data-stu-id="daa6d-997">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="daa6d-998">参数</span><span class="sxs-lookup"><span data-stu-id="daa6d-998">Parameters</span></span>

|<span data-ttu-id="daa6d-999">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-999">Name</span></span>| <span data-ttu-id="daa6d-1000">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-1000">Type</span></span>| <span data-ttu-id="daa6d-1001">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-1001">Attributes</span></span>| <span data-ttu-id="daa6d-1002">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-1002">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="daa6d-1003">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="daa6d-1003">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="daa6d-p169">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="daa6d-1007">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-1007">Object</span></span>| <span data-ttu-id="daa6d-1008">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1009">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1009">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="daa6d-1010">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-1010">Object</span></span>| <span data-ttu-id="daa6d-1011">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1011">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1012">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1012">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="daa6d-1013">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-1013">function</span></span>||<span data-ttu-id="daa6d-1014">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1014">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="daa6d-1015">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1015">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="daa6d-1016">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1016">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="daa6d-1017">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1017">Requirements</span></span>

|<span data-ttu-id="daa6d-1018">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1018">Requirement</span></span>| <span data-ttu-id="daa6d-1019">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-1019">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-1020">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-1020">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-1021">1.2</span><span class="sxs-lookup"><span data-stu-id="daa6d-1021">1.2</span></span>|
|[<span data-ttu-id="daa6d-1022">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-1022">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-1023">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-1023">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-1024">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-1024">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-1025">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-1025">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="daa6d-1026">返回：</span><span class="sxs-lookup"><span data-stu-id="daa6d-1026">Returns:</span></span>

<span data-ttu-id="daa6d-1027">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1027">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="daa6d-1028">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-1028">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="daa6d-1029">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-1029">Example</span></span>

```js
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

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="daa6d-1030">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="daa6d-1030">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="daa6d-1031">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1031">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="daa6d-p171">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-1035">参数</span><span class="sxs-lookup"><span data-stu-id="daa6d-1035">Parameters</span></span>

|<span data-ttu-id="daa6d-1036">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-1036">Name</span></span>| <span data-ttu-id="daa6d-1037">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-1037">Type</span></span>| <span data-ttu-id="daa6d-1038">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-1038">Attributes</span></span>| <span data-ttu-id="daa6d-1039">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-1039">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="daa6d-1040">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-1040">function</span></span>||<span data-ttu-id="daa6d-1041">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1041">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="daa6d-1042">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1042">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.5) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="daa6d-1043">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1043">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="daa6d-1044">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-1044">Object</span></span>| <span data-ttu-id="daa6d-1045">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1046">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1046">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="daa6d-1047">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1047">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="daa6d-1048">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1048">Requirements</span></span>

|<span data-ttu-id="daa6d-1049">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1049">Requirement</span></span>| <span data-ttu-id="daa6d-1050">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-1051">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-1052">1.0</span><span class="sxs-lookup"><span data-stu-id="daa6d-1052">1.0</span></span>|
|[<span data-ttu-id="daa6d-1053">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-1054">ReadItem</span></span>|
|[<span data-ttu-id="daa6d-1055">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-1056">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="daa6d-1056">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-1057">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-1057">Example</span></span>

<span data-ttu-id="daa6d-p174">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="daa6d-1061">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="daa6d-1061">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="daa6d-1062">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1062">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="daa6d-1063">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1063">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="daa6d-1064">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1064">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="daa6d-1065">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1065">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="daa6d-1066">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1066">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-1067">Parameters</span><span class="sxs-lookup"><span data-stu-id="daa6d-1067">Parameters</span></span>

|<span data-ttu-id="daa6d-1068">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-1068">Name</span></span>| <span data-ttu-id="daa6d-1069">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-1069">Type</span></span>| <span data-ttu-id="daa6d-1070">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-1070">Attributes</span></span>| <span data-ttu-id="daa6d-1071">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-1071">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="daa6d-1072">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-1072">String</span></span>||<span data-ttu-id="daa6d-1073">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1073">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="daa6d-1074">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-1074">Object</span></span>| <span data-ttu-id="daa6d-1075">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1076">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1076">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="daa6d-1077">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-1077">Object</span></span>| <span data-ttu-id="daa6d-1078">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1079">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1079">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="daa6d-1080">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-1080">function</span></span>| <span data-ttu-id="daa6d-1081">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1082">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1082">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="daa6d-1083">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1083">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="daa6d-1084">错误</span><span class="sxs-lookup"><span data-stu-id="daa6d-1084">Errors</span></span>

| <span data-ttu-id="daa6d-1085">错误代码</span><span class="sxs-lookup"><span data-stu-id="daa6d-1085">Error code</span></span> | <span data-ttu-id="daa6d-1086">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-1086">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="daa6d-1087">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1087">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="daa6d-1088">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1088">Requirements</span></span>

|<span data-ttu-id="daa6d-1089">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1089">Requirement</span></span>| <span data-ttu-id="daa6d-1090">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-1091">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-1092">1.1</span><span class="sxs-lookup"><span data-stu-id="daa6d-1092">1.1</span></span>|
|[<span data-ttu-id="daa6d-1093">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-1093">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="daa6d-1095">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-1095">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-1096">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-1096">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-1097">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-1097">Example</span></span>

<span data-ttu-id="daa6d-1098">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1098">The following code removes an attachment with an identifier of '0'.</span></span>

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

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="daa6d-1099">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="daa6d-1099">saveAsync([options], callback)</span></span>

<span data-ttu-id="daa6d-1100">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1100">Asynchronously saves an item.</span></span>

<span data-ttu-id="daa6d-1101">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1101">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="daa6d-1102">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1102">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="daa6d-1103">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1103">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-1104">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1104">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="daa6d-1105">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1105">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="daa6d-p178">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p178">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="daa6d-1109">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="daa6d-1109">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="daa6d-1110">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1110">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="daa6d-1111">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1111">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="daa6d-1112">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1112">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="daa6d-1113">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1113">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-1114">参数</span><span class="sxs-lookup"><span data-stu-id="daa6d-1114">Parameters</span></span>

|<span data-ttu-id="daa6d-1115">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-1115">Name</span></span>| <span data-ttu-id="daa6d-1116">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-1116">Type</span></span>| <span data-ttu-id="daa6d-1117">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-1117">Attributes</span></span>| <span data-ttu-id="daa6d-1118">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-1118">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="daa6d-1119">Object</span><span class="sxs-lookup"><span data-stu-id="daa6d-1119">Object</span></span>| <span data-ttu-id="daa6d-1120">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1121">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="daa6d-1122">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-1122">Object</span></span>| <span data-ttu-id="daa6d-1123">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1124">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="daa6d-1125">函数</span><span class="sxs-lookup"><span data-stu-id="daa6d-1125">function</span></span>||<span data-ttu-id="daa6d-1126">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1126">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="daa6d-1127">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1127">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="daa6d-1128">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1128">Requirements</span></span>

|<span data-ttu-id="daa6d-1129">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1129">Requirement</span></span>| <span data-ttu-id="daa6d-1130">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-1131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-1131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-1132">1.3</span><span class="sxs-lookup"><span data-stu-id="daa6d-1132">1.3</span></span>|
|[<span data-ttu-id="daa6d-1133">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-1133">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-1134">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-1134">ReadWriteItem</span></span>|
|[<span data-ttu-id="daa6d-1135">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-1135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-1136">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-1136">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="daa6d-1137">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-1137">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="daa6d-p180">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p180">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="daa6d-1140">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="daa6d-1140">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="daa6d-1141">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1141">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="daa6d-p181">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p181">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="daa6d-1145">参数</span><span class="sxs-lookup"><span data-stu-id="daa6d-1145">Parameters</span></span>

|<span data-ttu-id="daa6d-1146">名称</span><span class="sxs-lookup"><span data-stu-id="daa6d-1146">Name</span></span>| <span data-ttu-id="daa6d-1147">类型</span><span class="sxs-lookup"><span data-stu-id="daa6d-1147">Type</span></span>| <span data-ttu-id="daa6d-1148">属性</span><span class="sxs-lookup"><span data-stu-id="daa6d-1148">Attributes</span></span>| <span data-ttu-id="daa6d-1149">说明</span><span class="sxs-lookup"><span data-stu-id="daa6d-1149">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="daa6d-1150">字符串</span><span class="sxs-lookup"><span data-stu-id="daa6d-1150">String</span></span>||<span data-ttu-id="daa6d-p182">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="daa6d-p182">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="daa6d-1154">Object</span><span class="sxs-lookup"><span data-stu-id="daa6d-1154">Object</span></span>| <span data-ttu-id="daa6d-1155">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1155">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1156">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1156">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="daa6d-1157">对象</span><span class="sxs-lookup"><span data-stu-id="daa6d-1157">Object</span></span>| <span data-ttu-id="daa6d-1158">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1158">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1159">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1159">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="daa6d-1160">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="daa6d-1160">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="daa6d-1161">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="daa6d-1161">&lt;optional&gt;</span></span>|<span data-ttu-id="daa6d-1162">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1162">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="daa6d-1163">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1163">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="daa6d-1164">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1164">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="daa6d-1165">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1165">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="daa6d-1166">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1166">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="daa6d-1167">function</span><span class="sxs-lookup"><span data-stu-id="daa6d-1167">function</span></span>||<span data-ttu-id="daa6d-1168">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="daa6d-1168">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="daa6d-1169">Requirements</span><span class="sxs-lookup"><span data-stu-id="daa6d-1169">Requirements</span></span>

|<span data-ttu-id="daa6d-1170">要求</span><span class="sxs-lookup"><span data-stu-id="daa6d-1170">Requirement</span></span>| <span data-ttu-id="daa6d-1171">值</span><span class="sxs-lookup"><span data-stu-id="daa6d-1171">Value</span></span>|
|---|---|
|[<span data-ttu-id="daa6d-1172">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="daa6d-1172">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="daa6d-1173">1.2</span><span class="sxs-lookup"><span data-stu-id="daa6d-1173">1.2</span></span>|
|[<span data-ttu-id="daa6d-1174">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="daa6d-1174">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="daa6d-1175">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="daa6d-1175">ReadWriteItem</span></span>|
|[<span data-ttu-id="daa6d-1176">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="daa6d-1176">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="daa6d-1177">撰写</span><span class="sxs-lookup"><span data-stu-id="daa6d-1177">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="daa6d-1178">示例</span><span class="sxs-lookup"><span data-stu-id="daa6d-1178">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
