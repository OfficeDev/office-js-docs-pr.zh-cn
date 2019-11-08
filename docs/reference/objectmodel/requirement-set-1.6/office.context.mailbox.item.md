---
title: "\"Context\"-\"邮箱\"。项目-要求集1。6"
description: ''
ms.date: 11/06/2019
localization_priority: Normal
ms.openlocfilehash: 4aa9b5ae086b9879842a6f1cdd7125b74aa0c54d
ms.sourcegitcommit: 08c0b9ff319c391922fa43d3c2e9783cf6b53b1b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/08/2019
ms.locfileid: "38066142"
---
# <a name="item"></a><span data-ttu-id="4e408-102">item</span><span class="sxs-lookup"><span data-stu-id="4e408-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4e408-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4e408-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4e408-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="4e408-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-106">Requirements</span></span>

|<span data-ttu-id="4e408-107">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-107">Requirement</span></span>| <span data-ttu-id="4e408-108">值</span><span class="sxs-lookup"><span data-stu-id="4e408-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-110">1.0</span></span>|
|[<span data-ttu-id="4e408-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-112">受限</span><span class="sxs-lookup"><span data-stu-id="4e408-112">Restricted</span></span>|
|[<span data-ttu-id="4e408-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4e408-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="4e408-115">Members and methods</span></span>

| <span data-ttu-id="4e408-116">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-116">Member</span></span> | <span data-ttu-id="4e408-117">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4e408-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4e408-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="4e408-119">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-119">Member</span></span> |
| [<span data-ttu-id="4e408-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4e408-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="4e408-121">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-121">Member</span></span> |
| [<span data-ttu-id="4e408-122">body</span><span class="sxs-lookup"><span data-stu-id="4e408-122">body</span></span>](#body-body) | <span data-ttu-id="4e408-123">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-123">Member</span></span> |
| [<span data-ttu-id="4e408-124">cc</span><span class="sxs-lookup"><span data-stu-id="4e408-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4e408-125">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-125">Member</span></span> |
| [<span data-ttu-id="4e408-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4e408-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4e408-127">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-127">Member</span></span> |
| [<span data-ttu-id="4e408-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4e408-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4e408-129">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-129">Member</span></span> |
| [<span data-ttu-id="4e408-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4e408-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4e408-131">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-131">Member</span></span> |
| [<span data-ttu-id="4e408-132">end</span><span class="sxs-lookup"><span data-stu-id="4e408-132">end</span></span>](#end-datetime) | <span data-ttu-id="4e408-133">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-133">Member</span></span> |
| [<span data-ttu-id="4e408-134">from</span><span class="sxs-lookup"><span data-stu-id="4e408-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="4e408-135">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-135">Member</span></span> |
| [<span data-ttu-id="4e408-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4e408-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4e408-137">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-137">Member</span></span> |
| [<span data-ttu-id="4e408-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="4e408-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4e408-139">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-139">Member</span></span> |
| [<span data-ttu-id="4e408-140">itemId</span><span class="sxs-lookup"><span data-stu-id="4e408-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4e408-141">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-141">Member</span></span> |
| [<span data-ttu-id="4e408-142">itemType</span><span class="sxs-lookup"><span data-stu-id="4e408-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="4e408-143">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-143">Member</span></span> |
| [<span data-ttu-id="4e408-144">location</span><span class="sxs-lookup"><span data-stu-id="4e408-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="4e408-145">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-145">Member</span></span> |
| [<span data-ttu-id="4e408-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4e408-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4e408-147">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-147">Member</span></span> |
| [<span data-ttu-id="4e408-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4e408-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="4e408-149">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-149">Member</span></span> |
| [<span data-ttu-id="4e408-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4e408-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4e408-151">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-151">Member</span></span> |
| [<span data-ttu-id="4e408-152">organizer</span><span class="sxs-lookup"><span data-stu-id="4e408-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="4e408-153">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-153">Member</span></span> |
| [<span data-ttu-id="4e408-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4e408-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4e408-155">Member</span><span class="sxs-lookup"><span data-stu-id="4e408-155">Member</span></span> |
| [<span data-ttu-id="4e408-156">sender</span><span class="sxs-lookup"><span data-stu-id="4e408-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="4e408-157">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-157">Member</span></span> |
| [<span data-ttu-id="4e408-158">start</span><span class="sxs-lookup"><span data-stu-id="4e408-158">start</span></span>](#start-datetime) | <span data-ttu-id="4e408-159">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-159">Member</span></span> |
| [<span data-ttu-id="4e408-160">subject</span><span class="sxs-lookup"><span data-stu-id="4e408-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="4e408-161">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-161">Member</span></span> |
| [<span data-ttu-id="4e408-162">to</span><span class="sxs-lookup"><span data-stu-id="4e408-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4e408-163">成员</span><span class="sxs-lookup"><span data-stu-id="4e408-163">Member</span></span> |
| [<span data-ttu-id="4e408-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4e408-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4e408-165">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-165">Method</span></span> |
| [<span data-ttu-id="4e408-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4e408-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4e408-167">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-167">Method</span></span> |
| [<span data-ttu-id="4e408-168">close</span><span class="sxs-lookup"><span data-stu-id="4e408-168">close</span></span>](#close) | <span data-ttu-id="4e408-169">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-169">Method</span></span> |
| [<span data-ttu-id="4e408-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4e408-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="4e408-171">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-171">Method</span></span> |
| [<span data-ttu-id="4e408-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4e408-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="4e408-173">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-173">Method</span></span> |
| [<span data-ttu-id="4e408-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="4e408-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="4e408-175">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-175">Method</span></span> |
| [<span data-ttu-id="4e408-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4e408-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4e408-177">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-177">Method</span></span> |
| [<span data-ttu-id="4e408-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4e408-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4e408-179">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-179">Method</span></span> |
| [<span data-ttu-id="4e408-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4e408-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4e408-181">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-181">Method</span></span> |
| [<span data-ttu-id="4e408-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4e408-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4e408-183">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-183">Method</span></span> |
| [<span data-ttu-id="4e408-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4e408-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4e408-185">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-185">Method</span></span> |
| [<span data-ttu-id="4e408-186">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="4e408-186">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="4e408-187">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-187">Method</span></span> |
| [<span data-ttu-id="4e408-188">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="4e408-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="4e408-189">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-189">Method</span></span> |
| [<span data-ttu-id="4e408-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4e408-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4e408-191">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-191">Method</span></span> |
| [<span data-ttu-id="4e408-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4e408-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4e408-193">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-193">Method</span></span> |
| [<span data-ttu-id="4e408-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4e408-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4e408-195">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-195">Method</span></span> |
| [<span data-ttu-id="4e408-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4e408-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4e408-197">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4e408-198">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-198">Example</span></span>

<span data-ttu-id="4e408-199">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="4e408-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4e408-200">Members</span><span class="sxs-lookup"><span data-stu-id="4e408-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-16"></a><span data-ttu-id="4e408-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="4e408-201">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

<span data-ttu-id="4e408-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-204">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="4e408-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4e408-205">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="4e408-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-206">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-206">Type</span></span>

*   <span data-ttu-id="4e408-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span><span class="sxs-lookup"><span data-stu-id="4e408-207">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-208">Requirements</span></span>

|<span data-ttu-id="4e408-209">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-209">Requirement</span></span>| <span data-ttu-id="4e408-210">值</span><span class="sxs-lookup"><span data-stu-id="4e408-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-212">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-212">1.0</span></span>|
|[<span data-ttu-id="4e408-213">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-213">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-214">ReadItem</span></span>|
|[<span data-ttu-id="4e408-215">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-216">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-217">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-217">Example</span></span>

<span data-ttu-id="4e408-218">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="4e408-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4e408-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-219">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-220">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4e408-221">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-221">Compose mode only.</span></span>

<span data-ttu-id="4e408-222">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-222">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-223">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="4e408-223">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4e408-224">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-224">Get 500 members maximum.</span></span>
- <span data-ttu-id="4e408-225">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-225">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-226">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-226">Type</span></span>

*   [<span data-ttu-id="4e408-227">收件人</span><span class="sxs-lookup"><span data-stu-id="4e408-227">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4e408-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-228">Requirements</span></span>

|<span data-ttu-id="4e408-229">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-229">Requirement</span></span>| <span data-ttu-id="4e408-230">值</span><span class="sxs-lookup"><span data-stu-id="4e408-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-231">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-232">1.1</span><span class="sxs-lookup"><span data-stu-id="4e408-232">1.1</span></span>|
|[<span data-ttu-id="4e408-233">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-233">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-234">ReadItem</span></span>|
|[<span data-ttu-id="4e408-235">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-236">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-236">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-237">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-237">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-16"></a><span data-ttu-id="4e408-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-238">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-239">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-239">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-240">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-240">Type</span></span>

*   [<span data-ttu-id="4e408-241">Body</span><span class="sxs-lookup"><span data-stu-id="4e408-241">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4e408-242">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-242">Requirements</span></span>

|<span data-ttu-id="4e408-243">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-243">Requirement</span></span>| <span data-ttu-id="4e408-244">值</span><span class="sxs-lookup"><span data-stu-id="4e408-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-245">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-246">1.1</span><span class="sxs-lookup"><span data-stu-id="4e408-246">1.1</span></span>|
|[<span data-ttu-id="4e408-247">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-248">ReadItem</span></span>|
|[<span data-ttu-id="4e408-249">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-250">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-251">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-251">Example</span></span>

<span data-ttu-id="4e408-252">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="4e408-252">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4e408-253">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="4e408-253">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4e408-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-254">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-255">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4e408-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4e408-256">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-257">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-257">Read mode</span></span>

<span data-ttu-id="4e408-258">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-258">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="4e408-259">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-260">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-260">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-261">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-261">Compose mode</span></span>

<span data-ttu-id="4e408-262">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="4e408-263">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-263">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-264">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="4e408-264">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4e408-265">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-265">Get 500 members maximum.</span></span>
- <span data-ttu-id="4e408-266">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-266">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4e408-267">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-267">Type</span></span>

*   <span data-ttu-id="4e408-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-268">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-269">Requirements</span></span>

|<span data-ttu-id="4e408-270">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-270">Requirement</span></span>| <span data-ttu-id="4e408-271">值</span><span class="sxs-lookup"><span data-stu-id="4e408-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-272">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-272">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-273">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-273">1.0</span></span>|
|[<span data-ttu-id="4e408-274">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-274">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-275">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-275">ReadItem</span></span>|
|[<span data-ttu-id="4e408-276">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-276">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-277">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-277">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="4e408-278">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="4e408-278">(nullable) conversationId: String</span></span>

<span data-ttu-id="4e408-279">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="4e408-279">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4e408-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="4e408-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4e408-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="4e408-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-284">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-284">Type</span></span>

*   <span data-ttu-id="4e408-285">String</span><span class="sxs-lookup"><span data-stu-id="4e408-285">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-286">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-286">Requirements</span></span>

|<span data-ttu-id="4e408-287">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-287">Requirement</span></span>| <span data-ttu-id="4e408-288">值</span><span class="sxs-lookup"><span data-stu-id="4e408-288">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-289">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-289">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-290">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-290">1.0</span></span>|
|[<span data-ttu-id="4e408-291">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-291">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-292">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-292">ReadItem</span></span>|
|[<span data-ttu-id="4e408-293">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-293">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-294">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-294">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-295">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-295">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="4e408-296">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="4e408-296">dateTimeCreated: Date</span></span>

<span data-ttu-id="4e408-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-299">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-299">Type</span></span>

*   <span data-ttu-id="4e408-300">日期</span><span class="sxs-lookup"><span data-stu-id="4e408-300">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-301">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-301">Requirements</span></span>

|<span data-ttu-id="4e408-302">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-302">Requirement</span></span>| <span data-ttu-id="4e408-303">值</span><span class="sxs-lookup"><span data-stu-id="4e408-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-304">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-305">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-305">1.0</span></span>|
|[<span data-ttu-id="4e408-306">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-307">ReadItem</span></span>|
|[<span data-ttu-id="4e408-308">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-309">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-310">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-310">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="4e408-311">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="4e408-311">dateTimeModified: Date</span></span>

<span data-ttu-id="4e408-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-314">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-314">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-315">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-315">Type</span></span>

*   <span data-ttu-id="4e408-316">日期</span><span class="sxs-lookup"><span data-stu-id="4e408-316">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-317">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-317">Requirements</span></span>

|<span data-ttu-id="4e408-318">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-318">Requirement</span></span>| <span data-ttu-id="4e408-319">值</span><span class="sxs-lookup"><span data-stu-id="4e408-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-320">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-321">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-321">1.0</span></span>|
|[<span data-ttu-id="4e408-322">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-323">ReadItem</span></span>|
|[<span data-ttu-id="4e408-324">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-325">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-325">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-326">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-326">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="4e408-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-327">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-328">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4e408-328">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4e408-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4e408-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-331">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-331">Read mode</span></span>

<span data-ttu-id="4e408-332">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-332">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-333">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-333">Compose mode</span></span>

<span data-ttu-id="4e408-334">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-334">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4e408-335">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="4e408-335">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4e408-336">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="4e408-336">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4e408-337">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-337">Type</span></span>

*   <span data-ttu-id="4e408-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-338">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-339">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-339">Requirements</span></span>

|<span data-ttu-id="4e408-340">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-340">Requirement</span></span>| <span data-ttu-id="4e408-341">值</span><span class="sxs-lookup"><span data-stu-id="4e408-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-342">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-343">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-343">1.0</span></span>|
|[<span data-ttu-id="4e408-344">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-344">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-345">ReadItem</span></span>|
|[<span data-ttu-id="4e408-346">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-346">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-347">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-347">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="4e408-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-348">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-p114">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="4e408-p115">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="4e408-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-353">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="4e408-353">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-354">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-354">Type</span></span>

*   [<span data-ttu-id="4e408-355">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4e408-355">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="example"></a><span data-ttu-id="4e408-356">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="requirements"></a><span data-ttu-id="4e408-357">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-357">Requirements</span></span>

|<span data-ttu-id="4e408-358">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-358">Requirement</span></span>| <span data-ttu-id="4e408-359">值</span><span class="sxs-lookup"><span data-stu-id="4e408-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-360">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-361">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-361">1.0</span></span>|
|[<span data-ttu-id="4e408-362">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-363">ReadItem</span></span>|
|[<span data-ttu-id="4e408-364">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-365">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-365">Read</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="4e408-366">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="4e408-366">internetMessageId: String</span></span>

<span data-ttu-id="4e408-p116">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-369">Type</span><span class="sxs-lookup"><span data-stu-id="4e408-369">Type</span></span>

*   <span data-ttu-id="4e408-370">String</span><span class="sxs-lookup"><span data-stu-id="4e408-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-371">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-371">Requirements</span></span>

|<span data-ttu-id="4e408-372">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-372">Requirement</span></span>| <span data-ttu-id="4e408-373">值</span><span class="sxs-lookup"><span data-stu-id="4e408-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-374">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-374">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-375">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-375">1.0</span></span>|
|[<span data-ttu-id="4e408-376">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-376">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-377">ReadItem</span></span>|
|[<span data-ttu-id="4e408-378">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-378">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-379">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-380">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-380">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="4e408-381">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="4e408-381">itemClass: String</span></span>

<span data-ttu-id="4e408-p117">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4e408-p118">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="4e408-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="4e408-386">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-386">Type</span></span> | <span data-ttu-id="4e408-387">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-387">Description</span></span> | <span data-ttu-id="4e408-388">项目类</span><span class="sxs-lookup"><span data-stu-id="4e408-388">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="4e408-389">约会项目</span><span class="sxs-lookup"><span data-stu-id="4e408-389">Appointment items</span></span> | <span data-ttu-id="4e408-390">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="4e408-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="4e408-391">邮件项目</span><span class="sxs-lookup"><span data-stu-id="4e408-391">Message items</span></span> | <span data-ttu-id="4e408-392">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="4e408-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="4e408-393">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="4e408-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-394">Type</span><span class="sxs-lookup"><span data-stu-id="4e408-394">Type</span></span>

*   <span data-ttu-id="4e408-395">String</span><span class="sxs-lookup"><span data-stu-id="4e408-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-396">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-396">Requirements</span></span>

|<span data-ttu-id="4e408-397">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-397">Requirement</span></span>| <span data-ttu-id="4e408-398">值</span><span class="sxs-lookup"><span data-stu-id="4e408-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-399">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-400">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-400">1.0</span></span>|
|[<span data-ttu-id="4e408-401">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-401">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-402">ReadItem</span></span>|
|[<span data-ttu-id="4e408-403">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-403">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-404">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-405">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-405">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4e408-406">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="4e408-406">(nullable) itemId: String</span></span>

<span data-ttu-id="4e408-p119">获取当前项目的 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p119">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-409">`itemId` 属性返回的标识符与 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)相同。</span><span class="sxs-lookup"><span data-stu-id="4e408-409">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="4e408-410">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="4e408-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4e408-411">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="4e408-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4e408-412">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="4e408-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4e408-p121">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="4e408-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-415">Type</span><span class="sxs-lookup"><span data-stu-id="4e408-415">Type</span></span>

*   <span data-ttu-id="4e408-416">String</span><span class="sxs-lookup"><span data-stu-id="4e408-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-417">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-417">Requirements</span></span>

|<span data-ttu-id="4e408-418">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-418">Requirement</span></span>| <span data-ttu-id="4e408-419">值</span><span class="sxs-lookup"><span data-stu-id="4e408-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-420">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-421">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-421">1.0</span></span>|
|[<span data-ttu-id="4e408-422">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-422">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-423">ReadItem</span></span>|
|[<span data-ttu-id="4e408-424">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-424">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-425">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-426">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-426">Example</span></span>

<span data-ttu-id="4e408-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="4e408-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-16"></a><span data-ttu-id="4e408-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-429">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-430">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="4e408-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4e408-431">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="4e408-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-432">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-432">Type</span></span>

*   [<span data-ttu-id="4e408-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4e408-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4e408-434">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-434">Requirements</span></span>

|<span data-ttu-id="4e408-435">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-435">Requirement</span></span>| <span data-ttu-id="4e408-436">值</span><span class="sxs-lookup"><span data-stu-id="4e408-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-437">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-437">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-438">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-438">1.0</span></span>|
|[<span data-ttu-id="4e408-439">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-439">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-440">ReadItem</span></span>|
|[<span data-ttu-id="4e408-441">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-441">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-442">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-442">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-443">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-443">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-16"></a><span data-ttu-id="4e408-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-444">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-445">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="4e408-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-446">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-446">Read mode</span></span>

<span data-ttu-id="4e408-447">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="4e408-447">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-448">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-448">Compose mode</span></span>

<span data-ttu-id="4e408-449">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4e408-450">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-450">Type</span></span>

*   <span data-ttu-id="4e408-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-451">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-452">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-452">Requirements</span></span>

|<span data-ttu-id="4e408-453">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-453">Requirement</span></span>| <span data-ttu-id="4e408-454">值</span><span class="sxs-lookup"><span data-stu-id="4e408-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-455">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-456">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-456">1.0</span></span>|
|[<span data-ttu-id="4e408-457">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-457">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-458">ReadItem</span></span>|
|[<span data-ttu-id="4e408-459">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-459">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-460">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-460">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4e408-461">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="4e408-461">normalizedSubject: String</span></span>

<span data-ttu-id="4e408-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4e408-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="4e408-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-466">Type</span><span class="sxs-lookup"><span data-stu-id="4e408-466">Type</span></span>

*   <span data-ttu-id="4e408-467">String</span><span class="sxs-lookup"><span data-stu-id="4e408-467">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-468">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-468">Requirements</span></span>

|<span data-ttu-id="4e408-469">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-469">Requirement</span></span>| <span data-ttu-id="4e408-470">值</span><span class="sxs-lookup"><span data-stu-id="4e408-470">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-471">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-471">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-472">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-472">1.0</span></span>|
|[<span data-ttu-id="4e408-473">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-473">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-474">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-474">ReadItem</span></span>|
|[<span data-ttu-id="4e408-475">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-475">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-476">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-476">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-477">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-477">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-16"></a><span data-ttu-id="4e408-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-478">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-479">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="4e408-479">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-480">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-480">Type</span></span>

*   [<span data-ttu-id="4e408-481">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4e408-481">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4e408-482">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-482">Requirements</span></span>

|<span data-ttu-id="4e408-483">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-483">Requirement</span></span>| <span data-ttu-id="4e408-484">值</span><span class="sxs-lookup"><span data-stu-id="4e408-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-485">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-486">1.3</span><span class="sxs-lookup"><span data-stu-id="4e408-486">1.3</span></span>|
|[<span data-ttu-id="4e408-487">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-488">ReadItem</span></span>|
|[<span data-ttu-id="4e408-489">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-490">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-490">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-491">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-491">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4e408-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-492">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-493">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4e408-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4e408-494">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-495">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-495">Read mode</span></span>

<span data-ttu-id="4e408-496">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="4e408-497">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-498">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-498">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-499">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-499">Compose mode</span></span>

<span data-ttu-id="4e408-500">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-500">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="4e408-501">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-501">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-502">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="4e408-502">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4e408-503">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-503">Get 500 members maximum.</span></span>
- <span data-ttu-id="4e408-504">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-504">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4e408-505">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-505">Type</span></span>

*   <span data-ttu-id="4e408-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-506">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-507">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-507">Requirements</span></span>

|<span data-ttu-id="4e408-508">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-508">Requirement</span></span>| <span data-ttu-id="4e408-509">值</span><span class="sxs-lookup"><span data-stu-id="4e408-509">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-510">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-510">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-511">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-511">1.0</span></span>|
|[<span data-ttu-id="4e408-512">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-513">ReadItem</span></span>|
|[<span data-ttu-id="4e408-514">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-514">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-515">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-515">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="4e408-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-516">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-p128">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-519">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-519">Type</span></span>

*   [<span data-ttu-id="4e408-520">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4e408-520">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4e408-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-521">Requirements</span></span>

|<span data-ttu-id="4e408-522">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-522">Requirement</span></span>| <span data-ttu-id="4e408-523">值</span><span class="sxs-lookup"><span data-stu-id="4e408-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-524">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-525">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-525">1.0</span></span>|
|[<span data-ttu-id="4e408-526">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-527">ReadItem</span></span>|
|[<span data-ttu-id="4e408-528">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-529">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-529">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-530">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-530">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4e408-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-531">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-532">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4e408-532">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4e408-533">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-533">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-534">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-534">Read mode</span></span>

<span data-ttu-id="4e408-535">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-535">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="4e408-536">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-537">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-537">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-538">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-538">Compose mode</span></span>

<span data-ttu-id="4e408-539">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-539">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="4e408-540">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-540">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-541">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="4e408-541">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4e408-542">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-542">Get 500 members maximum.</span></span>
- <span data-ttu-id="4e408-543">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-543">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4e408-544">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-544">Type</span></span>

*   <span data-ttu-id="4e408-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-545">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-546">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-546">Requirements</span></span>

|<span data-ttu-id="4e408-547">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-547">Requirement</span></span>| <span data-ttu-id="4e408-548">值</span><span class="sxs-lookup"><span data-stu-id="4e408-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-549">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-550">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-550">1.0</span></span>|
|[<span data-ttu-id="4e408-551">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-552">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-552">ReadItem</span></span>|
|[<span data-ttu-id="4e408-553">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-554">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-554">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16"></a><span data-ttu-id="4e408-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-555">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-p132">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4e408-p133">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="4e408-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-560">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="4e408-560">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4e408-561">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-561">Type</span></span>

*   [<span data-ttu-id="4e408-562">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4e408-562">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="4e408-563">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-563">Requirements</span></span>

|<span data-ttu-id="4e408-564">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-564">Requirement</span></span>| <span data-ttu-id="4e408-565">值</span><span class="sxs-lookup"><span data-stu-id="4e408-565">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-566">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-566">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-567">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-567">1.0</span></span>|
|[<span data-ttu-id="4e408-568">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-568">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-569">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-569">ReadItem</span></span>|
|[<span data-ttu-id="4e408-570">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-570">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-571">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-571">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-572">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-572">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-16"></a><span data-ttu-id="4e408-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-573">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-574">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4e408-574">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4e408-p134">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4e408-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-577">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-577">Read mode</span></span>

<span data-ttu-id="4e408-578">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-578">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-579">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-579">Compose mode</span></span>

<span data-ttu-id="4e408-580">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-580">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4e408-581">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="4e408-581">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4e408-582">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="4e408-582">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.6#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4e408-583">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-583">Type</span></span>

*   <span data-ttu-id="4e408-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-584">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-585">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-585">Requirements</span></span>

|<span data-ttu-id="4e408-586">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-586">Requirement</span></span>| <span data-ttu-id="4e408-587">值</span><span class="sxs-lookup"><span data-stu-id="4e408-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-588">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-589">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-589">1.0</span></span>|
|[<span data-ttu-id="4e408-590">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-591">ReadItem</span></span>|
|[<span data-ttu-id="4e408-592">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-593">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-593">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-16"></a><span data-ttu-id="4e408-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-594">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-595">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="4e408-595">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4e408-596">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="4e408-596">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-597">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-597">Read mode</span></span>

<span data-ttu-id="4e408-p135">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="4e408-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-600">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-600">Compose mode</span></span>

<span data-ttu-id="4e408-601">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-601">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4e408-602">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-602">Type</span></span>

*   <span data-ttu-id="4e408-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-603">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-604">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-604">Requirements</span></span>

|<span data-ttu-id="4e408-605">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-605">Requirement</span></span>| <span data-ttu-id="4e408-606">值</span><span class="sxs-lookup"><span data-stu-id="4e408-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-607">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-608">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-608">1.0</span></span>|
|[<span data-ttu-id="4e408-609">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-610">ReadItem</span></span>|
|[<span data-ttu-id="4e408-611">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-612">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-612">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-16recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-16"></a><span data-ttu-id="4e408-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-613">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

<span data-ttu-id="4e408-614">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4e408-614">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4e408-615">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4e408-615">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4e408-616">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4e408-616">Read mode</span></span>

<span data-ttu-id="4e408-617">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-617">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="4e408-618">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-619">但是，在 Windows 和 Mac 上，您可以设置为最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-619">However, on Windows and Mac, you can set up to get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4e408-620">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4e408-620">Compose mode</span></span>

<span data-ttu-id="4e408-621">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-621">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="4e408-622">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-622">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="4e408-623">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="4e408-623">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="4e408-624">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-624">Get 500 members maximum.</span></span>
- <span data-ttu-id="4e408-625">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="4e408-625">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4e408-626">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-626">Type</span></span>

*   <span data-ttu-id="4e408-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-627">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.6)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-628">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-628">Requirements</span></span>

|<span data-ttu-id="4e408-629">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-629">Requirement</span></span>| <span data-ttu-id="4e408-630">值</span><span class="sxs-lookup"><span data-stu-id="4e408-630">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-631">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-631">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-632">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-632">1.0</span></span>|
|[<span data-ttu-id="4e408-633">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-633">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-634">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-634">ReadItem</span></span>|
|[<span data-ttu-id="4e408-635">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-635">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-636">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-636">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4e408-637">方法</span><span class="sxs-lookup"><span data-stu-id="4e408-637">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4e408-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e408-638">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4e408-639">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="4e408-639">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4e408-640">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="4e408-640">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4e408-641">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="4e408-641">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-642">参数</span><span class="sxs-lookup"><span data-stu-id="4e408-642">Parameters</span></span>

|<span data-ttu-id="4e408-643">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-643">Name</span></span>| <span data-ttu-id="4e408-644">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-644">Type</span></span>| <span data-ttu-id="4e408-645">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-645">Attributes</span></span>| <span data-ttu-id="4e408-646">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-646">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="4e408-647">String</span><span class="sxs-lookup"><span data-stu-id="4e408-647">String</span></span>||<span data-ttu-id="4e408-p139">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4e408-650">字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-650">String</span></span>||<span data-ttu-id="4e408-p140">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4e408-653">Object</span><span class="sxs-lookup"><span data-stu-id="4e408-653">Object</span></span>| <span data-ttu-id="4e408-654">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-654">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-655">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4e408-655">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="4e408-656">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-656">Object</span></span> | <span data-ttu-id="4e408-657">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-657">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-658">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="4e408-659">布尔值</span><span class="sxs-lookup"><span data-stu-id="4e408-659">Boolean</span></span> | <span data-ttu-id="4e408-660">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-660">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-661">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="4e408-661">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="4e408-662">函数</span><span class="sxs-lookup"><span data-stu-id="4e408-662">function</span></span>| <span data-ttu-id="4e408-663">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-663">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-664">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-664">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e408-665">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4e408-665">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4e408-666">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-666">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4e408-667">错误</span><span class="sxs-lookup"><span data-stu-id="4e408-667">Errors</span></span>

| <span data-ttu-id="4e408-668">错误代码</span><span class="sxs-lookup"><span data-stu-id="4e408-668">Error code</span></span> | <span data-ttu-id="4e408-669">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-669">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="4e408-670">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="4e408-670">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="4e408-671">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="4e408-671">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4e408-672">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="4e408-672">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4e408-673">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-673">Requirements</span></span>

|<span data-ttu-id="4e408-674">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-674">Requirement</span></span>| <span data-ttu-id="4e408-675">值</span><span class="sxs-lookup"><span data-stu-id="4e408-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-676">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-677">1.1</span><span class="sxs-lookup"><span data-stu-id="4e408-677">1.1</span></span>|
|[<span data-ttu-id="4e408-678">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-679">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e408-679">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e408-680">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-681">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-681">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e408-682">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-682">Examples</span></span>

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

<span data-ttu-id="4e408-683">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="4e408-683">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4e408-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e408-684">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4e408-685">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="4e408-685">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4e408-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="4e408-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4e408-689">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="4e408-689">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4e408-690">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="4e408-690">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-691">Parameters</span><span class="sxs-lookup"><span data-stu-id="4e408-691">Parameters</span></span>

|<span data-ttu-id="4e408-692">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-692">Name</span></span>| <span data-ttu-id="4e408-693">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-693">Type</span></span>| <span data-ttu-id="4e408-694">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-694">Attributes</span></span>| <span data-ttu-id="4e408-695">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-695">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="4e408-696">String</span><span class="sxs-lookup"><span data-stu-id="4e408-696">String</span></span>||<span data-ttu-id="4e408-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4e408-699">String</span><span class="sxs-lookup"><span data-stu-id="4e408-699">String</span></span>||<span data-ttu-id="4e408-700">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="4e408-700">The subject of the item to be attached.</span></span> <span data-ttu-id="4e408-701">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-701">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4e408-702">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-702">Object</span></span>| <span data-ttu-id="4e408-703">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-703">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-704">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4e408-704">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4e408-705">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-705">Object</span></span>| <span data-ttu-id="4e408-706">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-706">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-707">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-707">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4e408-708">函数</span><span class="sxs-lookup"><span data-stu-id="4e408-708">function</span></span>| <span data-ttu-id="4e408-709">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-709">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-710">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-710">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e408-711">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4e408-711">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4e408-712">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-712">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4e408-713">错误</span><span class="sxs-lookup"><span data-stu-id="4e408-713">Errors</span></span>

| <span data-ttu-id="4e408-714">错误代码</span><span class="sxs-lookup"><span data-stu-id="4e408-714">Error code</span></span> | <span data-ttu-id="4e408-715">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-715">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4e408-716">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="4e408-716">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4e408-717">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-717">Requirements</span></span>

|<span data-ttu-id="4e408-718">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-718">Requirement</span></span>| <span data-ttu-id="4e408-719">值</span><span class="sxs-lookup"><span data-stu-id="4e408-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-720">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-721">1.1</span><span class="sxs-lookup"><span data-stu-id="4e408-721">1.1</span></span>|
|[<span data-ttu-id="4e408-722">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-722">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-723">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e408-723">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e408-724">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-724">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-725">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-725">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-726">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-726">Example</span></span>

<span data-ttu-id="4e408-727">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="4e408-727">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="4e408-728">close()</span><span class="sxs-lookup"><span data-stu-id="4e408-728">close()</span></span>

<span data-ttu-id="4e408-729">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="4e408-729">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4e408-p144">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="4e408-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-732">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="4e408-732">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4e408-733">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="4e408-733">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-734">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-734">Requirements</span></span>

|<span data-ttu-id="4e408-735">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-735">Requirement</span></span>| <span data-ttu-id="4e408-736">值</span><span class="sxs-lookup"><span data-stu-id="4e408-736">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-737">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-737">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-738">1.3</span><span class="sxs-lookup"><span data-stu-id="4e408-738">1.3</span></span>|
|[<span data-ttu-id="4e408-739">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-739">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-740">受限</span><span class="sxs-lookup"><span data-stu-id="4e408-740">Restricted</span></span>|
|[<span data-ttu-id="4e408-741">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-741">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-742">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-742">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4e408-743">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4e408-743">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4e408-744">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="4e408-744">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-745">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-745">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4e408-746">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="4e408-746">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4e408-747">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="4e408-747">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4e408-p145">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="4e408-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-751">Parameters</span><span class="sxs-lookup"><span data-stu-id="4e408-751">Parameters</span></span>

| <span data-ttu-id="4e408-752">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-752">Name</span></span> | <span data-ttu-id="4e408-753">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-753">Type</span></span> | <span data-ttu-id="4e408-754">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-754">Attributes</span></span> | <span data-ttu-id="4e408-755">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-755">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="4e408-756">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="4e408-756">String &#124; Object</span></span>| |<span data-ttu-id="4e408-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4e408-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4e408-759">**或**</span><span class="sxs-lookup"><span data-stu-id="4e408-759">**OR**</span></span><br/><span data-ttu-id="4e408-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="4e408-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4e408-762">String</span><span class="sxs-lookup"><span data-stu-id="4e408-762">String</span></span> | <span data-ttu-id="4e408-763">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-763">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4e408-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4e408-766">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-766">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4e408-767">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-767">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-768">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-768">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4e408-769">String</span><span class="sxs-lookup"><span data-stu-id="4e408-769">String</span></span> | | <span data-ttu-id="4e408-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="4e408-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4e408-772">字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-772">String</span></span> | | <span data-ttu-id="4e408-773">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-773">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4e408-774">String</span><span class="sxs-lookup"><span data-stu-id="4e408-774">String</span></span> | | <span data-ttu-id="4e408-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="4e408-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="4e408-777">布尔</span><span class="sxs-lookup"><span data-stu-id="4e408-777">Boolean</span></span> | | <span data-ttu-id="4e408-p151">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="4e408-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4e408-780">String</span><span class="sxs-lookup"><span data-stu-id="4e408-780">String</span></span> | | <span data-ttu-id="4e408-p152">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4e408-784">函数</span><span class="sxs-lookup"><span data-stu-id="4e408-784">function</span></span> | <span data-ttu-id="4e408-785">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-785">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-786">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-786">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4e408-787">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-787">Requirements</span></span>

|<span data-ttu-id="4e408-788">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-788">Requirement</span></span>| <span data-ttu-id="4e408-789">值</span><span class="sxs-lookup"><span data-stu-id="4e408-789">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-790">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-790">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-791">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-791">1.0</span></span>|
|[<span data-ttu-id="4e408-792">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-792">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-793">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-793">ReadItem</span></span>|
|[<span data-ttu-id="4e408-794">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-794">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-795">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-795">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e408-796">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-796">Examples</span></span>

<span data-ttu-id="4e408-797">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-797">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4e408-798">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-798">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4e408-799">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-799">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4e408-800">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-800">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4e408-801">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-801">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4e408-802">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-802">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4e408-803">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4e408-803">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4e408-804">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="4e408-804">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-805">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-805">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4e408-806">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="4e408-806">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4e408-807">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="4e408-807">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4e408-p153">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="4e408-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-811">Parameters</span><span class="sxs-lookup"><span data-stu-id="4e408-811">Parameters</span></span>

| <span data-ttu-id="4e408-812">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-812">Name</span></span> | <span data-ttu-id="4e408-813">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-813">Type</span></span> | <span data-ttu-id="4e408-814">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-814">Attributes</span></span> | <span data-ttu-id="4e408-815">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-815">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="4e408-816">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="4e408-816">String &#124; Object</span></span>| | <span data-ttu-id="4e408-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4e408-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4e408-819">**或**</span><span class="sxs-lookup"><span data-stu-id="4e408-819">**OR**</span></span><br/><span data-ttu-id="4e408-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="4e408-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4e408-822">String</span><span class="sxs-lookup"><span data-stu-id="4e408-822">String</span></span> | <span data-ttu-id="4e408-823">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-823">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4e408-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4e408-826">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-826">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4e408-827">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-827">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-828">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-828">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4e408-829">String</span><span class="sxs-lookup"><span data-stu-id="4e408-829">String</span></span> | | <span data-ttu-id="4e408-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="4e408-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4e408-832">字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-832">String</span></span> | | <span data-ttu-id="4e408-833">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-833">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4e408-834">String</span><span class="sxs-lookup"><span data-stu-id="4e408-834">String</span></span> | | <span data-ttu-id="4e408-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="4e408-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="4e408-837">布尔</span><span class="sxs-lookup"><span data-stu-id="4e408-837">Boolean</span></span> | | <span data-ttu-id="4e408-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="4e408-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4e408-840">String</span><span class="sxs-lookup"><span data-stu-id="4e408-840">String</span></span> | | <span data-ttu-id="4e408-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4e408-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4e408-844">函数</span><span class="sxs-lookup"><span data-stu-id="4e408-844">function</span></span> | <span data-ttu-id="4e408-845">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-845">&lt;optional&gt;</span></span> | <span data-ttu-id="4e408-846">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4e408-847">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-847">Requirements</span></span>

|<span data-ttu-id="4e408-848">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-848">Requirement</span></span>| <span data-ttu-id="4e408-849">值</span><span class="sxs-lookup"><span data-stu-id="4e408-849">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-850">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-850">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-851">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-851">1.0</span></span>|
|[<span data-ttu-id="4e408-852">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-852">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-853">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-853">ReadItem</span></span>|
|[<span data-ttu-id="4e408-854">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-854">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-855">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-855">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e408-856">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-856">Examples</span></span>

<span data-ttu-id="4e408-857">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-857">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4e408-858">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-858">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4e408-859">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-859">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4e408-860">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-860">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4e408-861">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-861">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4e408-862">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="4e408-862">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="4e408-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="4e408-863">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="4e408-864">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="4e408-864">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-865">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-866">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-866">Requirements</span></span>

|<span data-ttu-id="4e408-867">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-867">Requirement</span></span>| <span data-ttu-id="4e408-868">值</span><span class="sxs-lookup"><span data-stu-id="4e408-868">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-869">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-869">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-870">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-870">1.0</span></span>|
|[<span data-ttu-id="4e408-871">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-871">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-872">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-872">ReadItem</span></span>|
|[<span data-ttu-id="4e408-873">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-873">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-874">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-874">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-875">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-875">Returns:</span></span>

<span data-ttu-id="4e408-876">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-876">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="4e408-877">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-877">Example</span></span>

<span data-ttu-id="4e408-878">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="4e408-878">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="4e408-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="4e408-879">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="4e408-880">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-880">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-881">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-881">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-882">Parameters</span><span class="sxs-lookup"><span data-stu-id="4e408-882">Parameters</span></span>

|<span data-ttu-id="4e408-883">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-883">Name</span></span>| <span data-ttu-id="4e408-884">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-884">Type</span></span>| <span data-ttu-id="4e408-885">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-885">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="4e408-886">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4e408-886">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.6)|<span data-ttu-id="4e408-887">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="4e408-887">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e408-888">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-888">Requirements</span></span>

|<span data-ttu-id="4e408-889">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-889">Requirement</span></span>| <span data-ttu-id="4e408-890">值</span><span class="sxs-lookup"><span data-stu-id="4e408-890">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-891">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-891">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-892">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-892">1.0</span></span>|
|[<span data-ttu-id="4e408-893">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-893">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-894">受限</span><span class="sxs-lookup"><span data-stu-id="4e408-894">Restricted</span></span>|
|[<span data-ttu-id="4e408-895">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-895">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-896">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-896">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-897">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-897">Returns:</span></span>

<span data-ttu-id="4e408-898">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="4e408-898">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4e408-899">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-899">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4e408-900">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="4e408-900">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4e408-901">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="4e408-901">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="4e408-902">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="4e408-902">Value of `entityType`</span></span> | <span data-ttu-id="4e408-903">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="4e408-903">Type of objects in returned array</span></span> | <span data-ttu-id="4e408-904">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-904">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="4e408-905">String</span><span class="sxs-lookup"><span data-stu-id="4e408-905">String</span></span> | <span data-ttu-id="4e408-906">**受限**</span><span class="sxs-lookup"><span data-stu-id="4e408-906">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="4e408-907">Contact</span><span class="sxs-lookup"><span data-stu-id="4e408-907">Contact</span></span> | <span data-ttu-id="4e408-908">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e408-908">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="4e408-909">String</span><span class="sxs-lookup"><span data-stu-id="4e408-909">String</span></span> | <span data-ttu-id="4e408-910">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e408-910">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="4e408-911">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4e408-911">MeetingSuggestion</span></span> | <span data-ttu-id="4e408-912">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e408-912">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="4e408-913">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4e408-913">PhoneNumber</span></span> | <span data-ttu-id="4e408-914">**受限**</span><span class="sxs-lookup"><span data-stu-id="4e408-914">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="4e408-915">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4e408-915">TaskSuggestion</span></span> | <span data-ttu-id="4e408-916">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4e408-916">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="4e408-917">字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-917">String</span></span> | <span data-ttu-id="4e408-918">**受限**</span><span class="sxs-lookup"><span data-stu-id="4e408-918">**Restricted**</span></span> |

<span data-ttu-id="4e408-919">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="4e408-919">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

##### <a name="example"></a><span data-ttu-id="4e408-920">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-920">Example</span></span>

<span data-ttu-id="4e408-921">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-921">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-16meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-16phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-16tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-16"></a><span data-ttu-id="4e408-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span><span class="sxs-lookup"><span data-stu-id="4e408-922">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))>}</span></span>

<span data-ttu-id="4e408-923">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="4e408-923">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-924">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-924">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4e408-925">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="4e408-925">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-926">参数</span><span class="sxs-lookup"><span data-stu-id="4e408-926">Parameters</span></span>

|<span data-ttu-id="4e408-927">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-927">Name</span></span>| <span data-ttu-id="4e408-928">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-928">Type</span></span>| <span data-ttu-id="4e408-929">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-929">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4e408-930">字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-930">String</span></span>|<span data-ttu-id="4e408-931">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="4e408-931">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e408-932">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-932">Requirements</span></span>

|<span data-ttu-id="4e408-933">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-933">Requirement</span></span>| <span data-ttu-id="4e408-934">值</span><span class="sxs-lookup"><span data-stu-id="4e408-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-935">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-936">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-936">1.0</span></span>|
|[<span data-ttu-id="4e408-937">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-938">ReadItem</span></span>|
|[<span data-ttu-id="4e408-939">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-940">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-941">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-941">Returns:</span></span>

<span data-ttu-id="4e408-p162">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4e408-944">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span><span class="sxs-lookup"><span data-stu-id="4e408-944">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.6)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.6)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.6)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.6))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="4e408-945">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4e408-945">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4e408-946">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="4e408-946">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-947">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-947">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4e408-p163">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="4e408-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4e408-951">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="4e408-951">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4e408-952">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="4e408-952">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4e408-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="4e408-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-956">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-956">Requirements</span></span>

|<span data-ttu-id="4e408-957">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-957">Requirement</span></span>| <span data-ttu-id="4e408-958">值</span><span class="sxs-lookup"><span data-stu-id="4e408-958">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-959">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-959">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-960">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-960">1.0</span></span>|
|[<span data-ttu-id="4e408-961">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-961">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-962">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-962">ReadItem</span></span>|
|[<span data-ttu-id="4e408-963">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-963">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-964">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-964">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-965">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-965">Returns:</span></span>

<span data-ttu-id="4e408-p165">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="4e408-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="4e408-968">类型：对象</span><span class="sxs-lookup"><span data-stu-id="4e408-968">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="4e408-969">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-969">Example</span></span>

<span data-ttu-id="4e408-970">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="4e408-970">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4e408-971">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="4e408-971">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4e408-972">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="4e408-972">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-973">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-973">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4e408-974">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="4e408-974">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4e408-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="4e408-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-977">参数</span><span class="sxs-lookup"><span data-stu-id="4e408-977">Parameters</span></span>

|<span data-ttu-id="4e408-978">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-978">Name</span></span>| <span data-ttu-id="4e408-979">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-979">Type</span></span>| <span data-ttu-id="4e408-980">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-980">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4e408-981">字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-981">String</span></span>|<span data-ttu-id="4e408-982">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="4e408-982">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e408-983">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-983">Requirements</span></span>

|<span data-ttu-id="4e408-984">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-984">Requirement</span></span>| <span data-ttu-id="4e408-985">值</span><span class="sxs-lookup"><span data-stu-id="4e408-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-986">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-987">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-987">1.0</span></span>|
|[<span data-ttu-id="4e408-988">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-988">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-989">ReadItem</span></span>|
|[<span data-ttu-id="4e408-990">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-990">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-991">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-991">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-992">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-992">Returns:</span></span>

<span data-ttu-id="4e408-993">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="4e408-993">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="4e408-994">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4e408-994">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="4e408-995">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-995">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4e408-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4e408-996">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4e408-997">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="4e408-997">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4e408-998">如果没有选定内容，但光标在正文或主题中，则该方法将返回所选数据的空字符串。</span><span class="sxs-lookup"><span data-stu-id="4e408-998">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data.</span></span> <span data-ttu-id="4e408-999">如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="4e408-999">If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-1000">在 Outlook 网页版中，如果未选中任何文本，但光标位于正文中，则该方法返回字符串“null”。</span><span class="sxs-lookup"><span data-stu-id="4e408-1000">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="4e408-1001">若要检查此情况，请参阅本节后面的示例。</span><span class="sxs-lookup"><span data-stu-id="4e408-1001">To check for this situation, see the example later in this section.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-1002">参数</span><span class="sxs-lookup"><span data-stu-id="4e408-1002">Parameters</span></span>

|<span data-ttu-id="4e408-1003">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-1003">Name</span></span>| <span data-ttu-id="4e408-1004">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-1004">Type</span></span>| <span data-ttu-id="4e408-1005">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-1005">Attributes</span></span>| <span data-ttu-id="4e408-1006">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-1006">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="4e408-1007">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4e408-1007">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4e408-p169">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="4e408-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="4e408-1011">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1011">Object</span></span>| <span data-ttu-id="4e408-1012">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1012">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1013">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4e408-1013">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4e408-1014">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1014">Object</span></span>| <span data-ttu-id="4e408-1015">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1015">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1016">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-1016">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4e408-1017">function</span><span class="sxs-lookup"><span data-stu-id="4e408-1017">function</span></span>||<span data-ttu-id="4e408-1018">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-1018">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4e408-1019">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="4e408-1019">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4e408-1020">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="4e408-1020">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e408-1021">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-1021">Requirements</span></span>

|<span data-ttu-id="4e408-1022">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1022">Requirement</span></span>| <span data-ttu-id="4e408-1023">值</span><span class="sxs-lookup"><span data-stu-id="4e408-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-1024">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-1025">1.2</span><span class="sxs-lookup"><span data-stu-id="4e408-1025">1.2</span></span>|
|[<span data-ttu-id="4e408-1026">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-1026">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-1027">ReadItem</span></span>|
|[<span data-ttu-id="4e408-1028">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-1028">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-1029">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-1029">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-1030">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-1030">Returns:</span></span>

<span data-ttu-id="4e408-1031">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="4e408-1031">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="4e408-1032">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-1032">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4e408-1033">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-1033">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  // Handle where Outlook on the web erroneously returns "null" instead of empty string.
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookWebApp'
      && asyncResult.value.endPosition === asyncResult.value.startPosition) {
    text = "";
  }

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-16"></a><span data-ttu-id="4e408-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="4e408-1034">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="4e408-1035">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="4e408-1035">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="4e408-1036">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="4e408-1036">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-1037">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-1037">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-1038">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1038">Requirements</span></span>

|<span data-ttu-id="4e408-1039">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1039">Requirement</span></span>| <span data-ttu-id="4e408-1040">值</span><span class="sxs-lookup"><span data-stu-id="4e408-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-1041">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="4e408-1042">1.6</span></span> |
|[<span data-ttu-id="4e408-1043">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-1043">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-1044">ReadItem</span></span>|
|[<span data-ttu-id="4e408-1045">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-1045">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-1046">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-1047">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-1047">Returns:</span></span>

<span data-ttu-id="4e408-1048">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="4e408-1048">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.6)</span></span>

##### <a name="example"></a><span data-ttu-id="4e408-1049">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-1049">Example</span></span>

<span data-ttu-id="4e408-1050">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="4e408-1050">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="4e408-1051">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4e408-1051">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="4e408-p172">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="4e408-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-1054">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-1054">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4e408-p173">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="4e408-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4e408-1058">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="4e408-1058">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4e408-1059">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="4e408-1059">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4e408-p174">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="4e408-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.6#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4e408-1063">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-1063">Requirements</span></span>

|<span data-ttu-id="4e408-1064">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1064">Requirement</span></span>| <span data-ttu-id="4e408-1065">值</span><span class="sxs-lookup"><span data-stu-id="4e408-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-1066">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-1067">1.6</span><span class="sxs-lookup"><span data-stu-id="4e408-1067">1.6</span></span> |
|[<span data-ttu-id="4e408-1068">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-1069">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-1069">ReadItem</span></span>|
|[<span data-ttu-id="4e408-1070">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-1071">阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-1071">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4e408-1072">返回：</span><span class="sxs-lookup"><span data-stu-id="4e408-1072">Returns:</span></span>

<span data-ttu-id="4e408-p175">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="4e408-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="4e408-1075">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-1075">Example</span></span>

<span data-ttu-id="4e408-1076">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="4e408-1076">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4e408-1077">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4e408-1077">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4e408-1078">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="4e408-1078">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4e408-p176">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="4e408-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-1082">参数</span><span class="sxs-lookup"><span data-stu-id="4e408-1082">Parameters</span></span>

|<span data-ttu-id="4e408-1083">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-1083">Name</span></span>| <span data-ttu-id="4e408-1084">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-1084">Type</span></span>| <span data-ttu-id="4e408-1085">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-1085">Attributes</span></span>| <span data-ttu-id="4e408-1086">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-1086">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4e408-1087">函数</span><span class="sxs-lookup"><span data-stu-id="4e408-1087">function</span></span>||<span data-ttu-id="4e408-1088">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-1088">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4e408-1089">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="4e408-1089">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.6) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4e408-1090">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="4e408-1090">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="4e408-1091">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1091">Object</span></span>| <span data-ttu-id="4e408-1092">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1093">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-1093">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4e408-1094">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="4e408-1094">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e408-1095">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-1095">Requirements</span></span>

|<span data-ttu-id="4e408-1096">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1096">Requirement</span></span>| <span data-ttu-id="4e408-1097">值</span><span class="sxs-lookup"><span data-stu-id="4e408-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-1098">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-1099">1.0</span><span class="sxs-lookup"><span data-stu-id="4e408-1099">1.0</span></span>|
|[<span data-ttu-id="4e408-1100">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4e408-1101">ReadItem</span></span>|
|[<span data-ttu-id="4e408-1102">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-1103">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4e408-1103">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-1104">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-1104">Example</span></span>

<span data-ttu-id="4e408-p179">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="4e408-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4e408-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4e408-1108">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4e408-1109">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="4e408-1109">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4e408-1110">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="4e408-1110">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4e408-1111">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="4e408-1111">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="4e408-1112">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="4e408-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4e408-1113">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="4e408-1113">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-1114">Parameters</span><span class="sxs-lookup"><span data-stu-id="4e408-1114">Parameters</span></span>

|<span data-ttu-id="4e408-1115">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-1115">Name</span></span>| <span data-ttu-id="4e408-1116">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-1116">Type</span></span>| <span data-ttu-id="4e408-1117">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-1117">Attributes</span></span>| <span data-ttu-id="4e408-1118">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="4e408-1119">String</span><span class="sxs-lookup"><span data-stu-id="4e408-1119">String</span></span>||<span data-ttu-id="4e408-1120">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="4e408-1120">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="4e408-1121">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1121">Object</span></span>| <span data-ttu-id="4e408-1122">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1123">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4e408-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4e408-1124">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1124">Object</span></span>| <span data-ttu-id="4e408-1125">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1126">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4e408-1127">函数</span><span class="sxs-lookup"><span data-stu-id="4e408-1127">function</span></span>| <span data-ttu-id="4e408-1128">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1129">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4e408-1130">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="4e408-1130">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4e408-1131">错误</span><span class="sxs-lookup"><span data-stu-id="4e408-1131">Errors</span></span>

| <span data-ttu-id="4e408-1132">错误代码</span><span class="sxs-lookup"><span data-stu-id="4e408-1132">Error code</span></span> | <span data-ttu-id="4e408-1133">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-1133">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="4e408-1134">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="4e408-1134">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4e408-1135">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-1135">Requirements</span></span>

|<span data-ttu-id="4e408-1136">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1136">Requirement</span></span>| <span data-ttu-id="4e408-1137">值</span><span class="sxs-lookup"><span data-stu-id="4e408-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-1138">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-1139">1.1</span><span class="sxs-lookup"><span data-stu-id="4e408-1139">1.1</span></span>|
|[<span data-ttu-id="4e408-1140">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-1141">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e408-1141">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e408-1142">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-1143">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-1143">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-1144">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-1144">Example</span></span>

<span data-ttu-id="4e408-1145">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="4e408-1145">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="4e408-1146">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4e408-1146">saveAsync([options], callback)</span></span>

<span data-ttu-id="4e408-1147">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="4e408-1147">Asynchronously saves an item.</span></span>

<span data-ttu-id="4e408-1148">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="4e408-1148">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="4e408-1149">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="4e408-1149">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="4e408-1150">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="4e408-1150">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-1151">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="4e408-1151">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4e408-1152">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="4e408-1152">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4e408-p183">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="4e408-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4e408-1156">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="4e408-1156">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4e408-1157">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="4e408-1157">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="4e408-1158">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="4e408-1158">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="4e408-1159">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="4e408-1159">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="4e408-1160">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="4e408-1160">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-1161">参数</span><span class="sxs-lookup"><span data-stu-id="4e408-1161">Parameters</span></span>

|<span data-ttu-id="4e408-1162">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-1162">Name</span></span>| <span data-ttu-id="4e408-1163">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-1163">Type</span></span>| <span data-ttu-id="4e408-1164">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-1164">Attributes</span></span>| <span data-ttu-id="4e408-1165">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-1165">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="4e408-1166">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1166">Object</span></span>| <span data-ttu-id="4e408-1167">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1167">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1168">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4e408-1168">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4e408-1169">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1169">Object</span></span>| <span data-ttu-id="4e408-1170">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1170">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1171">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-1171">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4e408-1172">函数</span><span class="sxs-lookup"><span data-stu-id="4e408-1172">function</span></span>||<span data-ttu-id="4e408-1173">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-1173">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4e408-1174">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4e408-1174">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4e408-1175">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1175">Requirements</span></span>

|<span data-ttu-id="4e408-1176">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1176">Requirement</span></span>| <span data-ttu-id="4e408-1177">值</span><span class="sxs-lookup"><span data-stu-id="4e408-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-1178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-1179">1.3</span><span class="sxs-lookup"><span data-stu-id="4e408-1179">1.3</span></span>|
|[<span data-ttu-id="4e408-1180">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-1181">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e408-1181">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e408-1182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-1183">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-1183">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4e408-1184">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-1184">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4e408-p185">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="4e408-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4e408-1187">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4e408-1187">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4e408-1188">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="4e408-1188">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4e408-p186">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="4e408-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4e408-1192">参数</span><span class="sxs-lookup"><span data-stu-id="4e408-1192">Parameters</span></span>

|<span data-ttu-id="4e408-1193">名称</span><span class="sxs-lookup"><span data-stu-id="4e408-1193">Name</span></span>| <span data-ttu-id="4e408-1194">类型</span><span class="sxs-lookup"><span data-stu-id="4e408-1194">Type</span></span>| <span data-ttu-id="4e408-1195">属性</span><span class="sxs-lookup"><span data-stu-id="4e408-1195">Attributes</span></span>| <span data-ttu-id="4e408-1196">说明</span><span class="sxs-lookup"><span data-stu-id="4e408-1196">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4e408-1197">字符串</span><span class="sxs-lookup"><span data-stu-id="4e408-1197">String</span></span>||<span data-ttu-id="4e408-p187">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="4e408-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="4e408-1201">Object</span><span class="sxs-lookup"><span data-stu-id="4e408-1201">Object</span></span>| <span data-ttu-id="4e408-1202">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1203">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4e408-1203">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4e408-1204">对象</span><span class="sxs-lookup"><span data-stu-id="4e408-1204">Object</span></span>| <span data-ttu-id="4e408-1205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1206">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4e408-1206">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4e408-1207">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4e408-1207">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4e408-1208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4e408-1208">&lt;optional&gt;</span></span>|<span data-ttu-id="4e408-1209">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="4e408-1209">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="4e408-1210">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="4e408-1210">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4e408-1211">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="4e408-1211">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="4e408-1212">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="4e408-1212">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4e408-1213">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="4e408-1213">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="4e408-1214">function</span><span class="sxs-lookup"><span data-stu-id="4e408-1214">function</span></span>||<span data-ttu-id="4e408-1215">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4e408-1215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4e408-1216">Requirements</span><span class="sxs-lookup"><span data-stu-id="4e408-1216">Requirements</span></span>

|<span data-ttu-id="4e408-1217">要求</span><span class="sxs-lookup"><span data-stu-id="4e408-1217">Requirement</span></span>| <span data-ttu-id="4e408-1218">值</span><span class="sxs-lookup"><span data-stu-id="4e408-1218">Value</span></span>|
|---|---|
|[<span data-ttu-id="4e408-1219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4e408-1219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4e408-1220">1.2</span><span class="sxs-lookup"><span data-stu-id="4e408-1220">1.2</span></span>|
|[<span data-ttu-id="4e408-1221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4e408-1221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4e408-1222">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4e408-1222">ReadWriteItem</span></span>|
|[<span data-ttu-id="4e408-1223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4e408-1223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4e408-1224">撰写</span><span class="sxs-lookup"><span data-stu-id="4e408-1224">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4e408-1225">示例</span><span class="sxs-lookup"><span data-stu-id="4e408-1225">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
