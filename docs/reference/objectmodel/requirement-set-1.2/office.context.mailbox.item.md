---
title: "\"Context\"-\"邮箱\"。项目-要求集1。2"
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 536c8b7bece6df6f9609406f3eccc50b330d7925
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268689"
---
# <a name="item"></a><span data-ttu-id="6d676-102">item</span><span class="sxs-lookup"><span data-stu-id="6d676-102">item</span></span>

### <span data-ttu-id="6d676-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="6d676-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="6d676-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="6d676-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="6d676-107">Requirements</span></span>

|<span data-ttu-id="6d676-108">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-108">Requirement</span></span>| <span data-ttu-id="6d676-109">值</span><span class="sxs-lookup"><span data-stu-id="6d676-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-111">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-111">1.0</span></span>|
|[<span data-ttu-id="6d676-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-113">受限</span><span class="sxs-lookup"><span data-stu-id="6d676-113">Restricted</span></span>|
|[<span data-ttu-id="6d676-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6d676-116">成员和方法</span><span class="sxs-lookup"><span data-stu-id="6d676-116">Members and methods</span></span>

| <span data-ttu-id="6d676-117">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-117">Member</span></span> | <span data-ttu-id="6d676-118">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6d676-119">attachments</span><span class="sxs-lookup"><span data-stu-id="6d676-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="6d676-120">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-120">Member</span></span> |
| [<span data-ttu-id="6d676-121">bcc</span><span class="sxs-lookup"><span data-stu-id="6d676-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="6d676-122">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-122">Member</span></span> |
| [<span data-ttu-id="6d676-123">body</span><span class="sxs-lookup"><span data-stu-id="6d676-123">body</span></span>](#body-body) | <span data-ttu-id="6d676-124">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-124">Member</span></span> |
| [<span data-ttu-id="6d676-125">cc</span><span class="sxs-lookup"><span data-stu-id="6d676-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6d676-126">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-126">Member</span></span> |
| [<span data-ttu-id="6d676-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="6d676-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="6d676-128">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-128">Member</span></span> |
| [<span data-ttu-id="6d676-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="6d676-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="6d676-130">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-130">Member</span></span> |
| [<span data-ttu-id="6d676-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="6d676-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="6d676-132">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-132">Member</span></span> |
| [<span data-ttu-id="6d676-133">end</span><span class="sxs-lookup"><span data-stu-id="6d676-133">end</span></span>](#end-datetime) | <span data-ttu-id="6d676-134">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-134">Member</span></span> |
| [<span data-ttu-id="6d676-135">from</span><span class="sxs-lookup"><span data-stu-id="6d676-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="6d676-136">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-136">Member</span></span> |
| [<span data-ttu-id="6d676-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="6d676-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="6d676-138">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-138">Member</span></span> |
| [<span data-ttu-id="6d676-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="6d676-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="6d676-140">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-140">Member</span></span> |
| [<span data-ttu-id="6d676-141">itemId</span><span class="sxs-lookup"><span data-stu-id="6d676-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="6d676-142">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-142">Member</span></span> |
| [<span data-ttu-id="6d676-143">itemType</span><span class="sxs-lookup"><span data-stu-id="6d676-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="6d676-144">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-144">Member</span></span> |
| [<span data-ttu-id="6d676-145">location</span><span class="sxs-lookup"><span data-stu-id="6d676-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="6d676-146">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-146">Member</span></span> |
| [<span data-ttu-id="6d676-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="6d676-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="6d676-148">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-148">Member</span></span> |
| [<span data-ttu-id="6d676-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="6d676-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6d676-150">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-150">Member</span></span> |
| [<span data-ttu-id="6d676-151">organizer</span><span class="sxs-lookup"><span data-stu-id="6d676-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="6d676-152">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-152">Member</span></span> |
| [<span data-ttu-id="6d676-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="6d676-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6d676-154">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-154">Member</span></span> |
| [<span data-ttu-id="6d676-155">sender</span><span class="sxs-lookup"><span data-stu-id="6d676-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="6d676-156">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-156">Member</span></span> |
| [<span data-ttu-id="6d676-157">start</span><span class="sxs-lookup"><span data-stu-id="6d676-157">start</span></span>](#start-datetime) | <span data-ttu-id="6d676-158">Member</span><span class="sxs-lookup"><span data-stu-id="6d676-158">Member</span></span> |
| [<span data-ttu-id="6d676-159">subject</span><span class="sxs-lookup"><span data-stu-id="6d676-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="6d676-160">成员</span><span class="sxs-lookup"><span data-stu-id="6d676-160">Member</span></span> |
| [<span data-ttu-id="6d676-161">to</span><span class="sxs-lookup"><span data-stu-id="6d676-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="6d676-162">成员</span><span class="sxs-lookup"><span data-stu-id="6d676-162">Member</span></span> |
| [<span data-ttu-id="6d676-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6d676-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="6d676-164">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-164">Method</span></span> |
| [<span data-ttu-id="6d676-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6d676-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="6d676-166">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-166">Method</span></span> |
| [<span data-ttu-id="6d676-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="6d676-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="6d676-168">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-168">Method</span></span> |
| [<span data-ttu-id="6d676-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="6d676-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="6d676-170">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-170">Method</span></span> |
| [<span data-ttu-id="6d676-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="6d676-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="6d676-172">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-172">Method</span></span> |
| [<span data-ttu-id="6d676-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="6d676-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="6d676-174">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-174">Method</span></span> |
| [<span data-ttu-id="6d676-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="6d676-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="6d676-176">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-176">Method</span></span> |
| [<span data-ttu-id="6d676-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="6d676-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="6d676-178">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-178">Method</span></span> |
| [<span data-ttu-id="6d676-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="6d676-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="6d676-180">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-180">Method</span></span> |
| [<span data-ttu-id="6d676-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6d676-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="6d676-182">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-182">Method</span></span> |
| [<span data-ttu-id="6d676-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="6d676-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="6d676-184">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-184">Method</span></span> |
| [<span data-ttu-id="6d676-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="6d676-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="6d676-186">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-186">Method</span></span> |
| [<span data-ttu-id="6d676-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="6d676-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="6d676-188">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="6d676-189">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-189">Example</span></span>

<span data-ttu-id="6d676-190">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="6d676-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="6d676-191">成员</span><span class="sxs-lookup"><span data-stu-id="6d676-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="6d676-192">附件: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="6d676-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="6d676-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-195">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="6d676-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="6d676-196">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="6d676-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-197">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-197">Type</span></span>

*   <span data-ttu-id="6d676-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="6d676-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-199">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-199">Requirements</span></span>

|<span data-ttu-id="6d676-200">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-200">Requirement</span></span>| <span data-ttu-id="6d676-201">值</span><span class="sxs-lookup"><span data-stu-id="6d676-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-202">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-203">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-203">1.0</span></span>|
|[<span data-ttu-id="6d676-204">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-205">ReadItem</span></span>|
|[<span data-ttu-id="6d676-206">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-207">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-208">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-208">Example</span></span>

<span data-ttu-id="6d676-209">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="6d676-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="6d676-210">密件抄送:[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-211">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="6d676-212">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-213">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-213">Type</span></span>

*   [<span data-ttu-id="6d676-214">收件人</span><span class="sxs-lookup"><span data-stu-id="6d676-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="6d676-215">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-215">Requirements</span></span>

|<span data-ttu-id="6d676-216">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-216">Requirement</span></span>| <span data-ttu-id="6d676-217">值</span><span class="sxs-lookup"><span data-stu-id="6d676-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-219">1.1</span><span class="sxs-lookup"><span data-stu-id="6d676-219">1.1</span></span>|
|[<span data-ttu-id="6d676-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-221">ReadItem</span></span>|
|[<span data-ttu-id="6d676-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-223">撰写</span><span class="sxs-lookup"><span data-stu-id="6d676-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-224">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-224">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="6d676-225">正文:[正文](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-226">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-227">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-227">Type</span></span>

*   [<span data-ttu-id="6d676-228">Body</span><span class="sxs-lookup"><span data-stu-id="6d676-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="6d676-229">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-229">Requirements</span></span>

|<span data-ttu-id="6d676-230">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-230">Requirement</span></span>| <span data-ttu-id="6d676-231">值</span><span class="sxs-lookup"><span data-stu-id="6d676-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-232">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-233">1.1</span><span class="sxs-lookup"><span data-stu-id="6d676-233">1.1</span></span>|
|[<span data-ttu-id="6d676-234">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-235">ReadItem</span></span>|
|[<span data-ttu-id="6d676-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-238">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-238">Example</span></span>

<span data-ttu-id="6d676-239">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="6d676-239">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="6d676-240">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="6d676-240">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="6d676-241"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="6d676-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-242">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="6d676-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="6d676-243">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-244">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-244">Read mode</span></span>

<span data-ttu-id="6d676-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="6d676-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-247">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-247">Compose mode</span></span>

<span data-ttu-id="6d676-248">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6d676-249">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-249">Type</span></span>

*   <span data-ttu-id="6d676-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-251">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-251">Requirements</span></span>

|<span data-ttu-id="6d676-252">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-252">Requirement</span></span>| <span data-ttu-id="6d676-253">值</span><span class="sxs-lookup"><span data-stu-id="6d676-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-254">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-255">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-255">1.0</span></span>|
|[<span data-ttu-id="6d676-256">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-257">ReadItem</span></span>|
|[<span data-ttu-id="6d676-258">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-259">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-259">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="6d676-260">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="6d676-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="6d676-261">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="6d676-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="6d676-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="6d676-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="6d676-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="6d676-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-266">Type</span><span class="sxs-lookup"><span data-stu-id="6d676-266">Type</span></span>

*   <span data-ttu-id="6d676-267">String</span><span class="sxs-lookup"><span data-stu-id="6d676-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-268">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-268">Requirements</span></span>

|<span data-ttu-id="6d676-269">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-269">Requirement</span></span>| <span data-ttu-id="6d676-270">值</span><span class="sxs-lookup"><span data-stu-id="6d676-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-272">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-272">1.0</span></span>|
|[<span data-ttu-id="6d676-273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-274">ReadItem</span></span>|
|[<span data-ttu-id="6d676-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-277">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-277">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="6d676-278">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="6d676-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="6d676-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-281">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-281">Type</span></span>

*   <span data-ttu-id="6d676-282">日期</span><span class="sxs-lookup"><span data-stu-id="6d676-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-283">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-283">Requirements</span></span>

|<span data-ttu-id="6d676-284">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-284">Requirement</span></span>| <span data-ttu-id="6d676-285">值</span><span class="sxs-lookup"><span data-stu-id="6d676-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-286">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-287">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-287">1.0</span></span>|
|[<span data-ttu-id="6d676-288">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-289">ReadItem</span></span>|
|[<span data-ttu-id="6d676-290">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-291">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-292">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-292">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="6d676-293">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="6d676-293">dateTimeModified: Date</span></span>

<span data-ttu-id="6d676-294">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="6d676-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="6d676-295">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-296">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="6d676-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-297">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-297">Type</span></span>

*   <span data-ttu-id="6d676-298">日期</span><span class="sxs-lookup"><span data-stu-id="6d676-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-299">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-299">Requirements</span></span>

|<span data-ttu-id="6d676-300">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-300">Requirement</span></span>| <span data-ttu-id="6d676-301">值</span><span class="sxs-lookup"><span data-stu-id="6d676-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-303">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-303">1.0</span></span>|
|[<span data-ttu-id="6d676-304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-305">ReadItem</span></span>|
|[<span data-ttu-id="6d676-306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-307">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-308">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-308">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="6d676-309">结束: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-310">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="6d676-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="6d676-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="6d676-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-313">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-313">Read mode</span></span>

<span data-ttu-id="6d676-314">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-314">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-315">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-315">Compose mode</span></span>

<span data-ttu-id="6d676-316">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="6d676-317">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="6d676-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="6d676-318">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="6d676-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="6d676-319">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-319">Type</span></span>

*   <span data-ttu-id="6d676-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-321">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-321">Requirements</span></span>

|<span data-ttu-id="6d676-322">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-322">Requirement</span></span>| <span data-ttu-id="6d676-323">值</span><span class="sxs-lookup"><span data-stu-id="6d676-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-324">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-325">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-325">1.0</span></span>|
|[<span data-ttu-id="6d676-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-327">ReadItem</span></span>|
|[<span data-ttu-id="6d676-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-329">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-329">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="6d676-330">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="6d676-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="6d676-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-335">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="6d676-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-336">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-336">Type</span></span>

*   [<span data-ttu-id="6d676-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6d676-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="6d676-338">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-338">Requirements</span></span>

|<span data-ttu-id="6d676-339">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-339">Requirement</span></span>| <span data-ttu-id="6d676-340">值</span><span class="sxs-lookup"><span data-stu-id="6d676-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-342">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-342">1.0</span></span>|
|[<span data-ttu-id="6d676-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-344">ReadItem</span></span>|
|[<span data-ttu-id="6d676-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-346">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-347">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-347">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="6d676-348">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="6d676-348">internetMessageId: String</span></span>

<span data-ttu-id="6d676-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-351">Type</span><span class="sxs-lookup"><span data-stu-id="6d676-351">Type</span></span>

*   <span data-ttu-id="6d676-352">String</span><span class="sxs-lookup"><span data-stu-id="6d676-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-353">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-353">Requirements</span></span>

|<span data-ttu-id="6d676-354">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-354">Requirement</span></span>| <span data-ttu-id="6d676-355">值</span><span class="sxs-lookup"><span data-stu-id="6d676-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-356">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-357">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-357">1.0</span></span>|
|[<span data-ttu-id="6d676-358">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-359">ReadItem</span></span>|
|[<span data-ttu-id="6d676-360">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-361">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-362">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-362">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="6d676-363">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="6d676-363">itemClass: String</span></span>

<span data-ttu-id="6d676-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="6d676-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="6d676-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="6d676-368">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-368">Type</span></span> | <span data-ttu-id="6d676-369">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-369">Description</span></span> | <span data-ttu-id="6d676-370">项目类</span><span class="sxs-lookup"><span data-stu-id="6d676-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="6d676-371">约会项目</span><span class="sxs-lookup"><span data-stu-id="6d676-371">Appointment items</span></span> | <span data-ttu-id="6d676-372">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="6d676-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="6d676-373">邮件项目</span><span class="sxs-lookup"><span data-stu-id="6d676-373">Message items</span></span> | <span data-ttu-id="6d676-374">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="6d676-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="6d676-375">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="6d676-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-376">Type</span><span class="sxs-lookup"><span data-stu-id="6d676-376">Type</span></span>

*   <span data-ttu-id="6d676-377">String</span><span class="sxs-lookup"><span data-stu-id="6d676-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-378">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-378">Requirements</span></span>

|<span data-ttu-id="6d676-379">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-379">Requirement</span></span>| <span data-ttu-id="6d676-380">值</span><span class="sxs-lookup"><span data-stu-id="6d676-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-381">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-382">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-382">1.0</span></span>|
|[<span data-ttu-id="6d676-383">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-384">ReadItem</span></span>|
|[<span data-ttu-id="6d676-385">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-386">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-387">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-387">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="6d676-388">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="6d676-388">(nullable) itemId: String</span></span>

<span data-ttu-id="6d676-389">获取当前项目的 Exchange Web 服务项目标识符。</span><span class="sxs-lookup"><span data-stu-id="6d676-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="6d676-390">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-391">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="6d676-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="6d676-392">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="6d676-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="6d676-393">在使用此值进行 REST API 调用之前, 应使用`Office.context.mailbox.convertToRestId`转换它, 这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="6d676-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="6d676-394">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="6d676-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-395">Type</span><span class="sxs-lookup"><span data-stu-id="6d676-395">Type</span></span>

*   <span data-ttu-id="6d676-396">String</span><span class="sxs-lookup"><span data-stu-id="6d676-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-397">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-397">Requirements</span></span>

|<span data-ttu-id="6d676-398">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-398">Requirement</span></span>| <span data-ttu-id="6d676-399">值</span><span class="sxs-lookup"><span data-stu-id="6d676-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-400">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-401">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-401">1.0</span></span>|
|[<span data-ttu-id="6d676-402">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-403">ReadItem</span></span>|
|[<span data-ttu-id="6d676-404">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-405">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-406">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-406">Example</span></span>

<span data-ttu-id="6d676-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="6d676-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="6d676-409">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-410">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="6d676-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="6d676-411">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="6d676-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-412">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-412">Type</span></span>

*   [<span data-ttu-id="6d676-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="6d676-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="6d676-414">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-414">Requirements</span></span>

|<span data-ttu-id="6d676-415">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-415">Requirement</span></span>| <span data-ttu-id="6d676-416">值</span><span class="sxs-lookup"><span data-stu-id="6d676-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-417">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-418">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-418">1.0</span></span>|
|[<span data-ttu-id="6d676-419">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-420">ReadItem</span></span>|
|[<span data-ttu-id="6d676-421">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-422">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-423">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-423">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="6d676-424">位置: 字符串 |[位置](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-425">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="6d676-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-426">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-426">Read mode</span></span>

<span data-ttu-id="6d676-427">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="6d676-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-428">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-428">Compose mode</span></span>

<span data-ttu-id="6d676-429">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6d676-430">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-430">Type</span></span>

*   <span data-ttu-id="6d676-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-432">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-432">Requirements</span></span>

|<span data-ttu-id="6d676-433">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-433">Requirement</span></span>| <span data-ttu-id="6d676-434">值</span><span class="sxs-lookup"><span data-stu-id="6d676-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-435">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-436">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-436">1.0</span></span>|
|[<span data-ttu-id="6d676-437">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-438">ReadItem</span></span>|
|[<span data-ttu-id="6d676-439">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-440">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-440">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="6d676-441">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="6d676-441">normalizedSubject: String</span></span>

<span data-ttu-id="6d676-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="6d676-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="6d676-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-446">Type</span><span class="sxs-lookup"><span data-stu-id="6d676-446">Type</span></span>

*   <span data-ttu-id="6d676-447">String</span><span class="sxs-lookup"><span data-stu-id="6d676-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-448">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-448">Requirements</span></span>

|<span data-ttu-id="6d676-449">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-449">Requirement</span></span>| <span data-ttu-id="6d676-450">值</span><span class="sxs-lookup"><span data-stu-id="6d676-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-451">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-452">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-452">1.0</span></span>|
|[<span data-ttu-id="6d676-453">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-454">ReadItem</span></span>|
|[<span data-ttu-id="6d676-455">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-456">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-457">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-457">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="6d676-458">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="6d676-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-459">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="6d676-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="6d676-460">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-461">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-461">Read mode</span></span>

<span data-ttu-id="6d676-462">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-463">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-463">Compose mode</span></span>

<span data-ttu-id="6d676-464">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6d676-465">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-465">Type</span></span>

*   <span data-ttu-id="6d676-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-467">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-467">Requirements</span></span>

|<span data-ttu-id="6d676-468">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-468">Requirement</span></span>| <span data-ttu-id="6d676-469">值</span><span class="sxs-lookup"><span data-stu-id="6d676-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-470">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-471">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-471">1.0</span></span>|
|[<span data-ttu-id="6d676-472">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-473">ReadItem</span></span>|
|[<span data-ttu-id="6d676-474">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-475">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-475">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="6d676-476">组织者: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-479">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-479">Type</span></span>

*   [<span data-ttu-id="6d676-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6d676-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="6d676-481">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-481">Requirements</span></span>

|<span data-ttu-id="6d676-482">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-482">Requirement</span></span>| <span data-ttu-id="6d676-483">值</span><span class="sxs-lookup"><span data-stu-id="6d676-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-485">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-485">1.0</span></span>|
|[<span data-ttu-id="6d676-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-487">ReadItem</span></span>|
|[<span data-ttu-id="6d676-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-489">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-490">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-490">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="6d676-491">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="6d676-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-492">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="6d676-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="6d676-493">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-494">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-494">Read mode</span></span>

<span data-ttu-id="6d676-495">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-496">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-496">Compose mode</span></span>

<span data-ttu-id="6d676-497">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="6d676-498">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-498">Type</span></span>

*   <span data-ttu-id="6d676-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-500">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-500">Requirements</span></span>

|<span data-ttu-id="6d676-501">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-501">Requirement</span></span>| <span data-ttu-id="6d676-502">值</span><span class="sxs-lookup"><span data-stu-id="6d676-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-504">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-504">1.0</span></span>|
|[<span data-ttu-id="6d676-505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-506">ReadItem</span></span>|
|[<span data-ttu-id="6d676-507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-508">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-508">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="6d676-509">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="6d676-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="6d676-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-514">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="6d676-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="6d676-515">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-515">Type</span></span>

*   [<span data-ttu-id="6d676-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="6d676-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="6d676-517">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-517">Requirements</span></span>

|<span data-ttu-id="6d676-518">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-518">Requirement</span></span>| <span data-ttu-id="6d676-519">值</span><span class="sxs-lookup"><span data-stu-id="6d676-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-520">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-521">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-521">1.0</span></span>|
|[<span data-ttu-id="6d676-522">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-523">ReadItem</span></span>|
|[<span data-ttu-id="6d676-524">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-525">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-526">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-526">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="6d676-527">开始日期: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-528">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="6d676-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="6d676-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="6d676-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-531">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-531">Read mode</span></span>

<span data-ttu-id="6d676-532">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-532">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-533">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-533">Compose mode</span></span>

<span data-ttu-id="6d676-534">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="6d676-535">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="6d676-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="6d676-536">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="6d676-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="6d676-537">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-537">Type</span></span>

*   <span data-ttu-id="6d676-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-539">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-539">Requirements</span></span>

|<span data-ttu-id="6d676-540">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-540">Requirement</span></span>| <span data-ttu-id="6d676-541">值</span><span class="sxs-lookup"><span data-stu-id="6d676-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-542">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-543">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-543">1.0</span></span>|
|[<span data-ttu-id="6d676-544">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-545">ReadItem</span></span>|
|[<span data-ttu-id="6d676-546">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-547">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-547">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="6d676-548">subject: String |[主题](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-549">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="6d676-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="6d676-550">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="6d676-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-551">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-551">Read mode</span></span>

<span data-ttu-id="6d676-p130">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="6d676-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-554">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-554">Compose mode</span></span>

<span data-ttu-id="6d676-555">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="6d676-556">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-556">Type</span></span>

*   <span data-ttu-id="6d676-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-558">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-558">Requirements</span></span>

|<span data-ttu-id="6d676-559">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-559">Requirement</span></span>| <span data-ttu-id="6d676-560">值</span><span class="sxs-lookup"><span data-stu-id="6d676-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-561">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-562">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-562">1.0</span></span>|
|[<span data-ttu-id="6d676-563">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-564">ReadItem</span></span>|
|[<span data-ttu-id="6d676-565">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-566">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-566">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="6d676-567">to: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="6d676-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="6d676-568">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="6d676-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="6d676-569">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="6d676-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="6d676-570">阅读模式</span><span class="sxs-lookup"><span data-stu-id="6d676-570">Read mode</span></span>

<span data-ttu-id="6d676-p132">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="6d676-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="6d676-573">撰写模式</span><span class="sxs-lookup"><span data-stu-id="6d676-573">Compose mode</span></span>

<span data-ttu-id="6d676-574">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="6d676-575">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-575">Type</span></span>

*   <span data-ttu-id="6d676-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-577">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-577">Requirements</span></span>

|<span data-ttu-id="6d676-578">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-578">Requirement</span></span>| <span data-ttu-id="6d676-579">值</span><span class="sxs-lookup"><span data-stu-id="6d676-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-580">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-581">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-581">1.0</span></span>|
|[<span data-ttu-id="6d676-582">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-583">ReadItem</span></span>|
|[<span data-ttu-id="6d676-584">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-585">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="6d676-586">方法</span><span class="sxs-lookup"><span data-stu-id="6d676-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="6d676-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6d676-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6d676-588">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="6d676-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="6d676-589">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="6d676-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="6d676-590">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="6d676-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-591">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-591">Parameters</span></span>

|<span data-ttu-id="6d676-592">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-592">Name</span></span>| <span data-ttu-id="6d676-593">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-593">Type</span></span>| <span data-ttu-id="6d676-594">属性</span><span class="sxs-lookup"><span data-stu-id="6d676-594">Attributes</span></span>| <span data-ttu-id="6d676-595">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="6d676-596">String</span><span class="sxs-lookup"><span data-stu-id="6d676-596">String</span></span>||<span data-ttu-id="6d676-p133">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6d676-599">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-599">String</span></span>||<span data-ttu-id="6d676-p134">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6d676-602">Object</span><span class="sxs-lookup"><span data-stu-id="6d676-602">Object</span></span>| <span data-ttu-id="6d676-603">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-603">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-604">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="6d676-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6d676-605">对象</span><span class="sxs-lookup"><span data-stu-id="6d676-605">Object</span></span>| <span data-ttu-id="6d676-606">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-606">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-607">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6d676-608">函数</span><span class="sxs-lookup"><span data-stu-id="6d676-608">function</span></span>| <span data-ttu-id="6d676-609">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-609">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-610">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6d676-611">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="6d676-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6d676-612">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6d676-613">错误</span><span class="sxs-lookup"><span data-stu-id="6d676-613">Errors</span></span>

| <span data-ttu-id="6d676-614">错误代码</span><span class="sxs-lookup"><span data-stu-id="6d676-614">Error code</span></span> | <span data-ttu-id="6d676-615">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="6d676-616">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="6d676-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="6d676-617">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="6d676-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6d676-618">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="6d676-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6d676-619">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-619">Requirements</span></span>

|<span data-ttu-id="6d676-620">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-620">Requirement</span></span>| <span data-ttu-id="6d676-621">值</span><span class="sxs-lookup"><span data-stu-id="6d676-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-622">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-623">1.1</span><span class="sxs-lookup"><span data-stu-id="6d676-623">1.1</span></span>|
|[<span data-ttu-id="6d676-624">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6d676-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="6d676-626">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-627">撰写</span><span class="sxs-lookup"><span data-stu-id="6d676-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-628">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-628">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="6d676-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6d676-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="6d676-630">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="6d676-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="6d676-p135">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="6d676-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="6d676-634">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="6d676-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="6d676-635">如果 Office 外接程序在 web 上的 Outlook 中运行, 则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是, 不支持这种情况, 建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="6d676-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-636">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-636">Parameters</span></span>

|<span data-ttu-id="6d676-637">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-637">Name</span></span>| <span data-ttu-id="6d676-638">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-638">Type</span></span>| <span data-ttu-id="6d676-639">属性</span><span class="sxs-lookup"><span data-stu-id="6d676-639">Attributes</span></span>| <span data-ttu-id="6d676-640">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="6d676-641">String</span><span class="sxs-lookup"><span data-stu-id="6d676-641">String</span></span>||<span data-ttu-id="6d676-p136">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="6d676-644">String</span><span class="sxs-lookup"><span data-stu-id="6d676-644">String</span></span>||<span data-ttu-id="6d676-645">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="6d676-645">The subject of the item to be attached.</span></span> <span data-ttu-id="6d676-646">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="6d676-647">对象</span><span class="sxs-lookup"><span data-stu-id="6d676-647">Object</span></span>| <span data-ttu-id="6d676-648">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-648">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-649">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="6d676-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6d676-650">对象</span><span class="sxs-lookup"><span data-stu-id="6d676-650">Object</span></span>| <span data-ttu-id="6d676-651">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-651">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-652">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6d676-653">函数</span><span class="sxs-lookup"><span data-stu-id="6d676-653">function</span></span>| <span data-ttu-id="6d676-654">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-654">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-655">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6d676-656">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="6d676-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="6d676-657">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6d676-658">错误</span><span class="sxs-lookup"><span data-stu-id="6d676-658">Errors</span></span>

| <span data-ttu-id="6d676-659">错误代码</span><span class="sxs-lookup"><span data-stu-id="6d676-659">Error code</span></span> | <span data-ttu-id="6d676-660">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="6d676-661">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="6d676-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6d676-662">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-662">Requirements</span></span>

|<span data-ttu-id="6d676-663">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-663">Requirement</span></span>| <span data-ttu-id="6d676-664">值</span><span class="sxs-lookup"><span data-stu-id="6d676-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-665">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-666">1.1</span><span class="sxs-lookup"><span data-stu-id="6d676-666">1.1</span></span>|
|[<span data-ttu-id="6d676-667">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6d676-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="6d676-669">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-670">撰写</span><span class="sxs-lookup"><span data-stu-id="6d676-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-671">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-671">Example</span></span>

<span data-ttu-id="6d676-672">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="6d676-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="6d676-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="6d676-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="6d676-674">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="6d676-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-675">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6d676-676">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="6d676-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6d676-677">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="6d676-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="6d676-678">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="6d676-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="6d676-679">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="6d676-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="6d676-680">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="6d676-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-681">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-681">Parameters</span></span>

|<span data-ttu-id="6d676-682">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-682">Name</span></span>| <span data-ttu-id="6d676-683">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-683">Type</span></span>| <span data-ttu-id="6d676-684">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6d676-685">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="6d676-685">String &#124; Object</span></span>| |<span data-ttu-id="6d676-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="6d676-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6d676-688">**或**</span><span class="sxs-lookup"><span data-stu-id="6d676-688">**OR**</span></span><br/><span data-ttu-id="6d676-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="6d676-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6d676-691">String</span><span class="sxs-lookup"><span data-stu-id="6d676-691">String</span></span> | <span data-ttu-id="6d676-692">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-692">&lt;optional&gt;</span></span> | <span data-ttu-id="6d676-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="6d676-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6d676-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6d676-696">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-696">&lt;optional&gt;</span></span> | <span data-ttu-id="6d676-697">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="6d676-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6d676-698">String</span><span class="sxs-lookup"><span data-stu-id="6d676-698">String</span></span> | | <span data-ttu-id="6d676-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="6d676-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6d676-701">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-701">String</span></span> | | <span data-ttu-id="6d676-702">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6d676-703">String</span><span class="sxs-lookup"><span data-stu-id="6d676-703">String</span></span> | | <span data-ttu-id="6d676-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="6d676-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6d676-706">String</span><span class="sxs-lookup"><span data-stu-id="6d676-706">String</span></span> | | <span data-ttu-id="6d676-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6d676-710">函数</span><span class="sxs-lookup"><span data-stu-id="6d676-710">function</span></span> | <span data-ttu-id="6d676-711">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-711">&lt;optional&gt;</span></span> | <span data-ttu-id="6d676-712">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6d676-713">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-713">Requirements</span></span>

|<span data-ttu-id="6d676-714">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-714">Requirement</span></span>| <span data-ttu-id="6d676-715">值</span><span class="sxs-lookup"><span data-stu-id="6d676-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-716">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-717">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-717">1.0</span></span>|
|[<span data-ttu-id="6d676-718">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-719">ReadItem</span></span>|
|[<span data-ttu-id="6d676-720">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-721">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6d676-722">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-722">Examples</span></span>

<span data-ttu-id="6d676-723">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="6d676-724">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-724">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="6d676-725">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-725">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6d676-726">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6d676-727">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6d676-728">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="6d676-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="6d676-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="6d676-730">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="6d676-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-731">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6d676-732">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="6d676-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="6d676-733">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="6d676-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="6d676-734">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="6d676-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="6d676-735">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="6d676-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="6d676-736">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="6d676-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-737">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-737">Parameters</span></span>

|<span data-ttu-id="6d676-738">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-738">Name</span></span>| <span data-ttu-id="6d676-739">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-739">Type</span></span>| <span data-ttu-id="6d676-740">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="6d676-741">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="6d676-741">String &#124; Object</span></span>| | <span data-ttu-id="6d676-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="6d676-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="6d676-744">**或**</span><span class="sxs-lookup"><span data-stu-id="6d676-744">**OR**</span></span><br/><span data-ttu-id="6d676-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="6d676-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="6d676-747">String</span><span class="sxs-lookup"><span data-stu-id="6d676-747">String</span></span> | <span data-ttu-id="6d676-748">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-748">&lt;optional&gt;</span></span> | <span data-ttu-id="6d676-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="6d676-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="6d676-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="6d676-752">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-752">&lt;optional&gt;</span></span> | <span data-ttu-id="6d676-753">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="6d676-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="6d676-754">String</span><span class="sxs-lookup"><span data-stu-id="6d676-754">String</span></span> | | <span data-ttu-id="6d676-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="6d676-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="6d676-757">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-757">String</span></span> | | <span data-ttu-id="6d676-758">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="6d676-759">String</span><span class="sxs-lookup"><span data-stu-id="6d676-759">String</span></span> | | <span data-ttu-id="6d676-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="6d676-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="6d676-762">String</span><span class="sxs-lookup"><span data-stu-id="6d676-762">String</span></span> | | <span data-ttu-id="6d676-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="6d676-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="6d676-766">函数</span><span class="sxs-lookup"><span data-stu-id="6d676-766">function</span></span> | <span data-ttu-id="6d676-767">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-767">&lt;optional&gt;</span></span> | <span data-ttu-id="6d676-768">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6d676-769">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-769">Requirements</span></span>

|<span data-ttu-id="6d676-770">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-770">Requirement</span></span>| <span data-ttu-id="6d676-771">值</span><span class="sxs-lookup"><span data-stu-id="6d676-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-772">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-773">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-773">1.0</span></span>|
|[<span data-ttu-id="6d676-774">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-775">ReadItem</span></span>|
|[<span data-ttu-id="6d676-776">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-777">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="6d676-778">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-778">Examples</span></span>

<span data-ttu-id="6d676-779">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="6d676-780">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-780">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="6d676-781">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-781">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="6d676-782">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="6d676-783">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="6d676-784">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="6d676-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="6d676-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="6d676-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="6d676-786">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="6d676-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-787">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-788">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-788">Requirements</span></span>

|<span data-ttu-id="6d676-789">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-789">Requirement</span></span>| <span data-ttu-id="6d676-790">值</span><span class="sxs-lookup"><span data-stu-id="6d676-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-791">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-792">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-792">1.0</span></span>|
|[<span data-ttu-id="6d676-793">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-794">ReadItem</span></span>|
|[<span data-ttu-id="6d676-795">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-796">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6d676-797">返回：</span><span class="sxs-lookup"><span data-stu-id="6d676-797">Returns:</span></span>

<span data-ttu-id="6d676-798">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="6d676-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="6d676-799">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-799">Example</span></span>

<span data-ttu-id="6d676-800">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="6d676-800">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="6d676-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="6d676-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="6d676-802">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="6d676-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-803">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-804">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-804">Parameters</span></span>

|<span data-ttu-id="6d676-805">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-805">Name</span></span>| <span data-ttu-id="6d676-806">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-806">Type</span></span>| <span data-ttu-id="6d676-807">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="6d676-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="6d676-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="6d676-809">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="6d676-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6d676-810">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-810">Requirements</span></span>

|<span data-ttu-id="6d676-811">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-811">Requirement</span></span>| <span data-ttu-id="6d676-812">值</span><span class="sxs-lookup"><span data-stu-id="6d676-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-813">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-814">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-814">1.0</span></span>|
|[<span data-ttu-id="6d676-815">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-816">受限</span><span class="sxs-lookup"><span data-stu-id="6d676-816">Restricted</span></span>|
|[<span data-ttu-id="6d676-817">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-818">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6d676-819">返回：</span><span class="sxs-lookup"><span data-stu-id="6d676-819">Returns:</span></span>

<span data-ttu-id="6d676-820">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="6d676-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="6d676-821">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="6d676-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="6d676-822">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="6d676-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="6d676-823">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="6d676-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="6d676-824">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="6d676-824">Value of `entityType`</span></span> | <span data-ttu-id="6d676-825">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="6d676-825">Type of objects in returned array</span></span> | <span data-ttu-id="6d676-826">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="6d676-827">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-827">String</span></span> | <span data-ttu-id="6d676-828">**受限**</span><span class="sxs-lookup"><span data-stu-id="6d676-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="6d676-829">Contact</span><span class="sxs-lookup"><span data-stu-id="6d676-829">Contact</span></span> | <span data-ttu-id="6d676-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6d676-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="6d676-831">String</span><span class="sxs-lookup"><span data-stu-id="6d676-831">String</span></span> | <span data-ttu-id="6d676-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6d676-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="6d676-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="6d676-833">MeetingSuggestion</span></span> | <span data-ttu-id="6d676-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6d676-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="6d676-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="6d676-835">PhoneNumber</span></span> | <span data-ttu-id="6d676-836">**受限**</span><span class="sxs-lookup"><span data-stu-id="6d676-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="6d676-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="6d676-837">TaskSuggestion</span></span> | <span data-ttu-id="6d676-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="6d676-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="6d676-839">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-839">String</span></span> | <span data-ttu-id="6d676-840">**受限**</span><span class="sxs-lookup"><span data-stu-id="6d676-840">**Restricted**</span></span> |

<span data-ttu-id="6d676-841">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="6d676-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="6d676-842">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-842">Example</span></span>

<span data-ttu-id="6d676-843">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="6d676-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="6d676-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="6d676-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="6d676-845">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="6d676-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-846">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6d676-847">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="6d676-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-848">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-848">Parameters</span></span>

|<span data-ttu-id="6d676-849">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-849">Name</span></span>| <span data-ttu-id="6d676-850">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-850">Type</span></span>| <span data-ttu-id="6d676-851">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6d676-852">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-852">String</span></span>|<span data-ttu-id="6d676-853">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="6d676-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6d676-854">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-854">Requirements</span></span>

|<span data-ttu-id="6d676-855">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-855">Requirement</span></span>| <span data-ttu-id="6d676-856">值</span><span class="sxs-lookup"><span data-stu-id="6d676-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-857">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-858">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-858">1.0</span></span>|
|[<span data-ttu-id="6d676-859">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-860">ReadItem</span></span>|
|[<span data-ttu-id="6d676-861">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-862">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6d676-863">返回：</span><span class="sxs-lookup"><span data-stu-id="6d676-863">Returns:</span></span>

<span data-ttu-id="6d676-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="6d676-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="6d676-866">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="6d676-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="6d676-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="6d676-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="6d676-868">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="6d676-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-869">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6d676-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="6d676-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="6d676-873">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="6d676-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="6d676-874">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="6d676-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="6d676-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="6d676-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d676-877">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-877">Requirements</span></span>

|<span data-ttu-id="6d676-878">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-878">Requirement</span></span>| <span data-ttu-id="6d676-879">值</span><span class="sxs-lookup"><span data-stu-id="6d676-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-880">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-881">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-881">1.0</span></span>|
|[<span data-ttu-id="6d676-882">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-883">ReadItem</span></span>|
|[<span data-ttu-id="6d676-884">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-885">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6d676-886">返回：</span><span class="sxs-lookup"><span data-stu-id="6d676-886">Returns:</span></span>

<span data-ttu-id="6d676-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="6d676-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="6d676-889">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="6d676-889">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6d676-890">对象</span><span class="sxs-lookup"><span data-stu-id="6d676-890">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6d676-891">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-891">Example</span></span>

<span data-ttu-id="6d676-892">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="6d676-892">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="6d676-893">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="6d676-893">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="6d676-894">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="6d676-894">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="6d676-895">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="6d676-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="6d676-896">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="6d676-896">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="6d676-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="6d676-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-899">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-899">Parameters</span></span>

|<span data-ttu-id="6d676-900">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-900">Name</span></span>| <span data-ttu-id="6d676-901">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-901">Type</span></span>| <span data-ttu-id="6d676-902">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-902">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="6d676-903">String</span><span class="sxs-lookup"><span data-stu-id="6d676-903">String</span></span>|<span data-ttu-id="6d676-904">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="6d676-904">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6d676-905">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-905">Requirements</span></span>

|<span data-ttu-id="6d676-906">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-906">Requirement</span></span>| <span data-ttu-id="6d676-907">值</span><span class="sxs-lookup"><span data-stu-id="6d676-907">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-908">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-908">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-909">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-909">1.0</span></span>|
|[<span data-ttu-id="6d676-910">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-910">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-911">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-911">ReadItem</span></span>|
|[<span data-ttu-id="6d676-912">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-912">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-913">阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-913">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="6d676-914">返回：</span><span class="sxs-lookup"><span data-stu-id="6d676-914">Returns:</span></span>

<span data-ttu-id="6d676-915">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="6d676-915">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="6d676-916">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="6d676-916">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6d676-917">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="6d676-917">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6d676-918">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-918">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="6d676-919">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="6d676-919">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="6d676-920">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="6d676-920">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="6d676-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="6d676-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-923">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-923">Parameters</span></span>

|<span data-ttu-id="6d676-924">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-924">Name</span></span>| <span data-ttu-id="6d676-925">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-925">Type</span></span>| <span data-ttu-id="6d676-926">属性</span><span class="sxs-lookup"><span data-stu-id="6d676-926">Attributes</span></span>| <span data-ttu-id="6d676-927">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-927">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="6d676-928">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6d676-928">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="6d676-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="6d676-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="6d676-932">Object</span><span class="sxs-lookup"><span data-stu-id="6d676-932">Object</span></span>| <span data-ttu-id="6d676-933">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-933">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-934">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="6d676-934">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6d676-935">Object</span><span class="sxs-lookup"><span data-stu-id="6d676-935">Object</span></span>| <span data-ttu-id="6d676-936">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-936">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-937">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-937">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6d676-938">function</span><span class="sxs-lookup"><span data-stu-id="6d676-938">function</span></span>||<span data-ttu-id="6d676-939">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-939">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6d676-940">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="6d676-940">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="6d676-941">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="6d676-941">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6d676-942">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-942">Requirements</span></span>

|<span data-ttu-id="6d676-943">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-943">Requirement</span></span>| <span data-ttu-id="6d676-944">值</span><span class="sxs-lookup"><span data-stu-id="6d676-944">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-945">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-945">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-946">1.2</span><span class="sxs-lookup"><span data-stu-id="6d676-946">1.2</span></span>|
|[<span data-ttu-id="6d676-947">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-947">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-948">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6d676-948">ReadWriteItem</span></span>|
|[<span data-ttu-id="6d676-949">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-949">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-950">撰写</span><span class="sxs-lookup"><span data-stu-id="6d676-950">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="6d676-951">返回：</span><span class="sxs-lookup"><span data-stu-id="6d676-951">Returns:</span></span>

<span data-ttu-id="6d676-952">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="6d676-952">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="6d676-953">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="6d676-953">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="6d676-954">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-954">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="6d676-955">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-955">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="6d676-956">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="6d676-956">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="6d676-957">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="6d676-957">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="6d676-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="6d676-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-961">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-961">Parameters</span></span>

|<span data-ttu-id="6d676-962">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-962">Name</span></span>| <span data-ttu-id="6d676-963">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-963">Type</span></span>| <span data-ttu-id="6d676-964">属性</span><span class="sxs-lookup"><span data-stu-id="6d676-964">Attributes</span></span>| <span data-ttu-id="6d676-965">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-965">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="6d676-966">函数</span><span class="sxs-lookup"><span data-stu-id="6d676-966">function</span></span>||<span data-ttu-id="6d676-967">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="6d676-968">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="6d676-968">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="6d676-969">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="6d676-969">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="6d676-970">Object</span><span class="sxs-lookup"><span data-stu-id="6d676-970">Object</span></span>| <span data-ttu-id="6d676-971">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-971">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-972">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-972">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="6d676-973">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="6d676-973">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6d676-974">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-974">Requirements</span></span>

|<span data-ttu-id="6d676-975">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-975">Requirement</span></span>| <span data-ttu-id="6d676-976">值</span><span class="sxs-lookup"><span data-stu-id="6d676-976">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-977">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-977">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-978">1.0</span><span class="sxs-lookup"><span data-stu-id="6d676-978">1.0</span></span>|
|[<span data-ttu-id="6d676-979">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-979">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-980">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6d676-980">ReadItem</span></span>|
|[<span data-ttu-id="6d676-981">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-981">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-982">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="6d676-982">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-983">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-983">Example</span></span>

<span data-ttu-id="6d676-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="6d676-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="6d676-987">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="6d676-987">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="6d676-988">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="6d676-988">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="6d676-989">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="6d676-989">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="6d676-990">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="6d676-990">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="6d676-991">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="6d676-991">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="6d676-992">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="6d676-992">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-993">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-993">Parameters</span></span>

|<span data-ttu-id="6d676-994">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-994">Name</span></span>| <span data-ttu-id="6d676-995">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-995">Type</span></span>| <span data-ttu-id="6d676-996">属性</span><span class="sxs-lookup"><span data-stu-id="6d676-996">Attributes</span></span>| <span data-ttu-id="6d676-997">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-997">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="6d676-998">String</span><span class="sxs-lookup"><span data-stu-id="6d676-998">String</span></span>||<span data-ttu-id="6d676-999">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="6d676-999">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="6d676-1000">对象</span><span class="sxs-lookup"><span data-stu-id="6d676-1000">Object</span></span>| <span data-ttu-id="6d676-1001">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-1002">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="6d676-1002">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6d676-1003">对象</span><span class="sxs-lookup"><span data-stu-id="6d676-1003">Object</span></span>| <span data-ttu-id="6d676-1004">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-1005">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-1005">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="6d676-1006">函数</span><span class="sxs-lookup"><span data-stu-id="6d676-1006">function</span></span>| <span data-ttu-id="6d676-1007">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-1008">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-1008">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="6d676-1009">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="6d676-1009">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="6d676-1010">错误</span><span class="sxs-lookup"><span data-stu-id="6d676-1010">Errors</span></span>

| <span data-ttu-id="6d676-1011">错误代码</span><span class="sxs-lookup"><span data-stu-id="6d676-1011">Error code</span></span> | <span data-ttu-id="6d676-1012">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-1012">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="6d676-1013">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="6d676-1013">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6d676-1014">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-1014">Requirements</span></span>

|<span data-ttu-id="6d676-1015">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-1015">Requirement</span></span>| <span data-ttu-id="6d676-1016">值</span><span class="sxs-lookup"><span data-stu-id="6d676-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-1017">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-1018">1.1</span><span class="sxs-lookup"><span data-stu-id="6d676-1018">1.1</span></span>|
|[<span data-ttu-id="6d676-1019">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-1019">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-1020">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6d676-1020">ReadWriteItem</span></span>|
|[<span data-ttu-id="6d676-1021">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-1021">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-1022">撰写</span><span class="sxs-lookup"><span data-stu-id="6d676-1022">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-1023">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-1023">Example</span></span>

<span data-ttu-id="6d676-1024">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="6d676-1024">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="6d676-1025">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="6d676-1025">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="6d676-1026">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="6d676-1026">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="6d676-p166">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="6d676-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="6d676-1030">参数</span><span class="sxs-lookup"><span data-stu-id="6d676-1030">Parameters</span></span>

|<span data-ttu-id="6d676-1031">名称</span><span class="sxs-lookup"><span data-stu-id="6d676-1031">Name</span></span>| <span data-ttu-id="6d676-1032">类型</span><span class="sxs-lookup"><span data-stu-id="6d676-1032">Type</span></span>| <span data-ttu-id="6d676-1033">属性</span><span class="sxs-lookup"><span data-stu-id="6d676-1033">Attributes</span></span>| <span data-ttu-id="6d676-1034">说明</span><span class="sxs-lookup"><span data-stu-id="6d676-1034">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="6d676-1035">字符串</span><span class="sxs-lookup"><span data-stu-id="6d676-1035">String</span></span>||<span data-ttu-id="6d676-p167">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="6d676-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="6d676-1039">Object</span><span class="sxs-lookup"><span data-stu-id="6d676-1039">Object</span></span>| <span data-ttu-id="6d676-1040">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-1041">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="6d676-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="6d676-1042">对象</span><span class="sxs-lookup"><span data-stu-id="6d676-1042">Object</span></span>| <span data-ttu-id="6d676-1043">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-1044">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="6d676-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="6d676-1045">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="6d676-1045">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="6d676-1046">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="6d676-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="6d676-1047">如果`text`为, 则当前样式应用于 web 上的 Outlook 和桌面客户端。</span><span class="sxs-lookup"><span data-stu-id="6d676-1047">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="6d676-1048">如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="6d676-1048">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="6d676-1049">如果`html`和字段支持 HTML (主题不), 则当前样式应用于 web 上的 outlook, 并且在 outlook 桌面客户端中应用了默认样式。</span><span class="sxs-lookup"><span data-stu-id="6d676-1049">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="6d676-1050">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="6d676-1050">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="6d676-1051">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="6d676-1051">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="6d676-1052">function</span><span class="sxs-lookup"><span data-stu-id="6d676-1052">function</span></span>||<span data-ttu-id="6d676-1053">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="6d676-1053">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="6d676-1054">Requirements</span><span class="sxs-lookup"><span data-stu-id="6d676-1054">Requirements</span></span>

|<span data-ttu-id="6d676-1055">要求</span><span class="sxs-lookup"><span data-stu-id="6d676-1055">Requirement</span></span>| <span data-ttu-id="6d676-1056">值</span><span class="sxs-lookup"><span data-stu-id="6d676-1056">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d676-1057">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="6d676-1057">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6d676-1058">1.2</span><span class="sxs-lookup"><span data-stu-id="6d676-1058">1.2</span></span>|
|[<span data-ttu-id="6d676-1059">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="6d676-1059">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6d676-1060">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="6d676-1060">ReadWriteItem</span></span>|
|[<span data-ttu-id="6d676-1061">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="6d676-1061">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6d676-1062">撰写</span><span class="sxs-lookup"><span data-stu-id="6d676-1062">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="6d676-1063">示例</span><span class="sxs-lookup"><span data-stu-id="6d676-1063">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
