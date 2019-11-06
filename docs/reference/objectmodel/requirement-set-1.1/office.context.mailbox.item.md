---
title: "\"Context\"-\"邮箱\"。项目-要求集1。1"
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 5cbf942ea9b1351e0f945a9ca5534a9ba090b79b
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001612"
---
# <a name="item"></a><span data-ttu-id="2822f-102">item</span><span class="sxs-lookup"><span data-stu-id="2822f-102">item</span></span>

### <span data-ttu-id="2822f-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="2822f-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="2822f-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="2822f-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-107">Requirements</span></span>

|<span data-ttu-id="2822f-108">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-108">Requirement</span></span>| <span data-ttu-id="2822f-109">值</span><span class="sxs-lookup"><span data-stu-id="2822f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-111">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-111">1.0</span></span>|
|[<span data-ttu-id="2822f-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-113">受限</span><span class="sxs-lookup"><span data-stu-id="2822f-113">Restricted</span></span>|
|[<span data-ttu-id="2822f-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2822f-116">成员和方法</span><span class="sxs-lookup"><span data-stu-id="2822f-116">Members and methods</span></span>

| <span data-ttu-id="2822f-117">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-117">Member</span></span> | <span data-ttu-id="2822f-118">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2822f-119">attachments</span><span class="sxs-lookup"><span data-stu-id="2822f-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="2822f-120">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-120">Member</span></span> |
| [<span data-ttu-id="2822f-121">bcc</span><span class="sxs-lookup"><span data-stu-id="2822f-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="2822f-122">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-122">Member</span></span> |
| [<span data-ttu-id="2822f-123">body</span><span class="sxs-lookup"><span data-stu-id="2822f-123">body</span></span>](#body-body) | <span data-ttu-id="2822f-124">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-124">Member</span></span> |
| [<span data-ttu-id="2822f-125">cc</span><span class="sxs-lookup"><span data-stu-id="2822f-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2822f-126">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-126">Member</span></span> |
| [<span data-ttu-id="2822f-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="2822f-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="2822f-128">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-128">Member</span></span> |
| [<span data-ttu-id="2822f-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="2822f-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="2822f-130">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-130">Member</span></span> |
| [<span data-ttu-id="2822f-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="2822f-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="2822f-132">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-132">Member</span></span> |
| [<span data-ttu-id="2822f-133">end</span><span class="sxs-lookup"><span data-stu-id="2822f-133">end</span></span>](#end-datetime) | <span data-ttu-id="2822f-134">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-134">Member</span></span> |
| [<span data-ttu-id="2822f-135">from</span><span class="sxs-lookup"><span data-stu-id="2822f-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="2822f-136">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-136">Member</span></span> |
| [<span data-ttu-id="2822f-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="2822f-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="2822f-138">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-138">Member</span></span> |
| [<span data-ttu-id="2822f-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="2822f-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="2822f-140">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-140">Member</span></span> |
| [<span data-ttu-id="2822f-141">itemId</span><span class="sxs-lookup"><span data-stu-id="2822f-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="2822f-142">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-142">Member</span></span> |
| [<span data-ttu-id="2822f-143">itemType</span><span class="sxs-lookup"><span data-stu-id="2822f-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="2822f-144">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-144">Member</span></span> |
| [<span data-ttu-id="2822f-145">location</span><span class="sxs-lookup"><span data-stu-id="2822f-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="2822f-146">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-146">Member</span></span> |
| [<span data-ttu-id="2822f-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="2822f-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="2822f-148">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-148">Member</span></span> |
| [<span data-ttu-id="2822f-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="2822f-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2822f-150">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-150">Member</span></span> |
| [<span data-ttu-id="2822f-151">organizer</span><span class="sxs-lookup"><span data-stu-id="2822f-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="2822f-152">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-152">Member</span></span> |
| [<span data-ttu-id="2822f-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="2822f-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2822f-154">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-154">Member</span></span> |
| [<span data-ttu-id="2822f-155">sender</span><span class="sxs-lookup"><span data-stu-id="2822f-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="2822f-156">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-156">Member</span></span> |
| [<span data-ttu-id="2822f-157">start</span><span class="sxs-lookup"><span data-stu-id="2822f-157">start</span></span>](#start-datetime) | <span data-ttu-id="2822f-158">Member</span><span class="sxs-lookup"><span data-stu-id="2822f-158">Member</span></span> |
| [<span data-ttu-id="2822f-159">subject</span><span class="sxs-lookup"><span data-stu-id="2822f-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="2822f-160">成员</span><span class="sxs-lookup"><span data-stu-id="2822f-160">Member</span></span> |
| [<span data-ttu-id="2822f-161">to</span><span class="sxs-lookup"><span data-stu-id="2822f-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2822f-162">成员</span><span class="sxs-lookup"><span data-stu-id="2822f-162">Member</span></span> |
| [<span data-ttu-id="2822f-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2822f-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="2822f-164">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-164">Method</span></span> |
| [<span data-ttu-id="2822f-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2822f-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="2822f-166">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-166">Method</span></span> |
| [<span data-ttu-id="2822f-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="2822f-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="2822f-168">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-168">Method</span></span> |
| [<span data-ttu-id="2822f-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="2822f-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="2822f-170">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-170">Method</span></span> |
| [<span data-ttu-id="2822f-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="2822f-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="2822f-172">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-172">Method</span></span> |
| [<span data-ttu-id="2822f-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="2822f-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2822f-174">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-174">Method</span></span> |
| [<span data-ttu-id="2822f-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="2822f-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2822f-176">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-176">Method</span></span> |
| [<span data-ttu-id="2822f-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2822f-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="2822f-178">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-178">Method</span></span> |
| [<span data-ttu-id="2822f-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="2822f-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="2822f-180">Method</span><span class="sxs-lookup"><span data-stu-id="2822f-180">Method</span></span> |
| [<span data-ttu-id="2822f-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2822f-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="2822f-182">方法</span><span class="sxs-lookup"><span data-stu-id="2822f-182">Method</span></span> |
| [<span data-ttu-id="2822f-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2822f-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="2822f-184">方法</span><span class="sxs-lookup"><span data-stu-id="2822f-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="2822f-185">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-185">Example</span></span>

<span data-ttu-id="2822f-186">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="2822f-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2822f-187">Members</span><span class="sxs-lookup"><span data-stu-id="2822f-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="2822f-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="2822f-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="2822f-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-191">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="2822f-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2822f-192">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="2822f-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-193">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-193">Type</span></span>

*   <span data-ttu-id="2822f-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="2822f-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-195">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-195">Requirements</span></span>

|<span data-ttu-id="2822f-196">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-196">Requirement</span></span>| <span data-ttu-id="2822f-197">值</span><span class="sxs-lookup"><span data-stu-id="2822f-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-199">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-199">1.0</span></span>|
|[<span data-ttu-id="2822f-200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-201">ReadItem</span></span>|
|[<span data-ttu-id="2822f-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-203">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-204">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-204">Example</span></span>

<span data-ttu-id="2822f-205">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="2822f-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="2822f-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-207">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2822f-208">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-208">Compose mode only.</span></span>

<span data-ttu-id="2822f-209">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-209">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-210">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2822f-210">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2822f-211">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-211">Get 500 members maximum.</span></span>
- <span data-ttu-id="2822f-212">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-212">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-213">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-213">Type</span></span>

*   [<span data-ttu-id="2822f-214">收件人</span><span class="sxs-lookup"><span data-stu-id="2822f-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="2822f-215">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-215">Requirements</span></span>

|<span data-ttu-id="2822f-216">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-216">Requirement</span></span>| <span data-ttu-id="2822f-217">值</span><span class="sxs-lookup"><span data-stu-id="2822f-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-219">1.1</span><span class="sxs-lookup"><span data-stu-id="2822f-219">1.1</span></span>|
|[<span data-ttu-id="2822f-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-221">ReadItem</span></span>|
|[<span data-ttu-id="2822f-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-223">撰写</span><span class="sxs-lookup"><span data-stu-id="2822f-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-224">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="2822f-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-226">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-227">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-227">Type</span></span>

*   [<span data-ttu-id="2822f-228">Body</span><span class="sxs-lookup"><span data-stu-id="2822f-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="2822f-229">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-229">Requirements</span></span>

|<span data-ttu-id="2822f-230">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-230">Requirement</span></span>| <span data-ttu-id="2822f-231">值</span><span class="sxs-lookup"><span data-stu-id="2822f-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-232">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-233">1.1</span><span class="sxs-lookup"><span data-stu-id="2822f-233">1.1</span></span>|
|[<span data-ttu-id="2822f-234">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-235">ReadItem</span></span>|
|[<span data-ttu-id="2822f-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-238">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-238">Example</span></span>

<span data-ttu-id="2822f-239">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="2822f-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="2822f-240">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="2822f-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="2822f-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-242">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2822f-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2822f-243">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-244">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-244">Read mode</span></span>

<span data-ttu-id="2822f-245">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="2822f-245">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="2822f-246">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-246">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-247">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-247">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-248">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-248">Compose mode</span></span>

<span data-ttu-id="2822f-249">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-249">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="2822f-250">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-251">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2822f-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2822f-252">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="2822f-253">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2822f-254">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-254">Type</span></span>

*   <span data-ttu-id="2822f-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-256">Requirements</span></span>

|<span data-ttu-id="2822f-257">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-257">Requirement</span></span>| <span data-ttu-id="2822f-258">值</span><span class="sxs-lookup"><span data-stu-id="2822f-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-260">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-260">1.0</span></span>|
|[<span data-ttu-id="2822f-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-262">ReadItem</span></span>|
|[<span data-ttu-id="2822f-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="2822f-265">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="2822f-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="2822f-266">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="2822f-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2822f-p110">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="2822f-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2822f-p111">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="2822f-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-271">Type</span><span class="sxs-lookup"><span data-stu-id="2822f-271">Type</span></span>

*   <span data-ttu-id="2822f-272">String</span><span class="sxs-lookup"><span data-stu-id="2822f-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-273">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-273">Requirements</span></span>

|<span data-ttu-id="2822f-274">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-274">Requirement</span></span>| <span data-ttu-id="2822f-275">值</span><span class="sxs-lookup"><span data-stu-id="2822f-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-276">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-277">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-277">1.0</span></span>|
|[<span data-ttu-id="2822f-278">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-279">ReadItem</span></span>|
|[<span data-ttu-id="2822f-280">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-281">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-282">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="2822f-283">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="2822f-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="2822f-p112">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-286">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-286">Type</span></span>

*   <span data-ttu-id="2822f-287">日期</span><span class="sxs-lookup"><span data-stu-id="2822f-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-288">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-288">Requirements</span></span>

|<span data-ttu-id="2822f-289">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-289">Requirement</span></span>| <span data-ttu-id="2822f-290">值</span><span class="sxs-lookup"><span data-stu-id="2822f-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-292">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-292">1.0</span></span>|
|[<span data-ttu-id="2822f-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-294">ReadItem</span></span>|
|[<span data-ttu-id="2822f-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-296">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-297">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="2822f-298">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="2822f-298">dateTimeModified: Date</span></span>

<span data-ttu-id="2822f-p113">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-301">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-302">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-302">Type</span></span>

*   <span data-ttu-id="2822f-303">日期</span><span class="sxs-lookup"><span data-stu-id="2822f-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-304">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-304">Requirements</span></span>

|<span data-ttu-id="2822f-305">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-305">Requirement</span></span>| <span data-ttu-id="2822f-306">值</span><span class="sxs-lookup"><span data-stu-id="2822f-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-307">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-308">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-308">1.0</span></span>|
|[<span data-ttu-id="2822f-309">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-310">ReadItem</span></span>|
|[<span data-ttu-id="2822f-311">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-312">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-313">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="2822f-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-315">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2822f-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2822f-p114">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2822f-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-318">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-318">Read mode</span></span>

<span data-ttu-id="2822f-319">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-320">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-320">Compose mode</span></span>

<span data-ttu-id="2822f-321">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2822f-322">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="2822f-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2822f-323">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="2822f-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2822f-324">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-324">Type</span></span>

*   <span data-ttu-id="2822f-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-326">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-326">Requirements</span></span>

|<span data-ttu-id="2822f-327">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-327">Requirement</span></span>| <span data-ttu-id="2822f-328">值</span><span class="sxs-lookup"><span data-stu-id="2822f-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-329">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-330">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-330">1.0</span></span>|
|[<span data-ttu-id="2822f-331">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-332">ReadItem</span></span>|
|[<span data-ttu-id="2822f-333">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-334">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="2822f-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-p115">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="2822f-p116">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="2822f-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-340">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="2822f-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-341">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-341">Type</span></span>

*   [<span data-ttu-id="2822f-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2822f-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="2822f-343">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-343">Requirements</span></span>

|<span data-ttu-id="2822f-344">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-344">Requirement</span></span>| <span data-ttu-id="2822f-345">值</span><span class="sxs-lookup"><span data-stu-id="2822f-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-346">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-347">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-347">1.0</span></span>|
|[<span data-ttu-id="2822f-348">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-349">ReadItem</span></span>|
|[<span data-ttu-id="2822f-350">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-351">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-352">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="2822f-353">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="2822f-353">internetMessageId: String</span></span>

<span data-ttu-id="2822f-p117">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-356">Type</span><span class="sxs-lookup"><span data-stu-id="2822f-356">Type</span></span>

*   <span data-ttu-id="2822f-357">String</span><span class="sxs-lookup"><span data-stu-id="2822f-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-358">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-358">Requirements</span></span>

|<span data-ttu-id="2822f-359">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-359">Requirement</span></span>| <span data-ttu-id="2822f-360">值</span><span class="sxs-lookup"><span data-stu-id="2822f-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-361">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-362">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-362">1.0</span></span>|
|[<span data-ttu-id="2822f-363">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-364">ReadItem</span></span>|
|[<span data-ttu-id="2822f-365">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-366">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-367">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="2822f-368">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="2822f-368">itemClass: String</span></span>

<span data-ttu-id="2822f-p118">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2822f-p119">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="2822f-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="2822f-373">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-373">Type</span></span> | <span data-ttu-id="2822f-374">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-374">Description</span></span> | <span data-ttu-id="2822f-375">项目类</span><span class="sxs-lookup"><span data-stu-id="2822f-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="2822f-376">约会项目</span><span class="sxs-lookup"><span data-stu-id="2822f-376">Appointment items</span></span> | <span data-ttu-id="2822f-377">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="2822f-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="2822f-378">邮件项目</span><span class="sxs-lookup"><span data-stu-id="2822f-378">Message items</span></span> | <span data-ttu-id="2822f-379">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="2822f-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="2822f-380">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="2822f-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-381">Type</span><span class="sxs-lookup"><span data-stu-id="2822f-381">Type</span></span>

*   <span data-ttu-id="2822f-382">String</span><span class="sxs-lookup"><span data-stu-id="2822f-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-383">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-383">Requirements</span></span>

|<span data-ttu-id="2822f-384">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-384">Requirement</span></span>| <span data-ttu-id="2822f-385">值</span><span class="sxs-lookup"><span data-stu-id="2822f-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-387">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-387">1.0</span></span>|
|[<span data-ttu-id="2822f-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-389">ReadItem</span></span>|
|[<span data-ttu-id="2822f-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-391">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-392">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2822f-393">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="2822f-393">(nullable) itemId: String</span></span>

<span data-ttu-id="2822f-394">获取当前项的[Exchange Web 服务项标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)。</span><span class="sxs-lookup"><span data-stu-id="2822f-394">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="2822f-395">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-395">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-396">`itemId`属性返回的标识符与[Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)相同。</span><span class="sxs-lookup"><span data-stu-id="2822f-396">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="2822f-397">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="2822f-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2822f-398">在使用此值进行 REST API 调用之前，应使用`Office.context.mailbox.convertToRestId`转换它，这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="2822f-398">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="2822f-399">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="2822f-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-400">Type</span><span class="sxs-lookup"><span data-stu-id="2822f-400">Type</span></span>

*   <span data-ttu-id="2822f-401">String</span><span class="sxs-lookup"><span data-stu-id="2822f-401">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-402">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-402">Requirements</span></span>

|<span data-ttu-id="2822f-403">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-403">Requirement</span></span>| <span data-ttu-id="2822f-404">值</span><span class="sxs-lookup"><span data-stu-id="2822f-404">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-405">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-405">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-406">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-406">1.0</span></span>|
|[<span data-ttu-id="2822f-407">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-407">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-408">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-408">ReadItem</span></span>|
|[<span data-ttu-id="2822f-409">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-409">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-410">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-410">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-411">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-411">Example</span></span>

<span data-ttu-id="2822f-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="2822f-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="2822f-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-414">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-415">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="2822f-415">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2822f-416">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="2822f-416">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-417">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-417">Type</span></span>

*   [<span data-ttu-id="2822f-418">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2822f-418">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="2822f-419">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-419">Requirements</span></span>

|<span data-ttu-id="2822f-420">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-420">Requirement</span></span>| <span data-ttu-id="2822f-421">值</span><span class="sxs-lookup"><span data-stu-id="2822f-421">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-422">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-423">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-423">1.0</span></span>|
|[<span data-ttu-id="2822f-424">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-424">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-425">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-425">ReadItem</span></span>|
|[<span data-ttu-id="2822f-426">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-426">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-427">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-427">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-428">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-428">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="2822f-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-429">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-430">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="2822f-430">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-431">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-431">Read mode</span></span>

<span data-ttu-id="2822f-432">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="2822f-432">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-433">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-433">Compose mode</span></span>

<span data-ttu-id="2822f-434">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-434">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2822f-435">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-435">Type</span></span>

*   <span data-ttu-id="2822f-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-436">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-437">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-437">Requirements</span></span>

|<span data-ttu-id="2822f-438">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-438">Requirement</span></span>| <span data-ttu-id="2822f-439">值</span><span class="sxs-lookup"><span data-stu-id="2822f-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-440">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-441">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-441">1.0</span></span>|
|[<span data-ttu-id="2822f-442">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-443">ReadItem</span></span>|
|[<span data-ttu-id="2822f-444">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-445">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-445">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2822f-446">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="2822f-446">normalizedSubject: String</span></span>

<span data-ttu-id="2822f-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2822f-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="2822f-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-451">Type</span><span class="sxs-lookup"><span data-stu-id="2822f-451">Type</span></span>

*   <span data-ttu-id="2822f-452">String</span><span class="sxs-lookup"><span data-stu-id="2822f-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-453">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-453">Requirements</span></span>

|<span data-ttu-id="2822f-454">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-454">Requirement</span></span>| <span data-ttu-id="2822f-455">值</span><span class="sxs-lookup"><span data-stu-id="2822f-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-456">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-457">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-457">1.0</span></span>|
|[<span data-ttu-id="2822f-458">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-459">ReadItem</span></span>|
|[<span data-ttu-id="2822f-460">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-461">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-462">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-462">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="2822f-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-463">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-464">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2822f-464">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2822f-465">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-465">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-466">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-466">Read mode</span></span>

<span data-ttu-id="2822f-467">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-467">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="2822f-468">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-468">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-469">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-469">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-470">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-470">Compose mode</span></span>

<span data-ttu-id="2822f-471">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-471">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="2822f-472">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-473">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2822f-473">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2822f-474">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-474">Get 500 members maximum.</span></span>
- <span data-ttu-id="2822f-475">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-475">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2822f-476">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-476">Type</span></span>

*   <span data-ttu-id="2822f-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-477">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-478">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-478">Requirements</span></span>

|<span data-ttu-id="2822f-479">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-479">Requirement</span></span>| <span data-ttu-id="2822f-480">值</span><span class="sxs-lookup"><span data-stu-id="2822f-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-481">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-482">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-482">1.0</span></span>|
|[<span data-ttu-id="2822f-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-484">ReadItem</span></span>|
|[<span data-ttu-id="2822f-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-486">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-486">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="2822f-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-487">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-p128">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-490">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-490">Type</span></span>

*   [<span data-ttu-id="2822f-491">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2822f-491">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="2822f-492">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-492">Requirements</span></span>

|<span data-ttu-id="2822f-493">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-493">Requirement</span></span>| <span data-ttu-id="2822f-494">值</span><span class="sxs-lookup"><span data-stu-id="2822f-494">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-495">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-495">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-496">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-496">1.0</span></span>|
|[<span data-ttu-id="2822f-497">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-497">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-498">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-498">ReadItem</span></span>|
|[<span data-ttu-id="2822f-499">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-499">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-500">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-500">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-501">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-501">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="2822f-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-502">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-503">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2822f-503">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2822f-504">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-504">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-505">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-505">Read mode</span></span>

<span data-ttu-id="2822f-506">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-506">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="2822f-507">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-507">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-508">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-508">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-509">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-509">Compose mode</span></span>

<span data-ttu-id="2822f-510">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-510">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="2822f-511">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-512">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2822f-512">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2822f-513">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-513">Get 500 members maximum.</span></span>
- <span data-ttu-id="2822f-514">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-514">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="2822f-515">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-515">Type</span></span>

*   <span data-ttu-id="2822f-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-516">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-517">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-517">Requirements</span></span>

|<span data-ttu-id="2822f-518">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-518">Requirement</span></span>| <span data-ttu-id="2822f-519">值</span><span class="sxs-lookup"><span data-stu-id="2822f-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-520">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-521">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-521">1.0</span></span>|
|[<span data-ttu-id="2822f-522">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-523">ReadItem</span></span>|
|[<span data-ttu-id="2822f-524">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-525">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-525">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="2822f-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-526">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-p132">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2822f-p133">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="2822f-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-531">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="2822f-531">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2822f-532">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-532">Type</span></span>

*   [<span data-ttu-id="2822f-533">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2822f-533">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="2822f-534">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-534">Requirements</span></span>

|<span data-ttu-id="2822f-535">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-535">Requirement</span></span>| <span data-ttu-id="2822f-536">值</span><span class="sxs-lookup"><span data-stu-id="2822f-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-537">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-538">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-538">1.0</span></span>|
|[<span data-ttu-id="2822f-539">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-540">ReadItem</span></span>|
|[<span data-ttu-id="2822f-541">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-542">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-542">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-543">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-543">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="2822f-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-544">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-545">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2822f-545">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2822f-p134">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2822f-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-548">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-548">Read mode</span></span>

<span data-ttu-id="2822f-549">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-549">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-550">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-550">Compose mode</span></span>

<span data-ttu-id="2822f-551">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-551">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2822f-552">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="2822f-552">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2822f-553">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="2822f-553">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2822f-554">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-554">Type</span></span>

*   <span data-ttu-id="2822f-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-555">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-556">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-556">Requirements</span></span>

|<span data-ttu-id="2822f-557">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-557">Requirement</span></span>| <span data-ttu-id="2822f-558">值</span><span class="sxs-lookup"><span data-stu-id="2822f-558">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-559">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-559">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-560">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-560">1.0</span></span>|
|[<span data-ttu-id="2822f-561">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-561">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-562">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-562">ReadItem</span></span>|
|[<span data-ttu-id="2822f-563">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-563">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-564">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-564">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="2822f-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-565">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-566">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="2822f-566">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2822f-567">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="2822f-567">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-568">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-568">Read mode</span></span>

<span data-ttu-id="2822f-p135">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="2822f-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-571">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-571">Compose mode</span></span>

<span data-ttu-id="2822f-572">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-572">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="2822f-573">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-573">Type</span></span>

*   <span data-ttu-id="2822f-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-574">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-575">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-575">Requirements</span></span>

|<span data-ttu-id="2822f-576">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-576">Requirement</span></span>| <span data-ttu-id="2822f-577">值</span><span class="sxs-lookup"><span data-stu-id="2822f-577">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-578">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-578">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-579">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-579">1.0</span></span>|
|[<span data-ttu-id="2822f-580">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-580">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-581">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-581">ReadItem</span></span>|
|[<span data-ttu-id="2822f-582">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-582">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-583">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-583">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="2822f-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-584">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="2822f-585">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2822f-585">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2822f-586">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2822f-586">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2822f-587">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2822f-587">Read mode</span></span>

<span data-ttu-id="2822f-588">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="2822f-588">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="2822f-589">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-589">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-590">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-590">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="2822f-591">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2822f-591">Compose mode</span></span>

<span data-ttu-id="2822f-592">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-592">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="2822f-593">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2822f-594">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2822f-594">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2822f-595">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-595">Get 500 members maximum.</span></span>
- <span data-ttu-id="2822f-596">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="2822f-596">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2822f-597">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-597">Type</span></span>

*   <span data-ttu-id="2822f-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-598">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-599">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-599">Requirements</span></span>

|<span data-ttu-id="2822f-600">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-600">Requirement</span></span>| <span data-ttu-id="2822f-601">值</span><span class="sxs-lookup"><span data-stu-id="2822f-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-602">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-602">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-603">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-603">1.0</span></span>|
|[<span data-ttu-id="2822f-604">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-604">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-605">ReadItem</span></span>|
|[<span data-ttu-id="2822f-606">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-606">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-607">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-607">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2822f-608">方法</span><span class="sxs-lookup"><span data-stu-id="2822f-608">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2822f-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2822f-609">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2822f-610">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="2822f-610">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2822f-611">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="2822f-611">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2822f-612">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="2822f-612">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-613">参数</span><span class="sxs-lookup"><span data-stu-id="2822f-613">Parameters</span></span>

|<span data-ttu-id="2822f-614">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-614">Name</span></span>| <span data-ttu-id="2822f-615">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-615">Type</span></span>| <span data-ttu-id="2822f-616">属性</span><span class="sxs-lookup"><span data-stu-id="2822f-616">Attributes</span></span>| <span data-ttu-id="2822f-617">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-617">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="2822f-618">String</span><span class="sxs-lookup"><span data-stu-id="2822f-618">String</span></span>||<span data-ttu-id="2822f-p139">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="2822f-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2822f-621">字符串</span><span class="sxs-lookup"><span data-stu-id="2822f-621">String</span></span>||<span data-ttu-id="2822f-p140">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2822f-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2822f-624">Object</span><span class="sxs-lookup"><span data-stu-id="2822f-624">Object</span></span>| <span data-ttu-id="2822f-625">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-625">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-626">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2822f-626">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2822f-627">对象</span><span class="sxs-lookup"><span data-stu-id="2822f-627">Object</span></span>| <span data-ttu-id="2822f-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-628">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-629">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-629">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2822f-630">函数</span><span class="sxs-lookup"><span data-stu-id="2822f-630">function</span></span>| <span data-ttu-id="2822f-631">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-631">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-632">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2822f-633">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2822f-633">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2822f-634">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-634">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2822f-635">错误</span><span class="sxs-lookup"><span data-stu-id="2822f-635">Errors</span></span>

| <span data-ttu-id="2822f-636">错误代码</span><span class="sxs-lookup"><span data-stu-id="2822f-636">Error code</span></span> | <span data-ttu-id="2822f-637">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-637">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="2822f-638">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="2822f-638">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="2822f-639">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="2822f-639">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2822f-640">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="2822f-640">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2822f-641">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-641">Requirements</span></span>

|<span data-ttu-id="2822f-642">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-642">Requirement</span></span>| <span data-ttu-id="2822f-643">值</span><span class="sxs-lookup"><span data-stu-id="2822f-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-644">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-645">1.1</span><span class="sxs-lookup"><span data-stu-id="2822f-645">1.1</span></span>|
|[<span data-ttu-id="2822f-646">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-646">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-647">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2822f-647">ReadWriteItem</span></span>|
|[<span data-ttu-id="2822f-648">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-648">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-649">撰写</span><span class="sxs-lookup"><span data-stu-id="2822f-649">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-650">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-650">Example</span></span>

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

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2822f-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2822f-651">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2822f-652">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="2822f-652">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2822f-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="2822f-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2822f-656">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="2822f-656">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2822f-657">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="2822f-657">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-658">Parameters</span><span class="sxs-lookup"><span data-stu-id="2822f-658">Parameters</span></span>

|<span data-ttu-id="2822f-659">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-659">Name</span></span>| <span data-ttu-id="2822f-660">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-660">Type</span></span>| <span data-ttu-id="2822f-661">属性</span><span class="sxs-lookup"><span data-stu-id="2822f-661">Attributes</span></span>| <span data-ttu-id="2822f-662">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-662">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="2822f-663">String</span><span class="sxs-lookup"><span data-stu-id="2822f-663">String</span></span>||<span data-ttu-id="2822f-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2822f-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2822f-666">String</span><span class="sxs-lookup"><span data-stu-id="2822f-666">String</span></span>||<span data-ttu-id="2822f-667">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="2822f-667">The subject of the item to be attached.</span></span> <span data-ttu-id="2822f-668">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2822f-668">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2822f-669">对象</span><span class="sxs-lookup"><span data-stu-id="2822f-669">Object</span></span>| <span data-ttu-id="2822f-670">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-670">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-671">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2822f-671">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2822f-672">对象</span><span class="sxs-lookup"><span data-stu-id="2822f-672">Object</span></span>| <span data-ttu-id="2822f-673">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-673">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-674">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-674">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2822f-675">函数</span><span class="sxs-lookup"><span data-stu-id="2822f-675">function</span></span>| <span data-ttu-id="2822f-676">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-676">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-677">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-677">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2822f-678">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2822f-678">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2822f-679">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-679">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2822f-680">错误</span><span class="sxs-lookup"><span data-stu-id="2822f-680">Errors</span></span>

| <span data-ttu-id="2822f-681">错误代码</span><span class="sxs-lookup"><span data-stu-id="2822f-681">Error code</span></span> | <span data-ttu-id="2822f-682">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-682">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2822f-683">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="2822f-683">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2822f-684">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-684">Requirements</span></span>

|<span data-ttu-id="2822f-685">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-685">Requirement</span></span>| <span data-ttu-id="2822f-686">值</span><span class="sxs-lookup"><span data-stu-id="2822f-686">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-687">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-687">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-688">1.1</span><span class="sxs-lookup"><span data-stu-id="2822f-688">1.1</span></span>|
|[<span data-ttu-id="2822f-689">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-689">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-690">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2822f-690">ReadWriteItem</span></span>|
|[<span data-ttu-id="2822f-691">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-691">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-692">撰写</span><span class="sxs-lookup"><span data-stu-id="2822f-692">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-693">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-693">Example</span></span>

<span data-ttu-id="2822f-694">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="2822f-694">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="2822f-695">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2822f-695">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="2822f-696">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="2822f-696">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-697">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-697">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2822f-698">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="2822f-698">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2822f-699">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="2822f-699">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-700">要求集1.1 中不支持在呼叫`displayReplyAllForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="2822f-700">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="2822f-701">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="2822f-701">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-702">Parameters</span><span class="sxs-lookup"><span data-stu-id="2822f-702">Parameters</span></span>

|<span data-ttu-id="2822f-703">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-703">Name</span></span>| <span data-ttu-id="2822f-704">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-704">Type</span></span>| <span data-ttu-id="2822f-705">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-705">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2822f-706">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="2822f-706">String &#124; Object</span></span>| |<span data-ttu-id="2822f-p145">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2822f-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2822f-709">**或**</span><span class="sxs-lookup"><span data-stu-id="2822f-709">**OR**</span></span><br/><span data-ttu-id="2822f-p146">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="2822f-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2822f-712">字符串</span><span class="sxs-lookup"><span data-stu-id="2822f-712">String</span></span> | <span data-ttu-id="2822f-713">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-713">&lt;optional&gt;</span></span> | <span data-ttu-id="2822f-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2822f-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="2822f-716">函数</span><span class="sxs-lookup"><span data-stu-id="2822f-716">function</span></span> | <span data-ttu-id="2822f-717">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-717">&lt;optional&gt;</span></span> | <span data-ttu-id="2822f-718">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-718">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2822f-719">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-719">Requirements</span></span>

|<span data-ttu-id="2822f-720">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-720">Requirement</span></span>| <span data-ttu-id="2822f-721">值</span><span class="sxs-lookup"><span data-stu-id="2822f-721">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-722">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-722">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-723">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-723">1.0</span></span>|
|[<span data-ttu-id="2822f-724">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-724">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-725">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-725">ReadItem</span></span>|
|[<span data-ttu-id="2822f-726">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-726">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-727">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-727">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2822f-728">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-728">Examples</span></span>

<span data-ttu-id="2822f-729">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-729">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2822f-730">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="2822f-730">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2822f-731">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="2822f-731">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2822f-732">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="2822f-732">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="2822f-733">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2822f-733">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="2822f-734">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="2822f-734">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-735">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-735">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2822f-736">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="2822f-736">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2822f-737">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="2822f-737">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-738">要求集1.1 中不支持在呼叫`displayReplyForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="2822f-738">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="2822f-739">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="2822f-739">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-740">Parameters</span><span class="sxs-lookup"><span data-stu-id="2822f-740">Parameters</span></span>

|<span data-ttu-id="2822f-741">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-741">Name</span></span>| <span data-ttu-id="2822f-742">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-742">Type</span></span>| <span data-ttu-id="2822f-743">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-743">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2822f-744">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="2822f-744">String &#124; Object</span></span>| | <span data-ttu-id="2822f-p149">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2822f-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2822f-747">**或**</span><span class="sxs-lookup"><span data-stu-id="2822f-747">**OR**</span></span><br/><span data-ttu-id="2822f-p150">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="2822f-p150">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2822f-750">字符串</span><span class="sxs-lookup"><span data-stu-id="2822f-750">String</span></span> | <span data-ttu-id="2822f-751">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-751">&lt;optional&gt;</span></span> | <span data-ttu-id="2822f-p151">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2822f-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="2822f-754">函数</span><span class="sxs-lookup"><span data-stu-id="2822f-754">function</span></span> | <span data-ttu-id="2822f-755">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-755">&lt;optional&gt;</span></span> | <span data-ttu-id="2822f-756">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2822f-757">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-757">Requirements</span></span>

|<span data-ttu-id="2822f-758">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-758">Requirement</span></span>| <span data-ttu-id="2822f-759">值</span><span class="sxs-lookup"><span data-stu-id="2822f-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-760">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-761">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-761">1.0</span></span>|
|[<span data-ttu-id="2822f-762">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-763">ReadItem</span></span>|
|[<span data-ttu-id="2822f-764">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-765">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2822f-766">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-766">Examples</span></span>

<span data-ttu-id="2822f-767">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-767">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2822f-768">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="2822f-768">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2822f-769">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="2822f-769">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2822f-770">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="2822f-770">Reply with a body and a callback.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="2822f-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="2822f-771">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="2822f-772">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="2822f-772">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-773">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-773">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-774">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-774">Requirements</span></span>

|<span data-ttu-id="2822f-775">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-775">Requirement</span></span>| <span data-ttu-id="2822f-776">值</span><span class="sxs-lookup"><span data-stu-id="2822f-776">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-777">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-777">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-778">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-778">1.0</span></span>|
|[<span data-ttu-id="2822f-779">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-779">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-780">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-780">ReadItem</span></span>|
|[<span data-ttu-id="2822f-781">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-781">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-782">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-782">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2822f-783">返回：</span><span class="sxs-lookup"><span data-stu-id="2822f-783">Returns:</span></span>

<span data-ttu-id="2822f-784">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="2822f-784">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="2822f-785">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-785">Example</span></span>

<span data-ttu-id="2822f-786">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="2822f-786">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="2822f-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="2822f-787">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="2822f-788">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="2822f-788">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-789">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-789">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-790">Parameters</span><span class="sxs-lookup"><span data-stu-id="2822f-790">Parameters</span></span>

|<span data-ttu-id="2822f-791">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-791">Name</span></span>| <span data-ttu-id="2822f-792">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-792">Type</span></span>| <span data-ttu-id="2822f-793">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-793">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="2822f-794">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2822f-794">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="2822f-795">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="2822f-795">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2822f-796">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-796">Requirements</span></span>

|<span data-ttu-id="2822f-797">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-797">Requirement</span></span>| <span data-ttu-id="2822f-798">值</span><span class="sxs-lookup"><span data-stu-id="2822f-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-799">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-800">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-800">1.0</span></span>|
|[<span data-ttu-id="2822f-801">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-802">受限</span><span class="sxs-lookup"><span data-stu-id="2822f-802">Restricted</span></span>|
|[<span data-ttu-id="2822f-803">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-804">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2822f-805">返回：</span><span class="sxs-lookup"><span data-stu-id="2822f-805">Returns:</span></span>

<span data-ttu-id="2822f-806">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="2822f-806">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2822f-807">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="2822f-807">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2822f-808">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="2822f-808">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2822f-809">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="2822f-809">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="2822f-810">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="2822f-810">Value of `entityType`</span></span> | <span data-ttu-id="2822f-811">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="2822f-811">Type of objects in returned array</span></span> | <span data-ttu-id="2822f-812">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-812">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="2822f-813">String</span><span class="sxs-lookup"><span data-stu-id="2822f-813">String</span></span> | <span data-ttu-id="2822f-814">**受限**</span><span class="sxs-lookup"><span data-stu-id="2822f-814">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="2822f-815">Contact</span><span class="sxs-lookup"><span data-stu-id="2822f-815">Contact</span></span> | <span data-ttu-id="2822f-816">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2822f-816">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="2822f-817">String</span><span class="sxs-lookup"><span data-stu-id="2822f-817">String</span></span> | <span data-ttu-id="2822f-818">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2822f-818">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="2822f-819">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2822f-819">MeetingSuggestion</span></span> | <span data-ttu-id="2822f-820">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2822f-820">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="2822f-821">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2822f-821">PhoneNumber</span></span> | <span data-ttu-id="2822f-822">**受限**</span><span class="sxs-lookup"><span data-stu-id="2822f-822">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="2822f-823">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2822f-823">TaskSuggestion</span></span> | <span data-ttu-id="2822f-824">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2822f-824">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="2822f-825">String</span><span class="sxs-lookup"><span data-stu-id="2822f-825">String</span></span> | <span data-ttu-id="2822f-826">**受限**</span><span class="sxs-lookup"><span data-stu-id="2822f-826">**Restricted**</span></span> |

<span data-ttu-id="2822f-827">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="2822f-827">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="2822f-828">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-828">Example</span></span>

<span data-ttu-id="2822f-829">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="2822f-829">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="2822f-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="2822f-830">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="2822f-831">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="2822f-831">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-832">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-832">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2822f-833">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="2822f-833">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-834">参数</span><span class="sxs-lookup"><span data-stu-id="2822f-834">Parameters</span></span>

|<span data-ttu-id="2822f-835">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-835">Name</span></span>| <span data-ttu-id="2822f-836">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-836">Type</span></span>| <span data-ttu-id="2822f-837">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-837">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2822f-838">字符串</span><span class="sxs-lookup"><span data-stu-id="2822f-838">String</span></span>|<span data-ttu-id="2822f-839">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="2822f-839">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2822f-840">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-840">Requirements</span></span>

|<span data-ttu-id="2822f-841">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-841">Requirement</span></span>| <span data-ttu-id="2822f-842">值</span><span class="sxs-lookup"><span data-stu-id="2822f-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-843">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-844">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-844">1.0</span></span>|
|[<span data-ttu-id="2822f-845">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-845">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-846">ReadItem</span></span>|
|[<span data-ttu-id="2822f-847">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-847">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-848">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2822f-849">返回：</span><span class="sxs-lookup"><span data-stu-id="2822f-849">Returns:</span></span>

<span data-ttu-id="2822f-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="2822f-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="2822f-852">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="2822f-852">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="2822f-853">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2822f-853">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2822f-854">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="2822f-854">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-855">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2822f-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="2822f-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2822f-859">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="2822f-859">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2822f-860">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="2822f-860">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="2822f-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="2822f-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2822f-863">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-863">Requirements</span></span>

|<span data-ttu-id="2822f-864">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-864">Requirement</span></span>| <span data-ttu-id="2822f-865">值</span><span class="sxs-lookup"><span data-stu-id="2822f-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-866">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-867">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-867">1.0</span></span>|
|[<span data-ttu-id="2822f-868">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-869">ReadItem</span></span>|
|[<span data-ttu-id="2822f-870">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-871">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2822f-872">返回：</span><span class="sxs-lookup"><span data-stu-id="2822f-872">Returns:</span></span>

<span data-ttu-id="2822f-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="2822f-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="2822f-875">类型：对象</span><span class="sxs-lookup"><span data-stu-id="2822f-875">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="2822f-876">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-876">Example</span></span>

<span data-ttu-id="2822f-877">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="2822f-877">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2822f-878">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="2822f-878">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2822f-879">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="2822f-879">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2822f-880">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2822f-880">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2822f-881">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="2822f-881">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2822f-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="2822f-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-884">Parameters</span><span class="sxs-lookup"><span data-stu-id="2822f-884">Parameters</span></span>

|<span data-ttu-id="2822f-885">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-885">Name</span></span>| <span data-ttu-id="2822f-886">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-886">Type</span></span>| <span data-ttu-id="2822f-887">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2822f-888">String</span><span class="sxs-lookup"><span data-stu-id="2822f-888">String</span></span>|<span data-ttu-id="2822f-889">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="2822f-889">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2822f-890">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-890">Requirements</span></span>

|<span data-ttu-id="2822f-891">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-891">Requirement</span></span>| <span data-ttu-id="2822f-892">值</span><span class="sxs-lookup"><span data-stu-id="2822f-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-893">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-894">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-894">1.0</span></span>|
|[<span data-ttu-id="2822f-895">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-896">ReadItem</span></span>|
|[<span data-ttu-id="2822f-897">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-898">阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2822f-899">返回：</span><span class="sxs-lookup"><span data-stu-id="2822f-899">Returns:</span></span>

<span data-ttu-id="2822f-900">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="2822f-900">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="2822f-901">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="2822f-901">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="2822f-902">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-902">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2822f-903">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2822f-903">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2822f-904">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="2822f-904">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2822f-p158">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="2822f-p158">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-908">参数</span><span class="sxs-lookup"><span data-stu-id="2822f-908">Parameters</span></span>

|<span data-ttu-id="2822f-909">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-909">Name</span></span>| <span data-ttu-id="2822f-910">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-910">Type</span></span>| <span data-ttu-id="2822f-911">属性</span><span class="sxs-lookup"><span data-stu-id="2822f-911">Attributes</span></span>| <span data-ttu-id="2822f-912">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-912">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2822f-913">函数</span><span class="sxs-lookup"><span data-stu-id="2822f-913">function</span></span>||<span data-ttu-id="2822f-914">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-914">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2822f-915">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="2822f-915">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2822f-916">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="2822f-916">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="2822f-917">对象</span><span class="sxs-lookup"><span data-stu-id="2822f-917">Object</span></span>| <span data-ttu-id="2822f-918">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-918">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-919">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-919">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2822f-920">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="2822f-920">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2822f-921">Requirements</span><span class="sxs-lookup"><span data-stu-id="2822f-921">Requirements</span></span>

|<span data-ttu-id="2822f-922">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-922">Requirement</span></span>| <span data-ttu-id="2822f-923">值</span><span class="sxs-lookup"><span data-stu-id="2822f-923">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-924">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-924">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-925">1.0</span><span class="sxs-lookup"><span data-stu-id="2822f-925">1.0</span></span>|
|[<span data-ttu-id="2822f-926">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-926">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-927">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2822f-927">ReadItem</span></span>|
|[<span data-ttu-id="2822f-928">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-928">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-929">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2822f-929">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-930">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-930">Example</span></span>

<span data-ttu-id="2822f-p161">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="2822f-p161">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2822f-934">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2822f-934">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2822f-935">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="2822f-935">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2822f-936">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="2822f-936">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="2822f-937">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="2822f-937">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="2822f-938">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="2822f-938">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="2822f-939">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="2822f-939">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2822f-940">Parameters</span><span class="sxs-lookup"><span data-stu-id="2822f-940">Parameters</span></span>

|<span data-ttu-id="2822f-941">名称</span><span class="sxs-lookup"><span data-stu-id="2822f-941">Name</span></span>| <span data-ttu-id="2822f-942">类型</span><span class="sxs-lookup"><span data-stu-id="2822f-942">Type</span></span>| <span data-ttu-id="2822f-943">属性</span><span class="sxs-lookup"><span data-stu-id="2822f-943">Attributes</span></span>| <span data-ttu-id="2822f-944">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-944">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="2822f-945">字符串</span><span class="sxs-lookup"><span data-stu-id="2822f-945">String</span></span>||<span data-ttu-id="2822f-946">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="2822f-946">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="2822f-947">对象</span><span class="sxs-lookup"><span data-stu-id="2822f-947">Object</span></span>| <span data-ttu-id="2822f-948">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-948">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-949">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2822f-949">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2822f-950">对象</span><span class="sxs-lookup"><span data-stu-id="2822f-950">Object</span></span>| <span data-ttu-id="2822f-951">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-951">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-952">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2822f-952">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2822f-953">函数</span><span class="sxs-lookup"><span data-stu-id="2822f-953">function</span></span>| <span data-ttu-id="2822f-954">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2822f-954">&lt;optional&gt;</span></span>|<span data-ttu-id="2822f-955">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2822f-955">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2822f-956">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="2822f-956">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2822f-957">错误</span><span class="sxs-lookup"><span data-stu-id="2822f-957">Errors</span></span>

| <span data-ttu-id="2822f-958">错误代码</span><span class="sxs-lookup"><span data-stu-id="2822f-958">Error code</span></span> | <span data-ttu-id="2822f-959">说明</span><span class="sxs-lookup"><span data-stu-id="2822f-959">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="2822f-960">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="2822f-960">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2822f-961">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-961">Requirements</span></span>

|<span data-ttu-id="2822f-962">要求</span><span class="sxs-lookup"><span data-stu-id="2822f-962">Requirement</span></span>| <span data-ttu-id="2822f-963">值</span><span class="sxs-lookup"><span data-stu-id="2822f-963">Value</span></span>|
|---|---|
|[<span data-ttu-id="2822f-964">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2822f-964">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2822f-965">1.1</span><span class="sxs-lookup"><span data-stu-id="2822f-965">1.1</span></span>|
|[<span data-ttu-id="2822f-966">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2822f-966">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2822f-967">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2822f-967">ReadWriteItem</span></span>|
|[<span data-ttu-id="2822f-968">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2822f-968">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2822f-969">撰写</span><span class="sxs-lookup"><span data-stu-id="2822f-969">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2822f-970">示例</span><span class="sxs-lookup"><span data-stu-id="2822f-970">Example</span></span>

<span data-ttu-id="2822f-971">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="2822f-971">The following code removes an attachment with an identifier of '0'.</span></span>

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
