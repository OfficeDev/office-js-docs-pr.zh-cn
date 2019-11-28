---
title: "\"Context\"-\"邮箱\"。项目-要求集1。2"
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: ab8c55d2f91b250b419c7c9c71fc044b6fa68279
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629207"
---
# <a name="item"></a><span data-ttu-id="1fadc-102">item</span><span class="sxs-lookup"><span data-stu-id="1fadc-102">item</span></span>

### <span data-ttu-id="1fadc-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="1fadc-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="1fadc-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-107">Requirements</span></span>

|<span data-ttu-id="1fadc-108">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-108">Requirement</span></span>| <span data-ttu-id="1fadc-109">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-111">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-111">1.0</span></span>|
|[<span data-ttu-id="1fadc-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-113">受限</span><span class="sxs-lookup"><span data-stu-id="1fadc-113">Restricted</span></span>|
|[<span data-ttu-id="1fadc-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1fadc-116">成员和方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-116">Members and methods</span></span>

| <span data-ttu-id="1fadc-117">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-117">Member</span></span> | <span data-ttu-id="1fadc-118">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1fadc-119">attachments</span><span class="sxs-lookup"><span data-stu-id="1fadc-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="1fadc-120">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-120">Member</span></span> |
| [<span data-ttu-id="1fadc-121">bcc</span><span class="sxs-lookup"><span data-stu-id="1fadc-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="1fadc-122">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-122">Member</span></span> |
| [<span data-ttu-id="1fadc-123">body</span><span class="sxs-lookup"><span data-stu-id="1fadc-123">body</span></span>](#body-body) | <span data-ttu-id="1fadc-124">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-124">Member</span></span> |
| [<span data-ttu-id="1fadc-125">cc</span><span class="sxs-lookup"><span data-stu-id="1fadc-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1fadc-126">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-126">Member</span></span> |
| [<span data-ttu-id="1fadc-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="1fadc-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1fadc-128">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-128">Member</span></span> |
| [<span data-ttu-id="1fadc-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1fadc-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1fadc-130">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-130">Member</span></span> |
| [<span data-ttu-id="1fadc-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1fadc-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1fadc-132">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-132">Member</span></span> |
| [<span data-ttu-id="1fadc-133">end</span><span class="sxs-lookup"><span data-stu-id="1fadc-133">end</span></span>](#end-datetime) | <span data-ttu-id="1fadc-134">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-134">Member</span></span> |
| [<span data-ttu-id="1fadc-135">from</span><span class="sxs-lookup"><span data-stu-id="1fadc-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="1fadc-136">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-136">Member</span></span> |
| [<span data-ttu-id="1fadc-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1fadc-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1fadc-138">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-138">Member</span></span> |
| [<span data-ttu-id="1fadc-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="1fadc-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1fadc-140">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-140">Member</span></span> |
| [<span data-ttu-id="1fadc-141">itemId</span><span class="sxs-lookup"><span data-stu-id="1fadc-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1fadc-142">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-142">Member</span></span> |
| [<span data-ttu-id="1fadc-143">itemType</span><span class="sxs-lookup"><span data-stu-id="1fadc-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="1fadc-144">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-144">Member</span></span> |
| [<span data-ttu-id="1fadc-145">location</span><span class="sxs-lookup"><span data-stu-id="1fadc-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="1fadc-146">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-146">Member</span></span> |
| [<span data-ttu-id="1fadc-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1fadc-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1fadc-148">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-148">Member</span></span> |
| [<span data-ttu-id="1fadc-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1fadc-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1fadc-150">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-150">Member</span></span> |
| [<span data-ttu-id="1fadc-151">organizer</span><span class="sxs-lookup"><span data-stu-id="1fadc-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="1fadc-152">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-152">Member</span></span> |
| [<span data-ttu-id="1fadc-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1fadc-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1fadc-154">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-154">Member</span></span> |
| [<span data-ttu-id="1fadc-155">sender</span><span class="sxs-lookup"><span data-stu-id="1fadc-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="1fadc-156">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-156">Member</span></span> |
| [<span data-ttu-id="1fadc-157">start</span><span class="sxs-lookup"><span data-stu-id="1fadc-157">start</span></span>](#start-datetime) | <span data-ttu-id="1fadc-158">Member</span><span class="sxs-lookup"><span data-stu-id="1fadc-158">Member</span></span> |
| [<span data-ttu-id="1fadc-159">subject</span><span class="sxs-lookup"><span data-stu-id="1fadc-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="1fadc-160">成员</span><span class="sxs-lookup"><span data-stu-id="1fadc-160">Member</span></span> |
| [<span data-ttu-id="1fadc-161">to</span><span class="sxs-lookup"><span data-stu-id="1fadc-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1fadc-162">成员</span><span class="sxs-lookup"><span data-stu-id="1fadc-162">Member</span></span> |
| [<span data-ttu-id="1fadc-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1fadc-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1fadc-164">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-164">Method</span></span> |
| [<span data-ttu-id="1fadc-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1fadc-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1fadc-166">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-166">Method</span></span> |
| [<span data-ttu-id="1fadc-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1fadc-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="1fadc-168">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-168">Method</span></span> |
| [<span data-ttu-id="1fadc-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1fadc-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="1fadc-170">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-170">Method</span></span> |
| [<span data-ttu-id="1fadc-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="1fadc-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="1fadc-172">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-172">Method</span></span> |
| [<span data-ttu-id="1fadc-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1fadc-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1fadc-174">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-174">Method</span></span> |
| [<span data-ttu-id="1fadc-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1fadc-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1fadc-176">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-176">Method</span></span> |
| [<span data-ttu-id="1fadc-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1fadc-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1fadc-178">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-178">Method</span></span> |
| [<span data-ttu-id="1fadc-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1fadc-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1fadc-180">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-180">Method</span></span> |
| [<span data-ttu-id="1fadc-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1fadc-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1fadc-182">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-182">Method</span></span> |
| [<span data-ttu-id="1fadc-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1fadc-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1fadc-184">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-184">Method</span></span> |
| [<span data-ttu-id="1fadc-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1fadc-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1fadc-186">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-186">Method</span></span> |
| [<span data-ttu-id="1fadc-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1fadc-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1fadc-188">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1fadc-189">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-189">Example</span></span>

<span data-ttu-id="1fadc-190">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="1fadc-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1fadc-191">Members</span><span class="sxs-lookup"><span data-stu-id="1fadc-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="1fadc-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="1fadc-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="1fadc-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-195">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="1fadc-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1fadc-196">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="1fadc-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-197">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-197">Type</span></span>

*   <span data-ttu-id="1fadc-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="1fadc-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-199">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-199">Requirements</span></span>

|<span data-ttu-id="1fadc-200">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-200">Requirement</span></span>| <span data-ttu-id="1fadc-201">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-202">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-203">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-203">1.0</span></span>|
|[<span data-ttu-id="1fadc-204">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-205">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-206">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-207">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-208">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-208">Example</span></span>

<span data-ttu-id="1fadc-209">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="1fadc-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="1fadc-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-211">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1fadc-212">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-212">Compose mode only.</span></span>

<span data-ttu-id="1fadc-213">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-213">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-214">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1fadc-214">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1fadc-215">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-215">Get 500 members maximum.</span></span>
- <span data-ttu-id="1fadc-216">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-216">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-217">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-217">Type</span></span>

*   [<span data-ttu-id="1fadc-218">收件人</span><span class="sxs-lookup"><span data-stu-id="1fadc-218">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="1fadc-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-219">Requirements</span></span>

|<span data-ttu-id="1fadc-220">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-220">Requirement</span></span>| <span data-ttu-id="1fadc-221">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-222">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-223">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadc-223">1.1</span></span>|
|[<span data-ttu-id="1fadc-224">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-224">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-225">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-227">撰写</span><span class="sxs-lookup"><span data-stu-id="1fadc-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-228">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-228">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="1fadc-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-229">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-230">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-231">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-231">Type</span></span>

*   [<span data-ttu-id="1fadc-232">Body</span><span class="sxs-lookup"><span data-stu-id="1fadc-232">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="1fadc-233">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-233">Requirements</span></span>

|<span data-ttu-id="1fadc-234">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-234">Requirement</span></span>| <span data-ttu-id="1fadc-235">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-236">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-237">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadc-237">1.1</span></span>|
|[<span data-ttu-id="1fadc-238">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-239">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-240">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-241">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-242">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-242">Example</span></span>

<span data-ttu-id="1fadc-243">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="1fadc-243">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="1fadc-244">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="1fadc-244">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="1fadc-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-245">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-246">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1fadc-246">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1fadc-247">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-247">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-248">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-248">Read mode</span></span>

<span data-ttu-id="1fadc-249">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-249">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="1fadc-250">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-251">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-251">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-252">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-252">Compose mode</span></span>

<span data-ttu-id="1fadc-253">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="1fadc-254">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-254">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-255">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1fadc-255">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1fadc-256">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-256">Get 500 members maximum.</span></span>
- <span data-ttu-id="1fadc-257">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-257">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1fadc-258">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-258">Type</span></span>

*   <span data-ttu-id="1fadc-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-259">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-260">Requirements</span></span>

|<span data-ttu-id="1fadc-261">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-261">Requirement</span></span>| <span data-ttu-id="1fadc-262">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-264">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-264">1.0</span></span>|
|[<span data-ttu-id="1fadc-265">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-266">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-267">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-268">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-268">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="1fadc-269">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="1fadc-269">(nullable) conversationId: String</span></span>

<span data-ttu-id="1fadc-270">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-270">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1fadc-p110">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p110">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1fadc-p111">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p111">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-275">Type</span><span class="sxs-lookup"><span data-stu-id="1fadc-275">Type</span></span>

*   <span data-ttu-id="1fadc-276">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-277">Requirements</span></span>

|<span data-ttu-id="1fadc-278">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-278">Requirement</span></span>| <span data-ttu-id="1fadc-279">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-281">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-281">1.0</span></span>|
|[<span data-ttu-id="1fadc-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-283">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-285">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-285">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-286">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-286">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="1fadc-287">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="1fadc-287">dateTimeCreated: Date</span></span>

<span data-ttu-id="1fadc-p112">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p112">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-290">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-290">Type</span></span>

*   <span data-ttu-id="1fadc-291">日期</span><span class="sxs-lookup"><span data-stu-id="1fadc-291">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-292">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-292">Requirements</span></span>

|<span data-ttu-id="1fadc-293">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-293">Requirement</span></span>| <span data-ttu-id="1fadc-294">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-294">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-295">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-295">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-296">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-296">1.0</span></span>|
|[<span data-ttu-id="1fadc-297">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-297">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-298">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-298">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-299">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-299">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-300">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-300">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-301">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-301">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="1fadc-302">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="1fadc-302">dateTimeModified: Date</span></span>

<span data-ttu-id="1fadc-p113">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p113">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-305">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-305">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-306">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-306">Type</span></span>

*   <span data-ttu-id="1fadc-307">日期</span><span class="sxs-lookup"><span data-stu-id="1fadc-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-308">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-308">Requirements</span></span>

|<span data-ttu-id="1fadc-309">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-309">Requirement</span></span>| <span data-ttu-id="1fadc-310">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-312">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-312">1.0</span></span>|
|[<span data-ttu-id="1fadc-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-314">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-316">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-317">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-317">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="1fadc-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-318">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-319">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1fadc-319">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1fadc-p114">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p114">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-322">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-322">Read mode</span></span>

<span data-ttu-id="1fadc-323">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-323">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-324">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-324">Compose mode</span></span>

<span data-ttu-id="1fadc-325">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-325">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1fadc-326">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="1fadc-326">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1fadc-327">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="1fadc-327">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1fadc-328">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-328">Type</span></span>

*   <span data-ttu-id="1fadc-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-329">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-330">Requirements</span></span>

|<span data-ttu-id="1fadc-331">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-331">Requirement</span></span>| <span data-ttu-id="1fadc-332">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-333">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-334">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-334">1.0</span></span>|
|[<span data-ttu-id="1fadc-335">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-336">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-337">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-338">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-338">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="1fadc-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-339">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-p115">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p115">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1fadc-p116">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p116">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-344">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-344">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-345">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-345">Type</span></span>

*   [<span data-ttu-id="1fadc-346">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1fadc-346">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="1fadc-347">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-347">Requirements</span></span>

|<span data-ttu-id="1fadc-348">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-348">Requirement</span></span>| <span data-ttu-id="1fadc-349">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-350">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-351">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-351">1.0</span></span>|
|[<span data-ttu-id="1fadc-352">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-353">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-354">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-355">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-355">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-356">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-356">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="1fadc-357">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="1fadc-357">internetMessageId: String</span></span>

<span data-ttu-id="1fadc-p117">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p117">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-360">Type</span><span class="sxs-lookup"><span data-stu-id="1fadc-360">Type</span></span>

*   <span data-ttu-id="1fadc-361">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-361">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-362">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-362">Requirements</span></span>

|<span data-ttu-id="1fadc-363">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-363">Requirement</span></span>| <span data-ttu-id="1fadc-364">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-364">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-365">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-365">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-366">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-366">1.0</span></span>|
|[<span data-ttu-id="1fadc-367">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-367">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-368">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-368">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-369">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-369">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-370">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-370">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-371">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-371">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="1fadc-372">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="1fadc-372">itemClass: String</span></span>

<span data-ttu-id="1fadc-p118">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p118">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1fadc-p119">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p119">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1fadc-377">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-377">Type</span></span> | <span data-ttu-id="1fadc-378">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-378">Description</span></span> | <span data-ttu-id="1fadc-379">项目类</span><span class="sxs-lookup"><span data-stu-id="1fadc-379">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1fadc-380">约会项目</span><span class="sxs-lookup"><span data-stu-id="1fadc-380">Appointment items</span></span> | <span data-ttu-id="1fadc-381">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="1fadc-381">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="1fadc-382">邮件项目</span><span class="sxs-lookup"><span data-stu-id="1fadc-382">Message items</span></span> | <span data-ttu-id="1fadc-383">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="1fadc-383">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1fadc-384">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-384">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-385">Type</span><span class="sxs-lookup"><span data-stu-id="1fadc-385">Type</span></span>

*   <span data-ttu-id="1fadc-386">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-386">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-387">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-387">Requirements</span></span>

|<span data-ttu-id="1fadc-388">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-388">Requirement</span></span>| <span data-ttu-id="1fadc-389">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-390">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-391">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-391">1.0</span></span>|
|[<span data-ttu-id="1fadc-392">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-393">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-394">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-395">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-395">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-396">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-396">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1fadc-397">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="1fadc-397">(nullable) itemId: String</span></span>

<span data-ttu-id="1fadc-p120">获取当前项目的 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p120">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-400">`itemId` 属性返回的标识符与 [Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)相同。</span><span class="sxs-lookup"><span data-stu-id="1fadc-400">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="1fadc-401">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="1fadc-401">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1fadc-402">在使用此值进行 REST API 调用之前，应使用`Office.context.mailbox.convertToRestId`转换它，这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="1fadc-402">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="1fadc-403">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="1fadc-403">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-404">Type</span><span class="sxs-lookup"><span data-stu-id="1fadc-404">Type</span></span>

*   <span data-ttu-id="1fadc-405">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-405">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-406">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-406">Requirements</span></span>

|<span data-ttu-id="1fadc-407">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-407">Requirement</span></span>| <span data-ttu-id="1fadc-408">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-409">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-410">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-410">1.0</span></span>|
|[<span data-ttu-id="1fadc-411">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-412">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-413">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-414">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-415">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-415">Example</span></span>

<span data-ttu-id="1fadc-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="1fadc-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-418">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-419">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="1fadc-419">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1fadc-420">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="1fadc-420">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-421">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-421">Type</span></span>

*   [<span data-ttu-id="1fadc-422">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1fadc-422">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="1fadc-423">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-423">Requirements</span></span>

|<span data-ttu-id="1fadc-424">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-424">Requirement</span></span>| <span data-ttu-id="1fadc-425">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-426">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-427">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-427">1.0</span></span>|
|[<span data-ttu-id="1fadc-428">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-429">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-430">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-431">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-432">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-432">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="1fadc-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-433">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-434">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="1fadc-434">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-435">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-435">Read mode</span></span>

<span data-ttu-id="1fadc-436">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="1fadc-436">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-437">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-437">Compose mode</span></span>

<span data-ttu-id="1fadc-438">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-438">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1fadc-439">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-439">Type</span></span>

*   <span data-ttu-id="1fadc-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-440">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-441">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-441">Requirements</span></span>

|<span data-ttu-id="1fadc-442">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-442">Requirement</span></span>| <span data-ttu-id="1fadc-443">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-444">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-445">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-445">1.0</span></span>|
|[<span data-ttu-id="1fadc-446">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-447">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-448">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-449">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-449">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1fadc-450">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="1fadc-450">normalizedSubject: String</span></span>

<span data-ttu-id="1fadc-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1fadc-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-455">Type</span><span class="sxs-lookup"><span data-stu-id="1fadc-455">Type</span></span>

*   <span data-ttu-id="1fadc-456">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-457">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-457">Requirements</span></span>

|<span data-ttu-id="1fadc-458">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-458">Requirement</span></span>| <span data-ttu-id="1fadc-459">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-460">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-461">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-461">1.0</span></span>|
|[<span data-ttu-id="1fadc-462">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-462">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-463">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-464">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-464">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-465">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-466">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="1fadc-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-467">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-468">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1fadc-468">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1fadc-469">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-469">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-470">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-470">Read mode</span></span>

<span data-ttu-id="1fadc-471">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-471">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="1fadc-472">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-472">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-473">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-473">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-474">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-474">Compose mode</span></span>

<span data-ttu-id="1fadc-475">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-475">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="1fadc-476">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-476">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-477">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1fadc-477">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1fadc-478">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-478">Get 500 members maximum.</span></span>
- <span data-ttu-id="1fadc-479">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-479">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1fadc-480">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-480">Type</span></span>

*   <span data-ttu-id="1fadc-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-481">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-482">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-482">Requirements</span></span>

|<span data-ttu-id="1fadc-483">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-483">Requirement</span></span>| <span data-ttu-id="1fadc-484">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-484">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-485">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-485">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-486">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-486">1.0</span></span>|
|[<span data-ttu-id="1fadc-487">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-487">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-488">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-488">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-489">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-489">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-490">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-490">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="1fadc-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-491">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-p128">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-494">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-494">Type</span></span>

*   [<span data-ttu-id="1fadc-495">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1fadc-495">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="1fadc-496">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-496">Requirements</span></span>

|<span data-ttu-id="1fadc-497">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-497">Requirement</span></span>| <span data-ttu-id="1fadc-498">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-499">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-500">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-500">1.0</span></span>|
|[<span data-ttu-id="1fadc-501">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-502">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-503">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-504">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-504">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-505">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-505">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="1fadc-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-506">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-507">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1fadc-507">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1fadc-508">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-508">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-509">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-509">Read mode</span></span>

<span data-ttu-id="1fadc-510">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-510">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="1fadc-511">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-511">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-512">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-512">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-513">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-513">Compose mode</span></span>

<span data-ttu-id="1fadc-514">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-514">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="1fadc-515">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-515">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-516">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1fadc-516">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1fadc-517">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-517">Get 500 members maximum.</span></span>
- <span data-ttu-id="1fadc-518">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-518">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="1fadc-519">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-519">Type</span></span>

*   <span data-ttu-id="1fadc-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-521">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-521">Requirements</span></span>

|<span data-ttu-id="1fadc-522">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-522">Requirement</span></span>| <span data-ttu-id="1fadc-523">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-524">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-525">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-525">1.0</span></span>|
|[<span data-ttu-id="1fadc-526">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-527">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-528">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-529">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="1fadc-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-p132">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1fadc-p133">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-535">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadc-536">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-536">Type</span></span>

*   [<span data-ttu-id="1fadc-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1fadc-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="1fadc-538">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-538">Requirements</span></span>

|<span data-ttu-id="1fadc-539">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-539">Requirement</span></span>| <span data-ttu-id="1fadc-540">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-541">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-542">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-542">1.0</span></span>|
|[<span data-ttu-id="1fadc-543">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-544">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-545">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-546">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-547">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="1fadc-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-549">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1fadc-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1fadc-p134">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-552">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-552">Read mode</span></span>

<span data-ttu-id="1fadc-553">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-554">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-554">Compose mode</span></span>

<span data-ttu-id="1fadc-555">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1fadc-556">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="1fadc-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="1fadc-557">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="1fadc-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1fadc-558">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-558">Type</span></span>

*   <span data-ttu-id="1fadc-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-560">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-560">Requirements</span></span>

|<span data-ttu-id="1fadc-561">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-561">Requirement</span></span>| <span data-ttu-id="1fadc-562">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-563">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-564">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-564">1.0</span></span>|
|[<span data-ttu-id="1fadc-565">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-566">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-568">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="1fadc-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-570">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="1fadc-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1fadc-571">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="1fadc-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-572">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-572">Read mode</span></span>

<span data-ttu-id="1fadc-p136">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p136">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-575">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-575">Compose mode</span></span>

<span data-ttu-id="1fadc-576">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="1fadc-577">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-577">Type</span></span>

*   <span data-ttu-id="1fadc-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-579">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-579">Requirements</span></span>

|<span data-ttu-id="1fadc-580">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-580">Requirement</span></span>| <span data-ttu-id="1fadc-581">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-582">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-583">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-583">1.0</span></span>|
|[<span data-ttu-id="1fadc-584">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-585">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-586">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-587">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="1fadc-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="1fadc-589">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1fadc-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1fadc-590">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1fadc-591">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-591">Read mode</span></span>

<span data-ttu-id="1fadc-592">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-592">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="1fadc-593">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-593">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-594">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-594">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="1fadc-595">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-595">Compose mode</span></span>

<span data-ttu-id="1fadc-596">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-596">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="1fadc-597">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-597">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1fadc-598">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1fadc-598">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1fadc-599">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-599">Get 500 members maximum.</span></span>
- <span data-ttu-id="1fadc-600">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1fadc-600">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1fadc-601">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-601">Type</span></span>

*   <span data-ttu-id="1fadc-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-602">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-603">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-603">Requirements</span></span>

|<span data-ttu-id="1fadc-604">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-604">Requirement</span></span>| <span data-ttu-id="1fadc-605">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-606">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-607">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-607">1.0</span></span>|
|[<span data-ttu-id="1fadc-608">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-609">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-610">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-611">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-611">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1fadc-612">方法</span><span class="sxs-lookup"><span data-stu-id="1fadc-612">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1fadc-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1fadc-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1fadc-614">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="1fadc-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1fadc-615">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="1fadc-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1fadc-616">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="1fadc-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-617">参数</span><span class="sxs-lookup"><span data-stu-id="1fadc-617">Parameters</span></span>

|<span data-ttu-id="1fadc-618">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-618">Name</span></span>| <span data-ttu-id="1fadc-619">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-619">Type</span></span>| <span data-ttu-id="1fadc-620">属性</span><span class="sxs-lookup"><span data-stu-id="1fadc-620">Attributes</span></span>| <span data-ttu-id="1fadc-621">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1fadc-622">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-622">String</span></span>||<span data-ttu-id="1fadc-p140">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p140">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1fadc-625">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-625">String</span></span>||<span data-ttu-id="1fadc-p141">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p141">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1fadc-628">Object</span><span class="sxs-lookup"><span data-stu-id="1fadc-628">Object</span></span>| <span data-ttu-id="1fadc-629">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-629">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-630">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1fadc-630">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fadc-631">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-631">Object</span></span>| <span data-ttu-id="1fadc-632">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-632">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-633">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-633">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1fadc-634">函数</span><span class="sxs-lookup"><span data-stu-id="1fadc-634">function</span></span>| <span data-ttu-id="1fadc-635">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-635">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-636">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-636">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1fadc-637">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1fadc-637">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1fadc-638">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-638">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fadc-639">错误</span><span class="sxs-lookup"><span data-stu-id="1fadc-639">Errors</span></span>

| <span data-ttu-id="1fadc-640">错误代码</span><span class="sxs-lookup"><span data-stu-id="1fadc-640">Error code</span></span> | <span data-ttu-id="1fadc-641">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-641">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1fadc-642">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="1fadc-642">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1fadc-643">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="1fadc-643">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1fadc-644">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="1fadc-644">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fadc-645">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-645">Requirements</span></span>

|<span data-ttu-id="1fadc-646">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-646">Requirement</span></span>| <span data-ttu-id="1fadc-647">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-648">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-649">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadc-649">1.1</span></span>|
|[<span data-ttu-id="1fadc-650">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-650">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-651">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-651">ReadWriteItem</span></span>|
|[<span data-ttu-id="1fadc-652">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-652">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-653">撰写</span><span class="sxs-lookup"><span data-stu-id="1fadc-653">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-654">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-654">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1fadc-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1fadc-655">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1fadc-656">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="1fadc-656">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1fadc-p142">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p142">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1fadc-660">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="1fadc-660">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1fadc-661">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="1fadc-661">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-662">Parameters</span><span class="sxs-lookup"><span data-stu-id="1fadc-662">Parameters</span></span>

|<span data-ttu-id="1fadc-663">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-663">Name</span></span>| <span data-ttu-id="1fadc-664">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-664">Type</span></span>| <span data-ttu-id="1fadc-665">属性</span><span class="sxs-lookup"><span data-stu-id="1fadc-665">Attributes</span></span>| <span data-ttu-id="1fadc-666">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-666">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1fadc-667">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-667">String</span></span>||<span data-ttu-id="1fadc-p143">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p143">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1fadc-670">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-670">String</span></span>||<span data-ttu-id="1fadc-671">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="1fadc-671">The subject of the item to be attached.</span></span> <span data-ttu-id="1fadc-672">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-672">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1fadc-673">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-673">Object</span></span>| <span data-ttu-id="1fadc-674">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-674">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-675">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1fadc-675">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fadc-676">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-676">Object</span></span>| <span data-ttu-id="1fadc-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-677">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-678">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-678">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1fadc-679">函数</span><span class="sxs-lookup"><span data-stu-id="1fadc-679">function</span></span>| <span data-ttu-id="1fadc-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-680">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-681">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-681">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1fadc-682">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1fadc-682">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1fadc-683">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-683">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fadc-684">错误</span><span class="sxs-lookup"><span data-stu-id="1fadc-684">Errors</span></span>

| <span data-ttu-id="1fadc-685">错误代码</span><span class="sxs-lookup"><span data-stu-id="1fadc-685">Error code</span></span> | <span data-ttu-id="1fadc-686">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-686">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1fadc-687">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="1fadc-687">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fadc-688">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-688">Requirements</span></span>

|<span data-ttu-id="1fadc-689">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-689">Requirement</span></span>| <span data-ttu-id="1fadc-690">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-691">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-692">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadc-692">1.1</span></span>|
|[<span data-ttu-id="1fadc-693">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-694">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-694">ReadWriteItem</span></span>|
|[<span data-ttu-id="1fadc-695">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-696">撰写</span><span class="sxs-lookup"><span data-stu-id="1fadc-696">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-697">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-697">Example</span></span>

<span data-ttu-id="1fadc-698">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="1fadc-698">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="1fadc-699">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1fadc-699">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="1fadc-700">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="1fadc-700">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-701">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-701">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fadc-702">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="1fadc-702">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1fadc-703">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="1fadc-703">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1fadc-p145">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-707">Parameters</span><span class="sxs-lookup"><span data-stu-id="1fadc-707">Parameters</span></span>

|<span data-ttu-id="1fadc-708">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-708">Name</span></span>| <span data-ttu-id="1fadc-709">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-709">Type</span></span>| <span data-ttu-id="1fadc-710">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-710">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1fadc-711">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-711">String &#124; Object</span></span>| |<span data-ttu-id="1fadc-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1fadc-714">**或**</span><span class="sxs-lookup"><span data-stu-id="1fadc-714">**OR**</span></span><br/><span data-ttu-id="1fadc-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1fadc-717">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-717">String</span></span> | <span data-ttu-id="1fadc-718">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-718">&lt;optional&gt;</span></span> | <span data-ttu-id="1fadc-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1fadc-721">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-721">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1fadc-722">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-722">&lt;optional&gt;</span></span> | <span data-ttu-id="1fadc-723">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-723">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1fadc-724">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-724">String</span></span> | | <span data-ttu-id="1fadc-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1fadc-727">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-727">String</span></span> | | <span data-ttu-id="1fadc-728">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-728">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1fadc-729">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-729">String</span></span> | | <span data-ttu-id="1fadc-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1fadc-732">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-732">String</span></span> | | <span data-ttu-id="1fadc-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1fadc-736">函数</span><span class="sxs-lookup"><span data-stu-id="1fadc-736">function</span></span> | <span data-ttu-id="1fadc-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-737">&lt;optional&gt;</span></span> | <span data-ttu-id="1fadc-738">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-738">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fadc-739">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-739">Requirements</span></span>

|<span data-ttu-id="1fadc-740">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-740">Requirement</span></span>| <span data-ttu-id="1fadc-741">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-742">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-743">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-743">1.0</span></span>|
|[<span data-ttu-id="1fadc-744">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-744">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-745">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-745">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-746">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-746">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-747">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-747">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1fadc-748">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-748">Examples</span></span>

<span data-ttu-id="1fadc-749">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-749">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1fadc-750">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-750">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1fadc-751">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-751">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1fadc-752">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-752">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1fadc-753">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-753">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1fadc-754">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-754">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="1fadc-755">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1fadc-755">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="1fadc-756">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="1fadc-756">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-757">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-757">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fadc-758">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="1fadc-758">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1fadc-759">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="1fadc-759">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1fadc-p152">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-763">Parameters</span><span class="sxs-lookup"><span data-stu-id="1fadc-763">Parameters</span></span>

|<span data-ttu-id="1fadc-764">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-764">Name</span></span>| <span data-ttu-id="1fadc-765">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-765">Type</span></span>| <span data-ttu-id="1fadc-766">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-766">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="1fadc-767">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-767">String &#124; Object</span></span>| | <span data-ttu-id="1fadc-p153">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1fadc-770">**或**</span><span class="sxs-lookup"><span data-stu-id="1fadc-770">**OR**</span></span><br/><span data-ttu-id="1fadc-p154">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1fadc-773">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-773">String</span></span> | <span data-ttu-id="1fadc-774">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-774">&lt;optional&gt;</span></span> | <span data-ttu-id="1fadc-p155">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1fadc-777">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-777">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1fadc-778">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-778">&lt;optional&gt;</span></span> | <span data-ttu-id="1fadc-779">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-779">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1fadc-780">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-780">String</span></span> | | <span data-ttu-id="1fadc-p156">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1fadc-783">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-783">String</span></span> | | <span data-ttu-id="1fadc-784">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-784">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1fadc-785">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-785">String</span></span> | | <span data-ttu-id="1fadc-p157">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1fadc-788">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-788">String</span></span> | | <span data-ttu-id="1fadc-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1fadc-792">函数</span><span class="sxs-lookup"><span data-stu-id="1fadc-792">function</span></span> | <span data-ttu-id="1fadc-793">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-793">&lt;optional&gt;</span></span> | <span data-ttu-id="1fadc-794">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-794">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fadc-795">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-795">Requirements</span></span>

|<span data-ttu-id="1fadc-796">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-796">Requirement</span></span>| <span data-ttu-id="1fadc-797">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-797">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-798">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-798">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-799">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-799">1.0</span></span>|
|[<span data-ttu-id="1fadc-800">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-800">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-801">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-801">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-802">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-802">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-803">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-803">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1fadc-804">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-804">Examples</span></span>

<span data-ttu-id="1fadc-805">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-805">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1fadc-806">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-806">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1fadc-807">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-807">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1fadc-808">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-808">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1fadc-809">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-809">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1fadc-810">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="1fadc-810">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="1fadc-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="1fadc-811">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="1fadc-812">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="1fadc-812">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-813">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-813">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-814">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-814">Requirements</span></span>

|<span data-ttu-id="1fadc-815">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-815">Requirement</span></span>| <span data-ttu-id="1fadc-816">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-817">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-818">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-818">1.0</span></span>|
|[<span data-ttu-id="1fadc-819">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-820">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-821">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-822">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fadc-823">返回：</span><span class="sxs-lookup"><span data-stu-id="1fadc-823">Returns:</span></span>

<span data-ttu-id="1fadc-824">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="1fadc-824">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="1fadc-825">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-825">Example</span></span>

<span data-ttu-id="1fadc-826">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="1fadc-826">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="1fadc-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="1fadc-827">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="1fadc-828">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-828">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-829">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-829">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-830">Parameters</span><span class="sxs-lookup"><span data-stu-id="1fadc-830">Parameters</span></span>

|<span data-ttu-id="1fadc-831">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-831">Name</span></span>| <span data-ttu-id="1fadc-832">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-832">Type</span></span>| <span data-ttu-id="1fadc-833">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-833">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1fadc-834">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1fadc-834">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="1fadc-835">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="1fadc-835">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fadc-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-836">Requirements</span></span>

|<span data-ttu-id="1fadc-837">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-837">Requirement</span></span>| <span data-ttu-id="1fadc-838">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-839">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-840">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-840">1.0</span></span>|
|[<span data-ttu-id="1fadc-841">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-842">受限</span><span class="sxs-lookup"><span data-stu-id="1fadc-842">Restricted</span></span>|
|[<span data-ttu-id="1fadc-843">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-844">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fadc-845">返回：</span><span class="sxs-lookup"><span data-stu-id="1fadc-845">Returns:</span></span>

<span data-ttu-id="1fadc-846">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="1fadc-846">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1fadc-847">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-847">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1fadc-848">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="1fadc-848">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1fadc-849">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="1fadc-849">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1fadc-850">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="1fadc-850">Value of `entityType`</span></span> | <span data-ttu-id="1fadc-851">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-851">Type of objects in returned array</span></span> | <span data-ttu-id="1fadc-852">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-852">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1fadc-853">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-853">String</span></span> | <span data-ttu-id="1fadc-854">**受限**</span><span class="sxs-lookup"><span data-stu-id="1fadc-854">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1fadc-855">Contact</span><span class="sxs-lookup"><span data-stu-id="1fadc-855">Contact</span></span> | <span data-ttu-id="1fadc-856">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fadc-856">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1fadc-857">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-857">String</span></span> | <span data-ttu-id="1fadc-858">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fadc-858">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1fadc-859">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1fadc-859">MeetingSuggestion</span></span> | <span data-ttu-id="1fadc-860">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fadc-860">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1fadc-861">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1fadc-861">PhoneNumber</span></span> | <span data-ttu-id="1fadc-862">**受限**</span><span class="sxs-lookup"><span data-stu-id="1fadc-862">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1fadc-863">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1fadc-863">TaskSuggestion</span></span> | <span data-ttu-id="1fadc-864">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1fadc-864">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1fadc-865">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-865">String</span></span> | <span data-ttu-id="1fadc-866">**受限**</span><span class="sxs-lookup"><span data-stu-id="1fadc-866">**Restricted**</span></span> |

<span data-ttu-id="1fadc-867">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="1fadc-867">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="1fadc-868">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-868">Example</span></span>

<span data-ttu-id="1fadc-869">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-869">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="1fadc-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="1fadc-870">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="1fadc-871">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="1fadc-871">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-872">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-872">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fadc-873">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="1fadc-873">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-874">参数</span><span class="sxs-lookup"><span data-stu-id="1fadc-874">Parameters</span></span>

|<span data-ttu-id="1fadc-875">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-875">Name</span></span>| <span data-ttu-id="1fadc-876">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-876">Type</span></span>| <span data-ttu-id="1fadc-877">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-877">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1fadc-878">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-878">String</span></span>|<span data-ttu-id="1fadc-879">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="1fadc-879">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fadc-880">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-880">Requirements</span></span>

|<span data-ttu-id="1fadc-881">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-881">Requirement</span></span>| <span data-ttu-id="1fadc-882">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-883">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-884">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-884">1.0</span></span>|
|[<span data-ttu-id="1fadc-885">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-885">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-886">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-886">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-887">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-887">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-888">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-888">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fadc-889">返回：</span><span class="sxs-lookup"><span data-stu-id="1fadc-889">Returns:</span></span>

<span data-ttu-id="1fadc-p160">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1fadc-892">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="1fadc-892">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="1fadc-893">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1fadc-893">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1fadc-894">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="1fadc-894">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-895">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-895">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fadc-p161">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1fadc-899">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="1fadc-899">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1fadc-900">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-900">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="1fadc-p162">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadc-903">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-903">Requirements</span></span>

|<span data-ttu-id="1fadc-904">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-904">Requirement</span></span>| <span data-ttu-id="1fadc-905">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-906">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-907">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-907">1.0</span></span>|
|[<span data-ttu-id="1fadc-908">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-908">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-909">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-910">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-910">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-911">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-911">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fadc-912">返回：</span><span class="sxs-lookup"><span data-stu-id="1fadc-912">Returns:</span></span>

<span data-ttu-id="1fadc-p163">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="1fadc-915">类型：对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-915">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="1fadc-916">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-916">Example</span></span>

<span data-ttu-id="1fadc-917">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="1fadc-917">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1fadc-918">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="1fadc-918">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1fadc-919">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="1fadc-919">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1fadc-920">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1fadc-920">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1fadc-921">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="1fadc-921">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1fadc-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-924">Parameters</span><span class="sxs-lookup"><span data-stu-id="1fadc-924">Parameters</span></span>

|<span data-ttu-id="1fadc-925">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-925">Name</span></span>| <span data-ttu-id="1fadc-926">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-926">Type</span></span>| <span data-ttu-id="1fadc-927">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-927">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1fadc-928">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-928">String</span></span>|<span data-ttu-id="1fadc-929">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="1fadc-929">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fadc-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-930">Requirements</span></span>

|<span data-ttu-id="1fadc-931">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-931">Requirement</span></span>| <span data-ttu-id="1fadc-932">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-933">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-934">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-934">1.0</span></span>|
|[<span data-ttu-id="1fadc-935">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-935">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-936">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-937">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-937">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-938">阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fadc-939">返回：</span><span class="sxs-lookup"><span data-stu-id="1fadc-939">Returns:</span></span>

<span data-ttu-id="1fadc-940">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="1fadc-940">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="1fadc-941">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1fadc-941">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="1fadc-942">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-942">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1fadc-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1fadc-943">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1fadc-944">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="1fadc-944">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1fadc-p165">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回空字符串。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p165">If there is no selection but the cursor is in the body or subject, the method returns an empty string for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-947">Parameters</span><span class="sxs-lookup"><span data-stu-id="1fadc-947">Parameters</span></span>

|<span data-ttu-id="1fadc-948">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-948">Name</span></span>| <span data-ttu-id="1fadc-949">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-949">Type</span></span>| <span data-ttu-id="1fadc-950">属性</span><span class="sxs-lookup"><span data-stu-id="1fadc-950">Attributes</span></span>| <span data-ttu-id="1fadc-951">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-951">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="1fadc-952">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1fadc-952">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1fadc-p166">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="1fadc-956">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-956">Object</span></span>| <span data-ttu-id="1fadc-957">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-957">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-958">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1fadc-958">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fadc-959">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-959">Object</span></span>| <span data-ttu-id="1fadc-960">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-960">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-961">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-961">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1fadc-962">function</span><span class="sxs-lookup"><span data-stu-id="1fadc-962">function</span></span>||<span data-ttu-id="1fadc-963">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-963">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1fadc-964">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-964">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1fadc-965">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="1fadc-965">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fadc-966">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-966">Requirements</span></span>

|<span data-ttu-id="1fadc-967">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-967">Requirement</span></span>| <span data-ttu-id="1fadc-968">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-968">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-969">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-969">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-970">1.2</span><span class="sxs-lookup"><span data-stu-id="1fadc-970">1.2</span></span>|
|[<span data-ttu-id="1fadc-971">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-971">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-972">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-972">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-973">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-973">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-974">撰写</span><span class="sxs-lookup"><span data-stu-id="1fadc-974">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1fadc-975">返回：</span><span class="sxs-lookup"><span data-stu-id="1fadc-975">Returns:</span></span>

<span data-ttu-id="1fadc-976">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="1fadc-976">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="1fadc-977">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-977">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1fadc-978">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-978">Example</span></span>

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  console.log("Selected text in " + prop + ": " + text);
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1fadc-979">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1fadc-979">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1fadc-980">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="1fadc-980">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1fadc-p168">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-984">参数</span><span class="sxs-lookup"><span data-stu-id="1fadc-984">Parameters</span></span>

|<span data-ttu-id="1fadc-985">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-985">Name</span></span>| <span data-ttu-id="1fadc-986">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-986">Type</span></span>| <span data-ttu-id="1fadc-987">属性</span><span class="sxs-lookup"><span data-stu-id="1fadc-987">Attributes</span></span>| <span data-ttu-id="1fadc-988">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-988">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1fadc-989">函数</span><span class="sxs-lookup"><span data-stu-id="1fadc-989">function</span></span>||<span data-ttu-id="1fadc-990">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-990">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1fadc-991">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="1fadc-991">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1fadc-992">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="1fadc-992">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1fadc-993">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-993">Object</span></span>| <span data-ttu-id="1fadc-994">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-994">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-995">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-995">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1fadc-996">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="1fadc-996">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1fadc-997">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-997">Requirements</span></span>

|<span data-ttu-id="1fadc-998">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-998">Requirement</span></span>| <span data-ttu-id="1fadc-999">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-1000">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="1fadc-1001">1.0</span></span>|
|[<span data-ttu-id="1fadc-1002">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-1003">ReadItem</span></span>|
|[<span data-ttu-id="1fadc-1004">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-1005">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1fadc-1005">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-1006">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-1006">Example</span></span>

<span data-ttu-id="1fadc-p171">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1fadc-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1fadc-1010">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1fadc-1011">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1011">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1fadc-1012">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1012">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="1fadc-1013">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1013">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="1fadc-1014">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1014">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="1fadc-1015">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1015">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-1016">Parameters</span><span class="sxs-lookup"><span data-stu-id="1fadc-1016">Parameters</span></span>

|<span data-ttu-id="1fadc-1017">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-1017">Name</span></span>| <span data-ttu-id="1fadc-1018">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-1018">Type</span></span>| <span data-ttu-id="1fadc-1019">属性</span><span class="sxs-lookup"><span data-stu-id="1fadc-1019">Attributes</span></span>| <span data-ttu-id="1fadc-1020">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-1020">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1fadc-1021">String</span><span class="sxs-lookup"><span data-stu-id="1fadc-1021">String</span></span>||<span data-ttu-id="1fadc-1022">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1022">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="1fadc-1023">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-1023">Object</span></span>| <span data-ttu-id="1fadc-1024">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-1024">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-1025">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1025">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fadc-1026">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-1026">Object</span></span>| <span data-ttu-id="1fadc-1027">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-1027">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-1028">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1028">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1fadc-1029">函数</span><span class="sxs-lookup"><span data-stu-id="1fadc-1029">function</span></span>| <span data-ttu-id="1fadc-1030">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-1030">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-1031">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1031">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1fadc-1032">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1032">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1fadc-1033">错误</span><span class="sxs-lookup"><span data-stu-id="1fadc-1033">Errors</span></span>

| <span data-ttu-id="1fadc-1034">错误代码</span><span class="sxs-lookup"><span data-stu-id="1fadc-1034">Error code</span></span> | <span data-ttu-id="1fadc-1035">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-1035">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1fadc-1036">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1036">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fadc-1037">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-1037">Requirements</span></span>

|<span data-ttu-id="1fadc-1038">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-1038">Requirement</span></span>| <span data-ttu-id="1fadc-1039">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-1039">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-1040">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-1040">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-1041">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadc-1041">1.1</span></span>|
|[<span data-ttu-id="1fadc-1042">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-1042">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-1043">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-1043">ReadWriteItem</span></span>|
|[<span data-ttu-id="1fadc-1044">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-1044">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-1045">撰写</span><span class="sxs-lookup"><span data-stu-id="1fadc-1045">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-1046">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-1046">Example</span></span>

<span data-ttu-id="1fadc-1047">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1047">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1fadc-1048">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1fadc-1048">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1fadc-1049">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1049">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1fadc-p173">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1fadc-1053">参数</span><span class="sxs-lookup"><span data-stu-id="1fadc-1053">Parameters</span></span>

|<span data-ttu-id="1fadc-1054">名称</span><span class="sxs-lookup"><span data-stu-id="1fadc-1054">Name</span></span>| <span data-ttu-id="1fadc-1055">类型</span><span class="sxs-lookup"><span data-stu-id="1fadc-1055">Type</span></span>| <span data-ttu-id="1fadc-1056">属性</span><span class="sxs-lookup"><span data-stu-id="1fadc-1056">Attributes</span></span>| <span data-ttu-id="1fadc-1057">说明</span><span class="sxs-lookup"><span data-stu-id="1fadc-1057">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1fadc-1058">字符串</span><span class="sxs-lookup"><span data-stu-id="1fadc-1058">String</span></span>||<span data-ttu-id="1fadc-p174">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="1fadc-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="1fadc-1062">Object</span><span class="sxs-lookup"><span data-stu-id="1fadc-1062">Object</span></span>| <span data-ttu-id="1fadc-1063">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-1064">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1fadc-1065">对象</span><span class="sxs-lookup"><span data-stu-id="1fadc-1065">Object</span></span>| <span data-ttu-id="1fadc-1066">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-1067">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="1fadc-1068">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1fadc-1068">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="1fadc-1069">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1fadc-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="1fadc-1070">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1070">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="1fadc-1071">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1071">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1fadc-1072">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1072">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="1fadc-1073">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1073">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1fadc-1074">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1074">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="1fadc-1075">function</span><span class="sxs-lookup"><span data-stu-id="1fadc-1075">function</span></span>||<span data-ttu-id="1fadc-1076">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1fadc-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1fadc-1077">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadc-1077">Requirements</span></span>

|<span data-ttu-id="1fadc-1078">要求</span><span class="sxs-lookup"><span data-stu-id="1fadc-1078">Requirement</span></span>| <span data-ttu-id="1fadc-1079">值</span><span class="sxs-lookup"><span data-stu-id="1fadc-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadc-1080">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1fadc-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1fadc-1081">1.2</span><span class="sxs-lookup"><span data-stu-id="1fadc-1081">1.2</span></span>|
|[<span data-ttu-id="1fadc-1082">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1fadc-1082">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1fadc-1083">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1fadc-1083">ReadWriteItem</span></span>|
|[<span data-ttu-id="1fadc-1084">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1fadc-1084">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1fadc-1085">撰写</span><span class="sxs-lookup"><span data-stu-id="1fadc-1085">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadc-1086">示例</span><span class="sxs-lookup"><span data-stu-id="1fadc-1086">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
