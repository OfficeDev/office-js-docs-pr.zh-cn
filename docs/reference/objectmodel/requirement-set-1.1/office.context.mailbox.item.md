---
title: "\"Context\"-\"邮箱\"。项目-要求集1。1"
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 20d3aaecc5e0c62f86a46ae29010a6462446bf1d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696440"
---
# <a name="item"></a><span data-ttu-id="0d6be-102">item</span><span class="sxs-lookup"><span data-stu-id="0d6be-102">item</span></span>

### <span data-ttu-id="0d6be-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="0d6be-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="0d6be-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6be-107">Requirements</span></span>

|<span data-ttu-id="0d6be-108">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-108">Requirement</span></span>| <span data-ttu-id="0d6be-109">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-111">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-111">1.0</span></span>|
|[<span data-ttu-id="0d6be-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-113">受限</span><span class="sxs-lookup"><span data-stu-id="0d6be-113">Restricted</span></span>|
|[<span data-ttu-id="0d6be-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0d6be-116">成员和方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-116">Members and methods</span></span>

| <span data-ttu-id="0d6be-117">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-117">Member</span></span> | <span data-ttu-id="0d6be-118">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0d6be-119">attachments</span><span class="sxs-lookup"><span data-stu-id="0d6be-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="0d6be-120">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-120">Member</span></span> |
| [<span data-ttu-id="0d6be-121">bcc</span><span class="sxs-lookup"><span data-stu-id="0d6be-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="0d6be-122">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-122">Member</span></span> |
| [<span data-ttu-id="0d6be-123">body</span><span class="sxs-lookup"><span data-stu-id="0d6be-123">body</span></span>](#body-body) | <span data-ttu-id="0d6be-124">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-124">Member</span></span> |
| [<span data-ttu-id="0d6be-125">cc</span><span class="sxs-lookup"><span data-stu-id="0d6be-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6be-126">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-126">Member</span></span> |
| [<span data-ttu-id="0d6be-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="0d6be-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0d6be-128">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-128">Member</span></span> |
| [<span data-ttu-id="0d6be-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0d6be-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0d6be-130">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-130">Member</span></span> |
| [<span data-ttu-id="0d6be-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0d6be-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0d6be-132">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-132">Member</span></span> |
| [<span data-ttu-id="0d6be-133">end</span><span class="sxs-lookup"><span data-stu-id="0d6be-133">end</span></span>](#end-datetime) | <span data-ttu-id="0d6be-134">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-134">Member</span></span> |
| [<span data-ttu-id="0d6be-135">from</span><span class="sxs-lookup"><span data-stu-id="0d6be-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="0d6be-136">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-136">Member</span></span> |
| [<span data-ttu-id="0d6be-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0d6be-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0d6be-138">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-138">Member</span></span> |
| [<span data-ttu-id="0d6be-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="0d6be-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0d6be-140">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-140">Member</span></span> |
| [<span data-ttu-id="0d6be-141">itemId</span><span class="sxs-lookup"><span data-stu-id="0d6be-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0d6be-142">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-142">Member</span></span> |
| [<span data-ttu-id="0d6be-143">itemType</span><span class="sxs-lookup"><span data-stu-id="0d6be-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="0d6be-144">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-144">Member</span></span> |
| [<span data-ttu-id="0d6be-145">location</span><span class="sxs-lookup"><span data-stu-id="0d6be-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="0d6be-146">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-146">Member</span></span> |
| [<span data-ttu-id="0d6be-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0d6be-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0d6be-148">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-148">Member</span></span> |
| [<span data-ttu-id="0d6be-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0d6be-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6be-150">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-150">Member</span></span> |
| [<span data-ttu-id="0d6be-151">organizer</span><span class="sxs-lookup"><span data-stu-id="0d6be-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="0d6be-152">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-152">Member</span></span> |
| [<span data-ttu-id="0d6be-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0d6be-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6be-154">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-154">Member</span></span> |
| [<span data-ttu-id="0d6be-155">sender</span><span class="sxs-lookup"><span data-stu-id="0d6be-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="0d6be-156">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-156">Member</span></span> |
| [<span data-ttu-id="0d6be-157">start</span><span class="sxs-lookup"><span data-stu-id="0d6be-157">start</span></span>](#start-datetime) | <span data-ttu-id="0d6be-158">Member</span><span class="sxs-lookup"><span data-stu-id="0d6be-158">Member</span></span> |
| [<span data-ttu-id="0d6be-159">subject</span><span class="sxs-lookup"><span data-stu-id="0d6be-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="0d6be-160">成员</span><span class="sxs-lookup"><span data-stu-id="0d6be-160">Member</span></span> |
| [<span data-ttu-id="0d6be-161">to</span><span class="sxs-lookup"><span data-stu-id="0d6be-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6be-162">成员</span><span class="sxs-lookup"><span data-stu-id="0d6be-162">Member</span></span> |
| [<span data-ttu-id="0d6be-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0d6be-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0d6be-164">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-164">Method</span></span> |
| [<span data-ttu-id="0d6be-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0d6be-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0d6be-166">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-166">Method</span></span> |
| [<span data-ttu-id="0d6be-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0d6be-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="0d6be-168">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-168">Method</span></span> |
| [<span data-ttu-id="0d6be-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0d6be-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="0d6be-170">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-170">Method</span></span> |
| [<span data-ttu-id="0d6be-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="0d6be-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="0d6be-172">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-172">Method</span></span> |
| [<span data-ttu-id="0d6be-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0d6be-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0d6be-174">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-174">Method</span></span> |
| [<span data-ttu-id="0d6be-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0d6be-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0d6be-176">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-176">Method</span></span> |
| [<span data-ttu-id="0d6be-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0d6be-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0d6be-178">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-178">Method</span></span> |
| [<span data-ttu-id="0d6be-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0d6be-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0d6be-180">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-180">Method</span></span> |
| [<span data-ttu-id="0d6be-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0d6be-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0d6be-182">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-182">Method</span></span> |
| [<span data-ttu-id="0d6be-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0d6be-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0d6be-184">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0d6be-185">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-185">Example</span></span>

<span data-ttu-id="0d6be-186">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="0d6be-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0d6be-187">成员</span><span class="sxs-lookup"><span data-stu-id="0d6be-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="0d6be-188">附件: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="0d6be-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="0d6be-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-191">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="0d6be-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0d6be-192">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="0d6be-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-193">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-193">Type</span></span>

*   <span data-ttu-id="0d6be-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="0d6be-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-195">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-195">Requirements</span></span>

|<span data-ttu-id="0d6be-196">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-196">Requirement</span></span>| <span data-ttu-id="0d6be-197">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-199">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-199">1.0</span></span>|
|[<span data-ttu-id="0d6be-200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-201">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-203">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-204">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-204">Example</span></span>

<span data-ttu-id="0d6be-205">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="0d6be-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="0d6be-206">密件抄送:[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-207">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0d6be-208">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-208">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-209">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-209">Type</span></span>

*   [<span data-ttu-id="0d6be-210">收件人</span><span class="sxs-lookup"><span data-stu-id="0d6be-210">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="0d6be-211">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-211">Requirements</span></span>

|<span data-ttu-id="0d6be-212">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-212">Requirement</span></span>| <span data-ttu-id="0d6be-213">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-215">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6be-215">1.1</span></span>|
|[<span data-ttu-id="0d6be-216">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-217">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-218">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-219">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6be-219">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-220">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-220">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="0d6be-221">正文:[正文](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-221">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-222">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-222">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-223">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-223">Type</span></span>

*   [<span data-ttu-id="0d6be-224">Body</span><span class="sxs-lookup"><span data-stu-id="0d6be-224">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="0d6be-225">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-225">Requirements</span></span>

|<span data-ttu-id="0d6be-226">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-226">Requirement</span></span>| <span data-ttu-id="0d6be-227">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-228">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-229">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6be-229">1.1</span></span>|
|[<span data-ttu-id="0d6be-230">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-231">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-232">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-233">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-234">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-234">Example</span></span>

<span data-ttu-id="0d6be-235">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="0d6be-235">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="0d6be-236">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="0d6be-236">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="0d6be-237"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="0d6be-237">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-238">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6be-238">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0d6be-239">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-239">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-240">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-240">Read mode</span></span>

<span data-ttu-id="0d6be-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-243">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-243">Compose mode</span></span>

<span data-ttu-id="0d6be-244">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-244">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6be-245">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-245">Type</span></span>

*   <span data-ttu-id="0d6be-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-247">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-247">Requirements</span></span>

|<span data-ttu-id="0d6be-248">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-248">Requirement</span></span>| <span data-ttu-id="0d6be-249">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-250">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-251">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-251">1.0</span></span>|
|[<span data-ttu-id="0d6be-252">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-253">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-254">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-255">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="0d6be-256">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="0d6be-256">(nullable) conversationId: String</span></span>

<span data-ttu-id="0d6be-257">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-257">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0d6be-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0d6be-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-262">Type</span><span class="sxs-lookup"><span data-stu-id="0d6be-262">Type</span></span>

*   <span data-ttu-id="0d6be-263">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-263">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-264">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-264">Requirements</span></span>

|<span data-ttu-id="0d6be-265">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-265">Requirement</span></span>| <span data-ttu-id="0d6be-266">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-268">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-268">1.0</span></span>|
|[<span data-ttu-id="0d6be-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-270">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-273">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-273">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="0d6be-274">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="0d6be-274">dateTimeCreated: Date</span></span>

<span data-ttu-id="0d6be-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-277">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-277">Type</span></span>

*   <span data-ttu-id="0d6be-278">日期</span><span class="sxs-lookup"><span data-stu-id="0d6be-278">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-279">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-279">Requirements</span></span>

|<span data-ttu-id="0d6be-280">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-280">Requirement</span></span>| <span data-ttu-id="0d6be-281">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-282">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-283">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-283">1.0</span></span>|
|[<span data-ttu-id="0d6be-284">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-285">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-286">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-287">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-288">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-288">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="0d6be-289">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="0d6be-289">dateTimeModified: Date</span></span>

<span data-ttu-id="0d6be-290">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6be-290">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="0d6be-291">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-291">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-292">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="0d6be-292">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-293">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-293">Type</span></span>

*   <span data-ttu-id="0d6be-294">日期</span><span class="sxs-lookup"><span data-stu-id="0d6be-294">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-295">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-295">Requirements</span></span>

|<span data-ttu-id="0d6be-296">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-296">Requirement</span></span>| <span data-ttu-id="0d6be-297">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-299">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-299">1.0</span></span>|
|[<span data-ttu-id="0d6be-300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-301">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-303">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-303">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-304">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-304">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="0d6be-305">结束: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-305">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-306">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6be-306">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0d6be-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-309">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-309">Read mode</span></span>

<span data-ttu-id="0d6be-310">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-310">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-311">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-311">Compose mode</span></span>

<span data-ttu-id="0d6be-312">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-312">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0d6be-313">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="0d6be-313">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0d6be-314">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="0d6be-314">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0d6be-315">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-315">Type</span></span>

*   <span data-ttu-id="0d6be-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-317">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-317">Requirements</span></span>

|<span data-ttu-id="0d6be-318">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-318">Requirement</span></span>| <span data-ttu-id="0d6be-319">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-320">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-321">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-321">1.0</span></span>|
|[<span data-ttu-id="0d6be-322">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-323">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-324">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-325">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-325">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="0d6be-326">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-326">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="0d6be-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-331">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="0d6be-331">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-332">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-332">Type</span></span>

*   [<span data-ttu-id="0d6be-333">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0d6be-333">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="0d6be-334">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-334">Requirements</span></span>

|<span data-ttu-id="0d6be-335">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-335">Requirement</span></span>| <span data-ttu-id="0d6be-336">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-337">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-338">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-338">1.0</span></span>|
|[<span data-ttu-id="0d6be-339">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-340">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-341">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-342">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-343">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-343">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="0d6be-344">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="0d6be-344">internetMessageId: String</span></span>

<span data-ttu-id="0d6be-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-347">Type</span><span class="sxs-lookup"><span data-stu-id="0d6be-347">Type</span></span>

*   <span data-ttu-id="0d6be-348">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-348">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-349">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-349">Requirements</span></span>

|<span data-ttu-id="0d6be-350">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-350">Requirement</span></span>| <span data-ttu-id="0d6be-351">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-351">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-352">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-353">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-353">1.0</span></span>|
|[<span data-ttu-id="0d6be-354">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-355">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-356">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-357">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-357">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-358">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-358">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="0d6be-359">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="0d6be-359">itemClass: String</span></span>

<span data-ttu-id="0d6be-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0d6be-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="0d6be-364">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-364">Type</span></span> | <span data-ttu-id="0d6be-365">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-365">Description</span></span> | <span data-ttu-id="0d6be-366">项目类</span><span class="sxs-lookup"><span data-stu-id="0d6be-366">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="0d6be-367">约会项目</span><span class="sxs-lookup"><span data-stu-id="0d6be-367">Appointment items</span></span> | <span data-ttu-id="0d6be-368">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="0d6be-368">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="0d6be-369">邮件项目</span><span class="sxs-lookup"><span data-stu-id="0d6be-369">Message items</span></span> | <span data-ttu-id="0d6be-370">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="0d6be-370">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="0d6be-371">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="0d6be-371">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-372">Type</span><span class="sxs-lookup"><span data-stu-id="0d6be-372">Type</span></span>

*   <span data-ttu-id="0d6be-373">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-374">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-374">Requirements</span></span>

|<span data-ttu-id="0d6be-375">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-375">Requirement</span></span>| <span data-ttu-id="0d6be-376">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-377">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-378">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-378">1.0</span></span>|
|[<span data-ttu-id="0d6be-379">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-380">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-381">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-382">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-383">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-383">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0d6be-384">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="0d6be-384">(nullable) itemId: String</span></span>

<span data-ttu-id="0d6be-385">获取当前项目的 Exchange Web 服务项目标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-385">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="0d6be-386">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-386">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-387">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="0d6be-387">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0d6be-388">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="0d6be-388">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0d6be-389">在使用此值进行 REST API 调用之前, 应使用`Office.context.mailbox.convertToRestId`转换它, 这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="0d6be-389">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="0d6be-390">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="0d6be-390">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-391">Type</span><span class="sxs-lookup"><span data-stu-id="0d6be-391">Type</span></span>

*   <span data-ttu-id="0d6be-392">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-392">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-393">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-393">Requirements</span></span>

|<span data-ttu-id="0d6be-394">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-394">Requirement</span></span>| <span data-ttu-id="0d6be-395">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-395">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-396">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-397">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-397">1.0</span></span>|
|[<span data-ttu-id="0d6be-398">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-398">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-399">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-399">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-400">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-400">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-401">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-401">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-402">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-402">Example</span></span>

<span data-ttu-id="0d6be-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="0d6be-405">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-405">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-406">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="0d6be-406">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0d6be-407">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="0d6be-407">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-408">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-408">Type</span></span>

*   [<span data-ttu-id="0d6be-409">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0d6be-409">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="0d6be-410">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-410">Requirements</span></span>

|<span data-ttu-id="0d6be-411">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-411">Requirement</span></span>| <span data-ttu-id="0d6be-412">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-414">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-414">1.0</span></span>|
|[<span data-ttu-id="0d6be-415">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-416">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-417">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-418">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-418">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-419">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-419">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="0d6be-420">位置: 字符串 |[位置](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-420">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-421">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="0d6be-421">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-422">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-422">Read mode</span></span>

<span data-ttu-id="0d6be-423">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="0d6be-423">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-424">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-424">Compose mode</span></span>

<span data-ttu-id="0d6be-425">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-425">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6be-426">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-426">Type</span></span>

*   <span data-ttu-id="0d6be-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-428">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-428">Requirements</span></span>

|<span data-ttu-id="0d6be-429">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-429">Requirement</span></span>| <span data-ttu-id="0d6be-430">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-430">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-431">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-431">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-432">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-432">1.0</span></span>|
|[<span data-ttu-id="0d6be-433">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-433">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-434">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-434">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-435">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-435">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-436">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-436">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0d6be-437">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="0d6be-437">normalizedSubject: String</span></span>

<span data-ttu-id="0d6be-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0d6be-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-442">Type</span><span class="sxs-lookup"><span data-stu-id="0d6be-442">Type</span></span>

*   <span data-ttu-id="0d6be-443">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-443">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-444">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-444">Requirements</span></span>

|<span data-ttu-id="0d6be-445">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-445">Requirement</span></span>| <span data-ttu-id="0d6be-446">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-448">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-448">1.0</span></span>|
|[<span data-ttu-id="0d6be-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-450">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-452">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-453">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-453">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="0d6be-454">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的数组</span><span class="sxs-lookup"><span data-stu-id="0d6be-454">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-455">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6be-455">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0d6be-456">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-456">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-457">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-457">Read mode</span></span>

<span data-ttu-id="0d6be-458">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-458">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-459">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-459">Compose mode</span></span>

<span data-ttu-id="0d6be-460">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-460">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6be-461">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-461">Type</span></span>

*   <span data-ttu-id="0d6be-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-463">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-463">Requirements</span></span>

|<span data-ttu-id="0d6be-464">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-464">Requirement</span></span>| <span data-ttu-id="0d6be-465">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-466">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-467">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-467">1.0</span></span>|
|[<span data-ttu-id="0d6be-468">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-469">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-470">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-471">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-471">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="0d6be-472">组织者: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-472">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-475">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-475">Type</span></span>

*   [<span data-ttu-id="0d6be-476">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0d6be-476">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="0d6be-477">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-477">Requirements</span></span>

|<span data-ttu-id="0d6be-478">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-478">Requirement</span></span>| <span data-ttu-id="0d6be-479">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-480">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-481">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-481">1.0</span></span>|
|[<span data-ttu-id="0d6be-482">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-482">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-483">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-484">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-484">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-485">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-485">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-486">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-486">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="0d6be-487">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的数组</span><span class="sxs-lookup"><span data-stu-id="0d6be-487">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-488">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6be-488">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0d6be-489">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-489">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-490">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-490">Read mode</span></span>

<span data-ttu-id="0d6be-491">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-491">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-492">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-492">Compose mode</span></span>

<span data-ttu-id="0d6be-493">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-493">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="0d6be-494">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-494">Type</span></span>

*   <span data-ttu-id="0d6be-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-496">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-496">Requirements</span></span>

|<span data-ttu-id="0d6be-497">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-497">Requirement</span></span>| <span data-ttu-id="0d6be-498">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-499">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-500">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-500">1.0</span></span>|
|[<span data-ttu-id="0d6be-501">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-502">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-503">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-504">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-504">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="0d6be-505">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-505">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0d6be-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-510">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="0d6be-510">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6be-511">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-511">Type</span></span>

*   [<span data-ttu-id="0d6be-512">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0d6be-512">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="0d6be-513">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-513">Requirements</span></span>

|<span data-ttu-id="0d6be-514">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-514">Requirement</span></span>| <span data-ttu-id="0d6be-515">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-515">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-516">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-516">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-517">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-517">1.0</span></span>|
|[<span data-ttu-id="0d6be-518">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-518">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-519">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-519">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-520">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-520">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-521">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-521">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-522">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-522">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="0d6be-523">开始日期: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-523">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-524">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6be-524">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0d6be-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-527">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-527">Read mode</span></span>

<span data-ttu-id="0d6be-528">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-528">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-529">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-529">Compose mode</span></span>

<span data-ttu-id="0d6be-530">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-530">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0d6be-531">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="0d6be-531">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0d6be-532">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="0d6be-532">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0d6be-533">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-533">Type</span></span>

*   <span data-ttu-id="0d6be-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-535">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-535">Requirements</span></span>

|<span data-ttu-id="0d6be-536">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-536">Requirement</span></span>| <span data-ttu-id="0d6be-537">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-538">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-539">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-539">1.0</span></span>|
|[<span data-ttu-id="0d6be-540">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-541">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-542">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-543">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-543">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="0d6be-544">subject: String |[主题](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-544">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-545">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="0d6be-545">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0d6be-546">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="0d6be-546">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-547">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-547">Read mode</span></span>

<span data-ttu-id="0d6be-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-550">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-550">Compose mode</span></span>

<span data-ttu-id="0d6be-551">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-551">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="0d6be-552">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-552">Type</span></span>

*   <span data-ttu-id="0d6be-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-554">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-554">Requirements</span></span>

|<span data-ttu-id="0d6be-555">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-555">Requirement</span></span>| <span data-ttu-id="0d6be-556">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-557">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-558">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-558">1.0</span></span>|
|[<span data-ttu-id="0d6be-559">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-560">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-561">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-562">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-562">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="0d6be-563">to: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的数组</span><span class="sxs-lookup"><span data-stu-id="0d6be-563">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="0d6be-564">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6be-564">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0d6be-565">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6be-565">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6be-566">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-566">Read mode</span></span>

<span data-ttu-id="0d6be-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6be-569">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-569">Compose mode</span></span>

<span data-ttu-id="0d6be-570">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-570">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6be-571">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-571">Type</span></span>

*   <span data-ttu-id="0d6be-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-573">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-573">Requirements</span></span>

|<span data-ttu-id="0d6be-574">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-574">Requirement</span></span>| <span data-ttu-id="0d6be-575">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-575">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-576">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-576">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-577">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-577">1.0</span></span>|
|[<span data-ttu-id="0d6be-578">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-578">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-579">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-579">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-580">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-580">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-581">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-581">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0d6be-582">方法</span><span class="sxs-lookup"><span data-stu-id="0d6be-582">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0d6be-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6be-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0d6be-584">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="0d6be-584">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0d6be-585">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="0d6be-585">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0d6be-586">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6be-586">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-587">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-587">Parameters</span></span>

|<span data-ttu-id="0d6be-588">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-588">Name</span></span>| <span data-ttu-id="0d6be-589">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-589">Type</span></span>| <span data-ttu-id="0d6be-590">属性</span><span class="sxs-lookup"><span data-stu-id="0d6be-590">Attributes</span></span>| <span data-ttu-id="0d6be-591">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-591">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="0d6be-592">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-592">String</span></span>||<span data-ttu-id="0d6be-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0d6be-595">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6be-595">String</span></span>||<span data-ttu-id="0d6be-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0d6be-598">Object</span><span class="sxs-lookup"><span data-stu-id="0d6be-598">Object</span></span>| <span data-ttu-id="0d6be-599">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-599">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-600">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6be-600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0d6be-601">对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-601">Object</span></span>| <span data-ttu-id="0d6be-602">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-602">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-603">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0d6be-604">函数</span><span class="sxs-lookup"><span data-stu-id="0d6be-604">function</span></span>| <span data-ttu-id="0d6be-605">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-605">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-606">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0d6be-607">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="0d6be-607">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0d6be-608">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-608">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0d6be-609">错误</span><span class="sxs-lookup"><span data-stu-id="0d6be-609">Errors</span></span>

| <span data-ttu-id="0d6be-610">错误代码</span><span class="sxs-lookup"><span data-stu-id="0d6be-610">Error code</span></span> | <span data-ttu-id="0d6be-611">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-611">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="0d6be-612">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="0d6be-612">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="0d6be-613">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="0d6be-613">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0d6be-614">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="0d6be-614">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0d6be-615">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-615">Requirements</span></span>

|<span data-ttu-id="0d6be-616">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-616">Requirement</span></span>| <span data-ttu-id="0d6be-617">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-618">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-619">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6be-619">1.1</span></span>|
|[<span data-ttu-id="0d6be-620">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-621">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6be-622">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-623">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6be-623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-624">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-624">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0d6be-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6be-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0d6be-626">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="0d6be-626">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0d6be-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0d6be-630">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6be-630">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0d6be-631">如果 Office 外接程序在 web 上的 Outlook 中运行, 则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是, 不支持这种情况, 建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="0d6be-631">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-632">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-632">Parameters</span></span>

|<span data-ttu-id="0d6be-633">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-633">Name</span></span>| <span data-ttu-id="0d6be-634">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-634">Type</span></span>| <span data-ttu-id="0d6be-635">属性</span><span class="sxs-lookup"><span data-stu-id="0d6be-635">Attributes</span></span>| <span data-ttu-id="0d6be-636">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-636">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="0d6be-637">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-637">String</span></span>||<span data-ttu-id="0d6be-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="0d6be-640">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-640">String</span></span>||<span data-ttu-id="0d6be-641">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="0d6be-641">The subject of the item to be attached.</span></span> <span data-ttu-id="0d6be-642">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-642">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="0d6be-643">对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-643">Object</span></span>| <span data-ttu-id="0d6be-644">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-644">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-645">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6be-645">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0d6be-646">对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-646">Object</span></span>| <span data-ttu-id="0d6be-647">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-647">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-648">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-648">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0d6be-649">函数</span><span class="sxs-lookup"><span data-stu-id="0d6be-649">function</span></span>| <span data-ttu-id="0d6be-650">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-650">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-651">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-651">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0d6be-652">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="0d6be-652">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0d6be-653">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-653">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0d6be-654">错误</span><span class="sxs-lookup"><span data-stu-id="0d6be-654">Errors</span></span>

| <span data-ttu-id="0d6be-655">错误代码</span><span class="sxs-lookup"><span data-stu-id="0d6be-655">Error code</span></span> | <span data-ttu-id="0d6be-656">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-656">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="0d6be-657">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="0d6be-657">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0d6be-658">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-658">Requirements</span></span>

|<span data-ttu-id="0d6be-659">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-659">Requirement</span></span>| <span data-ttu-id="0d6be-660">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-661">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-662">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6be-662">1.1</span></span>|
|[<span data-ttu-id="0d6be-663">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-664">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-664">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6be-665">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-666">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6be-666">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-667">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-667">Example</span></span>

<span data-ttu-id="0d6be-668">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6be-668">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="0d6be-669">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6be-669">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="0d6be-670">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="0d6be-670">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-671">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-671">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6be-672">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="0d6be-672">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0d6be-673">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="0d6be-673">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-674">要求集1.1 中不支持在呼叫`displayReplyAllForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="0d6be-674">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="0d6be-675">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="0d6be-675">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-676">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-676">Parameters</span></span>

|<span data-ttu-id="0d6be-677">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-677">Name</span></span>| <span data-ttu-id="0d6be-678">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-678">Type</span></span>| <span data-ttu-id="0d6be-679">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-679">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0d6be-680">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-680">String &#124; Object</span></span>| |<span data-ttu-id="0d6be-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0d6be-683">**或**</span><span class="sxs-lookup"><span data-stu-id="0d6be-683">**OR**</span></span><br/><span data-ttu-id="0d6be-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0d6be-686">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6be-686">String</span></span> | <span data-ttu-id="0d6be-687">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-687">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6be-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="0d6be-690">函数</span><span class="sxs-lookup"><span data-stu-id="0d6be-690">function</span></span> | <span data-ttu-id="0d6be-691">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-691">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6be-692">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-692">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0d6be-693">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-693">Requirements</span></span>

|<span data-ttu-id="0d6be-694">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-694">Requirement</span></span>| <span data-ttu-id="0d6be-695">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-695">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-696">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-696">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-697">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-697">1.0</span></span>|
|[<span data-ttu-id="0d6be-698">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-698">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-699">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-699">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-700">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-700">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-701">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-701">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0d6be-702">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-702">Examples</span></span>

<span data-ttu-id="0d6be-703">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-703">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0d6be-704">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6be-704">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0d6be-705">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6be-705">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0d6be-706">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="0d6be-706">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="0d6be-707">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6be-707">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="0d6be-708">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="0d6be-708">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-709">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-709">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6be-710">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="0d6be-710">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0d6be-711">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="0d6be-711">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-712">要求集1.1 中不支持在呼叫`displayReplyForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="0d6be-712">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="0d6be-713">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="0d6be-713">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-714">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-714">Parameters</span></span>

|<span data-ttu-id="0d6be-715">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-715">Name</span></span>| <span data-ttu-id="0d6be-716">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-716">Type</span></span>| <span data-ttu-id="0d6be-717">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-717">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="0d6be-718">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-718">String &#124; Object</span></span>| | <span data-ttu-id="0d6be-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0d6be-721">**或**</span><span class="sxs-lookup"><span data-stu-id="0d6be-721">**OR**</span></span><br/><span data-ttu-id="0d6be-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="0d6be-724">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6be-724">String</span></span> | <span data-ttu-id="0d6be-725">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-725">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6be-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="0d6be-728">函数</span><span class="sxs-lookup"><span data-stu-id="0d6be-728">function</span></span> | <span data-ttu-id="0d6be-729">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-729">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6be-730">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-730">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0d6be-731">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-731">Requirements</span></span>

|<span data-ttu-id="0d6be-732">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-732">Requirement</span></span>| <span data-ttu-id="0d6be-733">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-734">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-735">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-735">1.0</span></span>|
|[<span data-ttu-id="0d6be-736">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-737">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-738">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-739">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-739">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0d6be-740">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-740">Examples</span></span>

<span data-ttu-id="0d6be-741">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-741">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0d6be-742">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6be-742">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0d6be-743">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6be-743">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0d6be-744">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="0d6be-744">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="0d6be-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="0d6be-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="0d6be-746">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="0d6be-746">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-747">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-747">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-748">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-748">Requirements</span></span>

|<span data-ttu-id="0d6be-749">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-749">Requirement</span></span>| <span data-ttu-id="0d6be-750">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-750">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-751">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-752">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-752">1.0</span></span>|
|[<span data-ttu-id="0d6be-753">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-754">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-754">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-755">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-756">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-756">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6be-757">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6be-757">Returns:</span></span>

<span data-ttu-id="0d6be-758">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="0d6be-758">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="0d6be-759">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-759">Example</span></span>

<span data-ttu-id="0d6be-760">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="0d6be-760">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="0d6be-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="0d6be-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="0d6be-762">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6be-762">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-763">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-763">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-764">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-764">Parameters</span></span>

|<span data-ttu-id="0d6be-765">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-765">Name</span></span>| <span data-ttu-id="0d6be-766">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-766">Type</span></span>| <span data-ttu-id="0d6be-767">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-767">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="0d6be-768">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0d6be-768">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="0d6be-769">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="0d6be-769">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6be-770">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-770">Requirements</span></span>

|<span data-ttu-id="0d6be-771">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-771">Requirement</span></span>| <span data-ttu-id="0d6be-772">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-773">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-774">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-774">1.0</span></span>|
|[<span data-ttu-id="0d6be-775">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-775">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-776">受限</span><span class="sxs-lookup"><span data-stu-id="0d6be-776">Restricted</span></span>|
|[<span data-ttu-id="0d6be-777">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-777">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-778">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6be-779">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6be-779">Returns:</span></span>

<span data-ttu-id="0d6be-780">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="0d6be-780">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0d6be-781">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="0d6be-781">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="0d6be-782">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="0d6be-782">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0d6be-783">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="0d6be-783">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="0d6be-784">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="0d6be-784">Value of `entityType`</span></span> | <span data-ttu-id="0d6be-785">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-785">Type of objects in returned array</span></span> | <span data-ttu-id="0d6be-786">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-786">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="0d6be-787">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-787">String</span></span> | <span data-ttu-id="0d6be-788">**受限**</span><span class="sxs-lookup"><span data-stu-id="0d6be-788">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="0d6be-789">Contact</span><span class="sxs-lookup"><span data-stu-id="0d6be-789">Contact</span></span> | <span data-ttu-id="0d6be-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6be-790">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="0d6be-791">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-791">String</span></span> | <span data-ttu-id="0d6be-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6be-792">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="0d6be-793">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0d6be-793">MeetingSuggestion</span></span> | <span data-ttu-id="0d6be-794">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6be-794">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="0d6be-795">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0d6be-795">PhoneNumber</span></span> | <span data-ttu-id="0d6be-796">**受限**</span><span class="sxs-lookup"><span data-stu-id="0d6be-796">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="0d6be-797">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0d6be-797">TaskSuggestion</span></span> | <span data-ttu-id="0d6be-798">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6be-798">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="0d6be-799">String</span><span class="sxs-lookup"><span data-stu-id="0d6be-799">String</span></span> | <span data-ttu-id="0d6be-800">**受限**</span><span class="sxs-lookup"><span data-stu-id="0d6be-800">**Restricted**</span></span> |

<span data-ttu-id="0d6be-801">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="0d6be-801">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="0d6be-802">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-802">Example</span></span>

<span data-ttu-id="0d6be-803">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="0d6be-803">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="0d6be-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="0d6be-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="0d6be-805">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="0d6be-805">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-806">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-806">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6be-807">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="0d6be-807">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-808">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-808">Parameters</span></span>

|<span data-ttu-id="0d6be-809">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-809">Name</span></span>| <span data-ttu-id="0d6be-810">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-810">Type</span></span>| <span data-ttu-id="0d6be-811">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-811">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0d6be-812">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6be-812">String</span></span>|<span data-ttu-id="0d6be-813">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="0d6be-813">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6be-814">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-814">Requirements</span></span>

|<span data-ttu-id="0d6be-815">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-815">Requirement</span></span>| <span data-ttu-id="0d6be-816">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-817">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-818">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-818">1.0</span></span>|
|[<span data-ttu-id="0d6be-819">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-820">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-821">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-822">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6be-823">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6be-823">Returns:</span></span>

<span data-ttu-id="0d6be-p146">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="0d6be-826">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="0d6be-826">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="0d6be-827">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0d6be-827">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0d6be-828">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="0d6be-828">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-829">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-829">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6be-p147">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0d6be-833">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="0d6be-833">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0d6be-834">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="0d6be-834">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="0d6be-p148">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6be-837">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-837">Requirements</span></span>

|<span data-ttu-id="0d6be-838">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-838">Requirement</span></span>| <span data-ttu-id="0d6be-839">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-840">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-841">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-841">1.0</span></span>|
|[<span data-ttu-id="0d6be-842">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-843">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-844">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-845">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6be-846">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6be-846">Returns:</span></span>

<span data-ttu-id="0d6be-p149">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="0d6be-849">类型: 对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-849">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="0d6be-850">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-850">Example</span></span>

<span data-ttu-id="0d6be-851">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="0d6be-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0d6be-852">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="0d6be-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0d6be-853">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="0d6be-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6be-854">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6be-854">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6be-855">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="0d6be-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0d6be-p150">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-858">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-858">Parameters</span></span>

|<span data-ttu-id="0d6be-859">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-859">Name</span></span>| <span data-ttu-id="0d6be-860">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-860">Type</span></span>| <span data-ttu-id="0d6be-861">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="0d6be-862">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6be-862">String</span></span>|<span data-ttu-id="0d6be-863">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="0d6be-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6be-864">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-864">Requirements</span></span>

|<span data-ttu-id="0d6be-865">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-865">Requirement</span></span>| <span data-ttu-id="0d6be-866">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-867">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-868">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-868">1.0</span></span>|
|[<span data-ttu-id="0d6be-869">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-869">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-870">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-871">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-871">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-872">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6be-873">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6be-873">Returns:</span></span>

<span data-ttu-id="0d6be-874">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6be-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="0d6be-875">类型: Array. < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="0d6be-875">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="0d6be-876">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0d6be-877">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0d6be-877">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0d6be-878">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="0d6be-878">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0d6be-p151">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-882">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-882">Parameters</span></span>

|<span data-ttu-id="0d6be-883">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-883">Name</span></span>| <span data-ttu-id="0d6be-884">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-884">Type</span></span>| <span data-ttu-id="0d6be-885">属性</span><span class="sxs-lookup"><span data-stu-id="0d6be-885">Attributes</span></span>| <span data-ttu-id="0d6be-886">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-886">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="0d6be-887">函数</span><span class="sxs-lookup"><span data-stu-id="0d6be-887">function</span></span>||<span data-ttu-id="0d6be-888">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-888">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0d6be-889">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="0d6be-889">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0d6be-890">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="0d6be-890">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="0d6be-891">对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-891">Object</span></span>| <span data-ttu-id="0d6be-892">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-892">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-893">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-893">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="0d6be-894">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="0d6be-894">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6be-895">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-895">Requirements</span></span>

|<span data-ttu-id="0d6be-896">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-896">Requirement</span></span>| <span data-ttu-id="0d6be-897">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-897">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-898">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-898">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-899">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6be-899">1.0</span></span>|
|[<span data-ttu-id="0d6be-900">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-900">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-901">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-901">ReadItem</span></span>|
|[<span data-ttu-id="0d6be-902">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-902">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-903">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6be-903">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-904">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-904">Example</span></span>

<span data-ttu-id="0d6be-p154">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="0d6be-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0d6be-908">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6be-908">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0d6be-909">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="0d6be-909">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0d6be-910">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6be-910">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0d6be-911">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="0d6be-911">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="0d6be-912">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="0d6be-912">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0d6be-913">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="0d6be-913">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6be-914">参数</span><span class="sxs-lookup"><span data-stu-id="0d6be-914">Parameters</span></span>

|<span data-ttu-id="0d6be-915">名称</span><span class="sxs-lookup"><span data-stu-id="0d6be-915">Name</span></span>| <span data-ttu-id="0d6be-916">类型</span><span class="sxs-lookup"><span data-stu-id="0d6be-916">Type</span></span>| <span data-ttu-id="0d6be-917">属性</span><span class="sxs-lookup"><span data-stu-id="0d6be-917">Attributes</span></span>| <span data-ttu-id="0d6be-918">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-918">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="0d6be-919">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6be-919">String</span></span>||<span data-ttu-id="0d6be-920">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6be-920">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="0d6be-921">对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-921">Object</span></span>| <span data-ttu-id="0d6be-922">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-922">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-923">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6be-923">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="0d6be-924">对象</span><span class="sxs-lookup"><span data-stu-id="0d6be-924">Object</span></span>| <span data-ttu-id="0d6be-925">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-925">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-926">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6be-926">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="0d6be-927">函数</span><span class="sxs-lookup"><span data-stu-id="0d6be-927">function</span></span>| <span data-ttu-id="0d6be-928">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6be-928">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6be-929">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6be-929">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0d6be-930">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="0d6be-930">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0d6be-931">错误</span><span class="sxs-lookup"><span data-stu-id="0d6be-931">Errors</span></span>

| <span data-ttu-id="0d6be-932">错误代码</span><span class="sxs-lookup"><span data-stu-id="0d6be-932">Error code</span></span> | <span data-ttu-id="0d6be-933">说明</span><span class="sxs-lookup"><span data-stu-id="0d6be-933">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="0d6be-934">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="0d6be-934">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0d6be-935">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-935">Requirements</span></span>

|<span data-ttu-id="0d6be-936">要求</span><span class="sxs-lookup"><span data-stu-id="0d6be-936">Requirement</span></span>| <span data-ttu-id="0d6be-937">值</span><span class="sxs-lookup"><span data-stu-id="0d6be-937">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6be-938">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6be-938">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6be-939">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6be-939">1.1</span></span>|
|[<span data-ttu-id="0d6be-940">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6be-940">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6be-941">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6be-941">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6be-942">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6be-942">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6be-943">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6be-943">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6be-944">示例</span><span class="sxs-lookup"><span data-stu-id="0d6be-944">Example</span></span>

<span data-ttu-id="0d6be-945">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6be-945">The following code removes an attachment with an identifier of '0'.</span></span>

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
