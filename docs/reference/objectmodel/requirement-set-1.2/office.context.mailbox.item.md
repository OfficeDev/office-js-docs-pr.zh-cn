---
title: "\"Context\"-\"邮箱\"。项目-要求集1。2"
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: c765b0901c15adb7c3651ac279f224de05002023
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167345"
---
# <a name="item"></a><span data-ttu-id="357bc-102">item</span><span class="sxs-lookup"><span data-stu-id="357bc-102">item</span></span>

### <span data-ttu-id="357bc-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="357bc-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="357bc-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="357bc-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="357bc-107">Requirements</span></span>

|<span data-ttu-id="357bc-108">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-108">Requirement</span></span>| <span data-ttu-id="357bc-109">值</span><span class="sxs-lookup"><span data-stu-id="357bc-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-111">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-111">1.0</span></span>|
|[<span data-ttu-id="357bc-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-113">受限</span><span class="sxs-lookup"><span data-stu-id="357bc-113">Restricted</span></span>|
|[<span data-ttu-id="357bc-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="357bc-116">成员和方法</span><span class="sxs-lookup"><span data-stu-id="357bc-116">Members and methods</span></span>

| <span data-ttu-id="357bc-117">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-117">Member</span></span> | <span data-ttu-id="357bc-118">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="357bc-119">attachments</span><span class="sxs-lookup"><span data-stu-id="357bc-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="357bc-120">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-120">Member</span></span> |
| [<span data-ttu-id="357bc-121">bcc</span><span class="sxs-lookup"><span data-stu-id="357bc-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="357bc-122">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-122">Member</span></span> |
| [<span data-ttu-id="357bc-123">body</span><span class="sxs-lookup"><span data-stu-id="357bc-123">body</span></span>](#body-body) | <span data-ttu-id="357bc-124">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-124">Member</span></span> |
| [<span data-ttu-id="357bc-125">cc</span><span class="sxs-lookup"><span data-stu-id="357bc-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="357bc-126">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-126">Member</span></span> |
| [<span data-ttu-id="357bc-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="357bc-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="357bc-128">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-128">Member</span></span> |
| [<span data-ttu-id="357bc-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="357bc-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="357bc-130">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-130">Member</span></span> |
| [<span data-ttu-id="357bc-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="357bc-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="357bc-132">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-132">Member</span></span> |
| [<span data-ttu-id="357bc-133">end</span><span class="sxs-lookup"><span data-stu-id="357bc-133">end</span></span>](#end-datetime) | <span data-ttu-id="357bc-134">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-134">Member</span></span> |
| [<span data-ttu-id="357bc-135">from</span><span class="sxs-lookup"><span data-stu-id="357bc-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="357bc-136">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-136">Member</span></span> |
| [<span data-ttu-id="357bc-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="357bc-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="357bc-138">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-138">Member</span></span> |
| [<span data-ttu-id="357bc-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="357bc-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="357bc-140">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-140">Member</span></span> |
| [<span data-ttu-id="357bc-141">itemId</span><span class="sxs-lookup"><span data-stu-id="357bc-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="357bc-142">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-142">Member</span></span> |
| [<span data-ttu-id="357bc-143">itemType</span><span class="sxs-lookup"><span data-stu-id="357bc-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="357bc-144">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-144">Member</span></span> |
| [<span data-ttu-id="357bc-145">location</span><span class="sxs-lookup"><span data-stu-id="357bc-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="357bc-146">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-146">Member</span></span> |
| [<span data-ttu-id="357bc-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="357bc-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="357bc-148">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-148">Member</span></span> |
| [<span data-ttu-id="357bc-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="357bc-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="357bc-150">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-150">Member</span></span> |
| [<span data-ttu-id="357bc-151">organizer</span><span class="sxs-lookup"><span data-stu-id="357bc-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="357bc-152">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-152">Member</span></span> |
| [<span data-ttu-id="357bc-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="357bc-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="357bc-154">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-154">Member</span></span> |
| [<span data-ttu-id="357bc-155">sender</span><span class="sxs-lookup"><span data-stu-id="357bc-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="357bc-156">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-156">Member</span></span> |
| [<span data-ttu-id="357bc-157">start</span><span class="sxs-lookup"><span data-stu-id="357bc-157">start</span></span>](#start-datetime) | <span data-ttu-id="357bc-158">Member</span><span class="sxs-lookup"><span data-stu-id="357bc-158">Member</span></span> |
| [<span data-ttu-id="357bc-159">subject</span><span class="sxs-lookup"><span data-stu-id="357bc-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="357bc-160">成员</span><span class="sxs-lookup"><span data-stu-id="357bc-160">Member</span></span> |
| [<span data-ttu-id="357bc-161">to</span><span class="sxs-lookup"><span data-stu-id="357bc-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="357bc-162">成员</span><span class="sxs-lookup"><span data-stu-id="357bc-162">Member</span></span> |
| [<span data-ttu-id="357bc-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="357bc-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="357bc-164">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-164">Method</span></span> |
| [<span data-ttu-id="357bc-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="357bc-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="357bc-166">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-166">Method</span></span> |
| [<span data-ttu-id="357bc-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="357bc-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="357bc-168">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-168">Method</span></span> |
| [<span data-ttu-id="357bc-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="357bc-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="357bc-170">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-170">Method</span></span> |
| [<span data-ttu-id="357bc-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="357bc-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="357bc-172">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-172">Method</span></span> |
| [<span data-ttu-id="357bc-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="357bc-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="357bc-174">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-174">Method</span></span> |
| [<span data-ttu-id="357bc-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="357bc-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="357bc-176">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-176">Method</span></span> |
| [<span data-ttu-id="357bc-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="357bc-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="357bc-178">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-178">Method</span></span> |
| [<span data-ttu-id="357bc-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="357bc-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="357bc-180">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-180">Method</span></span> |
| [<span data-ttu-id="357bc-181">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="357bc-181">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="357bc-182">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-182">Method</span></span> |
| [<span data-ttu-id="357bc-183">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="357bc-183">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="357bc-184">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-184">Method</span></span> |
| [<span data-ttu-id="357bc-185">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="357bc-185">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="357bc-186">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-186">Method</span></span> |
| [<span data-ttu-id="357bc-187">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="357bc-187">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="357bc-188">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-188">Method</span></span> |

### <a name="example"></a><span data-ttu-id="357bc-189">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-189">Example</span></span>

<span data-ttu-id="357bc-190">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="357bc-190">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="357bc-191">成员</span><span class="sxs-lookup"><span data-stu-id="357bc-191">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="357bc-192">附件： Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="357bc-192">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="357bc-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-195">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="357bc-195">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="357bc-196">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="357bc-196">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-197">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-197">Type</span></span>

*   <span data-ttu-id="357bc-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="357bc-198">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-199">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-199">Requirements</span></span>

|<span data-ttu-id="357bc-200">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-200">Requirement</span></span>| <span data-ttu-id="357bc-201">值</span><span class="sxs-lookup"><span data-stu-id="357bc-201">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-202">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-202">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-203">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-203">1.0</span></span>|
|[<span data-ttu-id="357bc-204">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-204">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-205">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-205">ReadItem</span></span>|
|[<span data-ttu-id="357bc-206">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-206">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-207">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-207">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-208">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-208">Example</span></span>

<span data-ttu-id="357bc-209">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="357bc-209">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="357bc-210">密件抄送：[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-210">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-211">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-211">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="357bc-212">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-212">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-213">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-213">Type</span></span>

*   [<span data-ttu-id="357bc-214">收件人</span><span class="sxs-lookup"><span data-stu-id="357bc-214">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="357bc-215">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-215">Requirements</span></span>

|<span data-ttu-id="357bc-216">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-216">Requirement</span></span>| <span data-ttu-id="357bc-217">值</span><span class="sxs-lookup"><span data-stu-id="357bc-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-218">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-219">1.1</span><span class="sxs-lookup"><span data-stu-id="357bc-219">1.1</span></span>|
|[<span data-ttu-id="357bc-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-221">ReadItem</span></span>|
|[<span data-ttu-id="357bc-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-223">撰写</span><span class="sxs-lookup"><span data-stu-id="357bc-223">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-224">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-224">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="357bc-225">正文：[正文](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-225">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-226">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-226">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-227">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-227">Type</span></span>

*   [<span data-ttu-id="357bc-228">Body</span><span class="sxs-lookup"><span data-stu-id="357bc-228">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="357bc-229">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-229">Requirements</span></span>

|<span data-ttu-id="357bc-230">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-230">Requirement</span></span>| <span data-ttu-id="357bc-231">值</span><span class="sxs-lookup"><span data-stu-id="357bc-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-232">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-233">1.1</span><span class="sxs-lookup"><span data-stu-id="357bc-233">1.1</span></span>|
|[<span data-ttu-id="357bc-234">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-234">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-235">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-235">ReadItem</span></span>|
|[<span data-ttu-id="357bc-236">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-237">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-238">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-238">Example</span></span>

<span data-ttu-id="357bc-239">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="357bc-239">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="357bc-240">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="357bc-240">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="357bc-241"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的抄送： Array</span><span class="sxs-lookup"><span data-stu-id="357bc-241">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-242">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="357bc-242">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="357bc-243">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-243">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-244">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-244">Read mode</span></span>

<span data-ttu-id="357bc-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="357bc-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-247">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-247">Compose mode</span></span>

<span data-ttu-id="357bc-248">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-248">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="357bc-249">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-249">Type</span></span>

*   <span data-ttu-id="357bc-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-250">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-251">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-251">Requirements</span></span>

|<span data-ttu-id="357bc-252">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-252">Requirement</span></span>| <span data-ttu-id="357bc-253">值</span><span class="sxs-lookup"><span data-stu-id="357bc-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-254">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-255">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-255">1.0</span></span>|
|[<span data-ttu-id="357bc-256">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-257">ReadItem</span></span>|
|[<span data-ttu-id="357bc-258">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-259">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-259">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="357bc-260">（可以为 null） conversationId： String</span><span class="sxs-lookup"><span data-stu-id="357bc-260">(nullable) conversationId: String</span></span>

<span data-ttu-id="357bc-261">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="357bc-261">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="357bc-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="357bc-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="357bc-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="357bc-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-266">Type</span><span class="sxs-lookup"><span data-stu-id="357bc-266">Type</span></span>

*   <span data-ttu-id="357bc-267">String</span><span class="sxs-lookup"><span data-stu-id="357bc-267">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-268">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-268">Requirements</span></span>

|<span data-ttu-id="357bc-269">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-269">Requirement</span></span>| <span data-ttu-id="357bc-270">值</span><span class="sxs-lookup"><span data-stu-id="357bc-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-272">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-272">1.0</span></span>|
|[<span data-ttu-id="357bc-273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-274">ReadItem</span></span>|
|[<span data-ttu-id="357bc-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-277">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-277">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="357bc-278">dateTimeCreated： Date</span><span class="sxs-lookup"><span data-stu-id="357bc-278">dateTimeCreated: Date</span></span>

<span data-ttu-id="357bc-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-281">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-281">Type</span></span>

*   <span data-ttu-id="357bc-282">日期</span><span class="sxs-lookup"><span data-stu-id="357bc-282">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-283">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-283">Requirements</span></span>

|<span data-ttu-id="357bc-284">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-284">Requirement</span></span>| <span data-ttu-id="357bc-285">值</span><span class="sxs-lookup"><span data-stu-id="357bc-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-286">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-287">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-287">1.0</span></span>|
|[<span data-ttu-id="357bc-288">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-288">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-289">ReadItem</span></span>|
|[<span data-ttu-id="357bc-290">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-290">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-291">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-291">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-292">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-292">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="357bc-293">dateTimeModified： Date</span><span class="sxs-lookup"><span data-stu-id="357bc-293">dateTimeModified: Date</span></span>

<span data-ttu-id="357bc-294">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="357bc-294">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="357bc-295">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-295">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-296">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="357bc-296">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-297">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-297">Type</span></span>

*   <span data-ttu-id="357bc-298">日期</span><span class="sxs-lookup"><span data-stu-id="357bc-298">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-299">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-299">Requirements</span></span>

|<span data-ttu-id="357bc-300">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-300">Requirement</span></span>| <span data-ttu-id="357bc-301">值</span><span class="sxs-lookup"><span data-stu-id="357bc-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-303">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-303">1.0</span></span>|
|[<span data-ttu-id="357bc-304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-305">ReadItem</span></span>|
|[<span data-ttu-id="357bc-306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-307">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-308">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-308">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="357bc-309">结束：日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-309">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-310">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="357bc-310">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="357bc-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="357bc-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-313">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-313">Read mode</span></span>

<span data-ttu-id="357bc-314">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-314">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-315">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-315">Compose mode</span></span>

<span data-ttu-id="357bc-316">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-316">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="357bc-317">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="357bc-317">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="357bc-318">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="357bc-318">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="357bc-319">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-319">Type</span></span>

*   <span data-ttu-id="357bc-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-320">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-321">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-321">Requirements</span></span>

|<span data-ttu-id="357bc-322">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-322">Requirement</span></span>| <span data-ttu-id="357bc-323">值</span><span class="sxs-lookup"><span data-stu-id="357bc-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-324">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-325">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-325">1.0</span></span>|
|[<span data-ttu-id="357bc-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-326">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-327">ReadItem</span></span>|
|[<span data-ttu-id="357bc-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-328">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-329">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-329">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="357bc-330">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-330">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="357bc-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="357bc-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-335">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="357bc-335">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-336">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-336">Type</span></span>

*   [<span data-ttu-id="357bc-337">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="357bc-337">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="357bc-338">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-338">Requirements</span></span>

|<span data-ttu-id="357bc-339">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-339">Requirement</span></span>| <span data-ttu-id="357bc-340">值</span><span class="sxs-lookup"><span data-stu-id="357bc-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-342">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-342">1.0</span></span>|
|[<span data-ttu-id="357bc-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-344">ReadItem</span></span>|
|[<span data-ttu-id="357bc-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-346">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-346">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-347">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-347">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="357bc-348">internetMessageId： String</span><span class="sxs-lookup"><span data-stu-id="357bc-348">internetMessageId: String</span></span>

<span data-ttu-id="357bc-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-351">Type</span><span class="sxs-lookup"><span data-stu-id="357bc-351">Type</span></span>

*   <span data-ttu-id="357bc-352">String</span><span class="sxs-lookup"><span data-stu-id="357bc-352">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-353">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-353">Requirements</span></span>

|<span data-ttu-id="357bc-354">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-354">Requirement</span></span>| <span data-ttu-id="357bc-355">值</span><span class="sxs-lookup"><span data-stu-id="357bc-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-356">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-357">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-357">1.0</span></span>|
|[<span data-ttu-id="357bc-358">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-359">ReadItem</span></span>|
|[<span data-ttu-id="357bc-360">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-361">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-361">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-362">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-362">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="357bc-363">itemClass： String</span><span class="sxs-lookup"><span data-stu-id="357bc-363">itemClass: String</span></span>

<span data-ttu-id="357bc-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="357bc-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="357bc-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="357bc-368">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-368">Type</span></span> | <span data-ttu-id="357bc-369">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-369">Description</span></span> | <span data-ttu-id="357bc-370">项目类</span><span class="sxs-lookup"><span data-stu-id="357bc-370">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="357bc-371">约会项目</span><span class="sxs-lookup"><span data-stu-id="357bc-371">Appointment items</span></span> | <span data-ttu-id="357bc-372">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="357bc-372">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="357bc-373">邮件项目</span><span class="sxs-lookup"><span data-stu-id="357bc-373">Message items</span></span> | <span data-ttu-id="357bc-374">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="357bc-374">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="357bc-375">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="357bc-375">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-376">Type</span><span class="sxs-lookup"><span data-stu-id="357bc-376">Type</span></span>

*   <span data-ttu-id="357bc-377">String</span><span class="sxs-lookup"><span data-stu-id="357bc-377">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-378">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-378">Requirements</span></span>

|<span data-ttu-id="357bc-379">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-379">Requirement</span></span>| <span data-ttu-id="357bc-380">值</span><span class="sxs-lookup"><span data-stu-id="357bc-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-381">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-382">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-382">1.0</span></span>|
|[<span data-ttu-id="357bc-383">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-383">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-384">ReadItem</span></span>|
|[<span data-ttu-id="357bc-385">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-385">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-386">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-386">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-387">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-387">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="357bc-388">（可以为 null） itemId： String</span><span class="sxs-lookup"><span data-stu-id="357bc-388">(nullable) itemId: String</span></span>

<span data-ttu-id="357bc-389">获取当前项目的 Exchange Web 服务项目标识符。</span><span class="sxs-lookup"><span data-stu-id="357bc-389">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="357bc-390">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-390">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-391">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="357bc-391">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="357bc-392">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="357bc-392">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="357bc-393">在使用此值进行 REST API 调用之前，应使用`Office.context.mailbox.convertToRestId`转换它，这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="357bc-393">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="357bc-394">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="357bc-394">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-395">Type</span><span class="sxs-lookup"><span data-stu-id="357bc-395">Type</span></span>

*   <span data-ttu-id="357bc-396">String</span><span class="sxs-lookup"><span data-stu-id="357bc-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-397">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-397">Requirements</span></span>

|<span data-ttu-id="357bc-398">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-398">Requirement</span></span>| <span data-ttu-id="357bc-399">值</span><span class="sxs-lookup"><span data-stu-id="357bc-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-400">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-401">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-401">1.0</span></span>|
|[<span data-ttu-id="357bc-402">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-403">ReadItem</span></span>|
|[<span data-ttu-id="357bc-404">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-405">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-406">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-406">Example</span></span>

<span data-ttu-id="357bc-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="357bc-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="357bc-409">itemType： [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-409">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-410">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="357bc-410">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="357bc-411">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="357bc-411">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-412">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-412">Type</span></span>

*   [<span data-ttu-id="357bc-413">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="357bc-413">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="357bc-414">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-414">Requirements</span></span>

|<span data-ttu-id="357bc-415">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-415">Requirement</span></span>| <span data-ttu-id="357bc-416">值</span><span class="sxs-lookup"><span data-stu-id="357bc-416">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-417">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-417">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-418">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-418">1.0</span></span>|
|[<span data-ttu-id="357bc-419">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-419">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-420">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-420">ReadItem</span></span>|
|[<span data-ttu-id="357bc-421">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-421">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-422">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-422">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-423">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-423">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="357bc-424">位置：字符串 |[位置](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-424">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-425">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="357bc-425">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-426">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-426">Read mode</span></span>

<span data-ttu-id="357bc-427">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="357bc-427">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-428">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-428">Compose mode</span></span>

<span data-ttu-id="357bc-429">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-429">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="357bc-430">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-430">Type</span></span>

*   <span data-ttu-id="357bc-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-431">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-432">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-432">Requirements</span></span>

|<span data-ttu-id="357bc-433">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-433">Requirement</span></span>| <span data-ttu-id="357bc-434">值</span><span class="sxs-lookup"><span data-stu-id="357bc-434">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-435">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-435">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-436">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-436">1.0</span></span>|
|[<span data-ttu-id="357bc-437">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-437">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-438">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-438">ReadItem</span></span>|
|[<span data-ttu-id="357bc-439">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-439">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-440">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-440">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="357bc-441">normalizedSubject： String</span><span class="sxs-lookup"><span data-stu-id="357bc-441">normalizedSubject: String</span></span>

<span data-ttu-id="357bc-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="357bc-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="357bc-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-446">Type</span><span class="sxs-lookup"><span data-stu-id="357bc-446">Type</span></span>

*   <span data-ttu-id="357bc-447">String</span><span class="sxs-lookup"><span data-stu-id="357bc-447">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-448">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-448">Requirements</span></span>

|<span data-ttu-id="357bc-449">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-449">Requirement</span></span>| <span data-ttu-id="357bc-450">值</span><span class="sxs-lookup"><span data-stu-id="357bc-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-451">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-452">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-452">1.0</span></span>|
|[<span data-ttu-id="357bc-453">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-454">ReadItem</span></span>|
|[<span data-ttu-id="357bc-455">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-456">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-456">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-457">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-457">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="357bc-458">optionalAttendees： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="357bc-458">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-459">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="357bc-459">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="357bc-460">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-460">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-461">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-461">Read mode</span></span>

<span data-ttu-id="357bc-462">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-462">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-463">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-463">Compose mode</span></span>

<span data-ttu-id="357bc-464">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-464">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="357bc-465">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-465">Type</span></span>

*   <span data-ttu-id="357bc-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-466">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-467">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-467">Requirements</span></span>

|<span data-ttu-id="357bc-468">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-468">Requirement</span></span>| <span data-ttu-id="357bc-469">值</span><span class="sxs-lookup"><span data-stu-id="357bc-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-470">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-471">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-471">1.0</span></span>|
|[<span data-ttu-id="357bc-472">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-473">ReadItem</span></span>|
|[<span data-ttu-id="357bc-474">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-475">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-475">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="357bc-476">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-476">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-479">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-479">Type</span></span>

*   [<span data-ttu-id="357bc-480">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="357bc-480">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="357bc-481">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-481">Requirements</span></span>

|<span data-ttu-id="357bc-482">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-482">Requirement</span></span>| <span data-ttu-id="357bc-483">值</span><span class="sxs-lookup"><span data-stu-id="357bc-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-485">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-485">1.0</span></span>|
|[<span data-ttu-id="357bc-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-487">ReadItem</span></span>|
|[<span data-ttu-id="357bc-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-489">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-490">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-490">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="357bc-491">requiredAttendees： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="357bc-491">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-492">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="357bc-492">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="357bc-493">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-493">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-494">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-494">Read mode</span></span>

<span data-ttu-id="357bc-495">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-495">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-496">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-496">Compose mode</span></span>

<span data-ttu-id="357bc-497">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-497">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="357bc-498">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-498">Type</span></span>

*   <span data-ttu-id="357bc-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-499">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-500">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-500">Requirements</span></span>

|<span data-ttu-id="357bc-501">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-501">Requirement</span></span>| <span data-ttu-id="357bc-502">值</span><span class="sxs-lookup"><span data-stu-id="357bc-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-504">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-504">1.0</span></span>|
|[<span data-ttu-id="357bc-505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-506">ReadItem</span></span>|
|[<span data-ttu-id="357bc-507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-508">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-508">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="357bc-509">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-509">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="357bc-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="357bc-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-514">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="357bc-514">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="357bc-515">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-515">Type</span></span>

*   [<span data-ttu-id="357bc-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="357bc-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="357bc-517">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-517">Requirements</span></span>

|<span data-ttu-id="357bc-518">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-518">Requirement</span></span>| <span data-ttu-id="357bc-519">值</span><span class="sxs-lookup"><span data-stu-id="357bc-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-520">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-521">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-521">1.0</span></span>|
|[<span data-ttu-id="357bc-522">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-523">ReadItem</span></span>|
|[<span data-ttu-id="357bc-524">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-525">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-526">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-526">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="357bc-527">开始日期：日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-527">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-528">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="357bc-528">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="357bc-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="357bc-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-531">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-531">Read mode</span></span>

<span data-ttu-id="357bc-532">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-532">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-533">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-533">Compose mode</span></span>

<span data-ttu-id="357bc-534">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-534">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="357bc-535">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="357bc-535">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="357bc-536">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="357bc-536">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="357bc-537">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-537">Type</span></span>

*   <span data-ttu-id="357bc-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-538">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-539">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-539">Requirements</span></span>

|<span data-ttu-id="357bc-540">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-540">Requirement</span></span>| <span data-ttu-id="357bc-541">值</span><span class="sxs-lookup"><span data-stu-id="357bc-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-542">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-543">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-543">1.0</span></span>|
|[<span data-ttu-id="357bc-544">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-545">ReadItem</span></span>|
|[<span data-ttu-id="357bc-546">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-547">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-547">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="357bc-548">subject： String |[主题](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-548">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-549">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="357bc-549">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="357bc-550">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="357bc-550">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-551">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-551">Read mode</span></span>

<span data-ttu-id="357bc-p130">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="357bc-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-554">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-554">Compose mode</span></span>

<span data-ttu-id="357bc-555">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-555">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="357bc-556">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-556">Type</span></span>

*   <span data-ttu-id="357bc-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-557">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-558">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-558">Requirements</span></span>

|<span data-ttu-id="357bc-559">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-559">Requirement</span></span>| <span data-ttu-id="357bc-560">值</span><span class="sxs-lookup"><span data-stu-id="357bc-560">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-561">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-562">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-562">1.0</span></span>|
|[<span data-ttu-id="357bc-563">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-563">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-564">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-564">ReadItem</span></span>|
|[<span data-ttu-id="357bc-565">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-565">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-566">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-566">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="357bc-567">to： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="357bc-567">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="357bc-568">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="357bc-568">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="357bc-569">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="357bc-569">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="357bc-570">阅读模式</span><span class="sxs-lookup"><span data-stu-id="357bc-570">Read mode</span></span>

<span data-ttu-id="357bc-p132">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="357bc-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="357bc-573">撰写模式</span><span class="sxs-lookup"><span data-stu-id="357bc-573">Compose mode</span></span>

<span data-ttu-id="357bc-574">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-574">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="357bc-575">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-575">Type</span></span>

*   <span data-ttu-id="357bc-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-576">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-577">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-577">Requirements</span></span>

|<span data-ttu-id="357bc-578">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-578">Requirement</span></span>| <span data-ttu-id="357bc-579">值</span><span class="sxs-lookup"><span data-stu-id="357bc-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-580">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-581">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-581">1.0</span></span>|
|[<span data-ttu-id="357bc-582">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-583">ReadItem</span></span>|
|[<span data-ttu-id="357bc-584">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-585">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-585">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="357bc-586">方法</span><span class="sxs-lookup"><span data-stu-id="357bc-586">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="357bc-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="357bc-587">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="357bc-588">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="357bc-588">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="357bc-589">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="357bc-589">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="357bc-590">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="357bc-590">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-591">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-591">Parameters</span></span>

|<span data-ttu-id="357bc-592">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-592">Name</span></span>| <span data-ttu-id="357bc-593">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-593">Type</span></span>| <span data-ttu-id="357bc-594">属性</span><span class="sxs-lookup"><span data-stu-id="357bc-594">Attributes</span></span>| <span data-ttu-id="357bc-595">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-595">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="357bc-596">String</span><span class="sxs-lookup"><span data-stu-id="357bc-596">String</span></span>||<span data-ttu-id="357bc-p133">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="357bc-599">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-599">String</span></span>||<span data-ttu-id="357bc-p134">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="357bc-602">Object</span><span class="sxs-lookup"><span data-stu-id="357bc-602">Object</span></span>| <span data-ttu-id="357bc-603">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-603">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-604">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="357bc-604">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="357bc-605">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-605">Object</span></span>| <span data-ttu-id="357bc-606">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-606">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-607">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-607">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="357bc-608">函数</span><span class="sxs-lookup"><span data-stu-id="357bc-608">function</span></span>| <span data-ttu-id="357bc-609">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-609">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-610">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-610">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="357bc-611">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="357bc-611">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="357bc-612">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-612">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="357bc-613">错误</span><span class="sxs-lookup"><span data-stu-id="357bc-613">Errors</span></span>

| <span data-ttu-id="357bc-614">错误代码</span><span class="sxs-lookup"><span data-stu-id="357bc-614">Error code</span></span> | <span data-ttu-id="357bc-615">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-615">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="357bc-616">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="357bc-616">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="357bc-617">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="357bc-617">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="357bc-618">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="357bc-618">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="357bc-619">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-619">Requirements</span></span>

|<span data-ttu-id="357bc-620">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-620">Requirement</span></span>| <span data-ttu-id="357bc-621">值</span><span class="sxs-lookup"><span data-stu-id="357bc-621">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-622">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-622">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-623">1.1</span><span class="sxs-lookup"><span data-stu-id="357bc-623">1.1</span></span>|
|[<span data-ttu-id="357bc-624">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-624">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-625">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="357bc-625">ReadWriteItem</span></span>|
|[<span data-ttu-id="357bc-626">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-626">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-627">撰写</span><span class="sxs-lookup"><span data-stu-id="357bc-627">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-628">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-628">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="357bc-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="357bc-629">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="357bc-630">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="357bc-630">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="357bc-p135">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="357bc-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="357bc-634">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="357bc-634">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="357bc-635">如果 Office 外接程序在 web 上的 Outlook 中运行，则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是，不支持这种情况，建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="357bc-635">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-636">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-636">Parameters</span></span>

|<span data-ttu-id="357bc-637">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-637">Name</span></span>| <span data-ttu-id="357bc-638">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-638">Type</span></span>| <span data-ttu-id="357bc-639">属性</span><span class="sxs-lookup"><span data-stu-id="357bc-639">Attributes</span></span>| <span data-ttu-id="357bc-640">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-640">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="357bc-641">String</span><span class="sxs-lookup"><span data-stu-id="357bc-641">String</span></span>||<span data-ttu-id="357bc-p136">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="357bc-644">String</span><span class="sxs-lookup"><span data-stu-id="357bc-644">String</span></span>||<span data-ttu-id="357bc-645">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="357bc-645">The subject of the item to be attached.</span></span> <span data-ttu-id="357bc-646">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-646">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="357bc-647">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-647">Object</span></span>| <span data-ttu-id="357bc-648">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-648">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-649">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="357bc-649">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="357bc-650">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-650">Object</span></span>| <span data-ttu-id="357bc-651">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-651">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-652">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-652">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="357bc-653">函数</span><span class="sxs-lookup"><span data-stu-id="357bc-653">function</span></span>| <span data-ttu-id="357bc-654">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-654">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-655">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-655">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="357bc-656">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="357bc-656">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="357bc-657">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-657">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="357bc-658">错误</span><span class="sxs-lookup"><span data-stu-id="357bc-658">Errors</span></span>

| <span data-ttu-id="357bc-659">错误代码</span><span class="sxs-lookup"><span data-stu-id="357bc-659">Error code</span></span> | <span data-ttu-id="357bc-660">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-660">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="357bc-661">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="357bc-661">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="357bc-662">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-662">Requirements</span></span>

|<span data-ttu-id="357bc-663">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-663">Requirement</span></span>| <span data-ttu-id="357bc-664">值</span><span class="sxs-lookup"><span data-stu-id="357bc-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-665">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-666">1.1</span><span class="sxs-lookup"><span data-stu-id="357bc-666">1.1</span></span>|
|[<span data-ttu-id="357bc-667">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-668">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="357bc-668">ReadWriteItem</span></span>|
|[<span data-ttu-id="357bc-669">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-670">撰写</span><span class="sxs-lookup"><span data-stu-id="357bc-670">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-671">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-671">Example</span></span>

<span data-ttu-id="357bc-672">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="357bc-672">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="357bc-673">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="357bc-673">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="357bc-674">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="357bc-674">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-675">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-675">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="357bc-676">在 web 上的 Outlook 中，答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="357bc-676">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="357bc-677">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="357bc-677">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="357bc-678">如果在`formData.attachments`参数中指定了附件，则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="357bc-678">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="357bc-679">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="357bc-679">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="357bc-680">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="357bc-680">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-681">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-681">Parameters</span></span>

|<span data-ttu-id="357bc-682">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-682">Name</span></span>| <span data-ttu-id="357bc-683">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-683">Type</span></span>| <span data-ttu-id="357bc-684">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-684">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="357bc-685">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="357bc-685">String &#124; Object</span></span>| |<span data-ttu-id="357bc-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="357bc-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="357bc-688">**或**</span><span class="sxs-lookup"><span data-stu-id="357bc-688">**OR**</span></span><br/><span data-ttu-id="357bc-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="357bc-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="357bc-691">String</span><span class="sxs-lookup"><span data-stu-id="357bc-691">String</span></span> | <span data-ttu-id="357bc-692">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-692">&lt;optional&gt;</span></span> | <span data-ttu-id="357bc-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="357bc-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="357bc-695">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-695">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="357bc-696">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-696">&lt;optional&gt;</span></span> | <span data-ttu-id="357bc-697">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="357bc-697">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="357bc-698">String</span><span class="sxs-lookup"><span data-stu-id="357bc-698">String</span></span> | | <span data-ttu-id="357bc-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="357bc-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="357bc-701">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-701">String</span></span> | | <span data-ttu-id="357bc-702">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-702">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="357bc-703">String</span><span class="sxs-lookup"><span data-stu-id="357bc-703">String</span></span> | | <span data-ttu-id="357bc-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="357bc-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="357bc-706">String</span><span class="sxs-lookup"><span data-stu-id="357bc-706">String</span></span> | | <span data-ttu-id="357bc-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="357bc-710">函数</span><span class="sxs-lookup"><span data-stu-id="357bc-710">function</span></span> | <span data-ttu-id="357bc-711">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-711">&lt;optional&gt;</span></span> | <span data-ttu-id="357bc-712">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-712">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="357bc-713">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-713">Requirements</span></span>

|<span data-ttu-id="357bc-714">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-714">Requirement</span></span>| <span data-ttu-id="357bc-715">值</span><span class="sxs-lookup"><span data-stu-id="357bc-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-716">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-717">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-717">1.0</span></span>|
|[<span data-ttu-id="357bc-718">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-718">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-719">ReadItem</span></span>|
|[<span data-ttu-id="357bc-720">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-720">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-721">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-721">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="357bc-722">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-722">Examples</span></span>

<span data-ttu-id="357bc-723">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-723">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="357bc-724">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-724">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="357bc-725">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-725">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="357bc-726">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-726">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="357bc-727">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-727">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="357bc-728">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-728">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="357bc-729">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="357bc-729">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="357bc-730">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="357bc-730">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-731">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-731">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="357bc-732">在 web 上的 Outlook 中，答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="357bc-732">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="357bc-733">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="357bc-733">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="357bc-734">如果在`formData.attachments`参数中指定了附件，则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="357bc-734">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="357bc-735">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="357bc-735">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="357bc-736">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="357bc-736">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-737">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-737">Parameters</span></span>

|<span data-ttu-id="357bc-738">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-738">Name</span></span>| <span data-ttu-id="357bc-739">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-739">Type</span></span>| <span data-ttu-id="357bc-740">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-740">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="357bc-741">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="357bc-741">String &#124; Object</span></span>| | <span data-ttu-id="357bc-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="357bc-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="357bc-744">**或**</span><span class="sxs-lookup"><span data-stu-id="357bc-744">**OR**</span></span><br/><span data-ttu-id="357bc-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="357bc-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="357bc-747">String</span><span class="sxs-lookup"><span data-stu-id="357bc-747">String</span></span> | <span data-ttu-id="357bc-748">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-748">&lt;optional&gt;</span></span> | <span data-ttu-id="357bc-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="357bc-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="357bc-751">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-751">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="357bc-752">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-752">&lt;optional&gt;</span></span> | <span data-ttu-id="357bc-753">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="357bc-753">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="357bc-754">String</span><span class="sxs-lookup"><span data-stu-id="357bc-754">String</span></span> | | <span data-ttu-id="357bc-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="357bc-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="357bc-757">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-757">String</span></span> | | <span data-ttu-id="357bc-758">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-758">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="357bc-759">String</span><span class="sxs-lookup"><span data-stu-id="357bc-759">String</span></span> | | <span data-ttu-id="357bc-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="357bc-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="357bc-762">String</span><span class="sxs-lookup"><span data-stu-id="357bc-762">String</span></span> | | <span data-ttu-id="357bc-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="357bc-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="357bc-766">函数</span><span class="sxs-lookup"><span data-stu-id="357bc-766">function</span></span> | <span data-ttu-id="357bc-767">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-767">&lt;optional&gt;</span></span> | <span data-ttu-id="357bc-768">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="357bc-769">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-769">Requirements</span></span>

|<span data-ttu-id="357bc-770">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-770">Requirement</span></span>| <span data-ttu-id="357bc-771">值</span><span class="sxs-lookup"><span data-stu-id="357bc-771">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-772">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-772">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-773">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-773">1.0</span></span>|
|[<span data-ttu-id="357bc-774">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-774">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-775">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-775">ReadItem</span></span>|
|[<span data-ttu-id="357bc-776">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-776">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-777">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-777">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="357bc-778">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-778">Examples</span></span>

<span data-ttu-id="357bc-779">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-779">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="357bc-780">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-780">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="357bc-781">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-781">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="357bc-782">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-782">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="357bc-783">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-783">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="357bc-784">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="357bc-784">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="357bc-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="357bc-785">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="357bc-786">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="357bc-786">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-787">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-787">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-788">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-788">Requirements</span></span>

|<span data-ttu-id="357bc-789">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-789">Requirement</span></span>| <span data-ttu-id="357bc-790">值</span><span class="sxs-lookup"><span data-stu-id="357bc-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-791">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-792">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-792">1.0</span></span>|
|[<span data-ttu-id="357bc-793">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-793">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-794">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-794">ReadItem</span></span>|
|[<span data-ttu-id="357bc-795">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-795">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-796">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-796">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="357bc-797">返回：</span><span class="sxs-lookup"><span data-stu-id="357bc-797">Returns:</span></span>

<span data-ttu-id="357bc-798">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="357bc-798">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="357bc-799">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-799">Example</span></span>

<span data-ttu-id="357bc-800">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="357bc-800">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="357bc-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="357bc-801">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="357bc-802">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="357bc-802">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-803">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-803">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-804">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-804">Parameters</span></span>

|<span data-ttu-id="357bc-805">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-805">Name</span></span>| <span data-ttu-id="357bc-806">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-806">Type</span></span>| <span data-ttu-id="357bc-807">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-807">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="357bc-808">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="357bc-808">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="357bc-809">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="357bc-809">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="357bc-810">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-810">Requirements</span></span>

|<span data-ttu-id="357bc-811">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-811">Requirement</span></span>| <span data-ttu-id="357bc-812">值</span><span class="sxs-lookup"><span data-stu-id="357bc-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-813">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-814">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-814">1.0</span></span>|
|[<span data-ttu-id="357bc-815">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-815">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-816">受限</span><span class="sxs-lookup"><span data-stu-id="357bc-816">Restricted</span></span>|
|[<span data-ttu-id="357bc-817">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-817">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-818">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-818">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="357bc-819">返回：</span><span class="sxs-lookup"><span data-stu-id="357bc-819">Returns:</span></span>

<span data-ttu-id="357bc-820">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="357bc-820">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="357bc-821">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="357bc-821">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="357bc-822">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="357bc-822">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="357bc-823">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="357bc-823">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="357bc-824">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="357bc-824">Value of `entityType`</span></span> | <span data-ttu-id="357bc-825">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="357bc-825">Type of objects in returned array</span></span> | <span data-ttu-id="357bc-826">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-826">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="357bc-827">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-827">String</span></span> | <span data-ttu-id="357bc-828">**受限**</span><span class="sxs-lookup"><span data-stu-id="357bc-828">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="357bc-829">Contact</span><span class="sxs-lookup"><span data-stu-id="357bc-829">Contact</span></span> | <span data-ttu-id="357bc-830">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="357bc-830">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="357bc-831">String</span><span class="sxs-lookup"><span data-stu-id="357bc-831">String</span></span> | <span data-ttu-id="357bc-832">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="357bc-832">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="357bc-833">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="357bc-833">MeetingSuggestion</span></span> | <span data-ttu-id="357bc-834">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="357bc-834">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="357bc-835">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="357bc-835">PhoneNumber</span></span> | <span data-ttu-id="357bc-836">**受限**</span><span class="sxs-lookup"><span data-stu-id="357bc-836">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="357bc-837">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="357bc-837">TaskSuggestion</span></span> | <span data-ttu-id="357bc-838">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="357bc-838">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="357bc-839">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-839">String</span></span> | <span data-ttu-id="357bc-840">**受限**</span><span class="sxs-lookup"><span data-stu-id="357bc-840">**Restricted**</span></span> |

<span data-ttu-id="357bc-841">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="357bc-841">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

##### <a name="example"></a><span data-ttu-id="357bc-842">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-842">Example</span></span>

<span data-ttu-id="357bc-843">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="357bc-843">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-12meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-12phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-12tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="357bc-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span><span class="sxs-lookup"><span data-stu-id="357bc-844">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))>}</span></span>

<span data-ttu-id="357bc-845">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="357bc-845">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-846">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-846">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="357bc-847">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="357bc-847">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-848">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-848">Parameters</span></span>

|<span data-ttu-id="357bc-849">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-849">Name</span></span>| <span data-ttu-id="357bc-850">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-850">Type</span></span>| <span data-ttu-id="357bc-851">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-851">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="357bc-852">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-852">String</span></span>|<span data-ttu-id="357bc-853">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="357bc-853">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="357bc-854">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-854">Requirements</span></span>

|<span data-ttu-id="357bc-855">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-855">Requirement</span></span>| <span data-ttu-id="357bc-856">值</span><span class="sxs-lookup"><span data-stu-id="357bc-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-857">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-858">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-858">1.0</span></span>|
|[<span data-ttu-id="357bc-859">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-859">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-860">ReadItem</span></span>|
|[<span data-ttu-id="357bc-861">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-861">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-862">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-862">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="357bc-863">返回：</span><span class="sxs-lookup"><span data-stu-id="357bc-863">Returns:</span></span>

<span data-ttu-id="357bc-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="357bc-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="357bc-866">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span><span class="sxs-lookup"><span data-stu-id="357bc-866">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.2)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.2)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.2)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.2))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="357bc-867">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="357bc-867">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="357bc-868">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="357bc-868">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-869">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-869">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="357bc-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="357bc-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="357bc-873">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="357bc-873">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="357bc-874">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="357bc-874">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="357bc-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="357bc-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="357bc-877">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-877">Requirements</span></span>

|<span data-ttu-id="357bc-878">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-878">Requirement</span></span>| <span data-ttu-id="357bc-879">值</span><span class="sxs-lookup"><span data-stu-id="357bc-879">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-880">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-880">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-881">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-881">1.0</span></span>|
|[<span data-ttu-id="357bc-882">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-882">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-883">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-883">ReadItem</span></span>|
|[<span data-ttu-id="357bc-884">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-884">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-885">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-885">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="357bc-886">返回：</span><span class="sxs-lookup"><span data-stu-id="357bc-886">Returns:</span></span>

<span data-ttu-id="357bc-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="357bc-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="357bc-889">类型：对象</span><span class="sxs-lookup"><span data-stu-id="357bc-889">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="357bc-890">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-890">Example</span></span>

<span data-ttu-id="357bc-891">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="357bc-891">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="357bc-892">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="357bc-892">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="357bc-893">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="357bc-893">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="357bc-894">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="357bc-894">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="357bc-895">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="357bc-895">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="357bc-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="357bc-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-898">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-898">Parameters</span></span>

|<span data-ttu-id="357bc-899">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-899">Name</span></span>| <span data-ttu-id="357bc-900">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-900">Type</span></span>| <span data-ttu-id="357bc-901">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-901">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="357bc-902">String</span><span class="sxs-lookup"><span data-stu-id="357bc-902">String</span></span>|<span data-ttu-id="357bc-903">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="357bc-903">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="357bc-904">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-904">Requirements</span></span>

|<span data-ttu-id="357bc-905">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-905">Requirement</span></span>| <span data-ttu-id="357bc-906">值</span><span class="sxs-lookup"><span data-stu-id="357bc-906">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-907">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-907">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-908">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-908">1.0</span></span>|
|[<span data-ttu-id="357bc-909">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-909">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-910">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-910">ReadItem</span></span>|
|[<span data-ttu-id="357bc-911">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-911">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-912">阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-912">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="357bc-913">返回：</span><span class="sxs-lookup"><span data-stu-id="357bc-913">Returns:</span></span>

<span data-ttu-id="357bc-914">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="357bc-914">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="357bc-915">类型： Array. < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="357bc-915">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="357bc-916">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-916">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="357bc-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="357bc-917">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="357bc-918">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="357bc-918">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="357bc-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="357bc-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-921">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-921">Parameters</span></span>

|<span data-ttu-id="357bc-922">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-922">Name</span></span>| <span data-ttu-id="357bc-923">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-923">Type</span></span>| <span data-ttu-id="357bc-924">属性</span><span class="sxs-lookup"><span data-stu-id="357bc-924">Attributes</span></span>| <span data-ttu-id="357bc-925">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-925">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="357bc-926">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="357bc-926">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="357bc-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="357bc-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="357bc-930">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-930">Object</span></span>| <span data-ttu-id="357bc-931">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-931">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-932">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="357bc-932">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="357bc-933">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-933">Object</span></span>| <span data-ttu-id="357bc-934">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-934">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-935">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-935">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="357bc-936">function</span><span class="sxs-lookup"><span data-stu-id="357bc-936">function</span></span>||<span data-ttu-id="357bc-937">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="357bc-938">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="357bc-938">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="357bc-939">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="357bc-939">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="357bc-940">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-940">Requirements</span></span>

|<span data-ttu-id="357bc-941">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-941">Requirement</span></span>| <span data-ttu-id="357bc-942">值</span><span class="sxs-lookup"><span data-stu-id="357bc-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-943">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-944">1.2</span><span class="sxs-lookup"><span data-stu-id="357bc-944">1.2</span></span>|
|[<span data-ttu-id="357bc-945">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-946">ReadItem</span></span>|
|[<span data-ttu-id="357bc-947">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-948">撰写</span><span class="sxs-lookup"><span data-stu-id="357bc-948">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="357bc-949">返回：</span><span class="sxs-lookup"><span data-stu-id="357bc-949">Returns:</span></span>

<span data-ttu-id="357bc-950">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="357bc-950">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="357bc-951">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-951">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="357bc-952">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-952">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="357bc-953">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="357bc-953">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="357bc-954">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="357bc-954">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="357bc-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="357bc-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-958">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-958">Parameters</span></span>

|<span data-ttu-id="357bc-959">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-959">Name</span></span>| <span data-ttu-id="357bc-960">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-960">Type</span></span>| <span data-ttu-id="357bc-961">属性</span><span class="sxs-lookup"><span data-stu-id="357bc-961">Attributes</span></span>| <span data-ttu-id="357bc-962">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-962">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="357bc-963">函数</span><span class="sxs-lookup"><span data-stu-id="357bc-963">function</span></span>||<span data-ttu-id="357bc-964">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-964">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="357bc-965">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="357bc-965">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="357bc-966">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="357bc-966">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="357bc-967">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-967">Object</span></span>| <span data-ttu-id="357bc-968">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-968">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-969">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-969">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="357bc-970">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="357bc-970">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="357bc-971">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-971">Requirements</span></span>

|<span data-ttu-id="357bc-972">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-972">Requirement</span></span>| <span data-ttu-id="357bc-973">值</span><span class="sxs-lookup"><span data-stu-id="357bc-973">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-974">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-974">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-975">1.0</span><span class="sxs-lookup"><span data-stu-id="357bc-975">1.0</span></span>|
|[<span data-ttu-id="357bc-976">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-976">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-977">ReadItem</span><span class="sxs-lookup"><span data-stu-id="357bc-977">ReadItem</span></span>|
|[<span data-ttu-id="357bc-978">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-978">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-979">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="357bc-979">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-980">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-980">Example</span></span>

<span data-ttu-id="357bc-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="357bc-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="357bc-984">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="357bc-984">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="357bc-985">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="357bc-985">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="357bc-986">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="357bc-986">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="357bc-987">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="357bc-987">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="357bc-988">在 web 和移动设备上的 Outlook 中，附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="357bc-988">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="357bc-989">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="357bc-989">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-990">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-990">Parameters</span></span>

|<span data-ttu-id="357bc-991">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-991">Name</span></span>| <span data-ttu-id="357bc-992">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-992">Type</span></span>| <span data-ttu-id="357bc-993">属性</span><span class="sxs-lookup"><span data-stu-id="357bc-993">Attributes</span></span>| <span data-ttu-id="357bc-994">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-994">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="357bc-995">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-995">String</span></span>||<span data-ttu-id="357bc-996">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="357bc-996">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="357bc-997">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-997">Object</span></span>| <span data-ttu-id="357bc-998">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-998">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-999">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="357bc-999">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="357bc-1000">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-1000">Object</span></span>| <span data-ttu-id="357bc-1001">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-1001">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-1002">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-1002">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="357bc-1003">函数</span><span class="sxs-lookup"><span data-stu-id="357bc-1003">function</span></span>| <span data-ttu-id="357bc-1004">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-1005">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-1005">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="357bc-1006">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="357bc-1006">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="357bc-1007">错误</span><span class="sxs-lookup"><span data-stu-id="357bc-1007">Errors</span></span>

| <span data-ttu-id="357bc-1008">错误代码</span><span class="sxs-lookup"><span data-stu-id="357bc-1008">Error code</span></span> | <span data-ttu-id="357bc-1009">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-1009">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="357bc-1010">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="357bc-1010">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="357bc-1011">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-1011">Requirements</span></span>

|<span data-ttu-id="357bc-1012">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-1012">Requirement</span></span>| <span data-ttu-id="357bc-1013">值</span><span class="sxs-lookup"><span data-stu-id="357bc-1013">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-1014">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-1014">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-1015">1.1</span><span class="sxs-lookup"><span data-stu-id="357bc-1015">1.1</span></span>|
|[<span data-ttu-id="357bc-1016">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-1016">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-1017">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="357bc-1017">ReadWriteItem</span></span>|
|[<span data-ttu-id="357bc-1018">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-1018">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-1019">撰写</span><span class="sxs-lookup"><span data-stu-id="357bc-1019">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-1020">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-1020">Example</span></span>

<span data-ttu-id="357bc-1021">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="357bc-1021">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="357bc-1022">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="357bc-1022">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="357bc-1023">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="357bc-1023">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="357bc-p166">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="357bc-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="357bc-1027">参数</span><span class="sxs-lookup"><span data-stu-id="357bc-1027">Parameters</span></span>

|<span data-ttu-id="357bc-1028">名称</span><span class="sxs-lookup"><span data-stu-id="357bc-1028">Name</span></span>| <span data-ttu-id="357bc-1029">类型</span><span class="sxs-lookup"><span data-stu-id="357bc-1029">Type</span></span>| <span data-ttu-id="357bc-1030">属性</span><span class="sxs-lookup"><span data-stu-id="357bc-1030">Attributes</span></span>| <span data-ttu-id="357bc-1031">说明</span><span class="sxs-lookup"><span data-stu-id="357bc-1031">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="357bc-1032">字符串</span><span class="sxs-lookup"><span data-stu-id="357bc-1032">String</span></span>||<span data-ttu-id="357bc-p167">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="357bc-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="357bc-1036">Object</span><span class="sxs-lookup"><span data-stu-id="357bc-1036">Object</span></span>| <span data-ttu-id="357bc-1037">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-1038">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="357bc-1038">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="357bc-1039">对象</span><span class="sxs-lookup"><span data-stu-id="357bc-1039">Object</span></span>| <span data-ttu-id="357bc-1040">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-1041">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="357bc-1041">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="357bc-1042">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="357bc-1042">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="357bc-1043">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="357bc-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="357bc-1044">如果`text`为，则当前样式应用于 web 上的 Outlook 和桌面客户端。</span><span class="sxs-lookup"><span data-stu-id="357bc-1044">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="357bc-1045">如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="357bc-1045">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="357bc-1046">如果`html`和字段支持 HTML （主题不），则当前样式应用于 web 上的 outlook，并且在 outlook 桌面客户端中应用了默认样式。</span><span class="sxs-lookup"><span data-stu-id="357bc-1046">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="357bc-1047">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="357bc-1047">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="357bc-1048">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="357bc-1048">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="357bc-1049">function</span><span class="sxs-lookup"><span data-stu-id="357bc-1049">function</span></span>||<span data-ttu-id="357bc-1050">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="357bc-1050">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="357bc-1051">Requirements</span><span class="sxs-lookup"><span data-stu-id="357bc-1051">Requirements</span></span>

|<span data-ttu-id="357bc-1052">要求</span><span class="sxs-lookup"><span data-stu-id="357bc-1052">Requirement</span></span>| <span data-ttu-id="357bc-1053">值</span><span class="sxs-lookup"><span data-stu-id="357bc-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="357bc-1054">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="357bc-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="357bc-1055">1.2</span><span class="sxs-lookup"><span data-stu-id="357bc-1055">1.2</span></span>|
|[<span data-ttu-id="357bc-1056">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="357bc-1056">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="357bc-1057">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="357bc-1057">ReadWriteItem</span></span>|
|[<span data-ttu-id="357bc-1058">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="357bc-1058">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="357bc-1059">撰写</span><span class="sxs-lookup"><span data-stu-id="357bc-1059">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="357bc-1060">示例</span><span class="sxs-lookup"><span data-stu-id="357bc-1060">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
