---
title: "\"Context\"-\"邮箱\"。项目-要求集1。4"
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 4a8a97403c43e4af4ee5d840d7a4fd843d5bd398
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167338"
---
# <a name="item"></a><span data-ttu-id="4db53-102">item</span><span class="sxs-lookup"><span data-stu-id="4db53-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="4db53-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="4db53-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="4db53-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="4db53-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="4db53-106">Requirements</span></span>

|<span data-ttu-id="4db53-107">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-107">Requirement</span></span>| <span data-ttu-id="4db53-108">值</span><span class="sxs-lookup"><span data-stu-id="4db53-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-110">1.0</span></span>|
|[<span data-ttu-id="4db53-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-112">受限</span><span class="sxs-lookup"><span data-stu-id="4db53-112">Restricted</span></span>|
|[<span data-ttu-id="4db53-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4db53-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="4db53-115">Members and methods</span></span>

| <span data-ttu-id="4db53-116">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-116">Member</span></span> | <span data-ttu-id="4db53-117">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4db53-118">attachments</span><span class="sxs-lookup"><span data-stu-id="4db53-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="4db53-119">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-119">Member</span></span> |
| [<span data-ttu-id="4db53-120">bcc</span><span class="sxs-lookup"><span data-stu-id="4db53-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="4db53-121">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-121">Member</span></span> |
| [<span data-ttu-id="4db53-122">body</span><span class="sxs-lookup"><span data-stu-id="4db53-122">body</span></span>](#body-body) | <span data-ttu-id="4db53-123">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-123">Member</span></span> |
| [<span data-ttu-id="4db53-124">cc</span><span class="sxs-lookup"><span data-stu-id="4db53-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4db53-125">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-125">Member</span></span> |
| [<span data-ttu-id="4db53-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="4db53-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="4db53-127">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-127">Member</span></span> |
| [<span data-ttu-id="4db53-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="4db53-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="4db53-129">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-129">Member</span></span> |
| [<span data-ttu-id="4db53-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="4db53-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="4db53-131">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-131">Member</span></span> |
| [<span data-ttu-id="4db53-132">end</span><span class="sxs-lookup"><span data-stu-id="4db53-132">end</span></span>](#end-datetime) | <span data-ttu-id="4db53-133">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-133">Member</span></span> |
| [<span data-ttu-id="4db53-134">from</span><span class="sxs-lookup"><span data-stu-id="4db53-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="4db53-135">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-135">Member</span></span> |
| [<span data-ttu-id="4db53-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="4db53-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="4db53-137">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-137">Member</span></span> |
| [<span data-ttu-id="4db53-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="4db53-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="4db53-139">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-139">Member</span></span> |
| [<span data-ttu-id="4db53-140">itemId</span><span class="sxs-lookup"><span data-stu-id="4db53-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="4db53-141">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-141">Member</span></span> |
| [<span data-ttu-id="4db53-142">itemType</span><span class="sxs-lookup"><span data-stu-id="4db53-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="4db53-143">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-143">Member</span></span> |
| [<span data-ttu-id="4db53-144">location</span><span class="sxs-lookup"><span data-stu-id="4db53-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="4db53-145">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-145">Member</span></span> |
| [<span data-ttu-id="4db53-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="4db53-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="4db53-147">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-147">Member</span></span> |
| [<span data-ttu-id="4db53-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="4db53-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="4db53-149">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-149">Member</span></span> |
| [<span data-ttu-id="4db53-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="4db53-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4db53-151">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-151">Member</span></span> |
| [<span data-ttu-id="4db53-152">organizer</span><span class="sxs-lookup"><span data-stu-id="4db53-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="4db53-153">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-153">Member</span></span> |
| [<span data-ttu-id="4db53-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="4db53-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4db53-155">Member</span><span class="sxs-lookup"><span data-stu-id="4db53-155">Member</span></span> |
| [<span data-ttu-id="4db53-156">sender</span><span class="sxs-lookup"><span data-stu-id="4db53-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="4db53-157">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-157">Member</span></span> |
| [<span data-ttu-id="4db53-158">start</span><span class="sxs-lookup"><span data-stu-id="4db53-158">start</span></span>](#start-datetime) | <span data-ttu-id="4db53-159">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-159">Member</span></span> |
| [<span data-ttu-id="4db53-160">subject</span><span class="sxs-lookup"><span data-stu-id="4db53-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="4db53-161">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-161">Member</span></span> |
| [<span data-ttu-id="4db53-162">to</span><span class="sxs-lookup"><span data-stu-id="4db53-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="4db53-163">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-163">Member</span></span> |
| [<span data-ttu-id="4db53-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4db53-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="4db53-165">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-165">Method</span></span> |
| [<span data-ttu-id="4db53-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4db53-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="4db53-167">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-167">Method</span></span> |
| [<span data-ttu-id="4db53-168">close</span><span class="sxs-lookup"><span data-stu-id="4db53-168">close</span></span>](#close) | <span data-ttu-id="4db53-169">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-169">Method</span></span> |
| [<span data-ttu-id="4db53-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="4db53-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="4db53-171">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-171">Method</span></span> |
| [<span data-ttu-id="4db53-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="4db53-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="4db53-173">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-173">Method</span></span> |
| [<span data-ttu-id="4db53-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="4db53-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="4db53-175">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-175">Method</span></span> |
| [<span data-ttu-id="4db53-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="4db53-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4db53-177">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-177">Method</span></span> |
| [<span data-ttu-id="4db53-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="4db53-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="4db53-179">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-179">Method</span></span> |
| [<span data-ttu-id="4db53-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="4db53-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="4db53-181">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-181">Method</span></span> |
| [<span data-ttu-id="4db53-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="4db53-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="4db53-183">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-183">Method</span></span> |
| [<span data-ttu-id="4db53-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4db53-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="4db53-185">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-185">Method</span></span> |
| [<span data-ttu-id="4db53-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="4db53-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="4db53-187">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-187">Method</span></span> |
| [<span data-ttu-id="4db53-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="4db53-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="4db53-189">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-189">Method</span></span> |
| [<span data-ttu-id="4db53-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="4db53-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="4db53-191">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-191">Method</span></span> |
| [<span data-ttu-id="4db53-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="4db53-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="4db53-193">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="4db53-194">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-194">Example</span></span>

<span data-ttu-id="4db53-195">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="4db53-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4db53-196">成员</span><span class="sxs-lookup"><span data-stu-id="4db53-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-14"></a><span data-ttu-id="4db53-197">附件： Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="4db53-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

<span data-ttu-id="4db53-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-200">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="4db53-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4db53-201">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="4db53-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-202">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-202">Type</span></span>

*   <span data-ttu-id="4db53-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="4db53-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-204">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-204">Requirements</span></span>

|<span data-ttu-id="4db53-205">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-205">Requirement</span></span>| <span data-ttu-id="4db53-206">值</span><span class="sxs-lookup"><span data-stu-id="4db53-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-207">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-208">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-208">1.0</span></span>|
|[<span data-ttu-id="4db53-209">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-210">ReadItem</span></span>|
|[<span data-ttu-id="4db53-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-212">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-213">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-213">Example</span></span>

<span data-ttu-id="4db53-214">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="4db53-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="4db53-215">密件抄送：[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-216">获取一个对象，该对象提供用于获取或更新邮件的密件抄送（密件抄送）行的方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-216">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4db53-217">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-218">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-218">Type</span></span>

*   [<span data-ttu-id="4db53-219">收件人</span><span class="sxs-lookup"><span data-stu-id="4db53-219">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="4db53-220">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-220">Requirements</span></span>

|<span data-ttu-id="4db53-221">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-221">Requirement</span></span>| <span data-ttu-id="4db53-222">值</span><span class="sxs-lookup"><span data-stu-id="4db53-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-224">1.1</span><span class="sxs-lookup"><span data-stu-id="4db53-224">1.1</span></span>|
|[<span data-ttu-id="4db53-225">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-226">ReadItem</span></span>|
|[<span data-ttu-id="4db53-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-228">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-229">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-229">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-14"></a><span data-ttu-id="4db53-230">正文：[正文](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-230">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-231">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-232">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-232">Type</span></span>

*   [<span data-ttu-id="4db53-233">Body</span><span class="sxs-lookup"><span data-stu-id="4db53-233">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="4db53-234">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-234">Requirements</span></span>

|<span data-ttu-id="4db53-235">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-235">Requirement</span></span>| <span data-ttu-id="4db53-236">值</span><span class="sxs-lookup"><span data-stu-id="4db53-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-238">1.1</span><span class="sxs-lookup"><span data-stu-id="4db53-238">1.1</span></span>|
|[<span data-ttu-id="4db53-239">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-240">ReadItem</span></span>|
|[<span data-ttu-id="4db53-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-242">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-243">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-243">Example</span></span>

<span data-ttu-id="4db53-244">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="4db53-244">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4db53-245">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="4db53-245">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="4db53-246"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)的抄送： Array</span><span class="sxs-lookup"><span data-stu-id="4db53-246">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-247">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4db53-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4db53-248">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-249">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-249">Read mode</span></span>

<span data-ttu-id="4db53-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4db53-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-252">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-252">Compose mode</span></span>

<span data-ttu-id="4db53-253">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4db53-254">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-254">Type</span></span>

*   <span data-ttu-id="4db53-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-255">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-256">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-256">Requirements</span></span>

|<span data-ttu-id="4db53-257">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-257">Requirement</span></span>| <span data-ttu-id="4db53-258">值</span><span class="sxs-lookup"><span data-stu-id="4db53-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-260">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-260">1.0</span></span>|
|[<span data-ttu-id="4db53-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-262">ReadItem</span></span>|
|[<span data-ttu-id="4db53-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-264">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="4db53-265">（可以为 null） conversationId： String</span><span class="sxs-lookup"><span data-stu-id="4db53-265">(nullable) conversationId: String</span></span>

<span data-ttu-id="4db53-266">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="4db53-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4db53-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="4db53-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4db53-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="4db53-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-271">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-271">Type</span></span>

*   <span data-ttu-id="4db53-272">String</span><span class="sxs-lookup"><span data-stu-id="4db53-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-273">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-273">Requirements</span></span>

|<span data-ttu-id="4db53-274">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-274">Requirement</span></span>| <span data-ttu-id="4db53-275">值</span><span class="sxs-lookup"><span data-stu-id="4db53-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-276">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-277">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-277">1.0</span></span>|
|[<span data-ttu-id="4db53-278">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-279">ReadItem</span></span>|
|[<span data-ttu-id="4db53-280">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-281">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-282">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-282">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="4db53-283">dateTimeCreated： Date</span><span class="sxs-lookup"><span data-stu-id="4db53-283">dateTimeCreated: Date</span></span>

<span data-ttu-id="4db53-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-286">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-286">Type</span></span>

*   <span data-ttu-id="4db53-287">日期</span><span class="sxs-lookup"><span data-stu-id="4db53-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-288">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-288">Requirements</span></span>

|<span data-ttu-id="4db53-289">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-289">Requirement</span></span>| <span data-ttu-id="4db53-290">值</span><span class="sxs-lookup"><span data-stu-id="4db53-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-292">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-292">1.0</span></span>|
|[<span data-ttu-id="4db53-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-294">ReadItem</span></span>|
|[<span data-ttu-id="4db53-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-296">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-297">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-297">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="4db53-298">dateTimeModified： Date</span><span class="sxs-lookup"><span data-stu-id="4db53-298">dateTimeModified: Date</span></span>

<span data-ttu-id="4db53-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-301">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="4db53-301">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-302">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-302">Type</span></span>

*   <span data-ttu-id="4db53-303">日期</span><span class="sxs-lookup"><span data-stu-id="4db53-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-304">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-304">Requirements</span></span>

|<span data-ttu-id="4db53-305">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-305">Requirement</span></span>| <span data-ttu-id="4db53-306">值</span><span class="sxs-lookup"><span data-stu-id="4db53-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-307">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-308">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-308">1.0</span></span>|
|[<span data-ttu-id="4db53-309">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-310">ReadItem</span></span>|
|[<span data-ttu-id="4db53-311">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-312">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-313">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-313">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="4db53-314">结束：日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-314">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-315">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4db53-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4db53-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4db53-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-318">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-318">Read mode</span></span>

<span data-ttu-id="4db53-319">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-319">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-320">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-320">Compose mode</span></span>

<span data-ttu-id="4db53-321">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4db53-322">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="4db53-322">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4db53-323">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="4db53-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4db53-324">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-324">Type</span></span>

*   <span data-ttu-id="4db53-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-325">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-326">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-326">Requirements</span></span>

|<span data-ttu-id="4db53-327">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-327">Requirement</span></span>| <span data-ttu-id="4db53-328">值</span><span class="sxs-lookup"><span data-stu-id="4db53-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-329">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-330">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-330">1.0</span></span>|
|[<span data-ttu-id="4db53-331">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-332">ReadItem</span></span>|
|[<span data-ttu-id="4db53-333">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-334">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-334">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="4db53-335">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-335">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="4db53-p113">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="4db53-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-340">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="4db53-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-341">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-341">Type</span></span>

*   [<span data-ttu-id="4db53-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4db53-342">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="4db53-343">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-343">Requirements</span></span>

|<span data-ttu-id="4db53-344">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-344">Requirement</span></span>| <span data-ttu-id="4db53-345">值</span><span class="sxs-lookup"><span data-stu-id="4db53-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-346">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-347">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-347">1.0</span></span>|
|[<span data-ttu-id="4db53-348">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-349">ReadItem</span></span>|
|[<span data-ttu-id="4db53-350">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-351">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-352">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-352">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="4db53-353">internetMessageId： String</span><span class="sxs-lookup"><span data-stu-id="4db53-353">internetMessageId: String</span></span>

<span data-ttu-id="4db53-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-356">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-356">Type</span></span>

*   <span data-ttu-id="4db53-357">String</span><span class="sxs-lookup"><span data-stu-id="4db53-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-358">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-358">Requirements</span></span>

|<span data-ttu-id="4db53-359">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-359">Requirement</span></span>| <span data-ttu-id="4db53-360">值</span><span class="sxs-lookup"><span data-stu-id="4db53-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-361">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-362">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-362">1.0</span></span>|
|[<span data-ttu-id="4db53-363">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-364">ReadItem</span></span>|
|[<span data-ttu-id="4db53-365">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-366">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-367">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="4db53-368">itemClass： String</span><span class="sxs-lookup"><span data-stu-id="4db53-368">itemClass: String</span></span>

<span data-ttu-id="4db53-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4db53-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="4db53-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="4db53-373">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-373">Type</span></span> | <span data-ttu-id="4db53-374">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-374">Description</span></span> | <span data-ttu-id="4db53-375">项目类</span><span class="sxs-lookup"><span data-stu-id="4db53-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="4db53-376">约会项目</span><span class="sxs-lookup"><span data-stu-id="4db53-376">Appointment items</span></span> | <span data-ttu-id="4db53-377">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="4db53-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="4db53-378">邮件项目</span><span class="sxs-lookup"><span data-stu-id="4db53-378">Message items</span></span> | <span data-ttu-id="4db53-379">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="4db53-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="4db53-380">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="4db53-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-381">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-381">Type</span></span>

*   <span data-ttu-id="4db53-382">String</span><span class="sxs-lookup"><span data-stu-id="4db53-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-383">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-383">Requirements</span></span>

|<span data-ttu-id="4db53-384">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-384">Requirement</span></span>| <span data-ttu-id="4db53-385">值</span><span class="sxs-lookup"><span data-stu-id="4db53-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-387">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-387">1.0</span></span>|
|[<span data-ttu-id="4db53-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-389">ReadItem</span></span>|
|[<span data-ttu-id="4db53-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-391">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-392">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4db53-393">（可以为 null） itemId： String</span><span class="sxs-lookup"><span data-stu-id="4db53-393">(nullable) itemId: String</span></span>

<span data-ttu-id="4db53-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-396">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="4db53-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4db53-397">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="4db53-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4db53-398">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="4db53-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="4db53-399">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="4db53-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="4db53-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="4db53-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-402">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-402">Type</span></span>

*   <span data-ttu-id="4db53-403">String</span><span class="sxs-lookup"><span data-stu-id="4db53-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-404">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-404">Requirements</span></span>

|<span data-ttu-id="4db53-405">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-405">Requirement</span></span>| <span data-ttu-id="4db53-406">值</span><span class="sxs-lookup"><span data-stu-id="4db53-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-407">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-408">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-408">1.0</span></span>|
|[<span data-ttu-id="4db53-409">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-410">ReadItem</span></span>|
|[<span data-ttu-id="4db53-411">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-412">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-413">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-413">Example</span></span>

<span data-ttu-id="4db53-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="4db53-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-14"></a><span data-ttu-id="4db53-416">itemType： [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-416">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-417">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="4db53-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4db53-418">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="4db53-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-419">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-419">Type</span></span>

*   [<span data-ttu-id="4db53-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4db53-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="4db53-421">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-421">Requirements</span></span>

|<span data-ttu-id="4db53-422">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-422">Requirement</span></span>| <span data-ttu-id="4db53-423">值</span><span class="sxs-lookup"><span data-stu-id="4db53-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-425">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-425">1.0</span></span>|
|[<span data-ttu-id="4db53-426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-427">ReadItem</span></span>|
|[<span data-ttu-id="4db53-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-429">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-430">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-430">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-14"></a><span data-ttu-id="4db53-431">位置：字符串 |[位置](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-431">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-432">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="4db53-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-433">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-433">Read mode</span></span>

<span data-ttu-id="4db53-434">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="4db53-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-435">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-435">Compose mode</span></span>

<span data-ttu-id="4db53-436">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4db53-437">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-437">Type</span></span>

*   <span data-ttu-id="4db53-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-438">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-439">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-439">Requirements</span></span>

|<span data-ttu-id="4db53-440">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-440">Requirement</span></span>| <span data-ttu-id="4db53-441">值</span><span class="sxs-lookup"><span data-stu-id="4db53-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-442">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-443">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-443">1.0</span></span>|
|[<span data-ttu-id="4db53-444">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-445">ReadItem</span></span>|
|[<span data-ttu-id="4db53-446">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-447">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-447">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4db53-448">normalizedSubject： String</span><span class="sxs-lookup"><span data-stu-id="4db53-448">normalizedSubject: String</span></span>

<span data-ttu-id="4db53-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4db53-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="4db53-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-453">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-453">Type</span></span>

*   <span data-ttu-id="4db53-454">String</span><span class="sxs-lookup"><span data-stu-id="4db53-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-455">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-455">Requirements</span></span>

|<span data-ttu-id="4db53-456">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-456">Requirement</span></span>| <span data-ttu-id="4db53-457">值</span><span class="sxs-lookup"><span data-stu-id="4db53-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-458">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-459">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-459">1.0</span></span>|
|[<span data-ttu-id="4db53-460">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-461">ReadItem</span></span>|
|[<span data-ttu-id="4db53-462">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-463">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-464">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-464">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-14"></a><span data-ttu-id="4db53-465">notificationMessages： [notificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-465">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-466">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="4db53-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-467">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-467">Type</span></span>

*   [<span data-ttu-id="4db53-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="4db53-468">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="4db53-469">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-469">Requirements</span></span>

|<span data-ttu-id="4db53-470">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-470">Requirement</span></span>| <span data-ttu-id="4db53-471">值</span><span class="sxs-lookup"><span data-stu-id="4db53-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-472">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-473">1.3</span><span class="sxs-lookup"><span data-stu-id="4db53-473">1.3</span></span>|
|[<span data-ttu-id="4db53-474">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-475">ReadItem</span></span>|
|[<span data-ttu-id="4db53-476">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-477">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-478">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-478">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="4db53-479">optionalAttendees： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)的数组</span><span class="sxs-lookup"><span data-stu-id="4db53-479">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-480">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4db53-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4db53-481">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-482">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-482">Read mode</span></span>

<span data-ttu-id="4db53-483">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-484">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-484">Compose mode</span></span>

<span data-ttu-id="4db53-485">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4db53-486">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-486">Type</span></span>

*   <span data-ttu-id="4db53-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-487">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-488">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-488">Requirements</span></span>

|<span data-ttu-id="4db53-489">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-489">Requirement</span></span>| <span data-ttu-id="4db53-490">值</span><span class="sxs-lookup"><span data-stu-id="4db53-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-491">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-492">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-492">1.0</span></span>|
|[<span data-ttu-id="4db53-493">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-494">ReadItem</span></span>|
|[<span data-ttu-id="4db53-495">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-496">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-496">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="4db53-497">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-497">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-500">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-500">Type</span></span>

*   [<span data-ttu-id="4db53-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4db53-501">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="4db53-502">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-502">Requirements</span></span>

|<span data-ttu-id="4db53-503">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-503">Requirement</span></span>| <span data-ttu-id="4db53-504">值</span><span class="sxs-lookup"><span data-stu-id="4db53-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-505">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-506">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-506">1.0</span></span>|
|[<span data-ttu-id="4db53-507">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-508">ReadItem</span></span>|
|[<span data-ttu-id="4db53-509">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-510">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-511">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-511">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="4db53-512">requiredAttendees： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)的数组</span><span class="sxs-lookup"><span data-stu-id="4db53-512">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-513">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4db53-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4db53-514">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-515">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-515">Read mode</span></span>

<span data-ttu-id="4db53-516">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-517">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-517">Compose mode</span></span>

<span data-ttu-id="4db53-518">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4db53-519">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-519">Type</span></span>

*   <span data-ttu-id="4db53-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-520">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-521">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-521">Requirements</span></span>

|<span data-ttu-id="4db53-522">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-522">Requirement</span></span>| <span data-ttu-id="4db53-523">值</span><span class="sxs-lookup"><span data-stu-id="4db53-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-524">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-525">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-525">1.0</span></span>|
|[<span data-ttu-id="4db53-526">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-527">ReadItem</span></span>|
|[<span data-ttu-id="4db53-528">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-529">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-529">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="4db53-530">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-530">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4db53-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="4db53-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-535">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="4db53-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4db53-536">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-536">Type</span></span>

*   [<span data-ttu-id="4db53-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4db53-537">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="4db53-538">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-538">Requirements</span></span>

|<span data-ttu-id="4db53-539">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-539">Requirement</span></span>| <span data-ttu-id="4db53-540">值</span><span class="sxs-lookup"><span data-stu-id="4db53-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-541">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-542">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-542">1.0</span></span>|
|[<span data-ttu-id="4db53-543">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-544">ReadItem</span></span>|
|[<span data-ttu-id="4db53-545">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-546">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-547">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-547">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="4db53-548">开始日期：日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-548">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-549">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4db53-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4db53-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4db53-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-552">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-552">Read mode</span></span>

<span data-ttu-id="4db53-553">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-553">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-554">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-554">Compose mode</span></span>

<span data-ttu-id="4db53-555">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4db53-556">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="4db53-556">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4db53-557">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="4db53-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4db53-558">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-558">Type</span></span>

*   <span data-ttu-id="4db53-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-559">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-560">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-560">Requirements</span></span>

|<span data-ttu-id="4db53-561">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-561">Requirement</span></span>| <span data-ttu-id="4db53-562">值</span><span class="sxs-lookup"><span data-stu-id="4db53-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-563">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-564">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-564">1.0</span></span>|
|[<span data-ttu-id="4db53-565">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-566">ReadItem</span></span>|
|[<span data-ttu-id="4db53-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-568">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-568">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-14"></a><span data-ttu-id="4db53-569">subject： String |[主题](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-569">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-570">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="4db53-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4db53-571">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="4db53-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-572">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-572">Read mode</span></span>

<span data-ttu-id="4db53-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="4db53-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-575">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-575">Compose mode</span></span>

<span data-ttu-id="4db53-576">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4db53-577">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-577">Type</span></span>

*   <span data-ttu-id="4db53-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-578">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-579">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-579">Requirements</span></span>

|<span data-ttu-id="4db53-580">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-580">Requirement</span></span>| <span data-ttu-id="4db53-581">值</span><span class="sxs-lookup"><span data-stu-id="4db53-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-582">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-583">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-583">1.0</span></span>|
|[<span data-ttu-id="4db53-584">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-585">ReadItem</span></span>|
|[<span data-ttu-id="4db53-586">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-587">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-587">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="4db53-588">to： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)的数组</span><span class="sxs-lookup"><span data-stu-id="4db53-588">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="4db53-589">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4db53-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4db53-590">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4db53-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4db53-591">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4db53-591">Read mode</span></span>

<span data-ttu-id="4db53-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4db53-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4db53-594">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4db53-594">Compose mode</span></span>

<span data-ttu-id="4db53-595">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4db53-596">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-596">Type</span></span>

*   <span data-ttu-id="4db53-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-597">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-598">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-598">Requirements</span></span>

|<span data-ttu-id="4db53-599">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-599">Requirement</span></span>| <span data-ttu-id="4db53-600">值</span><span class="sxs-lookup"><span data-stu-id="4db53-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-601">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-602">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-602">1.0</span></span>|
|[<span data-ttu-id="4db53-603">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-604">ReadItem</span></span>|
|[<span data-ttu-id="4db53-605">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-606">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4db53-607">方法</span><span class="sxs-lookup"><span data-stu-id="4db53-607">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4db53-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4db53-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4db53-609">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="4db53-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4db53-610">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="4db53-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4db53-611">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="4db53-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-612">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-612">Parameters</span></span>

|<span data-ttu-id="4db53-613">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-613">Name</span></span>| <span data-ttu-id="4db53-614">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-614">Type</span></span>| <span data-ttu-id="4db53-615">属性</span><span class="sxs-lookup"><span data-stu-id="4db53-615">Attributes</span></span>| <span data-ttu-id="4db53-616">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="4db53-617">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-617">String</span></span>||<span data-ttu-id="4db53-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4db53-620">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-620">String</span></span>||<span data-ttu-id="4db53-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4db53-623">Object</span><span class="sxs-lookup"><span data-stu-id="4db53-623">Object</span></span>| <span data-ttu-id="4db53-624">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-624">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-625">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4db53-625">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4db53-626">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-626">Object</span></span>| <span data-ttu-id="4db53-627">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-627">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-628">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-628">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4db53-629">函数</span><span class="sxs-lookup"><span data-stu-id="4db53-629">function</span></span>| <span data-ttu-id="4db53-630">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-630">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-631">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4db53-632">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4db53-632">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4db53-633">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-633">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4db53-634">错误</span><span class="sxs-lookup"><span data-stu-id="4db53-634">Errors</span></span>

| <span data-ttu-id="4db53-635">错误代码</span><span class="sxs-lookup"><span data-stu-id="4db53-635">Error code</span></span> | <span data-ttu-id="4db53-636">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-636">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="4db53-637">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="4db53-637">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="4db53-638">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="4db53-638">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4db53-639">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="4db53-639">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4db53-640">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-640">Requirements</span></span>

|<span data-ttu-id="4db53-641">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-641">Requirement</span></span>| <span data-ttu-id="4db53-642">值</span><span class="sxs-lookup"><span data-stu-id="4db53-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-643">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-644">1.1</span><span class="sxs-lookup"><span data-stu-id="4db53-644">1.1</span></span>|
|[<span data-ttu-id="4db53-645">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-646">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4db53-646">ReadWriteItem</span></span>|
|[<span data-ttu-id="4db53-647">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-648">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-648">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-649">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-649">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4db53-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4db53-650">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4db53-651">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="4db53-651">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4db53-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="4db53-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4db53-655">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="4db53-655">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4db53-656">如果 Office 外接程序在 web 上的 Outlook 中运行，则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是，不支持这种情况，建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="4db53-656">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-657">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-657">Parameters</span></span>

|<span data-ttu-id="4db53-658">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-658">Name</span></span>| <span data-ttu-id="4db53-659">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-659">Type</span></span>| <span data-ttu-id="4db53-660">属性</span><span class="sxs-lookup"><span data-stu-id="4db53-660">Attributes</span></span>| <span data-ttu-id="4db53-661">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-661">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="4db53-662">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-662">String</span></span>||<span data-ttu-id="4db53-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4db53-665">String</span><span class="sxs-lookup"><span data-stu-id="4db53-665">String</span></span>||<span data-ttu-id="4db53-666">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="4db53-666">The subject of the item to be attached.</span></span> <span data-ttu-id="4db53-667">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-667">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4db53-668">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-668">Object</span></span>| <span data-ttu-id="4db53-669">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-669">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-670">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4db53-670">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4db53-671">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-671">Object</span></span>| <span data-ttu-id="4db53-672">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-672">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-673">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-673">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4db53-674">函数</span><span class="sxs-lookup"><span data-stu-id="4db53-674">function</span></span>| <span data-ttu-id="4db53-675">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-675">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-676">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-676">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4db53-677">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4db53-677">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4db53-678">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-678">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4db53-679">错误</span><span class="sxs-lookup"><span data-stu-id="4db53-679">Errors</span></span>

| <span data-ttu-id="4db53-680">错误代码</span><span class="sxs-lookup"><span data-stu-id="4db53-680">Error code</span></span> | <span data-ttu-id="4db53-681">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-681">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4db53-682">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="4db53-682">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4db53-683">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-683">Requirements</span></span>

|<span data-ttu-id="4db53-684">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-684">Requirement</span></span>| <span data-ttu-id="4db53-685">值</span><span class="sxs-lookup"><span data-stu-id="4db53-685">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-686">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-686">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-687">1.1</span><span class="sxs-lookup"><span data-stu-id="4db53-687">1.1</span></span>|
|[<span data-ttu-id="4db53-688">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-688">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-689">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4db53-689">ReadWriteItem</span></span>|
|[<span data-ttu-id="4db53-690">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-690">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-691">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-691">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-692">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-692">Example</span></span>

<span data-ttu-id="4db53-693">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="4db53-693">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="4db53-694">close()</span><span class="sxs-lookup"><span data-stu-id="4db53-694">close()</span></span>

<span data-ttu-id="4db53-695">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="4db53-695">Closes the current item that is being composed.</span></span>

<span data-ttu-id="4db53-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="4db53-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-698">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="4db53-698">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="4db53-699">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="4db53-699">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-700">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-700">Requirements</span></span>

|<span data-ttu-id="4db53-701">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-701">Requirement</span></span>| <span data-ttu-id="4db53-702">值</span><span class="sxs-lookup"><span data-stu-id="4db53-702">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-703">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-703">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-704">1.3</span><span class="sxs-lookup"><span data-stu-id="4db53-704">1.3</span></span>|
|[<span data-ttu-id="4db53-705">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-705">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-706">受限</span><span class="sxs-lookup"><span data-stu-id="4db53-706">Restricted</span></span>|
|[<span data-ttu-id="4db53-707">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-707">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-708">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-708">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4db53-709">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4db53-709">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4db53-710">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="4db53-710">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-711">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-711">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4db53-712">在 web 上的 Outlook 中，答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="4db53-712">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4db53-713">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="4db53-713">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4db53-714">如果在`formData.attachments`参数中指定了附件，则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="4db53-714">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="4db53-715">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="4db53-715">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="4db53-716">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="4db53-716">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-717">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-717">Parameters</span></span>

|<span data-ttu-id="4db53-718">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-718">Name</span></span>| <span data-ttu-id="4db53-719">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-719">Type</span></span>| <span data-ttu-id="4db53-720">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-720">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4db53-721">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="4db53-721">String &#124; Object</span></span>| |<span data-ttu-id="4db53-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4db53-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4db53-724">**或**</span><span class="sxs-lookup"><span data-stu-id="4db53-724">**OR**</span></span><br/><span data-ttu-id="4db53-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="4db53-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4db53-727">String</span><span class="sxs-lookup"><span data-stu-id="4db53-727">String</span></span> | <span data-ttu-id="4db53-728">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-728">&lt;optional&gt;</span></span> | <span data-ttu-id="4db53-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4db53-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4db53-731">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-731">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4db53-732">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-732">&lt;optional&gt;</span></span> | <span data-ttu-id="4db53-733">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="4db53-733">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4db53-734">String</span><span class="sxs-lookup"><span data-stu-id="4db53-734">String</span></span> | | <span data-ttu-id="4db53-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="4db53-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4db53-737">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-737">String</span></span> | | <span data-ttu-id="4db53-738">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-738">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4db53-739">String</span><span class="sxs-lookup"><span data-stu-id="4db53-739">String</span></span> | | <span data-ttu-id="4db53-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="4db53-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4db53-742">String</span><span class="sxs-lookup"><span data-stu-id="4db53-742">String</span></span> | | <span data-ttu-id="4db53-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4db53-746">函数</span><span class="sxs-lookup"><span data-stu-id="4db53-746">function</span></span> | <span data-ttu-id="4db53-747">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-747">&lt;optional&gt;</span></span> | <span data-ttu-id="4db53-748">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-748">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4db53-749">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-749">Requirements</span></span>

|<span data-ttu-id="4db53-750">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-750">Requirement</span></span>| <span data-ttu-id="4db53-751">值</span><span class="sxs-lookup"><span data-stu-id="4db53-751">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-752">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-752">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-753">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-753">1.0</span></span>|
|[<span data-ttu-id="4db53-754">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-754">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-755">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-755">ReadItem</span></span>|
|[<span data-ttu-id="4db53-756">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-756">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-757">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-757">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4db53-758">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-758">Examples</span></span>

<span data-ttu-id="4db53-759">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-759">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4db53-760">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-760">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4db53-761">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-761">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4db53-762">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-762">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4db53-763">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-763">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4db53-764">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-764">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4db53-765">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4db53-765">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4db53-766">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="4db53-766">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-767">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-767">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4db53-768">在 web 上的 Outlook 中，答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="4db53-768">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4db53-769">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="4db53-769">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4db53-770">如果在`formData.attachments`参数中指定了附件，则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="4db53-770">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="4db53-771">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="4db53-771">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="4db53-772">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="4db53-772">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-773">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-773">Parameters</span></span>

|<span data-ttu-id="4db53-774">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-774">Name</span></span>| <span data-ttu-id="4db53-775">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-775">Type</span></span>| <span data-ttu-id="4db53-776">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-776">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4db53-777">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="4db53-777">String &#124; Object</span></span>| | <span data-ttu-id="4db53-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4db53-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4db53-780">**或**</span><span class="sxs-lookup"><span data-stu-id="4db53-780">**OR**</span></span><br/><span data-ttu-id="4db53-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="4db53-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4db53-783">String</span><span class="sxs-lookup"><span data-stu-id="4db53-783">String</span></span> | <span data-ttu-id="4db53-784">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-784">&lt;optional&gt;</span></span> | <span data-ttu-id="4db53-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4db53-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4db53-787">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-787">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4db53-788">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-788">&lt;optional&gt;</span></span> | <span data-ttu-id="4db53-789">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="4db53-789">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4db53-790">String</span><span class="sxs-lookup"><span data-stu-id="4db53-790">String</span></span> | | <span data-ttu-id="4db53-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="4db53-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4db53-793">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-793">String</span></span> | | <span data-ttu-id="4db53-794">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-794">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4db53-795">String</span><span class="sxs-lookup"><span data-stu-id="4db53-795">String</span></span> | | <span data-ttu-id="4db53-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="4db53-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4db53-798">String</span><span class="sxs-lookup"><span data-stu-id="4db53-798">String</span></span> | | <span data-ttu-id="4db53-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4db53-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4db53-802">函数</span><span class="sxs-lookup"><span data-stu-id="4db53-802">function</span></span> | <span data-ttu-id="4db53-803">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-803">&lt;optional&gt;</span></span> | <span data-ttu-id="4db53-804">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-804">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4db53-805">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-805">Requirements</span></span>

|<span data-ttu-id="4db53-806">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-806">Requirement</span></span>| <span data-ttu-id="4db53-807">值</span><span class="sxs-lookup"><span data-stu-id="4db53-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-808">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-809">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-809">1.0</span></span>|
|[<span data-ttu-id="4db53-810">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-811">ReadItem</span></span>|
|[<span data-ttu-id="4db53-812">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-813">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-813">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4db53-814">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-814">Examples</span></span>

<span data-ttu-id="4db53-815">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-815">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4db53-816">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-816">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4db53-817">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-817">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4db53-818">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-818">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4db53-819">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-819">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4db53-820">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="4db53-820">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-14"></a><span data-ttu-id="4db53-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="4db53-821">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="4db53-822">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="4db53-822">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-823">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-823">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-824">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-824">Requirements</span></span>

|<span data-ttu-id="4db53-825">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-825">Requirement</span></span>| <span data-ttu-id="4db53-826">值</span><span class="sxs-lookup"><span data-stu-id="4db53-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-827">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-828">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-828">1.0</span></span>|
|[<span data-ttu-id="4db53-829">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-830">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-830">ReadItem</span></span>|
|[<span data-ttu-id="4db53-831">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-832">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-832">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4db53-833">返回：</span><span class="sxs-lookup"><span data-stu-id="4db53-833">Returns:</span></span>

<span data-ttu-id="4db53-834">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="4db53-834">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span></span>

##### <a name="example"></a><span data-ttu-id="4db53-835">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-835">Example</span></span>

<span data-ttu-id="4db53-836">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="4db53-836">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="4db53-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="4db53-837">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="4db53-838">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="4db53-838">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-839">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-839">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-840">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-840">Parameters</span></span>

|<span data-ttu-id="4db53-841">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-841">Name</span></span>| <span data-ttu-id="4db53-842">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-842">Type</span></span>| <span data-ttu-id="4db53-843">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-843">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="4db53-844">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4db53-844">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.4)|<span data-ttu-id="4db53-845">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="4db53-845">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4db53-846">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-846">Requirements</span></span>

|<span data-ttu-id="4db53-847">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-847">Requirement</span></span>| <span data-ttu-id="4db53-848">值</span><span class="sxs-lookup"><span data-stu-id="4db53-848">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-849">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-849">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-850">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-850">1.0</span></span>|
|[<span data-ttu-id="4db53-851">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-851">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-852">受限</span><span class="sxs-lookup"><span data-stu-id="4db53-852">Restricted</span></span>|
|[<span data-ttu-id="4db53-853">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-853">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-854">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-854">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4db53-855">返回：</span><span class="sxs-lookup"><span data-stu-id="4db53-855">Returns:</span></span>

<span data-ttu-id="4db53-856">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="4db53-856">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4db53-857">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="4db53-857">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4db53-858">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="4db53-858">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4db53-859">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="4db53-859">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="4db53-860">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="4db53-860">Value of `entityType`</span></span> | <span data-ttu-id="4db53-861">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="4db53-861">Type of objects in returned array</span></span> | <span data-ttu-id="4db53-862">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-862">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="4db53-863">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-863">String</span></span> | <span data-ttu-id="4db53-864">**受限**</span><span class="sxs-lookup"><span data-stu-id="4db53-864">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="4db53-865">Contact</span><span class="sxs-lookup"><span data-stu-id="4db53-865">Contact</span></span> | <span data-ttu-id="4db53-866">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4db53-866">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="4db53-867">String</span><span class="sxs-lookup"><span data-stu-id="4db53-867">String</span></span> | <span data-ttu-id="4db53-868">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4db53-868">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="4db53-869">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4db53-869">MeetingSuggestion</span></span> | <span data-ttu-id="4db53-870">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4db53-870">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="4db53-871">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4db53-871">PhoneNumber</span></span> | <span data-ttu-id="4db53-872">**受限**</span><span class="sxs-lookup"><span data-stu-id="4db53-872">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="4db53-873">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4db53-873">TaskSuggestion</span></span> | <span data-ttu-id="4db53-874">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4db53-874">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="4db53-875">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-875">String</span></span> | <span data-ttu-id="4db53-876">**受限**</span><span class="sxs-lookup"><span data-stu-id="4db53-876">**Restricted**</span></span> |

<span data-ttu-id="4db53-877">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="4db53-877">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

##### <a name="example"></a><span data-ttu-id="4db53-878">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-878">Example</span></span>

<span data-ttu-id="4db53-879">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="4db53-879">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="4db53-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="4db53-880">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="4db53-881">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="4db53-881">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-882">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-882">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4db53-883">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="4db53-883">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-884">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-884">Parameters</span></span>

|<span data-ttu-id="4db53-885">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-885">Name</span></span>| <span data-ttu-id="4db53-886">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-886">Type</span></span>| <span data-ttu-id="4db53-887">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-887">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4db53-888">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-888">String</span></span>|<span data-ttu-id="4db53-889">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="4db53-889">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4db53-890">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-890">Requirements</span></span>

|<span data-ttu-id="4db53-891">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-891">Requirement</span></span>| <span data-ttu-id="4db53-892">值</span><span class="sxs-lookup"><span data-stu-id="4db53-892">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-893">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-893">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-894">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-894">1.0</span></span>|
|[<span data-ttu-id="4db53-895">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-895">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-896">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-896">ReadItem</span></span>|
|[<span data-ttu-id="4db53-897">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-897">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-898">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-898">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4db53-899">返回：</span><span class="sxs-lookup"><span data-stu-id="4db53-899">Returns:</span></span>

<span data-ttu-id="4db53-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="4db53-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4db53-902">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="4db53-902">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="4db53-903">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4db53-903">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4db53-904">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="4db53-904">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-905">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-905">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4db53-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="4db53-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4db53-909">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="4db53-909">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4db53-910">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="4db53-910">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="4db53-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="4db53-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4db53-914">Requirements</span><span class="sxs-lookup"><span data-stu-id="4db53-914">Requirements</span></span>

|<span data-ttu-id="4db53-915">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-915">Requirement</span></span>| <span data-ttu-id="4db53-916">值</span><span class="sxs-lookup"><span data-stu-id="4db53-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-917">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-918">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-918">1.0</span></span>|
|[<span data-ttu-id="4db53-919">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-919">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-920">ReadItem</span></span>|
|[<span data-ttu-id="4db53-921">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-921">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-922">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-922">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4db53-923">返回：</span><span class="sxs-lookup"><span data-stu-id="4db53-923">Returns:</span></span>

<span data-ttu-id="4db53-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="4db53-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="4db53-926">类型：对象</span><span class="sxs-lookup"><span data-stu-id="4db53-926">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="4db53-927">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-927">Example</span></span>

<span data-ttu-id="4db53-928">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="4db53-928">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4db53-929">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="4db53-929">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4db53-930">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="4db53-930">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-931">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4db53-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4db53-932">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="4db53-932">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4db53-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="4db53-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-935">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-935">Parameters</span></span>

|<span data-ttu-id="4db53-936">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-936">Name</span></span>| <span data-ttu-id="4db53-937">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-937">Type</span></span>| <span data-ttu-id="4db53-938">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-938">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4db53-939">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-939">String</span></span>|<span data-ttu-id="4db53-940">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="4db53-940">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4db53-941">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-941">Requirements</span></span>

|<span data-ttu-id="4db53-942">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-942">Requirement</span></span>| <span data-ttu-id="4db53-943">值</span><span class="sxs-lookup"><span data-stu-id="4db53-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-944">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-945">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-945">1.0</span></span>|
|[<span data-ttu-id="4db53-946">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-947">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-947">ReadItem</span></span>|
|[<span data-ttu-id="4db53-948">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-949">阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-949">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4db53-950">返回：</span><span class="sxs-lookup"><span data-stu-id="4db53-950">Returns:</span></span>

<span data-ttu-id="4db53-951">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="4db53-951">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="4db53-952">类型： Array. < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="4db53-952">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="4db53-953">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-953">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4db53-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4db53-954">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4db53-955">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="4db53-955">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4db53-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="4db53-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-958">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-958">Parameters</span></span>

|<span data-ttu-id="4db53-959">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-959">Name</span></span>| <span data-ttu-id="4db53-960">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-960">Type</span></span>| <span data-ttu-id="4db53-961">属性</span><span class="sxs-lookup"><span data-stu-id="4db53-961">Attributes</span></span>| <span data-ttu-id="4db53-962">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-962">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="4db53-963">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4db53-963">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4db53-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="4db53-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="4db53-967">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-967">Object</span></span>| <span data-ttu-id="4db53-968">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-968">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-969">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4db53-969">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4db53-970">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-970">Object</span></span>| <span data-ttu-id="4db53-971">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-971">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-972">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-972">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4db53-973">function</span><span class="sxs-lookup"><span data-stu-id="4db53-973">function</span></span>||<span data-ttu-id="4db53-974">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-974">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4db53-975">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="4db53-975">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4db53-976">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="4db53-976">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4db53-977">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-977">Requirements</span></span>

|<span data-ttu-id="4db53-978">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-978">Requirement</span></span>| <span data-ttu-id="4db53-979">值</span><span class="sxs-lookup"><span data-stu-id="4db53-979">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-980">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-980">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-981">1.2</span><span class="sxs-lookup"><span data-stu-id="4db53-981">1.2</span></span>|
|[<span data-ttu-id="4db53-982">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-982">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-983">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-983">ReadItem</span></span>|
|[<span data-ttu-id="4db53-984">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-984">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-985">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-985">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4db53-986">返回：</span><span class="sxs-lookup"><span data-stu-id="4db53-986">Returns:</span></span>

<span data-ttu-id="4db53-987">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="4db53-987">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="4db53-988">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-988">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4db53-989">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-989">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4db53-990">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4db53-990">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4db53-991">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="4db53-991">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4db53-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="4db53-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-995">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-995">Parameters</span></span>

|<span data-ttu-id="4db53-996">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-996">Name</span></span>| <span data-ttu-id="4db53-997">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-997">Type</span></span>| <span data-ttu-id="4db53-998">属性</span><span class="sxs-lookup"><span data-stu-id="4db53-998">Attributes</span></span>| <span data-ttu-id="4db53-999">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-999">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4db53-1000">函数</span><span class="sxs-lookup"><span data-stu-id="4db53-1000">function</span></span>||<span data-ttu-id="4db53-1001">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-1001">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4db53-1002">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="4db53-1002">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4db53-1003">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="4db53-1003">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="4db53-1004">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-1004">Object</span></span>| <span data-ttu-id="4db53-1005">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1005">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1006">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-1006">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4db53-1007">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="4db53-1007">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4db53-1008">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-1008">Requirements</span></span>

|<span data-ttu-id="4db53-1009">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-1009">Requirement</span></span>| <span data-ttu-id="4db53-1010">值</span><span class="sxs-lookup"><span data-stu-id="4db53-1010">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-1011">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-1011">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-1012">1.0</span><span class="sxs-lookup"><span data-stu-id="4db53-1012">1.0</span></span>|
|[<span data-ttu-id="4db53-1013">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-1013">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-1014">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4db53-1014">ReadItem</span></span>|
|[<span data-ttu-id="4db53-1015">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-1015">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-1016">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4db53-1016">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-1017">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-1017">Example</span></span>

<span data-ttu-id="4db53-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="4db53-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4db53-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4db53-1021">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4db53-1022">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="4db53-1022">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4db53-1023">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="4db53-1023">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="4db53-1024">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="4db53-1024">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="4db53-1025">在 web 和移动设备上的 Outlook 中，附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="4db53-1025">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="4db53-1026">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="4db53-1026">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-1027">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-1027">Parameters</span></span>

|<span data-ttu-id="4db53-1028">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-1028">Name</span></span>| <span data-ttu-id="4db53-1029">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-1029">Type</span></span>| <span data-ttu-id="4db53-1030">属性</span><span class="sxs-lookup"><span data-stu-id="4db53-1030">Attributes</span></span>| <span data-ttu-id="4db53-1031">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-1031">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="4db53-1032">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-1032">String</span></span>||<span data-ttu-id="4db53-1033">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="4db53-1033">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="4db53-1034">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-1034">Object</span></span>| <span data-ttu-id="4db53-1035">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1036">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4db53-1036">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4db53-1037">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-1037">Object</span></span>| <span data-ttu-id="4db53-1038">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1038">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1039">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-1039">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4db53-1040">函数</span><span class="sxs-lookup"><span data-stu-id="4db53-1040">function</span></span>| <span data-ttu-id="4db53-1041">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1041">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1042">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-1042">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4db53-1043">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="4db53-1043">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4db53-1044">错误</span><span class="sxs-lookup"><span data-stu-id="4db53-1044">Errors</span></span>

| <span data-ttu-id="4db53-1045">错误代码</span><span class="sxs-lookup"><span data-stu-id="4db53-1045">Error code</span></span> | <span data-ttu-id="4db53-1046">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-1046">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="4db53-1047">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="4db53-1047">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4db53-1048">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-1048">Requirements</span></span>

|<span data-ttu-id="4db53-1049">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-1049">Requirement</span></span>| <span data-ttu-id="4db53-1050">值</span><span class="sxs-lookup"><span data-stu-id="4db53-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-1051">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-1052">1.1</span><span class="sxs-lookup"><span data-stu-id="4db53-1052">1.1</span></span>|
|[<span data-ttu-id="4db53-1053">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-1053">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-1054">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4db53-1054">ReadWriteItem</span></span>|
|[<span data-ttu-id="4db53-1055">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-1055">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-1056">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-1056">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-1057">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-1057">Example</span></span>

<span data-ttu-id="4db53-1058">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="4db53-1058">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="4db53-1059">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4db53-1059">saveAsync([options], callback)</span></span>

<span data-ttu-id="4db53-1060">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="4db53-1060">Asynchronously saves an item.</span></span>

<span data-ttu-id="4db53-1061">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="4db53-1061">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="4db53-1062">在 Outlook 网页或 Outlook 的联机模式中，将项目保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="4db53-1062">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="4db53-1063">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="4db53-1063">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-1064">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="4db53-1064">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="4db53-1065">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="4db53-1065">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="4db53-p168">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="4db53-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="4db53-1069">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="4db53-1069">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="4db53-1070">Mac 上的 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="4db53-1070">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="4db53-1071">在`saveAsync`撰写模式下从会议中调用时，此方法将失败。</span><span class="sxs-lookup"><span data-stu-id="4db53-1071">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="4db53-1072">若要解决此问题，请参阅[使用 OFFICE JS API 将会议保存为 Outlook For Mac 中的草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="4db53-1072">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="4db53-1073">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="4db53-1073">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-1074">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-1074">Parameters</span></span>

|<span data-ttu-id="4db53-1075">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-1075">Name</span></span>| <span data-ttu-id="4db53-1076">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-1076">Type</span></span>| <span data-ttu-id="4db53-1077">属性</span><span class="sxs-lookup"><span data-stu-id="4db53-1077">Attributes</span></span>| <span data-ttu-id="4db53-1078">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-1078">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="4db53-1079">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-1079">Object</span></span>| <span data-ttu-id="4db53-1080">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1080">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1081">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4db53-1081">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4db53-1082">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-1082">Object</span></span>| <span data-ttu-id="4db53-1083">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1084">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-1084">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="4db53-1085">函数</span><span class="sxs-lookup"><span data-stu-id="4db53-1085">function</span></span>||<span data-ttu-id="4db53-1086">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-1086">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4db53-1087">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4db53-1087">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4db53-1088">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-1088">Requirements</span></span>

|<span data-ttu-id="4db53-1089">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-1089">Requirement</span></span>| <span data-ttu-id="4db53-1090">值</span><span class="sxs-lookup"><span data-stu-id="4db53-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-1091">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-1092">1.3</span><span class="sxs-lookup"><span data-stu-id="4db53-1092">1.3</span></span>|
|[<span data-ttu-id="4db53-1093">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-1093">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4db53-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="4db53-1095">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-1095">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-1096">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-1096">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="4db53-1097">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-1097">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="4db53-p170">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="4db53-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4db53-1100">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4db53-1100">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4db53-1101">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="4db53-1101">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4db53-p171">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="4db53-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4db53-1105">参数</span><span class="sxs-lookup"><span data-stu-id="4db53-1105">Parameters</span></span>

|<span data-ttu-id="4db53-1106">名称</span><span class="sxs-lookup"><span data-stu-id="4db53-1106">Name</span></span>| <span data-ttu-id="4db53-1107">类型</span><span class="sxs-lookup"><span data-stu-id="4db53-1107">Type</span></span>| <span data-ttu-id="4db53-1108">属性</span><span class="sxs-lookup"><span data-stu-id="4db53-1108">Attributes</span></span>| <span data-ttu-id="4db53-1109">说明</span><span class="sxs-lookup"><span data-stu-id="4db53-1109">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4db53-1110">字符串</span><span class="sxs-lookup"><span data-stu-id="4db53-1110">String</span></span>||<span data-ttu-id="4db53-p172">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="4db53-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="4db53-1114">Object</span><span class="sxs-lookup"><span data-stu-id="4db53-1114">Object</span></span>| <span data-ttu-id="4db53-1115">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1115">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1116">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4db53-1116">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4db53-1117">对象</span><span class="sxs-lookup"><span data-stu-id="4db53-1117">Object</span></span>| <span data-ttu-id="4db53-1118">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1118">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1119">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4db53-1119">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="4db53-1120">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4db53-1120">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="4db53-1121">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4db53-1121">&lt;optional&gt;</span></span>|<span data-ttu-id="4db53-1122">如果`text`为，则当前样式应用于 web 上的 Outlook 和桌面客户端。</span><span class="sxs-lookup"><span data-stu-id="4db53-1122">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="4db53-1123">如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="4db53-1123">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4db53-1124">如果`html`和字段支持 HTML （主题不），则当前样式应用于 web 上的 outlook，并且在 outlook 桌面客户端中应用了默认样式。</span><span class="sxs-lookup"><span data-stu-id="4db53-1124">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="4db53-1125">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="4db53-1125">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4db53-1126">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="4db53-1126">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="4db53-1127">function</span><span class="sxs-lookup"><span data-stu-id="4db53-1127">function</span></span>||<span data-ttu-id="4db53-1128">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4db53-1128">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4db53-1129">Requirements</span><span class="sxs-lookup"><span data-stu-id="4db53-1129">Requirements</span></span>

|<span data-ttu-id="4db53-1130">要求</span><span class="sxs-lookup"><span data-stu-id="4db53-1130">Requirement</span></span>| <span data-ttu-id="4db53-1131">值</span><span class="sxs-lookup"><span data-stu-id="4db53-1131">Value</span></span>|
|---|---|
|[<span data-ttu-id="4db53-1132">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4db53-1132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4db53-1133">1.2</span><span class="sxs-lookup"><span data-stu-id="4db53-1133">1.2</span></span>|
|[<span data-ttu-id="4db53-1134">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4db53-1134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4db53-1135">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4db53-1135">ReadWriteItem</span></span>|
|[<span data-ttu-id="4db53-1136">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4db53-1136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4db53-1137">撰写</span><span class="sxs-lookup"><span data-stu-id="4db53-1137">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4db53-1138">示例</span><span class="sxs-lookup"><span data-stu-id="4db53-1138">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
