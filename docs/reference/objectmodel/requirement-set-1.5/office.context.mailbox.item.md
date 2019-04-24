---
title: Office.context.mailbox.item - 要求集 1.5
description: ''
ms.date: 04/12/2019
localization_priority: Priority
ms.openlocfilehash: 2f9394751180296d876d8c577d68adc1b5abb692
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451827"
---
# <a name="item"></a><span data-ttu-id="fd412-102">item</span><span class="sxs-lookup"><span data-stu-id="fd412-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="fd412-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="fd412-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="fd412-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="fd412-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="fd412-106">Requirements</span></span>

|<span data-ttu-id="fd412-107">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-107">Requirement</span></span>| <span data-ttu-id="fd412-108">值</span><span class="sxs-lookup"><span data-stu-id="fd412-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-110">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-110">1.0</span></span>|
|[<span data-ttu-id="fd412-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-112">受限</span><span class="sxs-lookup"><span data-stu-id="fd412-112">Restricted</span></span>|
|[<span data-ttu-id="fd412-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fd412-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="fd412-115">Members and methods</span></span>

| <span data-ttu-id="fd412-116">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-116">Member</span></span> | <span data-ttu-id="fd412-117">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fd412-118">attachments</span><span class="sxs-lookup"><span data-stu-id="fd412-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="fd412-119">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-119">Member</span></span> |
| [<span data-ttu-id="fd412-120">bcc</span><span class="sxs-lookup"><span data-stu-id="fd412-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="fd412-121">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-121">Member</span></span> |
| [<span data-ttu-id="fd412-122">body</span><span class="sxs-lookup"><span data-stu-id="fd412-122">body</span></span>](#body-body) | <span data-ttu-id="fd412-123">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-123">Member</span></span> |
| [<span data-ttu-id="fd412-124">cc</span><span class="sxs-lookup"><span data-stu-id="fd412-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fd412-125">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-125">Member</span></span> |
| [<span data-ttu-id="fd412-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="fd412-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="fd412-127">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-127">Member</span></span> |
| [<span data-ttu-id="fd412-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="fd412-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="fd412-129">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-129">Member</span></span> |
| [<span data-ttu-id="fd412-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="fd412-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="fd412-131">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-131">Member</span></span> |
| [<span data-ttu-id="fd412-132">end</span><span class="sxs-lookup"><span data-stu-id="fd412-132">end</span></span>](#end-datetime) | <span data-ttu-id="fd412-133">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-133">Member</span></span> |
| [<span data-ttu-id="fd412-134">from</span><span class="sxs-lookup"><span data-stu-id="fd412-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="fd412-135">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-135">Member</span></span> |
| [<span data-ttu-id="fd412-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="fd412-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="fd412-137">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-137">Member</span></span> |
| [<span data-ttu-id="fd412-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="fd412-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="fd412-139">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-139">Member</span></span> |
| [<span data-ttu-id="fd412-140">itemId</span><span class="sxs-lookup"><span data-stu-id="fd412-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="fd412-141">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-141">Member</span></span> |
| [<span data-ttu-id="fd412-142">itemType</span><span class="sxs-lookup"><span data-stu-id="fd412-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="fd412-143">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-143">Member</span></span> |
| [<span data-ttu-id="fd412-144">location</span><span class="sxs-lookup"><span data-stu-id="fd412-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="fd412-145">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-145">Member</span></span> |
| [<span data-ttu-id="fd412-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="fd412-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="fd412-147">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-147">Member</span></span> |
| [<span data-ttu-id="fd412-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="fd412-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="fd412-149">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-149">Member</span></span> |
| [<span data-ttu-id="fd412-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="fd412-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fd412-151">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-151">Member</span></span> |
| [<span data-ttu-id="fd412-152">organizer</span><span class="sxs-lookup"><span data-stu-id="fd412-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="fd412-153">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-153">Member</span></span> |
| [<span data-ttu-id="fd412-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="fd412-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fd412-155">Member</span><span class="sxs-lookup"><span data-stu-id="fd412-155">Member</span></span> |
| [<span data-ttu-id="fd412-156">sender</span><span class="sxs-lookup"><span data-stu-id="fd412-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="fd412-157">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-157">Member</span></span> |
| [<span data-ttu-id="fd412-158">start</span><span class="sxs-lookup"><span data-stu-id="fd412-158">start</span></span>](#start-datetime) | <span data-ttu-id="fd412-159">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-159">Member</span></span> |
| [<span data-ttu-id="fd412-160">subject</span><span class="sxs-lookup"><span data-stu-id="fd412-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="fd412-161">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-161">Member</span></span> |
| [<span data-ttu-id="fd412-162">to</span><span class="sxs-lookup"><span data-stu-id="fd412-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="fd412-163">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-163">Member</span></span> |
| [<span data-ttu-id="fd412-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fd412-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="fd412-165">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-165">Method</span></span> |
| [<span data-ttu-id="fd412-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fd412-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="fd412-167">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-167">Method</span></span> |
| [<span data-ttu-id="fd412-168">close</span><span class="sxs-lookup"><span data-stu-id="fd412-168">close</span></span>](#close) | <span data-ttu-id="fd412-169">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-169">Method</span></span> |
| [<span data-ttu-id="fd412-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="fd412-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="fd412-171">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-171">Method</span></span> |
| [<span data-ttu-id="fd412-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="fd412-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="fd412-173">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-173">Method</span></span> |
| [<span data-ttu-id="fd412-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="fd412-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="fd412-175">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-175">Method</span></span> |
| [<span data-ttu-id="fd412-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="fd412-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="fd412-177">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-177">Method</span></span> |
| [<span data-ttu-id="fd412-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="fd412-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="fd412-179">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-179">Method</span></span> |
| [<span data-ttu-id="fd412-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="fd412-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="fd412-181">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-181">Method</span></span> |
| [<span data-ttu-id="fd412-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="fd412-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="fd412-183">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-183">Method</span></span> |
| [<span data-ttu-id="fd412-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fd412-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="fd412-185">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-185">Method</span></span> |
| [<span data-ttu-id="fd412-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="fd412-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="fd412-187">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-187">Method</span></span> |
| [<span data-ttu-id="fd412-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="fd412-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="fd412-189">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-189">Method</span></span> |
| [<span data-ttu-id="fd412-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="fd412-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="fd412-191">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-191">Method</span></span> |
| [<span data-ttu-id="fd412-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="fd412-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="fd412-193">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="fd412-194">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-194">Example</span></span>

<span data-ttu-id="fd412-195">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="fd412-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="fd412-196">成员</span><span class="sxs-lookup"><span data-stu-id="fd412-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="fd412-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fd412-197">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="fd412-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-200">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="fd412-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="fd412-201">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="fd412-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-202">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-202">Type</span></span>

*   <span data-ttu-id="fd412-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="fd412-203">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-204">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-204">Requirements</span></span>

|<span data-ttu-id="fd412-205">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-205">Requirement</span></span>| <span data-ttu-id="fd412-206">值</span><span class="sxs-lookup"><span data-stu-id="fd412-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-207">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-208">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-208">1.0</span></span>|
|[<span data-ttu-id="fd412-209">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-210">ReadItem</span></span>|
|[<span data-ttu-id="fd412-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-212">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-213">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-213">Example</span></span>

<span data-ttu-id="fd412-214">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="fd412-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fd412-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-215">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fd412-216">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-216">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="fd412-217">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-217">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-218">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-218">Type</span></span>

*   [<span data-ttu-id="fd412-219">收件人</span><span class="sxs-lookup"><span data-stu-id="fd412-219">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="fd412-220">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-220">Requirements</span></span>

|<span data-ttu-id="fd412-221">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-221">Requirement</span></span>| <span data-ttu-id="fd412-222">值</span><span class="sxs-lookup"><span data-stu-id="fd412-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-223">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-223">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-224">1.1</span><span class="sxs-lookup"><span data-stu-id="fd412-224">1.1</span></span>|
|[<span data-ttu-id="fd412-225">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-225">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-226">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-226">ReadItem</span></span>|
|[<span data-ttu-id="fd412-227">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-227">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-228">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-228">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-229">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-229">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="fd412-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="fd412-230">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="fd412-231">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-231">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-232">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-232">Type</span></span>

*   [<span data-ttu-id="fd412-233">Body</span><span class="sxs-lookup"><span data-stu-id="fd412-233">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="fd412-234">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-234">Requirements</span></span>

|<span data-ttu-id="fd412-235">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-235">Requirement</span></span>| <span data-ttu-id="fd412-236">值</span><span class="sxs-lookup"><span data-stu-id="fd412-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-238">1.1</span><span class="sxs-lookup"><span data-stu-id="fd412-238">1.1</span></span>|
|[<span data-ttu-id="fd412-239">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-240">ReadItem</span></span>|
|[<span data-ttu-id="fd412-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-242">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-242">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-243">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-243">Example</span></span>

<span data-ttu-id="fd412-244">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="fd412-244">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="fd412-245">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="fd412-245">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fd412-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fd412-247">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="fd412-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="fd412-248">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-249">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-249">Read mode</span></span>

<span data-ttu-id="fd412-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="fd412-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-252">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-252">Compose mode</span></span>

<span data-ttu-id="fd412-253">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-253">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fd412-254">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-254">Type</span></span>

*   <span data-ttu-id="fd412-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-256">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-256">Requirements</span></span>

|<span data-ttu-id="fd412-257">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-257">Requirement</span></span>| <span data-ttu-id="fd412-258">值</span><span class="sxs-lookup"><span data-stu-id="fd412-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-260">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-260">1.0</span></span>|
|[<span data-ttu-id="fd412-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-262">ReadItem</span></span>|
|[<span data-ttu-id="fd412-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-264">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="fd412-265">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="fd412-265">(nullable) conversationId :String</span></span>

<span data-ttu-id="fd412-266">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="fd412-266">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="fd412-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="fd412-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="fd412-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="fd412-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-271">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-271">Type</span></span>

*   <span data-ttu-id="fd412-272">String</span><span class="sxs-lookup"><span data-stu-id="fd412-272">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-273">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-273">Requirements</span></span>

|<span data-ttu-id="fd412-274">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-274">Requirement</span></span>| <span data-ttu-id="fd412-275">值</span><span class="sxs-lookup"><span data-stu-id="fd412-275">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-276">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-277">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-277">1.0</span></span>|
|[<span data-ttu-id="fd412-278">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-279">ReadItem</span></span>|
|[<span data-ttu-id="fd412-280">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-281">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-281">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-282">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-282">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="fd412-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="fd412-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="fd412-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-286">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-286">Type</span></span>

*   <span data-ttu-id="fd412-287">日期</span><span class="sxs-lookup"><span data-stu-id="fd412-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-288">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-288">Requirements</span></span>

|<span data-ttu-id="fd412-289">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-289">Requirement</span></span>| <span data-ttu-id="fd412-290">值</span><span class="sxs-lookup"><span data-stu-id="fd412-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-292">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-292">1.0</span></span>|
|[<span data-ttu-id="fd412-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-293">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-294">ReadItem</span></span>|
|[<span data-ttu-id="fd412-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-295">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-296">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-297">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-297">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="fd412-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="fd412-298">dateTimeModified :Date</span></span>

<span data-ttu-id="fd412-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-301">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="fd412-301">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-302">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-302">Type</span></span>

*   <span data-ttu-id="fd412-303">日期</span><span class="sxs-lookup"><span data-stu-id="fd412-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-304">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-304">Requirements</span></span>

|<span data-ttu-id="fd412-305">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-305">Requirement</span></span>| <span data-ttu-id="fd412-306">值</span><span class="sxs-lookup"><span data-stu-id="fd412-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-307">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-308">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-308">1.0</span></span>|
|[<span data-ttu-id="fd412-309">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-310">ReadItem</span></span>|
|[<span data-ttu-id="fd412-311">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-312">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-313">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-313">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="fd412-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fd412-314">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="fd412-315">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="fd412-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="fd412-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="fd412-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-318">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-318">Read mode</span></span>

<span data-ttu-id="fd412-319">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-319">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-320">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-320">Compose mode</span></span>

<span data-ttu-id="fd412-321">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="fd412-322">使用 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="fd412-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="fd412-323">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="fd412-323">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="fd412-324">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-324">Type</span></span>

*   <span data-ttu-id="fd412-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fd412-325">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-326">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-326">Requirements</span></span>

|<span data-ttu-id="fd412-327">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-327">Requirement</span></span>| <span data-ttu-id="fd412-328">值</span><span class="sxs-lookup"><span data-stu-id="fd412-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-329">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-330">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-330">1.0</span></span>|
|[<span data-ttu-id="fd412-331">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-332">ReadItem</span></span>|
|[<span data-ttu-id="fd412-333">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-334">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-334">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="fd412-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fd412-335">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="fd412-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="fd412-p113">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="fd412-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-340">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="fd412-340">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-341">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-341">Type</span></span>

*   [<span data-ttu-id="fd412-342">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fd412-342">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fd412-343">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-343">Requirements</span></span>

|<span data-ttu-id="fd412-344">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-344">Requirement</span></span>| <span data-ttu-id="fd412-345">值</span><span class="sxs-lookup"><span data-stu-id="fd412-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-346">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-347">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-347">1.0</span></span>|
|[<span data-ttu-id="fd412-348">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-349">ReadItem</span></span>|
|[<span data-ttu-id="fd412-350">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-351">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-352">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-352">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="fd412-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="fd412-353">internetMessageId :String</span></span>

<span data-ttu-id="fd412-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-356">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-356">Type</span></span>

*   <span data-ttu-id="fd412-357">String</span><span class="sxs-lookup"><span data-stu-id="fd412-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-358">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-358">Requirements</span></span>

|<span data-ttu-id="fd412-359">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-359">Requirement</span></span>| <span data-ttu-id="fd412-360">值</span><span class="sxs-lookup"><span data-stu-id="fd412-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-361">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-362">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-362">1.0</span></span>|
|[<span data-ttu-id="fd412-363">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-364">ReadItem</span></span>|
|[<span data-ttu-id="fd412-365">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-366">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-367">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-367">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="fd412-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="fd412-368">itemClass :String</span></span>

<span data-ttu-id="fd412-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="fd412-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="fd412-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="fd412-373">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-373">Type</span></span> | <span data-ttu-id="fd412-374">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-374">Description</span></span> | <span data-ttu-id="fd412-375">项目类</span><span class="sxs-lookup"><span data-stu-id="fd412-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="fd412-376">约会项目</span><span class="sxs-lookup"><span data-stu-id="fd412-376">Appointment items</span></span> | <span data-ttu-id="fd412-377">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="fd412-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="fd412-378">邮件项目</span><span class="sxs-lookup"><span data-stu-id="fd412-378">Message items</span></span> | <span data-ttu-id="fd412-379">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="fd412-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="fd412-380">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="fd412-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-381">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-381">Type</span></span>

*   <span data-ttu-id="fd412-382">String</span><span class="sxs-lookup"><span data-stu-id="fd412-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-383">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-383">Requirements</span></span>

|<span data-ttu-id="fd412-384">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-384">Requirement</span></span>| <span data-ttu-id="fd412-385">值</span><span class="sxs-lookup"><span data-stu-id="fd412-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-387">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-387">1.0</span></span>|
|[<span data-ttu-id="fd412-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-389">ReadItem</span></span>|
|[<span data-ttu-id="fd412-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-391">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-392">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-392">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="fd412-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="fd412-393">(nullable) itemId :String</span></span>

<span data-ttu-id="fd412-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-396">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="fd412-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="fd412-397">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="fd412-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="fd412-398">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="fd412-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="fd412-399">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="fd412-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="fd412-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="fd412-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-402">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-402">Type</span></span>

*   <span data-ttu-id="fd412-403">String</span><span class="sxs-lookup"><span data-stu-id="fd412-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-404">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-404">Requirements</span></span>

|<span data-ttu-id="fd412-405">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-405">Requirement</span></span>| <span data-ttu-id="fd412-406">值</span><span class="sxs-lookup"><span data-stu-id="fd412-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-407">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-408">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-408">1.0</span></span>|
|[<span data-ttu-id="fd412-409">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-409">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-410">ReadItem</span></span>|
|[<span data-ttu-id="fd412-411">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-411">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-412">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-413">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-413">Example</span></span>

<span data-ttu-id="fd412-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="fd412-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="fd412-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="fd412-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="fd412-417">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="fd412-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="fd412-418">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="fd412-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-419">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-419">Type</span></span>

*   [<span data-ttu-id="fd412-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="fd412-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="fd412-421">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-421">Requirements</span></span>

|<span data-ttu-id="fd412-422">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-422">Requirement</span></span>| <span data-ttu-id="fd412-423">值</span><span class="sxs-lookup"><span data-stu-id="fd412-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-425">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-425">1.0</span></span>|
|[<span data-ttu-id="fd412-426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-426">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-427">ReadItem</span></span>|
|[<span data-ttu-id="fd412-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-429">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-429">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-430">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-430">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="fd412-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="fd412-431">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="fd412-432">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="fd412-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-433">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-433">Read mode</span></span>

<span data-ttu-id="fd412-434">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="fd412-434">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-435">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-435">Compose mode</span></span>

<span data-ttu-id="fd412-436">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fd412-437">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-437">Type</span></span>

*   <span data-ttu-id="fd412-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="fd412-438">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-439">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-439">Requirements</span></span>

|<span data-ttu-id="fd412-440">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-440">Requirement</span></span>| <span data-ttu-id="fd412-441">值</span><span class="sxs-lookup"><span data-stu-id="fd412-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-442">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-443">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-443">1.0</span></span>|
|[<span data-ttu-id="fd412-444">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-445">ReadItem</span></span>|
|[<span data-ttu-id="fd412-446">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-447">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-447">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="fd412-448">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="fd412-448">normalizedSubject :String</span></span>

<span data-ttu-id="fd412-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="fd412-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="fd412-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-453">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-453">Type</span></span>

*   <span data-ttu-id="fd412-454">String</span><span class="sxs-lookup"><span data-stu-id="fd412-454">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-455">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-455">Requirements</span></span>

|<span data-ttu-id="fd412-456">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-456">Requirement</span></span>| <span data-ttu-id="fd412-457">值</span><span class="sxs-lookup"><span data-stu-id="fd412-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-458">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-459">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-459">1.0</span></span>|
|[<span data-ttu-id="fd412-460">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-460">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-461">ReadItem</span></span>|
|[<span data-ttu-id="fd412-462">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-462">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-463">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-463">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-464">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-464">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="fd412-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="fd412-465">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="fd412-466">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="fd412-466">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-467">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-467">Type</span></span>

*   [<span data-ttu-id="fd412-468">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="fd412-468">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="fd412-469">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-469">Requirements</span></span>

|<span data-ttu-id="fd412-470">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-470">Requirement</span></span>| <span data-ttu-id="fd412-471">值</span><span class="sxs-lookup"><span data-stu-id="fd412-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-472">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-472">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-473">1.3</span><span class="sxs-lookup"><span data-stu-id="fd412-473">1.3</span></span>|
|[<span data-ttu-id="fd412-474">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-474">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-475">ReadItem</span></span>|
|[<span data-ttu-id="fd412-476">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-476">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-477">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-477">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-478">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-478">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fd412-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fd412-480">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="fd412-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="fd412-481">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-482">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-482">Read mode</span></span>

<span data-ttu-id="fd412-483">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-484">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-484">Compose mode</span></span>

<span data-ttu-id="fd412-485">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fd412-486">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-486">Type</span></span>

*   <span data-ttu-id="fd412-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-488">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-488">Requirements</span></span>

|<span data-ttu-id="fd412-489">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-489">Requirement</span></span>| <span data-ttu-id="fd412-490">值</span><span class="sxs-lookup"><span data-stu-id="fd412-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-491">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-492">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-492">1.0</span></span>|
|[<span data-ttu-id="fd412-493">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-493">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-494">ReadItem</span></span>|
|[<span data-ttu-id="fd412-495">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-495">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-496">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-496">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="fd412-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fd412-497">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="fd412-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-500">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-500">Type</span></span>

*   [<span data-ttu-id="fd412-501">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fd412-501">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fd412-502">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-502">Requirements</span></span>

|<span data-ttu-id="fd412-503">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-503">Requirement</span></span>| <span data-ttu-id="fd412-504">值</span><span class="sxs-lookup"><span data-stu-id="fd412-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-505">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-506">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-506">1.0</span></span>|
|[<span data-ttu-id="fd412-507">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-508">ReadItem</span></span>|
|[<span data-ttu-id="fd412-509">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-510">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-510">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-511">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-511">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fd412-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-512">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fd412-513">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="fd412-513">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="fd412-514">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-514">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-515">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-515">Read mode</span></span>

<span data-ttu-id="fd412-516">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-516">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-517">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-517">Compose mode</span></span>

<span data-ttu-id="fd412-518">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-518">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="fd412-519">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-519">Type</span></span>

*   <span data-ttu-id="fd412-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-520">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-521">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-521">Requirements</span></span>

|<span data-ttu-id="fd412-522">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-522">Requirement</span></span>| <span data-ttu-id="fd412-523">值</span><span class="sxs-lookup"><span data-stu-id="fd412-523">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-524">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-525">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-525">1.0</span></span>|
|[<span data-ttu-id="fd412-526">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-526">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-527">ReadItem</span></span>|
|[<span data-ttu-id="fd412-528">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-528">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-529">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-529">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="fd412-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="fd412-530">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="fd412-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="fd412-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="fd412-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-535">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="fd412-535">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="fd412-536">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-536">Type</span></span>

*   [<span data-ttu-id="fd412-537">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="fd412-537">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="fd412-538">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-538">Requirements</span></span>

|<span data-ttu-id="fd412-539">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-539">Requirement</span></span>| <span data-ttu-id="fd412-540">值</span><span class="sxs-lookup"><span data-stu-id="fd412-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-541">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-542">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-542">1.0</span></span>|
|[<span data-ttu-id="fd412-543">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-544">ReadItem</span></span>|
|[<span data-ttu-id="fd412-545">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-546">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-547">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-547">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="fd412-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fd412-548">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="fd412-549">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="fd412-549">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="fd412-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="fd412-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-552">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-552">Read mode</span></span>

<span data-ttu-id="fd412-553">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-553">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-554">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-554">Compose mode</span></span>

<span data-ttu-id="fd412-555">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-555">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="fd412-556">使用 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="fd412-556">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="fd412-557">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="fd412-557">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="fd412-558">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-558">Type</span></span>

*   <span data-ttu-id="fd412-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="fd412-559">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-560">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-560">Requirements</span></span>

|<span data-ttu-id="fd412-561">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-561">Requirement</span></span>| <span data-ttu-id="fd412-562">值</span><span class="sxs-lookup"><span data-stu-id="fd412-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-563">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-564">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-564">1.0</span></span>|
|[<span data-ttu-id="fd412-565">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-566">ReadItem</span></span>|
|[<span data-ttu-id="fd412-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-568">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-568">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="fd412-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fd412-569">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="fd412-570">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="fd412-570">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="fd412-571">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="fd412-571">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-572">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-572">Read mode</span></span>

<span data-ttu-id="fd412-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="fd412-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-575">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-575">Compose mode</span></span>

<span data-ttu-id="fd412-576">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-576">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="fd412-577">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-577">Type</span></span>

*   <span data-ttu-id="fd412-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="fd412-578">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-579">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-579">Requirements</span></span>

|<span data-ttu-id="fd412-580">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-580">Requirement</span></span>| <span data-ttu-id="fd412-581">值</span><span class="sxs-lookup"><span data-stu-id="fd412-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-582">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-583">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-583">1.0</span></span>|
|[<span data-ttu-id="fd412-584">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-584">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-585">ReadItem</span></span>|
|[<span data-ttu-id="fd412-586">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-586">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-587">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-587">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="fd412-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-588">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="fd412-589">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="fd412-589">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="fd412-590">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="fd412-590">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="fd412-591">阅读模式</span><span class="sxs-lookup"><span data-stu-id="fd412-591">Read mode</span></span>

<span data-ttu-id="fd412-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="fd412-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="fd412-594">撰写模式</span><span class="sxs-lookup"><span data-stu-id="fd412-594">Compose mode</span></span>

<span data-ttu-id="fd412-595">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-595">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="fd412-596">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-596">Type</span></span>

*   <span data-ttu-id="fd412-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="fd412-597">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-598">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-598">Requirements</span></span>

|<span data-ttu-id="fd412-599">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-599">Requirement</span></span>| <span data-ttu-id="fd412-600">值</span><span class="sxs-lookup"><span data-stu-id="fd412-600">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-601">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-601">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-602">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-602">1.0</span></span>|
|[<span data-ttu-id="fd412-603">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-603">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-604">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-604">ReadItem</span></span>|
|[<span data-ttu-id="fd412-605">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-605">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-606">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-606">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="fd412-607">方法</span><span class="sxs-lookup"><span data-stu-id="fd412-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="fd412-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fd412-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fd412-609">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="fd412-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="fd412-610">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="fd412-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="fd412-611">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="fd412-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-612">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-612">Parameters</span></span>

|<span data-ttu-id="fd412-613">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-613">Name</span></span>| <span data-ttu-id="fd412-614">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-614">Type</span></span>| <span data-ttu-id="fd412-615">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-615">Attributes</span></span>| <span data-ttu-id="fd412-616">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="fd412-617">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-617">String</span></span>||<span data-ttu-id="fd412-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fd412-620">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-620">String</span></span>||<span data-ttu-id="fd412-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fd412-623">Object</span><span class="sxs-lookup"><span data-stu-id="fd412-623">Object</span></span>| <span data-ttu-id="fd412-624">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-624">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-625">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="fd412-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="fd412-626">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-626">Object</span></span> | <span data-ttu-id="fd412-627">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-627">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-628">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="fd412-629">布尔值</span><span class="sxs-lookup"><span data-stu-id="fd412-629">Boolean</span></span> | <span data-ttu-id="fd412-630">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-630">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-631">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="fd412-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="fd412-632">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-632">function</span></span>| <span data-ttu-id="fd412-633">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-633">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-634">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fd412-635">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="fd412-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fd412-636">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fd412-637">错误</span><span class="sxs-lookup"><span data-stu-id="fd412-637">Errors</span></span>

| <span data-ttu-id="fd412-638">错误代码</span><span class="sxs-lookup"><span data-stu-id="fd412-638">Error code</span></span> | <span data-ttu-id="fd412-639">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="fd412-640">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="fd412-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="fd412-641">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="fd412-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fd412-642">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="fd412-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd412-643">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-643">Requirements</span></span>

|<span data-ttu-id="fd412-644">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-644">Requirement</span></span>| <span data-ttu-id="fd412-645">值</span><span class="sxs-lookup"><span data-stu-id="fd412-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-646">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-647">1.1</span><span class="sxs-lookup"><span data-stu-id="fd412-647">1.1</span></span>|
|[<span data-ttu-id="fd412-648">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-648">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fd412-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="fd412-650">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-650">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-651">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fd412-652">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-652">Examples</span></span>

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

<span data-ttu-id="fd412-653">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="fd412-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="fd412-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fd412-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="fd412-655">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="fd412-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="fd412-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="fd412-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="fd412-659">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="fd412-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="fd412-660">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="fd412-660">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-661">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-661">Parameters</span></span>

|<span data-ttu-id="fd412-662">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-662">Name</span></span>| <span data-ttu-id="fd412-663">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-663">Type</span></span>| <span data-ttu-id="fd412-664">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-664">Attributes</span></span>| <span data-ttu-id="fd412-665">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="fd412-666">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-666">String</span></span>||<span data-ttu-id="fd412-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="fd412-669">String</span><span class="sxs-lookup"><span data-stu-id="fd412-669">String</span></span>||<span data-ttu-id="fd412-670">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="fd412-670">The subject of the item to be attached.</span></span> <span data-ttu-id="fd412-671">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-671">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="fd412-672">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-672">Object</span></span>| <span data-ttu-id="fd412-673">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-673">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-674">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="fd412-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fd412-675">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-675">Object</span></span>| <span data-ttu-id="fd412-676">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-676">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-677">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fd412-678">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-678">function</span></span>| <span data-ttu-id="fd412-679">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-679">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-680">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fd412-681">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="fd412-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="fd412-682">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fd412-683">错误</span><span class="sxs-lookup"><span data-stu-id="fd412-683">Errors</span></span>

| <span data-ttu-id="fd412-684">错误代码</span><span class="sxs-lookup"><span data-stu-id="fd412-684">Error code</span></span> | <span data-ttu-id="fd412-685">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="fd412-686">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="fd412-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd412-687">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-687">Requirements</span></span>

|<span data-ttu-id="fd412-688">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-688">Requirement</span></span>| <span data-ttu-id="fd412-689">值</span><span class="sxs-lookup"><span data-stu-id="fd412-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-690">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-691">1.1</span><span class="sxs-lookup"><span data-stu-id="fd412-691">1.1</span></span>|
|[<span data-ttu-id="fd412-692">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-692">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fd412-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="fd412-694">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-694">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-695">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-696">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-696">Example</span></span>

<span data-ttu-id="fd412-697">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="fd412-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="fd412-698">close()</span><span class="sxs-lookup"><span data-stu-id="fd412-698">close()</span></span>

<span data-ttu-id="fd412-699">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="fd412-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="fd412-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="fd412-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-702">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="fd412-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="fd412-703">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="fd412-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-704">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-704">Requirements</span></span>

|<span data-ttu-id="fd412-705">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-705">Requirement</span></span>| <span data-ttu-id="fd412-706">值</span><span class="sxs-lookup"><span data-stu-id="fd412-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-707">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-708">1.3</span><span class="sxs-lookup"><span data-stu-id="fd412-708">1.3</span></span>|
|[<span data-ttu-id="fd412-709">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-709">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-710">受限</span><span class="sxs-lookup"><span data-stu-id="fd412-710">Restricted</span></span>|
|[<span data-ttu-id="fd412-711">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-711">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-712">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-712">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="fd412-713">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="fd412-713">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="fd412-714">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="fd412-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-715">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-715">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd412-716">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="fd412-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fd412-717">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="fd412-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="fd412-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="fd412-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-721">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-721">Parameters</span></span>

| <span data-ttu-id="fd412-722">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-722">Name</span></span> | <span data-ttu-id="fd412-723">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-723">Type</span></span> | <span data-ttu-id="fd412-724">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-724">Attributes</span></span> | <span data-ttu-id="fd412-725">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="fd412-726">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="fd412-726">String &#124; Object</span></span>| |<span data-ttu-id="fd412-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="fd412-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fd412-729">**或**</span><span class="sxs-lookup"><span data-stu-id="fd412-729">**OR**</span></span><br/><span data-ttu-id="fd412-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="fd412-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fd412-732">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-732">String</span></span> | <span data-ttu-id="fd412-733">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-733">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="fd412-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fd412-736">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fd412-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-737">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-738">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="fd412-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fd412-739">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-739">String</span></span> | | <span data-ttu-id="fd412-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="fd412-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fd412-742">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-742">String</span></span> | | <span data-ttu-id="fd412-743">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fd412-744">String</span><span class="sxs-lookup"><span data-stu-id="fd412-744">String</span></span> | | <span data-ttu-id="fd412-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="fd412-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="fd412-747">布尔</span><span class="sxs-lookup"><span data-stu-id="fd412-747">Boolean</span></span> | | <span data-ttu-id="fd412-p144">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="fd412-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fd412-750">String</span><span class="sxs-lookup"><span data-stu-id="fd412-750">String</span></span> | | <span data-ttu-id="fd412-p145">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fd412-754">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-754">function</span></span> | <span data-ttu-id="fd412-755">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-755">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-756">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd412-757">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-757">Requirements</span></span>

|<span data-ttu-id="fd412-758">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-758">Requirement</span></span>| <span data-ttu-id="fd412-759">值</span><span class="sxs-lookup"><span data-stu-id="fd412-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-760">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-761">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-761">1.0</span></span>|
|[<span data-ttu-id="fd412-762">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-762">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-763">ReadItem</span></span>|
|[<span data-ttu-id="fd412-764">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-764">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-765">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fd412-766">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-766">Examples</span></span>

<span data-ttu-id="fd412-767">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="fd412-768">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-768">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="fd412-769">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-769">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fd412-770">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fd412-771">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fd412-772">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="fd412-773">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="fd412-773">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="fd412-774">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="fd412-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-775">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd412-776">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="fd412-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="fd412-777">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="fd412-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="fd412-p146">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="fd412-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-781">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-781">Parameters</span></span>

| <span data-ttu-id="fd412-782">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-782">Name</span></span> | <span data-ttu-id="fd412-783">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-783">Type</span></span> | <span data-ttu-id="fd412-784">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-784">Attributes</span></span> | <span data-ttu-id="fd412-785">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="fd412-786">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="fd412-786">String &#124; Object</span></span>| | <span data-ttu-id="fd412-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="fd412-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="fd412-789">**或**</span><span class="sxs-lookup"><span data-stu-id="fd412-789">**OR**</span></span><br/><span data-ttu-id="fd412-p148">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="fd412-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="fd412-792">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-792">String</span></span> | <span data-ttu-id="fd412-793">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-793">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-p149">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="fd412-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="fd412-796">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="fd412-797">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-797">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-798">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="fd412-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="fd412-799">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-799">String</span></span> | | <span data-ttu-id="fd412-p150">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="fd412-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="fd412-802">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-802">String</span></span> | | <span data-ttu-id="fd412-803">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="fd412-804">String</span><span class="sxs-lookup"><span data-stu-id="fd412-804">String</span></span> | | <span data-ttu-id="fd412-p151">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="fd412-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="fd412-807">布尔</span><span class="sxs-lookup"><span data-stu-id="fd412-807">Boolean</span></span> | | <span data-ttu-id="fd412-p152">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="fd412-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="fd412-810">String</span><span class="sxs-lookup"><span data-stu-id="fd412-810">String</span></span> | | <span data-ttu-id="fd412-p153">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="fd412-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="fd412-814">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-814">function</span></span> | <span data-ttu-id="fd412-815">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-815">&lt;optional&gt;</span></span> | <span data-ttu-id="fd412-816">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd412-817">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-817">Requirements</span></span>

|<span data-ttu-id="fd412-818">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-818">Requirement</span></span>| <span data-ttu-id="fd412-819">值</span><span class="sxs-lookup"><span data-stu-id="fd412-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-820">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-821">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-821">1.0</span></span>|
|[<span data-ttu-id="fd412-822">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-822">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-823">ReadItem</span></span>|
|[<span data-ttu-id="fd412-824">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-824">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-825">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="fd412-826">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-826">Examples</span></span>

<span data-ttu-id="fd412-827">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="fd412-828">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-828">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="fd412-829">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-829">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="fd412-830">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="fd412-831">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="fd412-832">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="fd412-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="fd412-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="fd412-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="fd412-834">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="fd412-834">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-835">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-835">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-836">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-836">Requirements</span></span>

|<span data-ttu-id="fd412-837">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-837">Requirement</span></span>| <span data-ttu-id="fd412-838">值</span><span class="sxs-lookup"><span data-stu-id="fd412-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-839">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-840">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-840">1.0</span></span>|
|[<span data-ttu-id="fd412-841">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-841">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-842">ReadItem</span></span>|
|[<span data-ttu-id="fd412-843">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-843">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-844">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd412-845">返回：</span><span class="sxs-lookup"><span data-stu-id="fd412-845">Returns:</span></span>

<span data-ttu-id="fd412-846">类型：[Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="fd412-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="fd412-847">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-847">Example</span></span>

<span data-ttu-id="fd412-848">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="fd412-848">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="fd412-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fd412-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fd412-850">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="fd412-850">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-851">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-851">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-852">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-852">Parameters</span></span>

|<span data-ttu-id="fd412-853">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-853">Name</span></span>| <span data-ttu-id="fd412-854">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-854">Type</span></span>| <span data-ttu-id="fd412-855">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="fd412-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="fd412-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="fd412-857">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="fd412-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd412-858">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-858">Requirements</span></span>

|<span data-ttu-id="fd412-859">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-859">Requirement</span></span>| <span data-ttu-id="fd412-860">值</span><span class="sxs-lookup"><span data-stu-id="fd412-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-861">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-862">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-862">1.0</span></span>|
|[<span data-ttu-id="fd412-863">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-863">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-864">受限</span><span class="sxs-lookup"><span data-stu-id="fd412-864">Restricted</span></span>|
|[<span data-ttu-id="fd412-865">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-865">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-866">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd412-867">返回：</span><span class="sxs-lookup"><span data-stu-id="fd412-867">Returns:</span></span>

<span data-ttu-id="fd412-868">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="fd412-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="fd412-869">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="fd412-869">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="fd412-870">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="fd412-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="fd412-871">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="fd412-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="fd412-872">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="fd412-872">Value of `entityType`</span></span> | <span data-ttu-id="fd412-873">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="fd412-873">Type of objects in returned array</span></span> | <span data-ttu-id="fd412-874">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="fd412-875">String</span><span class="sxs-lookup"><span data-stu-id="fd412-875">String</span></span> | <span data-ttu-id="fd412-876">**受限**</span><span class="sxs-lookup"><span data-stu-id="fd412-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="fd412-877">Contact</span><span class="sxs-lookup"><span data-stu-id="fd412-877">Contact</span></span> | <span data-ttu-id="fd412-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fd412-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="fd412-879">String</span><span class="sxs-lookup"><span data-stu-id="fd412-879">String</span></span> | <span data-ttu-id="fd412-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fd412-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="fd412-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="fd412-881">MeetingSuggestion</span></span> | <span data-ttu-id="fd412-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fd412-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="fd412-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="fd412-883">PhoneNumber</span></span> | <span data-ttu-id="fd412-884">**受限**</span><span class="sxs-lookup"><span data-stu-id="fd412-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="fd412-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="fd412-885">TaskSuggestion</span></span> | <span data-ttu-id="fd412-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="fd412-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="fd412-887">String</span><span class="sxs-lookup"><span data-stu-id="fd412-887">String</span></span> | <span data-ttu-id="fd412-888">**受限**</span><span class="sxs-lookup"><span data-stu-id="fd412-888">**Restricted**</span></span> |

<span data-ttu-id="fd412-889">类型：Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fd412-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="fd412-890">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-890">Example</span></span>

<span data-ttu-id="fd412-891">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="fd412-891">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="fd412-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="fd412-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="fd412-893">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="fd412-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-894">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-894">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd412-895">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="fd412-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-896">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-896">Parameters</span></span>

|<span data-ttu-id="fd412-897">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-897">Name</span></span>| <span data-ttu-id="fd412-898">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-898">Type</span></span>| <span data-ttu-id="fd412-899">描述</span><span class="sxs-lookup"><span data-stu-id="fd412-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fd412-900">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-900">String</span></span>|<span data-ttu-id="fd412-901">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="fd412-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd412-902">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-902">Requirements</span></span>

|<span data-ttu-id="fd412-903">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-903">Requirement</span></span>| <span data-ttu-id="fd412-904">值</span><span class="sxs-lookup"><span data-stu-id="fd412-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-905">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-906">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-906">1.0</span></span>|
|[<span data-ttu-id="fd412-907">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-907">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-908">ReadItem</span></span>|
|[<span data-ttu-id="fd412-909">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-909">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-910">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd412-911">返回：</span><span class="sxs-lookup"><span data-stu-id="fd412-911">Returns:</span></span>

<span data-ttu-id="fd412-p155">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="fd412-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="fd412-914">类型：Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="fd412-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="fd412-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="fd412-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="fd412-916">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="fd412-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-917">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-917">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd412-p156">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="fd412-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="fd412-921">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="fd412-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="fd412-922">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="fd412-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="fd412-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="fd412-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="fd412-926">Requirements</span><span class="sxs-lookup"><span data-stu-id="fd412-926">Requirements</span></span>

|<span data-ttu-id="fd412-927">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-927">Requirement</span></span>| <span data-ttu-id="fd412-928">值</span><span class="sxs-lookup"><span data-stu-id="fd412-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-929">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-930">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-930">1.0</span></span>|
|[<span data-ttu-id="fd412-931">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-931">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-932">ReadItem</span></span>|
|[<span data-ttu-id="fd412-933">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-933">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-934">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd412-935">返回：</span><span class="sxs-lookup"><span data-stu-id="fd412-935">Returns:</span></span>

<span data-ttu-id="fd412-p158">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="fd412-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="fd412-938">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="fd412-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fd412-939">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fd412-940">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-940">Example</span></span>

<span data-ttu-id="fd412-941">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="fd412-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="fd412-942">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="fd412-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="fd412-943">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="fd412-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-944">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="fd412-944">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fd412-945">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="fd412-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="fd412-p159">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="fd412-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-948">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-948">Parameters</span></span>

|<span data-ttu-id="fd412-949">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-949">Name</span></span>| <span data-ttu-id="fd412-950">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-950">Type</span></span>| <span data-ttu-id="fd412-951">描述</span><span class="sxs-lookup"><span data-stu-id="fd412-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="fd412-952">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-952">String</span></span>|<span data-ttu-id="fd412-953">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="fd412-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd412-954">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-954">Requirements</span></span>

|<span data-ttu-id="fd412-955">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-955">Requirement</span></span>| <span data-ttu-id="fd412-956">值</span><span class="sxs-lookup"><span data-stu-id="fd412-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-957">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-958">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-958">1.0</span></span>|
|[<span data-ttu-id="fd412-959">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-959">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-960">ReadItem</span></span>|
|[<span data-ttu-id="fd412-961">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-961">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-962">阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd412-963">返回：</span><span class="sxs-lookup"><span data-stu-id="fd412-963">Returns:</span></span>

<span data-ttu-id="fd412-964">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="fd412-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="fd412-965">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="fd412-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fd412-966">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="fd412-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fd412-967">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-967">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="fd412-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="fd412-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="fd412-969">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="fd412-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="fd412-p160">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="fd412-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-972">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-972">Parameters</span></span>

|<span data-ttu-id="fd412-973">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-973">Name</span></span>| <span data-ttu-id="fd412-974">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-974">Type</span></span>| <span data-ttu-id="fd412-975">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-975">Attributes</span></span>| <span data-ttu-id="fd412-976">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="fd412-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fd412-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="fd412-p161">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="fd412-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="fd412-981">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-981">Object</span></span>| <span data-ttu-id="fd412-982">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-982">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-983">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="fd412-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fd412-984">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-984">Object</span></span>| <span data-ttu-id="fd412-985">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-985">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-986">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fd412-987">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-987">function</span></span>||<span data-ttu-id="fd412-988">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fd412-989">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="fd412-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="fd412-990">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="fd412-990">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd412-991">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-991">Requirements</span></span>

|<span data-ttu-id="fd412-992">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-992">Requirement</span></span>| <span data-ttu-id="fd412-993">值</span><span class="sxs-lookup"><span data-stu-id="fd412-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-994">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-995">1.2</span><span class="sxs-lookup"><span data-stu-id="fd412-995">1.2</span></span>|
|[<span data-ttu-id="fd412-996">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-996">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fd412-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="fd412-998">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-998">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-999">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="fd412-1000">返回：</span><span class="sxs-lookup"><span data-stu-id="fd412-1000">Returns:</span></span>

<span data-ttu-id="fd412-1001">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="fd412-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="fd412-1002">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="fd412-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="fd412-1003">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="fd412-1004">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="fd412-1005">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="fd412-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="fd412-1006">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="fd412-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="fd412-p163">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="fd412-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-1010">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-1010">Parameters</span></span>

|<span data-ttu-id="fd412-1011">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-1011">Name</span></span>| <span data-ttu-id="fd412-1012">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-1012">Type</span></span>| <span data-ttu-id="fd412-1013">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-1013">Attributes</span></span>| <span data-ttu-id="fd412-1014">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="fd412-1015">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-1015">function</span></span>||<span data-ttu-id="fd412-1016">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fd412-1017">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="fd412-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="fd412-1018">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="fd412-1018">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="fd412-1019">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-1019">Object</span></span>| <span data-ttu-id="fd412-1020">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1021">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-1021">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="fd412-1022">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="fd412-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd412-1023">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-1023">Requirements</span></span>

|<span data-ttu-id="fd412-1024">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-1024">Requirement</span></span>| <span data-ttu-id="fd412-1025">值</span><span class="sxs-lookup"><span data-stu-id="fd412-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-1026">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="fd412-1027">1.0</span></span>|
|[<span data-ttu-id="fd412-1028">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-1028">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fd412-1029">ReadItem</span></span>|
|[<span data-ttu-id="fd412-1030">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-1030">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-1031">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="fd412-1031">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-1032">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-1032">Example</span></span>

<span data-ttu-id="fd412-p166">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="fd412-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="fd412-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="fd412-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="fd412-1037">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="fd412-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="fd412-p167">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="fd412-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-1042">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-1042">Parameters</span></span>

|<span data-ttu-id="fd412-1043">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-1043">Name</span></span>| <span data-ttu-id="fd412-1044">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-1044">Type</span></span>| <span data-ttu-id="fd412-1045">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-1045">Attributes</span></span>| <span data-ttu-id="fd412-1046">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="fd412-1047">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-1047">String</span></span>||<span data-ttu-id="fd412-1048">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="fd412-1048">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="fd412-1049">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-1049">Object</span></span>| <span data-ttu-id="fd412-1050">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1050">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1051">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="fd412-1051">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fd412-1052">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-1052">Object</span></span>| <span data-ttu-id="fd412-1053">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1053">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1054">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-1054">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fd412-1055">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-1055">function</span></span>| <span data-ttu-id="fd412-1056">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1056">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1057">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-1057">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="fd412-1058">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="fd412-1058">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="fd412-1059">错误</span><span class="sxs-lookup"><span data-stu-id="fd412-1059">Errors</span></span>

| <span data-ttu-id="fd412-1060">错误代码</span><span class="sxs-lookup"><span data-stu-id="fd412-1060">Error code</span></span> | <span data-ttu-id="fd412-1061">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-1061">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="fd412-1062">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="fd412-1062">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd412-1063">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-1063">Requirements</span></span>

|<span data-ttu-id="fd412-1064">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-1064">Requirement</span></span>| <span data-ttu-id="fd412-1065">值</span><span class="sxs-lookup"><span data-stu-id="fd412-1065">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-1066">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-1066">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-1067">1.1</span><span class="sxs-lookup"><span data-stu-id="fd412-1067">1.1</span></span>|
|[<span data-ttu-id="fd412-1068">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-1068">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-1069">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fd412-1069">ReadWriteItem</span></span>|
|[<span data-ttu-id="fd412-1070">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-1070">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-1071">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-1071">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-1072">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-1072">Example</span></span>

<span data-ttu-id="fd412-1073">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="fd412-1073">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="fd412-1074">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="fd412-1074">saveAsync([options], callback)</span></span>

<span data-ttu-id="fd412-1075">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="fd412-1075">Asynchronously saves an item.</span></span>

<span data-ttu-id="fd412-p168">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="fd412-p168">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-1079">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="fd412-1079">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="fd412-1080">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="fd412-1080">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="fd412-p170">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="fd412-p170">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="fd412-1084">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="fd412-1084">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="fd412-1085">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="fd412-1085">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="fd412-1086">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="fd412-1086">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="fd412-1087">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="fd412-1087">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-1088">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-1088">Parameters</span></span>

|<span data-ttu-id="fd412-1089">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-1089">Name</span></span>| <span data-ttu-id="fd412-1090">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-1090">Type</span></span>| <span data-ttu-id="fd412-1091">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-1091">Attributes</span></span>| <span data-ttu-id="fd412-1092">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-1092">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="fd412-1093">Object</span><span class="sxs-lookup"><span data-stu-id="fd412-1093">Object</span></span>| <span data-ttu-id="fd412-1094">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1094">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1095">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="fd412-1095">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fd412-1096">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-1096">Object</span></span>| <span data-ttu-id="fd412-1097">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1097">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1098">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-1098">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="fd412-1099">函数</span><span class="sxs-lookup"><span data-stu-id="fd412-1099">function</span></span>||<span data-ttu-id="fd412-1100">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-1100">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="fd412-1101">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="fd412-1101">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fd412-1102">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-1102">Requirements</span></span>

|<span data-ttu-id="fd412-1103">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-1103">Requirement</span></span>| <span data-ttu-id="fd412-1104">值</span><span class="sxs-lookup"><span data-stu-id="fd412-1104">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-1105">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-1105">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-1106">1.3</span><span class="sxs-lookup"><span data-stu-id="fd412-1106">1.3</span></span>|
|[<span data-ttu-id="fd412-1107">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-1107">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-1108">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fd412-1108">ReadWriteItem</span></span>|
|[<span data-ttu-id="fd412-1109">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-1109">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-1110">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-1110">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="fd412-1111">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-1111">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="fd412-p172">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="fd412-p172">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="fd412-1114">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="fd412-1114">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="fd412-1115">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="fd412-1115">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="fd412-p173">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="fd412-p173">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="fd412-1119">参数</span><span class="sxs-lookup"><span data-stu-id="fd412-1119">Parameters</span></span>

|<span data-ttu-id="fd412-1120">名称</span><span class="sxs-lookup"><span data-stu-id="fd412-1120">Name</span></span>| <span data-ttu-id="fd412-1121">类型</span><span class="sxs-lookup"><span data-stu-id="fd412-1121">Type</span></span>| <span data-ttu-id="fd412-1122">属性</span><span class="sxs-lookup"><span data-stu-id="fd412-1122">Attributes</span></span>| <span data-ttu-id="fd412-1123">说明</span><span class="sxs-lookup"><span data-stu-id="fd412-1123">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="fd412-1124">字符串</span><span class="sxs-lookup"><span data-stu-id="fd412-1124">String</span></span>||<span data-ttu-id="fd412-p174">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="fd412-p174">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="fd412-1128">Object</span><span class="sxs-lookup"><span data-stu-id="fd412-1128">Object</span></span>| <span data-ttu-id="fd412-1129">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1129">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1130">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="fd412-1130">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="fd412-1131">对象</span><span class="sxs-lookup"><span data-stu-id="fd412-1131">Object</span></span>| <span data-ttu-id="fd412-1132">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1132">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-1133">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="fd412-1133">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="fd412-1134">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="fd412-1134">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="fd412-1135">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="fd412-1135">&lt;optional&gt;</span></span>|<span data-ttu-id="fd412-p175">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="fd412-p175">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="fd412-p176">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="fd412-p176">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="fd412-1140">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="fd412-1140">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="fd412-1141">function</span><span class="sxs-lookup"><span data-stu-id="fd412-1141">function</span></span>||<span data-ttu-id="fd412-1142">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="fd412-1142">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="fd412-1143">Requirements</span><span class="sxs-lookup"><span data-stu-id="fd412-1143">Requirements</span></span>

|<span data-ttu-id="fd412-1144">要求</span><span class="sxs-lookup"><span data-stu-id="fd412-1144">Requirement</span></span>| <span data-ttu-id="fd412-1145">值</span><span class="sxs-lookup"><span data-stu-id="fd412-1145">Value</span></span>|
|---|---|
|[<span data-ttu-id="fd412-1146">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="fd412-1146">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fd412-1147">1.2</span><span class="sxs-lookup"><span data-stu-id="fd412-1147">1.2</span></span>|
|[<span data-ttu-id="fd412-1148">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="fd412-1148">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fd412-1149">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="fd412-1149">ReadWriteItem</span></span>|
|[<span data-ttu-id="fd412-1150">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="fd412-1150">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="fd412-1151">撰写</span><span class="sxs-lookup"><span data-stu-id="fd412-1151">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="fd412-1152">示例</span><span class="sxs-lookup"><span data-stu-id="fd412-1152">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
