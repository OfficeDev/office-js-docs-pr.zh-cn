---
title: "\"Context\"-\"邮箱\"。项目-要求集1。1"
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: d3242f2bdabf464c262fdb8e6efd8695dc7ee330
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268500"
---
# <a name="item"></a><span data-ttu-id="494ed-102">item</span><span class="sxs-lookup"><span data-stu-id="494ed-102">item</span></span>

### <span data-ttu-id="494ed-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="494ed-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="494ed-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="494ed-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="494ed-107">Requirements</span></span>

|<span data-ttu-id="494ed-108">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-108">Requirement</span></span>| <span data-ttu-id="494ed-109">值</span><span class="sxs-lookup"><span data-stu-id="494ed-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-111">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-111">1.0</span></span>|
|[<span data-ttu-id="494ed-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-113">受限</span><span class="sxs-lookup"><span data-stu-id="494ed-113">Restricted</span></span>|
|[<span data-ttu-id="494ed-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-115">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="494ed-116">成员和方法</span><span class="sxs-lookup"><span data-stu-id="494ed-116">Members and methods</span></span>

| <span data-ttu-id="494ed-117">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-117">Member</span></span> | <span data-ttu-id="494ed-118">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-118">Type</span></span> |
|--------|------|
| [<span data-ttu-id="494ed-119">attachments</span><span class="sxs-lookup"><span data-stu-id="494ed-119">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="494ed-120">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-120">Member</span></span> |
| [<span data-ttu-id="494ed-121">bcc</span><span class="sxs-lookup"><span data-stu-id="494ed-121">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="494ed-122">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-122">Member</span></span> |
| [<span data-ttu-id="494ed-123">body</span><span class="sxs-lookup"><span data-stu-id="494ed-123">body</span></span>](#body-body) | <span data-ttu-id="494ed-124">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-124">Member</span></span> |
| [<span data-ttu-id="494ed-125">cc</span><span class="sxs-lookup"><span data-stu-id="494ed-125">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="494ed-126">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-126">Member</span></span> |
| [<span data-ttu-id="494ed-127">conversationId</span><span class="sxs-lookup"><span data-stu-id="494ed-127">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="494ed-128">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-128">Member</span></span> |
| [<span data-ttu-id="494ed-129">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="494ed-129">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="494ed-130">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-130">Member</span></span> |
| [<span data-ttu-id="494ed-131">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="494ed-131">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="494ed-132">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-132">Member</span></span> |
| [<span data-ttu-id="494ed-133">end</span><span class="sxs-lookup"><span data-stu-id="494ed-133">end</span></span>](#end-datetime) | <span data-ttu-id="494ed-134">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-134">Member</span></span> |
| [<span data-ttu-id="494ed-135">from</span><span class="sxs-lookup"><span data-stu-id="494ed-135">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="494ed-136">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-136">Member</span></span> |
| [<span data-ttu-id="494ed-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="494ed-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="494ed-138">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-138">Member</span></span> |
| [<span data-ttu-id="494ed-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="494ed-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="494ed-140">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-140">Member</span></span> |
| [<span data-ttu-id="494ed-141">itemId</span><span class="sxs-lookup"><span data-stu-id="494ed-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="494ed-142">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-142">Member</span></span> |
| [<span data-ttu-id="494ed-143">itemType</span><span class="sxs-lookup"><span data-stu-id="494ed-143">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="494ed-144">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-144">Member</span></span> |
| [<span data-ttu-id="494ed-145">location</span><span class="sxs-lookup"><span data-stu-id="494ed-145">location</span></span>](#location-stringlocation) | <span data-ttu-id="494ed-146">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-146">Member</span></span> |
| [<span data-ttu-id="494ed-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="494ed-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="494ed-148">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-148">Member</span></span> |
| [<span data-ttu-id="494ed-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="494ed-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="494ed-150">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-150">Member</span></span> |
| [<span data-ttu-id="494ed-151">organizer</span><span class="sxs-lookup"><span data-stu-id="494ed-151">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="494ed-152">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-152">Member</span></span> |
| [<span data-ttu-id="494ed-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="494ed-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="494ed-154">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-154">Member</span></span> |
| [<span data-ttu-id="494ed-155">sender</span><span class="sxs-lookup"><span data-stu-id="494ed-155">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="494ed-156">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-156">Member</span></span> |
| [<span data-ttu-id="494ed-157">start</span><span class="sxs-lookup"><span data-stu-id="494ed-157">start</span></span>](#start-datetime) | <span data-ttu-id="494ed-158">Member</span><span class="sxs-lookup"><span data-stu-id="494ed-158">Member</span></span> |
| [<span data-ttu-id="494ed-159">subject</span><span class="sxs-lookup"><span data-stu-id="494ed-159">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="494ed-160">成员</span><span class="sxs-lookup"><span data-stu-id="494ed-160">Member</span></span> |
| [<span data-ttu-id="494ed-161">to</span><span class="sxs-lookup"><span data-stu-id="494ed-161">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="494ed-162">成员</span><span class="sxs-lookup"><span data-stu-id="494ed-162">Member</span></span> |
| [<span data-ttu-id="494ed-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="494ed-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="494ed-164">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-164">Method</span></span> |
| [<span data-ttu-id="494ed-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="494ed-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="494ed-166">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-166">Method</span></span> |
| [<span data-ttu-id="494ed-167">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="494ed-167">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="494ed-168">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-168">Method</span></span> |
| [<span data-ttu-id="494ed-169">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="494ed-169">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="494ed-170">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-170">Method</span></span> |
| [<span data-ttu-id="494ed-171">getEntities</span><span class="sxs-lookup"><span data-stu-id="494ed-171">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="494ed-172">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-172">Method</span></span> |
| [<span data-ttu-id="494ed-173">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="494ed-173">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="494ed-174">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-174">Method</span></span> |
| [<span data-ttu-id="494ed-175">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="494ed-175">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="494ed-176">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-176">Method</span></span> |
| [<span data-ttu-id="494ed-177">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="494ed-177">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="494ed-178">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-178">Method</span></span> |
| [<span data-ttu-id="494ed-179">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="494ed-179">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="494ed-180">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-180">Method</span></span> |
| [<span data-ttu-id="494ed-181">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="494ed-181">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="494ed-182">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-182">Method</span></span> |
| [<span data-ttu-id="494ed-183">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="494ed-183">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="494ed-184">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-184">Method</span></span> |

### <a name="example"></a><span data-ttu-id="494ed-185">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-185">Example</span></span>

<span data-ttu-id="494ed-186">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="494ed-186">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="494ed-187">成员</span><span class="sxs-lookup"><span data-stu-id="494ed-187">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-11"></a><span data-ttu-id="494ed-188">附件: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="494ed-188">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

<span data-ttu-id="494ed-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-191">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="494ed-191">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="494ed-192">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="494ed-192">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-193">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-193">Type</span></span>

*   <span data-ttu-id="494ed-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span><span class="sxs-lookup"><span data-stu-id="494ed-194">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.1)></span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-195">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-195">Requirements</span></span>

|<span data-ttu-id="494ed-196">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-196">Requirement</span></span>| <span data-ttu-id="494ed-197">值</span><span class="sxs-lookup"><span data-stu-id="494ed-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-199">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-199">1.0</span></span>|
|[<span data-ttu-id="494ed-200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-201">ReadItem</span></span>|
|[<span data-ttu-id="494ed-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-203">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-203">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-204">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-204">Example</span></span>

<span data-ttu-id="494ed-205">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="494ed-205">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="494ed-206">密件抄送:[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-206">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-207">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-207">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="494ed-208">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-208">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-209">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-209">Type</span></span>

*   [<span data-ttu-id="494ed-210">收件人</span><span class="sxs-lookup"><span data-stu-id="494ed-210">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="494ed-211">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-211">Requirements</span></span>

|<span data-ttu-id="494ed-212">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-212">Requirement</span></span>| <span data-ttu-id="494ed-213">值</span><span class="sxs-lookup"><span data-stu-id="494ed-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-215">1.1</span><span class="sxs-lookup"><span data-stu-id="494ed-215">1.1</span></span>|
|[<span data-ttu-id="494ed-216">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-217">ReadItem</span></span>|
|[<span data-ttu-id="494ed-218">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-219">撰写</span><span class="sxs-lookup"><span data-stu-id="494ed-219">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-220">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-220">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-11"></a><span data-ttu-id="494ed-221">正文:[正文](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-221">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-222">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-222">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-223">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-223">Type</span></span>

*   [<span data-ttu-id="494ed-224">Body</span><span class="sxs-lookup"><span data-stu-id="494ed-224">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="494ed-225">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-225">Requirements</span></span>

|<span data-ttu-id="494ed-226">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-226">Requirement</span></span>| <span data-ttu-id="494ed-227">值</span><span class="sxs-lookup"><span data-stu-id="494ed-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-228">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-229">1.1</span><span class="sxs-lookup"><span data-stu-id="494ed-229">1.1</span></span>|
|[<span data-ttu-id="494ed-230">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-231">ReadItem</span></span>|
|[<span data-ttu-id="494ed-232">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-233">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-234">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-234">Example</span></span>

<span data-ttu-id="494ed-235">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="494ed-235">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="494ed-236">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="494ed-236">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="494ed-237"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="494ed-237">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-238">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="494ed-238">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="494ed-239">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-239">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-240">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-240">Read mode</span></span>

<span data-ttu-id="494ed-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="494ed-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-243">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-243">Compose mode</span></span>

<span data-ttu-id="494ed-244">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-244">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="494ed-245">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-245">Type</span></span>

*   <span data-ttu-id="494ed-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-246">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-247">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-247">Requirements</span></span>

|<span data-ttu-id="494ed-248">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-248">Requirement</span></span>| <span data-ttu-id="494ed-249">值</span><span class="sxs-lookup"><span data-stu-id="494ed-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-250">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-251">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-251">1.0</span></span>|
|[<span data-ttu-id="494ed-252">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-253">ReadItem</span></span>|
|[<span data-ttu-id="494ed-254">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-255">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-255">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="494ed-256">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="494ed-256">(nullable) conversationId: String</span></span>

<span data-ttu-id="494ed-257">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="494ed-257">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="494ed-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="494ed-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="494ed-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="494ed-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-262">Type</span><span class="sxs-lookup"><span data-stu-id="494ed-262">Type</span></span>

*   <span data-ttu-id="494ed-263">String</span><span class="sxs-lookup"><span data-stu-id="494ed-263">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-264">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-264">Requirements</span></span>

|<span data-ttu-id="494ed-265">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-265">Requirement</span></span>| <span data-ttu-id="494ed-266">值</span><span class="sxs-lookup"><span data-stu-id="494ed-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-268">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-268">1.0</span></span>|
|[<span data-ttu-id="494ed-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-270">ReadItem</span></span>|
|[<span data-ttu-id="494ed-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-273">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-273">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="494ed-274">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="494ed-274">dateTimeCreated: Date</span></span>

<span data-ttu-id="494ed-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-277">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-277">Type</span></span>

*   <span data-ttu-id="494ed-278">日期</span><span class="sxs-lookup"><span data-stu-id="494ed-278">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-279">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-279">Requirements</span></span>

|<span data-ttu-id="494ed-280">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-280">Requirement</span></span>| <span data-ttu-id="494ed-281">值</span><span class="sxs-lookup"><span data-stu-id="494ed-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-282">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-283">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-283">1.0</span></span>|
|[<span data-ttu-id="494ed-284">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-285">ReadItem</span></span>|
|[<span data-ttu-id="494ed-286">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-287">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-288">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-288">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="494ed-289">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="494ed-289">dateTimeModified: Date</span></span>

<span data-ttu-id="494ed-290">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="494ed-290">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="494ed-291">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-291">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-292">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="494ed-292">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-293">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-293">Type</span></span>

*   <span data-ttu-id="494ed-294">日期</span><span class="sxs-lookup"><span data-stu-id="494ed-294">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-295">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-295">Requirements</span></span>

|<span data-ttu-id="494ed-296">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-296">Requirement</span></span>| <span data-ttu-id="494ed-297">值</span><span class="sxs-lookup"><span data-stu-id="494ed-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-299">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-299">1.0</span></span>|
|[<span data-ttu-id="494ed-300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-301">ReadItem</span></span>|
|[<span data-ttu-id="494ed-302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-303">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-303">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-304">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-304">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="494ed-305">结束: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-305">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-306">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="494ed-306">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="494ed-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="494ed-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-309">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-309">Read mode</span></span>

<span data-ttu-id="494ed-310">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-310">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-311">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-311">Compose mode</span></span>

<span data-ttu-id="494ed-312">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-312">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="494ed-313">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="494ed-313">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="494ed-314">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="494ed-314">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="494ed-315">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-315">Type</span></span>

*   <span data-ttu-id="494ed-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-316">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-317">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-317">Requirements</span></span>

|<span data-ttu-id="494ed-318">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-318">Requirement</span></span>| <span data-ttu-id="494ed-319">值</span><span class="sxs-lookup"><span data-stu-id="494ed-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-320">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-321">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-321">1.0</span></span>|
|[<span data-ttu-id="494ed-322">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-323">ReadItem</span></span>|
|[<span data-ttu-id="494ed-324">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-325">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-325">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="494ed-326">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-326">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="494ed-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="494ed-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-331">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="494ed-331">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-332">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-332">Type</span></span>

*   [<span data-ttu-id="494ed-333">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="494ed-333">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="494ed-334">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-334">Requirements</span></span>

|<span data-ttu-id="494ed-335">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-335">Requirement</span></span>| <span data-ttu-id="494ed-336">值</span><span class="sxs-lookup"><span data-stu-id="494ed-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-337">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-338">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-338">1.0</span></span>|
|[<span data-ttu-id="494ed-339">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-340">ReadItem</span></span>|
|[<span data-ttu-id="494ed-341">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-342">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-343">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-343">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="494ed-344">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="494ed-344">internetMessageId: String</span></span>

<span data-ttu-id="494ed-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-347">Type</span><span class="sxs-lookup"><span data-stu-id="494ed-347">Type</span></span>

*   <span data-ttu-id="494ed-348">String</span><span class="sxs-lookup"><span data-stu-id="494ed-348">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-349">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-349">Requirements</span></span>

|<span data-ttu-id="494ed-350">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-350">Requirement</span></span>| <span data-ttu-id="494ed-351">值</span><span class="sxs-lookup"><span data-stu-id="494ed-351">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-352">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-353">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-353">1.0</span></span>|
|[<span data-ttu-id="494ed-354">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-355">ReadItem</span></span>|
|[<span data-ttu-id="494ed-356">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-357">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-357">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-358">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-358">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="494ed-359">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="494ed-359">itemClass: String</span></span>

<span data-ttu-id="494ed-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="494ed-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="494ed-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="494ed-364">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-364">Type</span></span> | <span data-ttu-id="494ed-365">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-365">Description</span></span> | <span data-ttu-id="494ed-366">项目类</span><span class="sxs-lookup"><span data-stu-id="494ed-366">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="494ed-367">约会项目</span><span class="sxs-lookup"><span data-stu-id="494ed-367">Appointment items</span></span> | <span data-ttu-id="494ed-368">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="494ed-368">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="494ed-369">邮件项目</span><span class="sxs-lookup"><span data-stu-id="494ed-369">Message items</span></span> | <span data-ttu-id="494ed-370">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="494ed-370">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="494ed-371">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="494ed-371">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-372">Type</span><span class="sxs-lookup"><span data-stu-id="494ed-372">Type</span></span>

*   <span data-ttu-id="494ed-373">String</span><span class="sxs-lookup"><span data-stu-id="494ed-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-374">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-374">Requirements</span></span>

|<span data-ttu-id="494ed-375">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-375">Requirement</span></span>| <span data-ttu-id="494ed-376">值</span><span class="sxs-lookup"><span data-stu-id="494ed-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-377">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-378">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-378">1.0</span></span>|
|[<span data-ttu-id="494ed-379">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-380">ReadItem</span></span>|
|[<span data-ttu-id="494ed-381">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-382">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-383">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-383">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="494ed-384">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="494ed-384">(nullable) itemId: String</span></span>

<span data-ttu-id="494ed-385">获取当前项目的 Exchange Web 服务项目标识符。</span><span class="sxs-lookup"><span data-stu-id="494ed-385">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="494ed-386">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-386">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-387">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="494ed-387">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="494ed-388">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="494ed-388">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="494ed-389">在使用此值进行 REST API 调用之前, 应使用`Office.context.mailbox.convertToRestId`转换它, 这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="494ed-389">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="494ed-390">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="494ed-390">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-391">Type</span><span class="sxs-lookup"><span data-stu-id="494ed-391">Type</span></span>

*   <span data-ttu-id="494ed-392">String</span><span class="sxs-lookup"><span data-stu-id="494ed-392">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-393">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-393">Requirements</span></span>

|<span data-ttu-id="494ed-394">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-394">Requirement</span></span>| <span data-ttu-id="494ed-395">值</span><span class="sxs-lookup"><span data-stu-id="494ed-395">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-396">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-397">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-397">1.0</span></span>|
|[<span data-ttu-id="494ed-398">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-398">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-399">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-399">ReadItem</span></span>|
|[<span data-ttu-id="494ed-400">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-400">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-401">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-401">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-402">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-402">Example</span></span>

<span data-ttu-id="494ed-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="494ed-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-11"></a><span data-ttu-id="494ed-405">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-405">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-406">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="494ed-406">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="494ed-407">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="494ed-407">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-408">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-408">Type</span></span>

*   [<span data-ttu-id="494ed-409">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="494ed-409">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="494ed-410">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-410">Requirements</span></span>

|<span data-ttu-id="494ed-411">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-411">Requirement</span></span>| <span data-ttu-id="494ed-412">值</span><span class="sxs-lookup"><span data-stu-id="494ed-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-414">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-414">1.0</span></span>|
|[<span data-ttu-id="494ed-415">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-415">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-416">ReadItem</span></span>|
|[<span data-ttu-id="494ed-417">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-418">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-418">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-419">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-419">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-11"></a><span data-ttu-id="494ed-420">位置: 字符串 |[位置](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-420">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-421">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="494ed-421">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-422">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-422">Read mode</span></span>

<span data-ttu-id="494ed-423">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="494ed-423">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-424">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-424">Compose mode</span></span>

<span data-ttu-id="494ed-425">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-425">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="494ed-426">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-426">Type</span></span>

*   <span data-ttu-id="494ed-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-427">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-428">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-428">Requirements</span></span>

|<span data-ttu-id="494ed-429">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-429">Requirement</span></span>| <span data-ttu-id="494ed-430">值</span><span class="sxs-lookup"><span data-stu-id="494ed-430">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-431">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-431">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-432">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-432">1.0</span></span>|
|[<span data-ttu-id="494ed-433">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-433">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-434">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-434">ReadItem</span></span>|
|[<span data-ttu-id="494ed-435">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-435">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-436">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-436">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="494ed-437">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="494ed-437">normalizedSubject: String</span></span>

<span data-ttu-id="494ed-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="494ed-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="494ed-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-442">Type</span><span class="sxs-lookup"><span data-stu-id="494ed-442">Type</span></span>

*   <span data-ttu-id="494ed-443">String</span><span class="sxs-lookup"><span data-stu-id="494ed-443">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-444">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-444">Requirements</span></span>

|<span data-ttu-id="494ed-445">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-445">Requirement</span></span>| <span data-ttu-id="494ed-446">值</span><span class="sxs-lookup"><span data-stu-id="494ed-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-448">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-448">1.0</span></span>|
|[<span data-ttu-id="494ed-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-450">ReadItem</span></span>|
|[<span data-ttu-id="494ed-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-452">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-453">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-453">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="494ed-454">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的数组</span><span class="sxs-lookup"><span data-stu-id="494ed-454">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-455">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="494ed-455">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="494ed-456">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-456">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-457">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-457">Read mode</span></span>

<span data-ttu-id="494ed-458">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-458">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-459">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-459">Compose mode</span></span>

<span data-ttu-id="494ed-460">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-460">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="494ed-461">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-461">Type</span></span>

*   <span data-ttu-id="494ed-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-462">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-463">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-463">Requirements</span></span>

|<span data-ttu-id="494ed-464">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-464">Requirement</span></span>| <span data-ttu-id="494ed-465">值</span><span class="sxs-lookup"><span data-stu-id="494ed-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-466">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-467">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-467">1.0</span></span>|
|[<span data-ttu-id="494ed-468">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-469">ReadItem</span></span>|
|[<span data-ttu-id="494ed-470">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-471">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-471">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="494ed-472">组织者: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-472">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-475">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-475">Type</span></span>

*   [<span data-ttu-id="494ed-476">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="494ed-476">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="494ed-477">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-477">Requirements</span></span>

|<span data-ttu-id="494ed-478">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-478">Requirement</span></span>| <span data-ttu-id="494ed-479">值</span><span class="sxs-lookup"><span data-stu-id="494ed-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-480">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-481">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-481">1.0</span></span>|
|[<span data-ttu-id="494ed-482">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-482">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-483">ReadItem</span></span>|
|[<span data-ttu-id="494ed-484">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-484">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-485">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-485">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-486">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-486">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="494ed-487">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的数组</span><span class="sxs-lookup"><span data-stu-id="494ed-487">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-488">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="494ed-488">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="494ed-489">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-489">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-490">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-490">Read mode</span></span>

<span data-ttu-id="494ed-491">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-491">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-492">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-492">Compose mode</span></span>

<span data-ttu-id="494ed-493">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-493">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="494ed-494">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-494">Type</span></span>

*   <span data-ttu-id="494ed-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-495">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-496">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-496">Requirements</span></span>

|<span data-ttu-id="494ed-497">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-497">Requirement</span></span>| <span data-ttu-id="494ed-498">值</span><span class="sxs-lookup"><span data-stu-id="494ed-498">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-499">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-499">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-500">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-500">1.0</span></span>|
|[<span data-ttu-id="494ed-501">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-501">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-502">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-502">ReadItem</span></span>|
|[<span data-ttu-id="494ed-503">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-503">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-504">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-504">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11"></a><span data-ttu-id="494ed-505">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-505">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="494ed-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="494ed-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-510">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="494ed-510">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="494ed-511">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-511">Type</span></span>

*   [<span data-ttu-id="494ed-512">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="494ed-512">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="494ed-513">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-513">Requirements</span></span>

|<span data-ttu-id="494ed-514">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-514">Requirement</span></span>| <span data-ttu-id="494ed-515">值</span><span class="sxs-lookup"><span data-stu-id="494ed-515">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-516">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-516">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-517">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-517">1.0</span></span>|
|[<span data-ttu-id="494ed-518">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-518">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-519">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-519">ReadItem</span></span>|
|[<span data-ttu-id="494ed-520">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-520">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-521">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-521">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-522">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-522">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-11"></a><span data-ttu-id="494ed-523">开始日期: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-523">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-524">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="494ed-524">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="494ed-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="494ed-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-527">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-527">Read mode</span></span>

<span data-ttu-id="494ed-528">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-528">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-529">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-529">Compose mode</span></span>

<span data-ttu-id="494ed-530">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-530">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="494ed-531">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="494ed-531">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="494ed-532">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="494ed-532">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.1#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="494ed-533">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-533">Type</span></span>

*   <span data-ttu-id="494ed-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-534">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-535">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-535">Requirements</span></span>

|<span data-ttu-id="494ed-536">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-536">Requirement</span></span>| <span data-ttu-id="494ed-537">值</span><span class="sxs-lookup"><span data-stu-id="494ed-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-538">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-539">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-539">1.0</span></span>|
|[<span data-ttu-id="494ed-540">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-541">ReadItem</span></span>|
|[<span data-ttu-id="494ed-542">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-543">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-543">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-11"></a><span data-ttu-id="494ed-544">subject: String |[主题](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-544">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-545">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="494ed-545">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="494ed-546">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="494ed-546">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-547">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-547">Read mode</span></span>

<span data-ttu-id="494ed-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="494ed-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-550">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-550">Compose mode</span></span>

<span data-ttu-id="494ed-551">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-551">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="494ed-552">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-552">Type</span></span>

*   <span data-ttu-id="494ed-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-553">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-554">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-554">Requirements</span></span>

|<span data-ttu-id="494ed-555">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-555">Requirement</span></span>| <span data-ttu-id="494ed-556">值</span><span class="sxs-lookup"><span data-stu-id="494ed-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-557">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-558">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-558">1.0</span></span>|
|[<span data-ttu-id="494ed-559">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-560">ReadItem</span></span>|
|[<span data-ttu-id="494ed-561">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-562">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-562">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-11recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-11"></a><span data-ttu-id="494ed-563">to: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)的数组</span><span class="sxs-lookup"><span data-stu-id="494ed-563">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

<span data-ttu-id="494ed-564">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="494ed-564">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="494ed-565">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="494ed-565">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="494ed-566">阅读模式</span><span class="sxs-lookup"><span data-stu-id="494ed-566">Read mode</span></span>

<span data-ttu-id="494ed-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="494ed-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="494ed-569">撰写模式</span><span class="sxs-lookup"><span data-stu-id="494ed-569">Compose mode</span></span>

<span data-ttu-id="494ed-570">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-570">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="494ed-571">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-571">Type</span></span>

*   <span data-ttu-id="494ed-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-572">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.1)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1)</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-573">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-573">Requirements</span></span>

|<span data-ttu-id="494ed-574">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-574">Requirement</span></span>| <span data-ttu-id="494ed-575">值</span><span class="sxs-lookup"><span data-stu-id="494ed-575">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-576">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-576">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-577">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-577">1.0</span></span>|
|[<span data-ttu-id="494ed-578">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-578">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-579">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-579">ReadItem</span></span>|
|[<span data-ttu-id="494ed-580">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-580">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-581">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-581">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="494ed-582">方法</span><span class="sxs-lookup"><span data-stu-id="494ed-582">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="494ed-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="494ed-583">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="494ed-584">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="494ed-584">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="494ed-585">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="494ed-585">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="494ed-586">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="494ed-586">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-587">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-587">Parameters</span></span>

|<span data-ttu-id="494ed-588">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-588">Name</span></span>| <span data-ttu-id="494ed-589">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-589">Type</span></span>| <span data-ttu-id="494ed-590">属性</span><span class="sxs-lookup"><span data-stu-id="494ed-590">Attributes</span></span>| <span data-ttu-id="494ed-591">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-591">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="494ed-592">String</span><span class="sxs-lookup"><span data-stu-id="494ed-592">String</span></span>||<span data-ttu-id="494ed-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="494ed-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="494ed-595">字符串</span><span class="sxs-lookup"><span data-stu-id="494ed-595">String</span></span>||<span data-ttu-id="494ed-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="494ed-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="494ed-598">Object</span><span class="sxs-lookup"><span data-stu-id="494ed-598">Object</span></span>| <span data-ttu-id="494ed-599">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-599">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-600">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="494ed-600">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="494ed-601">对象</span><span class="sxs-lookup"><span data-stu-id="494ed-601">Object</span></span>| <span data-ttu-id="494ed-602">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-602">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-603">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-603">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="494ed-604">函数</span><span class="sxs-lookup"><span data-stu-id="494ed-604">function</span></span>| <span data-ttu-id="494ed-605">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-605">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-606">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="494ed-607">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="494ed-607">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="494ed-608">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-608">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="494ed-609">错误</span><span class="sxs-lookup"><span data-stu-id="494ed-609">Errors</span></span>

| <span data-ttu-id="494ed-610">错误代码</span><span class="sxs-lookup"><span data-stu-id="494ed-610">Error code</span></span> | <span data-ttu-id="494ed-611">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-611">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="494ed-612">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="494ed-612">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="494ed-613">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="494ed-613">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="494ed-614">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="494ed-614">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="494ed-615">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-615">Requirements</span></span>

|<span data-ttu-id="494ed-616">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-616">Requirement</span></span>| <span data-ttu-id="494ed-617">值</span><span class="sxs-lookup"><span data-stu-id="494ed-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-618">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-619">1.1</span><span class="sxs-lookup"><span data-stu-id="494ed-619">1.1</span></span>|
|[<span data-ttu-id="494ed-620">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="494ed-621">ReadWriteItem</span></span>|
|[<span data-ttu-id="494ed-622">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-623">撰写</span><span class="sxs-lookup"><span data-stu-id="494ed-623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-624">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-624">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="494ed-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="494ed-625">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="494ed-626">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="494ed-626">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="494ed-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="494ed-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="494ed-630">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="494ed-630">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="494ed-631">如果 Office 外接程序在 web 上的 Outlook 中运行, 则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是, 不支持这种情况, 建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="494ed-631">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-632">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-632">Parameters</span></span>

|<span data-ttu-id="494ed-633">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-633">Name</span></span>| <span data-ttu-id="494ed-634">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-634">Type</span></span>| <span data-ttu-id="494ed-635">属性</span><span class="sxs-lookup"><span data-stu-id="494ed-635">Attributes</span></span>| <span data-ttu-id="494ed-636">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-636">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="494ed-637">String</span><span class="sxs-lookup"><span data-stu-id="494ed-637">String</span></span>||<span data-ttu-id="494ed-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="494ed-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="494ed-640">String</span><span class="sxs-lookup"><span data-stu-id="494ed-640">String</span></span>||<span data-ttu-id="494ed-641">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="494ed-641">The subject of the item to be attached.</span></span> <span data-ttu-id="494ed-642">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="494ed-642">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="494ed-643">对象</span><span class="sxs-lookup"><span data-stu-id="494ed-643">Object</span></span>| <span data-ttu-id="494ed-644">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-644">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-645">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="494ed-645">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="494ed-646">对象</span><span class="sxs-lookup"><span data-stu-id="494ed-646">Object</span></span>| <span data-ttu-id="494ed-647">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-647">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-648">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-648">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="494ed-649">函数</span><span class="sxs-lookup"><span data-stu-id="494ed-649">function</span></span>| <span data-ttu-id="494ed-650">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-650">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-651">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-651">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="494ed-652">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="494ed-652">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="494ed-653">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-653">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="494ed-654">错误</span><span class="sxs-lookup"><span data-stu-id="494ed-654">Errors</span></span>

| <span data-ttu-id="494ed-655">错误代码</span><span class="sxs-lookup"><span data-stu-id="494ed-655">Error code</span></span> | <span data-ttu-id="494ed-656">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-656">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="494ed-657">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="494ed-657">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="494ed-658">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-658">Requirements</span></span>

|<span data-ttu-id="494ed-659">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-659">Requirement</span></span>| <span data-ttu-id="494ed-660">值</span><span class="sxs-lookup"><span data-stu-id="494ed-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-661">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-662">1.1</span><span class="sxs-lookup"><span data-stu-id="494ed-662">1.1</span></span>|
|[<span data-ttu-id="494ed-663">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-664">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="494ed-664">ReadWriteItem</span></span>|
|[<span data-ttu-id="494ed-665">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-666">撰写</span><span class="sxs-lookup"><span data-stu-id="494ed-666">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-667">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-667">Example</span></span>

<span data-ttu-id="494ed-668">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="494ed-668">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="494ed-669">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="494ed-669">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="494ed-670">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="494ed-670">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-671">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-671">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="494ed-672">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="494ed-672">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="494ed-673">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="494ed-673">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-674">要求集1.1 中不支持在呼叫`displayReplyAllForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="494ed-674">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="494ed-675">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="494ed-675">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-676">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-676">Parameters</span></span>

|<span data-ttu-id="494ed-677">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-677">Name</span></span>| <span data-ttu-id="494ed-678">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-678">Type</span></span>| <span data-ttu-id="494ed-679">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-679">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="494ed-680">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="494ed-680">String &#124; Object</span></span>| |<span data-ttu-id="494ed-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="494ed-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="494ed-683">**或**</span><span class="sxs-lookup"><span data-stu-id="494ed-683">**OR**</span></span><br/><span data-ttu-id="494ed-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="494ed-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="494ed-686">字符串</span><span class="sxs-lookup"><span data-stu-id="494ed-686">String</span></span> | <span data-ttu-id="494ed-687">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-687">&lt;optional&gt;</span></span> | <span data-ttu-id="494ed-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="494ed-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="494ed-690">函数</span><span class="sxs-lookup"><span data-stu-id="494ed-690">function</span></span> | <span data-ttu-id="494ed-691">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-691">&lt;optional&gt;</span></span> | <span data-ttu-id="494ed-692">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-692">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="494ed-693">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-693">Requirements</span></span>

|<span data-ttu-id="494ed-694">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-694">Requirement</span></span>| <span data-ttu-id="494ed-695">值</span><span class="sxs-lookup"><span data-stu-id="494ed-695">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-696">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-696">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-697">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-697">1.0</span></span>|
|[<span data-ttu-id="494ed-698">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-698">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-699">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-699">ReadItem</span></span>|
|[<span data-ttu-id="494ed-700">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-700">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-701">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-701">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="494ed-702">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-702">Examples</span></span>

<span data-ttu-id="494ed-703">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-703">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="494ed-704">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="494ed-704">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="494ed-705">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="494ed-705">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="494ed-706">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="494ed-706">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="494ed-707">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="494ed-707">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="494ed-708">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="494ed-708">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-709">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-709">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="494ed-710">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="494ed-710">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="494ed-711">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="494ed-711">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-712">要求集1.1 中不支持在呼叫`displayReplyForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="494ed-712">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="494ed-713">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="494ed-713">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-714">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-714">Parameters</span></span>

|<span data-ttu-id="494ed-715">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-715">Name</span></span>| <span data-ttu-id="494ed-716">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-716">Type</span></span>| <span data-ttu-id="494ed-717">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-717">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="494ed-718">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="494ed-718">String &#124; Object</span></span>| | <span data-ttu-id="494ed-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="494ed-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="494ed-721">**或**</span><span class="sxs-lookup"><span data-stu-id="494ed-721">**OR**</span></span><br/><span data-ttu-id="494ed-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="494ed-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="494ed-724">字符串</span><span class="sxs-lookup"><span data-stu-id="494ed-724">String</span></span> | <span data-ttu-id="494ed-725">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-725">&lt;optional&gt;</span></span> | <span data-ttu-id="494ed-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="494ed-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="494ed-728">函数</span><span class="sxs-lookup"><span data-stu-id="494ed-728">function</span></span> | <span data-ttu-id="494ed-729">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-729">&lt;optional&gt;</span></span> | <span data-ttu-id="494ed-730">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-730">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="494ed-731">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-731">Requirements</span></span>

|<span data-ttu-id="494ed-732">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-732">Requirement</span></span>| <span data-ttu-id="494ed-733">值</span><span class="sxs-lookup"><span data-stu-id="494ed-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-734">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-735">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-735">1.0</span></span>|
|[<span data-ttu-id="494ed-736">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-737">ReadItem</span></span>|
|[<span data-ttu-id="494ed-738">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-739">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-739">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="494ed-740">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-740">Examples</span></span>

<span data-ttu-id="494ed-741">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-741">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="494ed-742">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="494ed-742">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="494ed-743">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="494ed-743">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="494ed-744">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="494ed-744">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-11"></a><span data-ttu-id="494ed-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span><span class="sxs-lookup"><span data-stu-id="494ed-745">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)}</span></span>

<span data-ttu-id="494ed-746">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="494ed-746">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-747">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-747">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-748">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-748">Requirements</span></span>

|<span data-ttu-id="494ed-749">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-749">Requirement</span></span>| <span data-ttu-id="494ed-750">值</span><span class="sxs-lookup"><span data-stu-id="494ed-750">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-751">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-751">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-752">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-752">1.0</span></span>|
|[<span data-ttu-id="494ed-753">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-753">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-754">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-754">ReadItem</span></span>|
|[<span data-ttu-id="494ed-755">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-755">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-756">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-756">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="494ed-757">返回：</span><span class="sxs-lookup"><span data-stu-id="494ed-757">Returns:</span></span>

<span data-ttu-id="494ed-758">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="494ed-758">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.1)</span></span>

##### <a name="example"></a><span data-ttu-id="494ed-759">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-759">Example</span></span>

<span data-ttu-id="494ed-760">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="494ed-760">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="494ed-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="494ed-761">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="494ed-762">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="494ed-762">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-763">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-763">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-764">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-764">Parameters</span></span>

|<span data-ttu-id="494ed-765">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-765">Name</span></span>| <span data-ttu-id="494ed-766">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-766">Type</span></span>| <span data-ttu-id="494ed-767">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-767">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="494ed-768">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="494ed-768">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.MailboxEnums.entitytype?view=outlook-js-1.1)|<span data-ttu-id="494ed-769">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="494ed-769">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="494ed-770">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-770">Requirements</span></span>

|<span data-ttu-id="494ed-771">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-771">Requirement</span></span>| <span data-ttu-id="494ed-772">值</span><span class="sxs-lookup"><span data-stu-id="494ed-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-773">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-774">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-774">1.0</span></span>|
|[<span data-ttu-id="494ed-775">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-775">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-776">受限</span><span class="sxs-lookup"><span data-stu-id="494ed-776">Restricted</span></span>|
|[<span data-ttu-id="494ed-777">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-777">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-778">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="494ed-779">返回：</span><span class="sxs-lookup"><span data-stu-id="494ed-779">Returns:</span></span>

<span data-ttu-id="494ed-780">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="494ed-780">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="494ed-781">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="494ed-781">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="494ed-782">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="494ed-782">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="494ed-783">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="494ed-783">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="494ed-784">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="494ed-784">Value of `entityType`</span></span> | <span data-ttu-id="494ed-785">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="494ed-785">Type of objects in returned array</span></span> | <span data-ttu-id="494ed-786">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-786">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="494ed-787">String</span><span class="sxs-lookup"><span data-stu-id="494ed-787">String</span></span> | <span data-ttu-id="494ed-788">**受限**</span><span class="sxs-lookup"><span data-stu-id="494ed-788">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="494ed-789">Contact</span><span class="sxs-lookup"><span data-stu-id="494ed-789">Contact</span></span> | <span data-ttu-id="494ed-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="494ed-790">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="494ed-791">String</span><span class="sxs-lookup"><span data-stu-id="494ed-791">String</span></span> | <span data-ttu-id="494ed-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="494ed-792">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="494ed-793">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="494ed-793">MeetingSuggestion</span></span> | <span data-ttu-id="494ed-794">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="494ed-794">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="494ed-795">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="494ed-795">PhoneNumber</span></span> | <span data-ttu-id="494ed-796">**受限**</span><span class="sxs-lookup"><span data-stu-id="494ed-796">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="494ed-797">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="494ed-797">TaskSuggestion</span></span> | <span data-ttu-id="494ed-798">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="494ed-798">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="494ed-799">String</span><span class="sxs-lookup"><span data-stu-id="494ed-799">String</span></span> | <span data-ttu-id="494ed-800">**受限**</span><span class="sxs-lookup"><span data-stu-id="494ed-800">**Restricted**</span></span> |

<span data-ttu-id="494ed-801">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="494ed-801">Type:  Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


##### <a name="example"></a><span data-ttu-id="494ed-802">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-802">Example</span></span>

<span data-ttu-id="494ed-803">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="494ed-803">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-11meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-11phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-11tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-11"></a><span data-ttu-id="494ed-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span><span class="sxs-lookup"><span data-stu-id="494ed-804">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))>}</span></span>

<span data-ttu-id="494ed-805">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="494ed-805">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-806">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-806">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="494ed-807">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="494ed-807">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-808">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-808">Parameters</span></span>

|<span data-ttu-id="494ed-809">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-809">Name</span></span>| <span data-ttu-id="494ed-810">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-810">Type</span></span>| <span data-ttu-id="494ed-811">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-811">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="494ed-812">字符串</span><span class="sxs-lookup"><span data-stu-id="494ed-812">String</span></span>|<span data-ttu-id="494ed-813">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="494ed-813">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="494ed-814">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-814">Requirements</span></span>

|<span data-ttu-id="494ed-815">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-815">Requirement</span></span>| <span data-ttu-id="494ed-816">值</span><span class="sxs-lookup"><span data-stu-id="494ed-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-817">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-818">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-818">1.0</span></span>|
|[<span data-ttu-id="494ed-819">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-820">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-820">ReadItem</span></span>|
|[<span data-ttu-id="494ed-821">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-822">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-822">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="494ed-823">返回：</span><span class="sxs-lookup"><span data-stu-id="494ed-823">Returns:</span></span>

<span data-ttu-id="494ed-p146">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="494ed-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="494ed-826">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span><span class="sxs-lookup"><span data-stu-id="494ed-826">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.1)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.1)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.1)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.1))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="494ed-827">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="494ed-827">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="494ed-828">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="494ed-828">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-829">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-829">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="494ed-p147">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="494ed-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="494ed-833">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="494ed-833">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="494ed-834">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="494ed-834">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="494ed-p148">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="494ed-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="494ed-837">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-837">Requirements</span></span>

|<span data-ttu-id="494ed-838">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-838">Requirement</span></span>| <span data-ttu-id="494ed-839">值</span><span class="sxs-lookup"><span data-stu-id="494ed-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-840">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-841">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-841">1.0</span></span>|
|[<span data-ttu-id="494ed-842">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-842">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-843">ReadItem</span></span>|
|[<span data-ttu-id="494ed-844">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-844">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-845">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-845">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="494ed-846">返回：</span><span class="sxs-lookup"><span data-stu-id="494ed-846">Returns:</span></span>

<span data-ttu-id="494ed-p149">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="494ed-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="494ed-849">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="494ed-849">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="494ed-850">对象</span><span class="sxs-lookup"><span data-stu-id="494ed-850">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="494ed-851">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-851">Example</span></span>

<span data-ttu-id="494ed-852">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="494ed-852">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="494ed-853">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="494ed-853">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="494ed-854">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="494ed-854">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="494ed-855">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="494ed-855">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="494ed-856">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="494ed-856">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="494ed-p150">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="494ed-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-859">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-859">Parameters</span></span>

|<span data-ttu-id="494ed-860">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-860">Name</span></span>| <span data-ttu-id="494ed-861">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-861">Type</span></span>| <span data-ttu-id="494ed-862">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-862">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="494ed-863">字符串</span><span class="sxs-lookup"><span data-stu-id="494ed-863">String</span></span>|<span data-ttu-id="494ed-864">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="494ed-864">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="494ed-865">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-865">Requirements</span></span>

|<span data-ttu-id="494ed-866">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-866">Requirement</span></span>| <span data-ttu-id="494ed-867">值</span><span class="sxs-lookup"><span data-stu-id="494ed-867">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-868">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-868">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-869">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-869">1.0</span></span>|
|[<span data-ttu-id="494ed-870">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-870">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-871">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-871">ReadItem</span></span>|
|[<span data-ttu-id="494ed-872">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-872">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-873">阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-873">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="494ed-874">返回：</span><span class="sxs-lookup"><span data-stu-id="494ed-874">Returns:</span></span>

<span data-ttu-id="494ed-875">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="494ed-875">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="494ed-876">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="494ed-876">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="494ed-877">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="494ed-877">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="494ed-878">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-878">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="494ed-879">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="494ed-879">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="494ed-880">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="494ed-880">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="494ed-p151">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="494ed-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-884">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-884">Parameters</span></span>

|<span data-ttu-id="494ed-885">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-885">Name</span></span>| <span data-ttu-id="494ed-886">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-886">Type</span></span>| <span data-ttu-id="494ed-887">属性</span><span class="sxs-lookup"><span data-stu-id="494ed-887">Attributes</span></span>| <span data-ttu-id="494ed-888">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-888">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="494ed-889">函数</span><span class="sxs-lookup"><span data-stu-id="494ed-889">function</span></span>||<span data-ttu-id="494ed-890">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-890">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="494ed-891">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="494ed-891">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.1) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="494ed-892">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="494ed-892">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="494ed-893">对象</span><span class="sxs-lookup"><span data-stu-id="494ed-893">Object</span></span>| <span data-ttu-id="494ed-894">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-894">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-895">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-895">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="494ed-896">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="494ed-896">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="494ed-897">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-897">Requirements</span></span>

|<span data-ttu-id="494ed-898">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-898">Requirement</span></span>| <span data-ttu-id="494ed-899">值</span><span class="sxs-lookup"><span data-stu-id="494ed-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-900">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-901">1.0</span><span class="sxs-lookup"><span data-stu-id="494ed-901">1.0</span></span>|
|[<span data-ttu-id="494ed-902">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-902">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-903">ReadItem</span><span class="sxs-lookup"><span data-stu-id="494ed-903">ReadItem</span></span>|
|[<span data-ttu-id="494ed-904">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-904">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-905">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="494ed-905">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-906">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-906">Example</span></span>

<span data-ttu-id="494ed-p154">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="494ed-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="494ed-910">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="494ed-910">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="494ed-911">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="494ed-911">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="494ed-912">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="494ed-912">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="494ed-913">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="494ed-913">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="494ed-914">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="494ed-914">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="494ed-915">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="494ed-915">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="494ed-916">参数</span><span class="sxs-lookup"><span data-stu-id="494ed-916">Parameters</span></span>

|<span data-ttu-id="494ed-917">名称</span><span class="sxs-lookup"><span data-stu-id="494ed-917">Name</span></span>| <span data-ttu-id="494ed-918">类型</span><span class="sxs-lookup"><span data-stu-id="494ed-918">Type</span></span>| <span data-ttu-id="494ed-919">属性</span><span class="sxs-lookup"><span data-stu-id="494ed-919">Attributes</span></span>| <span data-ttu-id="494ed-920">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-920">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="494ed-921">字符串</span><span class="sxs-lookup"><span data-stu-id="494ed-921">String</span></span>||<span data-ttu-id="494ed-922">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="494ed-922">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="494ed-923">对象</span><span class="sxs-lookup"><span data-stu-id="494ed-923">Object</span></span>| <span data-ttu-id="494ed-924">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-924">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-925">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="494ed-925">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="494ed-926">对象</span><span class="sxs-lookup"><span data-stu-id="494ed-926">Object</span></span>| <span data-ttu-id="494ed-927">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-927">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-928">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="494ed-928">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="494ed-929">函数</span><span class="sxs-lookup"><span data-stu-id="494ed-929">function</span></span>| <span data-ttu-id="494ed-930">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="494ed-930">&lt;optional&gt;</span></span>|<span data-ttu-id="494ed-931">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="494ed-931">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="494ed-932">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="494ed-932">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="494ed-933">错误</span><span class="sxs-lookup"><span data-stu-id="494ed-933">Errors</span></span>

| <span data-ttu-id="494ed-934">错误代码</span><span class="sxs-lookup"><span data-stu-id="494ed-934">Error code</span></span> | <span data-ttu-id="494ed-935">说明</span><span class="sxs-lookup"><span data-stu-id="494ed-935">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="494ed-936">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="494ed-936">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="494ed-937">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-937">Requirements</span></span>

|<span data-ttu-id="494ed-938">要求</span><span class="sxs-lookup"><span data-stu-id="494ed-938">Requirement</span></span>| <span data-ttu-id="494ed-939">值</span><span class="sxs-lookup"><span data-stu-id="494ed-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="494ed-940">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="494ed-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="494ed-941">1.1</span><span class="sxs-lookup"><span data-stu-id="494ed-941">1.1</span></span>|
|[<span data-ttu-id="494ed-942">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="494ed-942">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="494ed-943">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="494ed-943">ReadWriteItem</span></span>|
|[<span data-ttu-id="494ed-944">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="494ed-944">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="494ed-945">撰写</span><span class="sxs-lookup"><span data-stu-id="494ed-945">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="494ed-946">示例</span><span class="sxs-lookup"><span data-stu-id="494ed-946">Example</span></span>

<span data-ttu-id="494ed-947">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="494ed-947">The following code removes an attachment with an identifier of '0'.</span></span>

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
