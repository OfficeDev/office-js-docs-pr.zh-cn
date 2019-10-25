---
title: "\"Context\"-\"邮箱\"。项目-要求集1。4"
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 7981f27f41724dff01ca56f7d0109cd86e0b2c61
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682639"
---
# <a name="item"></a><span data-ttu-id="2ab52-102">item</span><span class="sxs-lookup"><span data-stu-id="2ab52-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="2ab52-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="2ab52-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="2ab52-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-106">Requirements</span></span>

|<span data-ttu-id="2ab52-107">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-107">Requirement</span></span>| <span data-ttu-id="2ab52-108">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-110">1.0</span></span>|
|[<span data-ttu-id="2ab52-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-112">受限</span><span class="sxs-lookup"><span data-stu-id="2ab52-112">Restricted</span></span>|
|[<span data-ttu-id="2ab52-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2ab52-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-115">Members and methods</span></span>

| <span data-ttu-id="2ab52-116">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-116">Member</span></span> | <span data-ttu-id="2ab52-117">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2ab52-118">attachments</span><span class="sxs-lookup"><span data-stu-id="2ab52-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="2ab52-119">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-119">Member</span></span> |
| [<span data-ttu-id="2ab52-120">bcc</span><span class="sxs-lookup"><span data-stu-id="2ab52-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="2ab52-121">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-121">Member</span></span> |
| [<span data-ttu-id="2ab52-122">body</span><span class="sxs-lookup"><span data-stu-id="2ab52-122">body</span></span>](#body-body) | <span data-ttu-id="2ab52-123">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-123">Member</span></span> |
| [<span data-ttu-id="2ab52-124">cc</span><span class="sxs-lookup"><span data-stu-id="2ab52-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2ab52-125">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-125">Member</span></span> |
| [<span data-ttu-id="2ab52-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="2ab52-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="2ab52-127">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-127">Member</span></span> |
| [<span data-ttu-id="2ab52-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="2ab52-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="2ab52-129">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-129">Member</span></span> |
| [<span data-ttu-id="2ab52-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="2ab52-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="2ab52-131">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-131">Member</span></span> |
| [<span data-ttu-id="2ab52-132">end</span><span class="sxs-lookup"><span data-stu-id="2ab52-132">end</span></span>](#end-datetime) | <span data-ttu-id="2ab52-133">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-133">Member</span></span> |
| [<span data-ttu-id="2ab52-134">from</span><span class="sxs-lookup"><span data-stu-id="2ab52-134">from</span></span>](#from-emailaddressdetails) | <span data-ttu-id="2ab52-135">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-135">Member</span></span> |
| [<span data-ttu-id="2ab52-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="2ab52-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="2ab52-137">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-137">Member</span></span> |
| [<span data-ttu-id="2ab52-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="2ab52-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="2ab52-139">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-139">Member</span></span> |
| [<span data-ttu-id="2ab52-140">itemId</span><span class="sxs-lookup"><span data-stu-id="2ab52-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="2ab52-141">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-141">Member</span></span> |
| [<span data-ttu-id="2ab52-142">itemType</span><span class="sxs-lookup"><span data-stu-id="2ab52-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="2ab52-143">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-143">Member</span></span> |
| [<span data-ttu-id="2ab52-144">location</span><span class="sxs-lookup"><span data-stu-id="2ab52-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="2ab52-145">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-145">Member</span></span> |
| [<span data-ttu-id="2ab52-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="2ab52-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="2ab52-147">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-147">Member</span></span> |
| [<span data-ttu-id="2ab52-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="2ab52-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="2ab52-149">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-149">Member</span></span> |
| [<span data-ttu-id="2ab52-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="2ab52-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2ab52-151">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-151">Member</span></span> |
| [<span data-ttu-id="2ab52-152">organizer</span><span class="sxs-lookup"><span data-stu-id="2ab52-152">organizer</span></span>](#organizer-emailaddressdetails) | <span data-ttu-id="2ab52-153">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-153">Member</span></span> |
| [<span data-ttu-id="2ab52-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="2ab52-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2ab52-155">Member</span><span class="sxs-lookup"><span data-stu-id="2ab52-155">Member</span></span> |
| [<span data-ttu-id="2ab52-156">sender</span><span class="sxs-lookup"><span data-stu-id="2ab52-156">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="2ab52-157">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-157">Member</span></span> |
| [<span data-ttu-id="2ab52-158">start</span><span class="sxs-lookup"><span data-stu-id="2ab52-158">start</span></span>](#start-datetime) | <span data-ttu-id="2ab52-159">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-159">Member</span></span> |
| [<span data-ttu-id="2ab52-160">subject</span><span class="sxs-lookup"><span data-stu-id="2ab52-160">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="2ab52-161">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-161">Member</span></span> |
| [<span data-ttu-id="2ab52-162">to</span><span class="sxs-lookup"><span data-stu-id="2ab52-162">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="2ab52-163">成员</span><span class="sxs-lookup"><span data-stu-id="2ab52-163">Member</span></span> |
| [<span data-ttu-id="2ab52-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2ab52-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="2ab52-165">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-165">Method</span></span> |
| [<span data-ttu-id="2ab52-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2ab52-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="2ab52-167">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-167">Method</span></span> |
| [<span data-ttu-id="2ab52-168">close</span><span class="sxs-lookup"><span data-stu-id="2ab52-168">close</span></span>](#close) | <span data-ttu-id="2ab52-169">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-169">Method</span></span> |
| [<span data-ttu-id="2ab52-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="2ab52-170">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="2ab52-171">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-171">Method</span></span> |
| [<span data-ttu-id="2ab52-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="2ab52-172">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="2ab52-173">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-173">Method</span></span> |
| [<span data-ttu-id="2ab52-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="2ab52-174">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="2ab52-175">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-175">Method</span></span> |
| [<span data-ttu-id="2ab52-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="2ab52-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2ab52-177">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-177">Method</span></span> |
| [<span data-ttu-id="2ab52-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="2ab52-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="2ab52-179">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-179">Method</span></span> |
| [<span data-ttu-id="2ab52-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="2ab52-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="2ab52-181">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-181">Method</span></span> |
| [<span data-ttu-id="2ab52-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="2ab52-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="2ab52-183">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-183">Method</span></span> |
| [<span data-ttu-id="2ab52-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2ab52-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="2ab52-185">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-185">Method</span></span> |
| [<span data-ttu-id="2ab52-186">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="2ab52-186">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="2ab52-187">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-187">Method</span></span> |
| [<span data-ttu-id="2ab52-188">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="2ab52-188">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="2ab52-189">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-189">Method</span></span> |
| [<span data-ttu-id="2ab52-190">saveAsync</span><span class="sxs-lookup"><span data-stu-id="2ab52-190">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="2ab52-191">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-191">Method</span></span> |
| [<span data-ttu-id="2ab52-192">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="2ab52-192">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="2ab52-193">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-193">Method</span></span> |

### <a name="example"></a><span data-ttu-id="2ab52-194">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-194">Example</span></span>

<span data-ttu-id="2ab52-195">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="2ab52-195">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="2ab52-196">Members</span><span class="sxs-lookup"><span data-stu-id="2ab52-196">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-14"></a><span data-ttu-id="2ab52-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="2ab52-197">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

<span data-ttu-id="2ab52-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-200">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="2ab52-200">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="2ab52-201">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="2ab52-201">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-202">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-202">Type</span></span>

*   <span data-ttu-id="2ab52-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span><span class="sxs-lookup"><span data-stu-id="2ab52-203">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.4)></span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-204">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-204">Requirements</span></span>

|<span data-ttu-id="2ab52-205">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-205">Requirement</span></span>| <span data-ttu-id="2ab52-206">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-207">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-208">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-208">1.0</span></span>|
|[<span data-ttu-id="2ab52-209">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-209">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-210">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-210">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-211">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-211">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-212">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-212">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-213">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-213">Example</span></span>

<span data-ttu-id="2ab52-214">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="2ab52-214">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="2ab52-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-215">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-216">获取一个对象，该对象提供用于获取或更新邮件的密件抄送（密件抄送）行的方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-216">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="2ab52-217">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-217">Compose mode only.</span></span>

<span data-ttu-id="2ab52-218">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-218">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-219">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2ab52-219">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2ab52-220">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-220">Get 500 members maximum.</span></span>
- <span data-ttu-id="2ab52-221">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-221">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-222">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-222">Type</span></span>

*   [<span data-ttu-id="2ab52-223">收件人</span><span class="sxs-lookup"><span data-stu-id="2ab52-223">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="2ab52-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-224">Requirements</span></span>

|<span data-ttu-id="2ab52-225">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-225">Requirement</span></span>| <span data-ttu-id="2ab52-226">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-228">1.1</span><span class="sxs-lookup"><span data-stu-id="2ab52-228">1.1</span></span>|
|[<span data-ttu-id="2ab52-229">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-230">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-231">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-232">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-233">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-233">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-14"></a><span data-ttu-id="2ab52-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-234">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-235">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-236">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-236">Type</span></span>

*   [<span data-ttu-id="2ab52-237">Body</span><span class="sxs-lookup"><span data-stu-id="2ab52-237">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="2ab52-238">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-238">Requirements</span></span>

|<span data-ttu-id="2ab52-239">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-239">Requirement</span></span>| <span data-ttu-id="2ab52-240">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-242">1.1</span><span class="sxs-lookup"><span data-stu-id="2ab52-242">1.1</span></span>|
|[<span data-ttu-id="2ab52-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-244">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-247">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-247">Example</span></span>

<span data-ttu-id="2ab52-248">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="2ab52-248">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="2ab52-249">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="2ab52-249">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="2ab52-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-250">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-251">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2ab52-251">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="2ab52-252">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-252">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-253">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-253">Read mode</span></span>

<span data-ttu-id="2ab52-254">`cc` 属性返回包含邮件的`EmailAddressDetails`行上所列的每个收件人的 \*\*\*\* 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-254">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="2ab52-255">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-255">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-256">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-256">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-257">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-257">Compose mode</span></span>

<span data-ttu-id="2ab52-258">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-258">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="2ab52-259">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-259">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-260">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2ab52-260">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2ab52-261">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-261">Get 500 members maximum.</span></span>
- <span data-ttu-id="2ab52-262">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-262">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2ab52-263">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-263">Type</span></span>

*   <span data-ttu-id="2ab52-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-264">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-265">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-265">Requirements</span></span>

|<span data-ttu-id="2ab52-266">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-266">Requirement</span></span>| <span data-ttu-id="2ab52-267">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-269">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-269">1.0</span></span>|
|[<span data-ttu-id="2ab52-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-271">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-273">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="2ab52-274">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="2ab52-274">(nullable) conversationId: String</span></span>

<span data-ttu-id="2ab52-275">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="2ab52-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="2ab52-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-280">Type</span><span class="sxs-lookup"><span data-stu-id="2ab52-280">Type</span></span>

*   <span data-ttu-id="2ab52-281">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-282">Requirements</span></span>

|<span data-ttu-id="2ab52-283">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-283">Requirement</span></span>| <span data-ttu-id="2ab52-284">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-285">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-286">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-286">1.0</span></span>|
|[<span data-ttu-id="2ab52-287">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-287">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-288">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-289">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-289">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-290">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-291">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-291">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="2ab52-292">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="2ab52-292">dateTimeCreated: Date</span></span>

<span data-ttu-id="2ab52-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-295">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-295">Type</span></span>

*   <span data-ttu-id="2ab52-296">日期</span><span class="sxs-lookup"><span data-stu-id="2ab52-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-297">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-297">Requirements</span></span>

|<span data-ttu-id="2ab52-298">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-298">Requirement</span></span>| <span data-ttu-id="2ab52-299">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-300">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-301">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-301">1.0</span></span>|
|[<span data-ttu-id="2ab52-302">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-302">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-303">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-304">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-304">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-305">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-306">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-306">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="2ab52-307">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="2ab52-307">dateTimeModified: Date</span></span>

<span data-ttu-id="2ab52-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-310">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-310">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-311">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-311">Type</span></span>

*   <span data-ttu-id="2ab52-312">日期</span><span class="sxs-lookup"><span data-stu-id="2ab52-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-313">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-313">Requirements</span></span>

|<span data-ttu-id="2ab52-314">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-314">Requirement</span></span>| <span data-ttu-id="2ab52-315">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-316">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-317">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-317">1.0</span></span>|
|[<span data-ttu-id="2ab52-318">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-319">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-320">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-321">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-322">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-322">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="2ab52-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-323">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-324">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2ab52-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="2ab52-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-327">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-327">Read mode</span></span>

<span data-ttu-id="2ab52-328">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-328">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-329">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-329">Compose mode</span></span>

<span data-ttu-id="2ab52-330">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="2ab52-331">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="2ab52-331">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2ab52-332">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="2ab52-332">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2ab52-333">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-333">Type</span></span>

*   <span data-ttu-id="2ab52-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-334">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-335">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-335">Requirements</span></span>

|<span data-ttu-id="2ab52-336">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-336">Requirement</span></span>| <span data-ttu-id="2ab52-337">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-338">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-339">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-339">1.0</span></span>|
|[<span data-ttu-id="2ab52-340">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-341">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-342">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-343">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-343">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="2ab52-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-344">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-p114">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p114">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="2ab52-p115">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p115">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-349">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-350">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-350">Type</span></span>

*   [<span data-ttu-id="2ab52-351">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2ab52-351">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="2ab52-352">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-352">Requirements</span></span>

|<span data-ttu-id="2ab52-353">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-353">Requirement</span></span>| <span data-ttu-id="2ab52-354">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-355">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-356">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-356">1.0</span></span>|
|[<span data-ttu-id="2ab52-357">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-357">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-358">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-359">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-359">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-360">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-360">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-361">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-361">Example</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="2ab52-362">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="2ab52-362">internetMessageId: String</span></span>

<span data-ttu-id="2ab52-p116">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-365">Type</span><span class="sxs-lookup"><span data-stu-id="2ab52-365">Type</span></span>

*   <span data-ttu-id="2ab52-366">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-366">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-367">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-367">Requirements</span></span>

|<span data-ttu-id="2ab52-368">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-368">Requirement</span></span>| <span data-ttu-id="2ab52-369">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-370">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-371">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-371">1.0</span></span>|
|[<span data-ttu-id="2ab52-372">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-373">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-374">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-374">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-375">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-375">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-376">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-376">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="2ab52-377">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="2ab52-377">itemClass: String</span></span>

<span data-ttu-id="2ab52-p117">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="2ab52-p118">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="2ab52-382">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-382">Type</span></span> | <span data-ttu-id="2ab52-383">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-383">Description</span></span> | <span data-ttu-id="2ab52-384">项目类</span><span class="sxs-lookup"><span data-stu-id="2ab52-384">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="2ab52-385">约会项目</span><span class="sxs-lookup"><span data-stu-id="2ab52-385">Appointment items</span></span> | <span data-ttu-id="2ab52-386">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="2ab52-386">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="2ab52-387">邮件项目</span><span class="sxs-lookup"><span data-stu-id="2ab52-387">Message items</span></span> | <span data-ttu-id="2ab52-388">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="2ab52-388">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="2ab52-389">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-389">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-390">Type</span><span class="sxs-lookup"><span data-stu-id="2ab52-390">Type</span></span>

*   <span data-ttu-id="2ab52-391">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-391">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-392">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-392">Requirements</span></span>

|<span data-ttu-id="2ab52-393">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-393">Requirement</span></span>| <span data-ttu-id="2ab52-394">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-395">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-396">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-396">1.0</span></span>|
|[<span data-ttu-id="2ab52-397">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-397">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-398">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-399">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-399">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-400">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-400">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-401">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-401">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="2ab52-402">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="2ab52-402">(nullable) itemId: String</span></span>

<span data-ttu-id="2ab52-p119">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-405">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="2ab52-405">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="2ab52-406">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="2ab52-406">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="2ab52-407">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="2ab52-407">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="2ab52-408">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="2ab52-408">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="2ab52-p121">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-411">Type</span><span class="sxs-lookup"><span data-stu-id="2ab52-411">Type</span></span>

*   <span data-ttu-id="2ab52-412">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-412">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-413">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-413">Requirements</span></span>

|<span data-ttu-id="2ab52-414">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-414">Requirement</span></span>| <span data-ttu-id="2ab52-415">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-415">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-416">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-417">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-417">1.0</span></span>|
|[<span data-ttu-id="2ab52-418">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-418">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-419">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-419">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-420">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-420">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-421">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-421">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-422">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-422">Example</span></span>

<span data-ttu-id="2ab52-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-14"></a><span data-ttu-id="2ab52-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-425">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-426">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="2ab52-426">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="2ab52-427">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="2ab52-427">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-428">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-428">Type</span></span>

*   [<span data-ttu-id="2ab52-429">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="2ab52-429">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="2ab52-430">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-430">Requirements</span></span>

|<span data-ttu-id="2ab52-431">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-431">Requirement</span></span>| <span data-ttu-id="2ab52-432">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-432">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-433">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-433">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-434">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-434">1.0</span></span>|
|[<span data-ttu-id="2ab52-435">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-435">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-436">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-436">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-437">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-437">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-438">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-438">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-439">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-439">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-14"></a><span data-ttu-id="2ab52-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-440">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-441">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="2ab52-441">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-442">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-442">Read mode</span></span>

<span data-ttu-id="2ab52-443">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="2ab52-443">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-444">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-444">Compose mode</span></span>

<span data-ttu-id="2ab52-445">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-445">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2ab52-446">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-446">Type</span></span>

*   <span data-ttu-id="2ab52-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-447">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-448">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-448">Requirements</span></span>

|<span data-ttu-id="2ab52-449">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-449">Requirement</span></span>| <span data-ttu-id="2ab52-450">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-450">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-451">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-451">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-452">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-452">1.0</span></span>|
|[<span data-ttu-id="2ab52-453">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-453">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-454">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-454">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-455">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-455">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-456">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-456">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="2ab52-457">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="2ab52-457">normalizedSubject: String</span></span>

<span data-ttu-id="2ab52-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="2ab52-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-462">Type</span><span class="sxs-lookup"><span data-stu-id="2ab52-462">Type</span></span>

*   <span data-ttu-id="2ab52-463">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-464">Requirements</span></span>

|<span data-ttu-id="2ab52-465">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-465">Requirement</span></span>| <span data-ttu-id="2ab52-466">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-468">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-468">1.0</span></span>|
|[<span data-ttu-id="2ab52-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-470">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-472">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-473">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-473">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-14"></a><span data-ttu-id="2ab52-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-474">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-475">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="2ab52-475">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-476">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-476">Type</span></span>

*   [<span data-ttu-id="2ab52-477">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="2ab52-477">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="2ab52-478">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-478">Requirements</span></span>

|<span data-ttu-id="2ab52-479">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-479">Requirement</span></span>| <span data-ttu-id="2ab52-480">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-481">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-482">1.3</span><span class="sxs-lookup"><span data-stu-id="2ab52-482">1.3</span></span>|
|[<span data-ttu-id="2ab52-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-484">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-486">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-486">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-487">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-487">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="2ab52-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-488">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-489">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2ab52-489">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="2ab52-490">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-490">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-491">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-491">Read mode</span></span>

<span data-ttu-id="2ab52-492">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-492">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="2ab52-493">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-493">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-494">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-494">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-495">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-495">Compose mode</span></span>

<span data-ttu-id="2ab52-496">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-496">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="2ab52-497">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-497">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-498">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2ab52-498">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2ab52-499">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-499">Get 500 members maximum.</span></span>
- <span data-ttu-id="2ab52-500">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-500">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2ab52-501">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-501">Type</span></span>

*   <span data-ttu-id="2ab52-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-502">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-503">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-503">Requirements</span></span>

|<span data-ttu-id="2ab52-504">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-504">Requirement</span></span>| <span data-ttu-id="2ab52-505">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-506">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-507">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-507">1.0</span></span>|
|[<span data-ttu-id="2ab52-508">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-509">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-510">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-511">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-511">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="2ab52-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-512">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-p128">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p128">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-515">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-515">Type</span></span>

*   [<span data-ttu-id="2ab52-516">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2ab52-516">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="2ab52-517">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-517">Requirements</span></span>

|<span data-ttu-id="2ab52-518">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-518">Requirement</span></span>| <span data-ttu-id="2ab52-519">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-519">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-520">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-520">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-521">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-521">1.0</span></span>|
|[<span data-ttu-id="2ab52-522">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-522">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-523">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-523">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-524">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-524">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-525">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-525">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-526">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-526">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="2ab52-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-527">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-528">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2ab52-528">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="2ab52-529">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-529">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-530">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-530">Read mode</span></span>

<span data-ttu-id="2ab52-531">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-531">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="2ab52-532">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-532">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-533">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-533">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-534">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-534">Compose mode</span></span>

<span data-ttu-id="2ab52-535">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-535">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="2ab52-536">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-536">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-537">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2ab52-537">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2ab52-538">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-538">Get 500 members maximum.</span></span>
- <span data-ttu-id="2ab52-539">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-539">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="2ab52-540">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-540">Type</span></span>

*   <span data-ttu-id="2ab52-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-541">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-542">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-542">Requirements</span></span>

|<span data-ttu-id="2ab52-543">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-543">Requirement</span></span>| <span data-ttu-id="2ab52-544">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-545">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-545">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-546">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-546">1.0</span></span>|
|[<span data-ttu-id="2ab52-547">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-547">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-548">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-549">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-549">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-550">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-550">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14"></a><span data-ttu-id="2ab52-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-551">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-p132">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p132">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="2ab52-p133">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p133">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-556">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-556">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="2ab52-557">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-557">Type</span></span>

*   [<span data-ttu-id="2ab52-558">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="2ab52-558">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)

##### <a name="requirements"></a><span data-ttu-id="2ab52-559">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-559">Requirements</span></span>

|<span data-ttu-id="2ab52-560">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-560">Requirement</span></span>| <span data-ttu-id="2ab52-561">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-561">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-562">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-563">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-563">1.0</span></span>|
|[<span data-ttu-id="2ab52-564">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-565">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-566">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-566">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-567">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-567">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-568">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-568">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-14"></a><span data-ttu-id="2ab52-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-569">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-570">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2ab52-570">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="2ab52-p134">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p134">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-573">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-573">Read mode</span></span>

<span data-ttu-id="2ab52-574">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-574">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-575">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-575">Compose mode</span></span>

<span data-ttu-id="2ab52-576">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-576">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="2ab52-577">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="2ab52-577">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="2ab52-578">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="2ab52-578">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.4#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="2ab52-579">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-579">Type</span></span>

*   <span data-ttu-id="2ab52-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-580">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-581">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-581">Requirements</span></span>

|<span data-ttu-id="2ab52-582">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-582">Requirement</span></span>| <span data-ttu-id="2ab52-583">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-584">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-585">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-585">1.0</span></span>|
|[<span data-ttu-id="2ab52-586">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-586">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-587">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-588">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-589">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-589">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-14"></a><span data-ttu-id="2ab52-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-590">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-591">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="2ab52-591">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="2ab52-592">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="2ab52-592">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-593">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-593">Read mode</span></span>

<span data-ttu-id="2ab52-p135">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p135">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-596">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-596">Compose mode</span></span>

<span data-ttu-id="2ab52-597">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-597">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="2ab52-598">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-598">Type</span></span>

*   <span data-ttu-id="2ab52-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-599">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-600">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-600">Requirements</span></span>

|<span data-ttu-id="2ab52-601">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-601">Requirement</span></span>| <span data-ttu-id="2ab52-602">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-603">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-604">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-604">1.0</span></span>|
|[<span data-ttu-id="2ab52-605">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-605">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-606">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-607">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-607">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-608">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-608">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-14recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-14"></a><span data-ttu-id="2ab52-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-609">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

<span data-ttu-id="2ab52-610">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="2ab52-610">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="2ab52-611">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-611">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="2ab52-612">阅读模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-612">Read mode</span></span>

<span data-ttu-id="2ab52-613">`to` 属性返回包含邮件的`EmailAddressDetails`行上所列的每个收件人的 \*\*\*\* 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-613">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="2ab52-614">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-614">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-615">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-615">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="2ab52-616">撰写模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-616">Compose mode</span></span>

<span data-ttu-id="2ab52-617">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-617">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="2ab52-618">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-618">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="2ab52-619">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="2ab52-619">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="2ab52-620">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="2ab52-620">Get 500 members maximum.</span></span>
- <span data-ttu-id="2ab52-621">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-621">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="2ab52-622">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-622">Type</span></span>

*   <span data-ttu-id="2ab52-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-623">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.4)</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-624">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-624">Requirements</span></span>

|<span data-ttu-id="2ab52-625">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-625">Requirement</span></span>| <span data-ttu-id="2ab52-626">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-627">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-628">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-628">1.0</span></span>|
|[<span data-ttu-id="2ab52-629">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-630">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-631">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-632">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-632">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2ab52-633">方法</span><span class="sxs-lookup"><span data-stu-id="2ab52-633">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="2ab52-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2ab52-634">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2ab52-635">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="2ab52-635">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="2ab52-636">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="2ab52-636">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="2ab52-637">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="2ab52-637">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-638">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-638">Parameters</span></span>

|<span data-ttu-id="2ab52-639">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-639">Name</span></span>| <span data-ttu-id="2ab52-640">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-640">Type</span></span>| <span data-ttu-id="2ab52-641">属性</span><span class="sxs-lookup"><span data-stu-id="2ab52-641">Attributes</span></span>| <span data-ttu-id="2ab52-642">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-642">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="2ab52-643">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-643">String</span></span>||<span data-ttu-id="2ab52-p139">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p139">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2ab52-646">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-646">String</span></span>||<span data-ttu-id="2ab52-p140">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p140">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2ab52-649">Object</span><span class="sxs-lookup"><span data-stu-id="2ab52-649">Object</span></span>| <span data-ttu-id="2ab52-650">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-650">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-651">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2ab52-651">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2ab52-652">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-652">Object</span></span>| <span data-ttu-id="2ab52-653">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-653">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-654">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-654">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2ab52-655">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-655">function</span></span>| <span data-ttu-id="2ab52-656">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-656">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-657">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-657">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2ab52-658">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2ab52-658">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2ab52-659">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-659">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2ab52-660">错误</span><span class="sxs-lookup"><span data-stu-id="2ab52-660">Errors</span></span>

| <span data-ttu-id="2ab52-661">错误代码</span><span class="sxs-lookup"><span data-stu-id="2ab52-661">Error code</span></span> | <span data-ttu-id="2ab52-662">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-662">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="2ab52-663">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="2ab52-663">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="2ab52-664">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="2ab52-664">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2ab52-665">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="2ab52-665">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2ab52-666">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-666">Requirements</span></span>

|<span data-ttu-id="2ab52-667">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-667">Requirement</span></span>| <span data-ttu-id="2ab52-668">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-669">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-670">1.1</span><span class="sxs-lookup"><span data-stu-id="2ab52-670">1.1</span></span>|
|[<span data-ttu-id="2ab52-671">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-671">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-672">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-672">ReadWriteItem</span></span>|
|[<span data-ttu-id="2ab52-673">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-673">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-674">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-674">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-675">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-675">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="2ab52-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2ab52-676">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="2ab52-677">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="2ab52-677">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="2ab52-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="2ab52-681">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="2ab52-681">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="2ab52-682">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="2ab52-682">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-683">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-683">Parameters</span></span>

|<span data-ttu-id="2ab52-684">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-684">Name</span></span>| <span data-ttu-id="2ab52-685">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-685">Type</span></span>| <span data-ttu-id="2ab52-686">属性</span><span class="sxs-lookup"><span data-stu-id="2ab52-686">Attributes</span></span>| <span data-ttu-id="2ab52-687">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-687">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="2ab52-688">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-688">String</span></span>||<span data-ttu-id="2ab52-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="2ab52-691">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-691">String</span></span>||<span data-ttu-id="2ab52-692">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="2ab52-692">The subject of the item to be attached.</span></span> <span data-ttu-id="2ab52-693">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-693">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="2ab52-694">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-694">Object</span></span>| <span data-ttu-id="2ab52-695">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-695">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-696">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2ab52-696">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2ab52-697">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-697">Object</span></span>| <span data-ttu-id="2ab52-698">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-698">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-699">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-699">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2ab52-700">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-700">function</span></span>| <span data-ttu-id="2ab52-701">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-701">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-702">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2ab52-703">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2ab52-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="2ab52-704">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-704">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2ab52-705">错误</span><span class="sxs-lookup"><span data-stu-id="2ab52-705">Errors</span></span>

| <span data-ttu-id="2ab52-706">错误代码</span><span class="sxs-lookup"><span data-stu-id="2ab52-706">Error code</span></span> | <span data-ttu-id="2ab52-707">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-707">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="2ab52-708">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="2ab52-708">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2ab52-709">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-709">Requirements</span></span>

|<span data-ttu-id="2ab52-710">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-710">Requirement</span></span>| <span data-ttu-id="2ab52-711">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-712">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-713">1.1</span><span class="sxs-lookup"><span data-stu-id="2ab52-713">1.1</span></span>|
|[<span data-ttu-id="2ab52-714">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-714">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-715">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-715">ReadWriteItem</span></span>|
|[<span data-ttu-id="2ab52-716">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-716">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-717">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-717">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-718">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-718">Example</span></span>

<span data-ttu-id="2ab52-719">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="2ab52-719">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="2ab52-720">close()</span><span class="sxs-lookup"><span data-stu-id="2ab52-720">close()</span></span>

<span data-ttu-id="2ab52-721">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="2ab52-721">Closes the current item that is being composed.</span></span>

<span data-ttu-id="2ab52-p144">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-724">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="2ab52-724">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="2ab52-725">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="2ab52-725">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-726">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-726">Requirements</span></span>

|<span data-ttu-id="2ab52-727">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-727">Requirement</span></span>| <span data-ttu-id="2ab52-728">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-729">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-730">1.3</span><span class="sxs-lookup"><span data-stu-id="2ab52-730">1.3</span></span>|
|[<span data-ttu-id="2ab52-731">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-732">受限</span><span class="sxs-lookup"><span data-stu-id="2ab52-732">Restricted</span></span>|
|[<span data-ttu-id="2ab52-733">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-734">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-734">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="2ab52-735">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2ab52-735">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="2ab52-736">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="2ab52-736">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-737">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-737">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2ab52-738">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="2ab52-738">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2ab52-739">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="2ab52-739">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="2ab52-p145">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p145">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-743">Parameters</span><span class="sxs-lookup"><span data-stu-id="2ab52-743">Parameters</span></span>

|<span data-ttu-id="2ab52-744">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-744">Name</span></span>| <span data-ttu-id="2ab52-745">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-745">Type</span></span>| <span data-ttu-id="2ab52-746">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-746">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2ab52-747">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-747">String &#124; Object</span></span>| |<span data-ttu-id="2ab52-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2ab52-750">**或**</span><span class="sxs-lookup"><span data-stu-id="2ab52-750">**OR**</span></span><br/><span data-ttu-id="2ab52-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2ab52-753">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-753">String</span></span> | <span data-ttu-id="2ab52-754">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-754">&lt;optional&gt;</span></span> | <span data-ttu-id="2ab52-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="2ab52-757">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-757">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2ab52-758">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-758">&lt;optional&gt;</span></span> | <span data-ttu-id="2ab52-759">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-759">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="2ab52-760">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-760">String</span></span> | | <span data-ttu-id="2ab52-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="2ab52-763">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-763">String</span></span> | | <span data-ttu-id="2ab52-764">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-764">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="2ab52-765">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-765">String</span></span> | | <span data-ttu-id="2ab52-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="2ab52-768">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-768">String</span></span> | | <span data-ttu-id="2ab52-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="2ab52-772">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-772">function</span></span> | <span data-ttu-id="2ab52-773">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-773">&lt;optional&gt;</span></span> | <span data-ttu-id="2ab52-774">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-774">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2ab52-775">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-775">Requirements</span></span>

|<span data-ttu-id="2ab52-776">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-776">Requirement</span></span>| <span data-ttu-id="2ab52-777">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-777">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-778">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-778">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-779">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-779">1.0</span></span>|
|[<span data-ttu-id="2ab52-780">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-780">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-781">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-781">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-782">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-782">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-783">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-783">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2ab52-784">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-784">Examples</span></span>

<span data-ttu-id="2ab52-785">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-785">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="2ab52-786">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-786">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="2ab52-787">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-787">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2ab52-788">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-788">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2ab52-789">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-789">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2ab52-790">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-790">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="2ab52-791">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="2ab52-791">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="2ab52-792">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="2ab52-792">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-793">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-793">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2ab52-794">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="2ab52-794">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="2ab52-795">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="2ab52-795">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="2ab52-p152">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p152">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-799">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-799">Parameters</span></span>

|<span data-ttu-id="2ab52-800">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-800">Name</span></span>| <span data-ttu-id="2ab52-801">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-801">Type</span></span>| <span data-ttu-id="2ab52-802">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-802">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="2ab52-803">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-803">String &#124; Object</span></span>| | <span data-ttu-id="2ab52-p153">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="2ab52-806">**或**</span><span class="sxs-lookup"><span data-stu-id="2ab52-806">**OR**</span></span><br/><span data-ttu-id="2ab52-p154">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p154">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="2ab52-809">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-809">String</span></span> | <span data-ttu-id="2ab52-810">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-810">&lt;optional&gt;</span></span> | <span data-ttu-id="2ab52-p155">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p155">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="2ab52-813">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-813">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2ab52-814">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-814">&lt;optional&gt;</span></span> | <span data-ttu-id="2ab52-815">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-815">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="2ab52-816">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-816">String</span></span> | | <span data-ttu-id="2ab52-p156">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p156">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="2ab52-819">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-819">String</span></span> | | <span data-ttu-id="2ab52-820">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-820">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="2ab52-821">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-821">String</span></span> | | <span data-ttu-id="2ab52-p157">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p157">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="2ab52-824">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-824">String</span></span> | | <span data-ttu-id="2ab52-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="2ab52-828">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-828">function</span></span> | <span data-ttu-id="2ab52-829">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-829">&lt;optional&gt;</span></span> | <span data-ttu-id="2ab52-830">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-830">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2ab52-831">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-831">Requirements</span></span>

|<span data-ttu-id="2ab52-832">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-832">Requirement</span></span>| <span data-ttu-id="2ab52-833">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-834">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-835">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-835">1.0</span></span>|
|[<span data-ttu-id="2ab52-836">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-837">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-837">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-838">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-839">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-839">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="2ab52-840">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-840">Examples</span></span>

<span data-ttu-id="2ab52-841">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-841">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="2ab52-842">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-842">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="2ab52-843">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-843">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="2ab52-844">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-844">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="2ab52-845">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-845">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="2ab52-846">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="2ab52-846">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-14"></a><span data-ttu-id="2ab52-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="2ab52-847">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="2ab52-848">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="2ab52-848">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-849">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-849">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-850">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-850">Requirements</span></span>

|<span data-ttu-id="2ab52-851">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-851">Requirement</span></span>| <span data-ttu-id="2ab52-852">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-852">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-853">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-853">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-854">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-854">1.0</span></span>|
|[<span data-ttu-id="2ab52-855">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-855">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-856">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-856">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-857">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-857">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-858">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-858">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2ab52-859">返回：</span><span class="sxs-lookup"><span data-stu-id="2ab52-859">Returns:</span></span>

<span data-ttu-id="2ab52-860">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="2ab52-860">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.4)</span></span>

##### <a name="example"></a><span data-ttu-id="2ab52-861">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-861">Example</span></span>

<span data-ttu-id="2ab52-862">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="2ab52-862">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="2ab52-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="2ab52-863">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="2ab52-864">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-864">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-865">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-865">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-866">Parameters</span><span class="sxs-lookup"><span data-stu-id="2ab52-866">Parameters</span></span>

|<span data-ttu-id="2ab52-867">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-867">Name</span></span>| <span data-ttu-id="2ab52-868">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-868">Type</span></span>| <span data-ttu-id="2ab52-869">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-869">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="2ab52-870">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="2ab52-870">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.4)|<span data-ttu-id="2ab52-871">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="2ab52-871">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2ab52-872">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-872">Requirements</span></span>

|<span data-ttu-id="2ab52-873">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-873">Requirement</span></span>| <span data-ttu-id="2ab52-874">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-874">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-875">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-875">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-876">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-876">1.0</span></span>|
|[<span data-ttu-id="2ab52-877">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-877">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-878">受限</span><span class="sxs-lookup"><span data-stu-id="2ab52-878">Restricted</span></span>|
|[<span data-ttu-id="2ab52-879">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-879">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-880">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-880">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2ab52-881">返回：</span><span class="sxs-lookup"><span data-stu-id="2ab52-881">Returns:</span></span>

<span data-ttu-id="2ab52-882">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="2ab52-882">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="2ab52-883">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-883">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="2ab52-884">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="2ab52-884">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="2ab52-885">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="2ab52-885">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="2ab52-886">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="2ab52-886">Value of `entityType`</span></span> | <span data-ttu-id="2ab52-887">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-887">Type of objects in returned array</span></span> | <span data-ttu-id="2ab52-888">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-888">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="2ab52-889">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-889">String</span></span> | <span data-ttu-id="2ab52-890">**受限**</span><span class="sxs-lookup"><span data-stu-id="2ab52-890">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="2ab52-891">Contact</span><span class="sxs-lookup"><span data-stu-id="2ab52-891">Contact</span></span> | <span data-ttu-id="2ab52-892">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2ab52-892">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="2ab52-893">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-893">String</span></span> | <span data-ttu-id="2ab52-894">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2ab52-894">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="2ab52-895">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="2ab52-895">MeetingSuggestion</span></span> | <span data-ttu-id="2ab52-896">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2ab52-896">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="2ab52-897">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="2ab52-897">PhoneNumber</span></span> | <span data-ttu-id="2ab52-898">**受限**</span><span class="sxs-lookup"><span data-stu-id="2ab52-898">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="2ab52-899">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="2ab52-899">TaskSuggestion</span></span> | <span data-ttu-id="2ab52-900">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="2ab52-900">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="2ab52-901">String</span><span class="sxs-lookup"><span data-stu-id="2ab52-901">String</span></span> | <span data-ttu-id="2ab52-902">**受限**</span><span class="sxs-lookup"><span data-stu-id="2ab52-902">**Restricted**</span></span> |

<span data-ttu-id="2ab52-903">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="2ab52-903">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

##### <a name="example"></a><span data-ttu-id="2ab52-904">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-904">Example</span></span>

<span data-ttu-id="2ab52-905">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-905">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-14meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-14phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-14tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-14"></a><span data-ttu-id="2ab52-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span><span class="sxs-lookup"><span data-stu-id="2ab52-906">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))>}</span></span>

<span data-ttu-id="2ab52-907">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="2ab52-907">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-908">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-908">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2ab52-909">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="2ab52-909">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-910">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-910">Parameters</span></span>

|<span data-ttu-id="2ab52-911">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-911">Name</span></span>| <span data-ttu-id="2ab52-912">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-912">Type</span></span>| <span data-ttu-id="2ab52-913">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-913">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2ab52-914">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-914">String</span></span>|<span data-ttu-id="2ab52-915">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="2ab52-915">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2ab52-916">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-916">Requirements</span></span>

|<span data-ttu-id="2ab52-917">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-917">Requirement</span></span>| <span data-ttu-id="2ab52-918">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-918">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-919">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-919">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-920">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-920">1.0</span></span>|
|[<span data-ttu-id="2ab52-921">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-921">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-922">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-922">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-923">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-923">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-924">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-924">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2ab52-925">返回：</span><span class="sxs-lookup"><span data-stu-id="2ab52-925">Returns:</span></span>

<span data-ttu-id="2ab52-p160">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="2ab52-928">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span><span class="sxs-lookup"><span data-stu-id="2ab52-928">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.4)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.4)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.4)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.4))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="2ab52-929">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="2ab52-929">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="2ab52-930">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-930">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-931">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-931">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2ab52-p161">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="2ab52-935">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="2ab52-935">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="2ab52-936">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-936">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="2ab52-p162">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.4#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ab52-940">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-940">Requirements</span></span>

|<span data-ttu-id="2ab52-941">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-941">Requirement</span></span>| <span data-ttu-id="2ab52-942">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-942">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-943">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-943">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-944">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-944">1.0</span></span>|
|[<span data-ttu-id="2ab52-945">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-945">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-946">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-946">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-947">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-947">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-948">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-948">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2ab52-949">返回：</span><span class="sxs-lookup"><span data-stu-id="2ab52-949">Returns:</span></span>

<span data-ttu-id="2ab52-p163">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="2ab52-952">类型：对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-952">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="2ab52-953">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-953">Example</span></span>

<span data-ttu-id="2ab52-954">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="2ab52-954">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="2ab52-955">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="2ab52-955">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="2ab52-956">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="2ab52-956">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-957">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2ab52-958">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="2ab52-958">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="2ab52-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-961">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-961">Parameters</span></span>

|<span data-ttu-id="2ab52-962">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-962">Name</span></span>| <span data-ttu-id="2ab52-963">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-963">Type</span></span>| <span data-ttu-id="2ab52-964">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-964">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="2ab52-965">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-965">String</span></span>|<span data-ttu-id="2ab52-966">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="2ab52-966">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2ab52-967">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-967">Requirements</span></span>

|<span data-ttu-id="2ab52-968">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-968">Requirement</span></span>| <span data-ttu-id="2ab52-969">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-969">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-970">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-970">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-971">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-971">1.0</span></span>|
|[<span data-ttu-id="2ab52-972">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-972">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-973">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-973">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-974">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-974">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-975">阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-975">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2ab52-976">返回：</span><span class="sxs-lookup"><span data-stu-id="2ab52-976">Returns:</span></span>

<span data-ttu-id="2ab52-977">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="2ab52-977">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="2ab52-978">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="2ab52-978">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="2ab52-979">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-979">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="2ab52-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="2ab52-980">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="2ab52-981">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="2ab52-981">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="2ab52-p165">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-984">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-984">Parameters</span></span>

|<span data-ttu-id="2ab52-985">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-985">Name</span></span>| <span data-ttu-id="2ab52-986">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-986">Type</span></span>| <span data-ttu-id="2ab52-987">属性</span><span class="sxs-lookup"><span data-stu-id="2ab52-987">Attributes</span></span>| <span data-ttu-id="2ab52-988">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-988">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="2ab52-989">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2ab52-989">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="2ab52-p166">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="2ab52-993">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-993">Object</span></span>| <span data-ttu-id="2ab52-994">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-994">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-995">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2ab52-995">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2ab52-996">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-996">Object</span></span>| <span data-ttu-id="2ab52-997">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-997">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-998">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-998">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2ab52-999">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-999">function</span></span>||<span data-ttu-id="2ab52-1000">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1000">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2ab52-1001">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1001">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="2ab52-1002">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1002">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2ab52-1003">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-1003">Requirements</span></span>

|<span data-ttu-id="2ab52-1004">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-1004">Requirement</span></span>| <span data-ttu-id="2ab52-1005">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-1006">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-1007">1.2</span><span class="sxs-lookup"><span data-stu-id="2ab52-1007">1.2</span></span>|
|[<span data-ttu-id="2ab52-1008">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-1008">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-1009">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-1010">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-1010">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-1011">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-1011">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="2ab52-1012">返回：</span><span class="sxs-lookup"><span data-stu-id="2ab52-1012">Returns:</span></span>

<span data-ttu-id="2ab52-1013">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1013">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="2ab52-1014">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-1014">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2ab52-1015">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-1015">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="2ab52-1016">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2ab52-1016">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="2ab52-1017">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1017">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="2ab52-p168">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-1021">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-1021">Parameters</span></span>

|<span data-ttu-id="2ab52-1022">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-1022">Name</span></span>| <span data-ttu-id="2ab52-1023">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-1023">Type</span></span>| <span data-ttu-id="2ab52-1024">属性</span><span class="sxs-lookup"><span data-stu-id="2ab52-1024">Attributes</span></span>| <span data-ttu-id="2ab52-1025">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-1025">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2ab52-1026">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-1026">function</span></span>||<span data-ttu-id="2ab52-1027">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2ab52-1028">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1028">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.4) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="2ab52-1029">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1029">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="2ab52-1030">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-1030">Object</span></span>| <span data-ttu-id="2ab52-1031">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1031">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1032">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1032">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="2ab52-1033">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1033">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2ab52-1034">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-1034">Requirements</span></span>

|<span data-ttu-id="2ab52-1035">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-1035">Requirement</span></span>| <span data-ttu-id="2ab52-1036">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-1036">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-1037">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-1037">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-1038">1.0</span><span class="sxs-lookup"><span data-stu-id="2ab52-1038">1.0</span></span>|
|[<span data-ttu-id="2ab52-1039">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-1039">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-1040">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-1040">ReadItem</span></span>|
|[<span data-ttu-id="2ab52-1041">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-1041">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-1042">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="2ab52-1042">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-1043">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-1043">Example</span></span>

<span data-ttu-id="2ab52-p171">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="2ab52-1047">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2ab52-1047">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="2ab52-1048">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1048">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="2ab52-1049">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1049">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="2ab52-1050">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1050">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="2ab52-1051">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1051">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="2ab52-1052">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1052">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-1053">Parameters</span><span class="sxs-lookup"><span data-stu-id="2ab52-1053">Parameters</span></span>

|<span data-ttu-id="2ab52-1054">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-1054">Name</span></span>| <span data-ttu-id="2ab52-1055">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-1055">Type</span></span>| <span data-ttu-id="2ab52-1056">属性</span><span class="sxs-lookup"><span data-stu-id="2ab52-1056">Attributes</span></span>| <span data-ttu-id="2ab52-1057">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-1057">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="2ab52-1058">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-1058">String</span></span>||<span data-ttu-id="2ab52-1059">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1059">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="2ab52-1060">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-1060">Object</span></span>| <span data-ttu-id="2ab52-1061">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1062">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1062">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2ab52-1063">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-1063">Object</span></span>| <span data-ttu-id="2ab52-1064">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1065">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1065">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="2ab52-1066">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-1066">function</span></span>| <span data-ttu-id="2ab52-1067">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1068">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1068">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="2ab52-1069">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1069">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2ab52-1070">错误</span><span class="sxs-lookup"><span data-stu-id="2ab52-1070">Errors</span></span>

| <span data-ttu-id="2ab52-1071">错误代码</span><span class="sxs-lookup"><span data-stu-id="2ab52-1071">Error code</span></span> | <span data-ttu-id="2ab52-1072">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-1072">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="2ab52-1073">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1073">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2ab52-1074">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-1074">Requirements</span></span>

|<span data-ttu-id="2ab52-1075">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-1075">Requirement</span></span>| <span data-ttu-id="2ab52-1076">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-1076">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-1077">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-1077">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-1078">1.1</span><span class="sxs-lookup"><span data-stu-id="2ab52-1078">1.1</span></span>|
|[<span data-ttu-id="2ab52-1079">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-1079">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-1080">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-1080">ReadWriteItem</span></span>|
|[<span data-ttu-id="2ab52-1081">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-1081">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-1082">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-1082">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-1083">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-1083">Example</span></span>

<span data-ttu-id="2ab52-1084">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1084">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="2ab52-1085">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2ab52-1085">saveAsync([options], callback)</span></span>

<span data-ttu-id="2ab52-1086">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1086">Asynchronously saves an item.</span></span>

<span data-ttu-id="2ab52-1087">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1087">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="2ab52-1088">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1088">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="2ab52-1089">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1089">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-1090">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1090">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="2ab52-1091">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1091">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="2ab52-p175">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="2ab52-1095">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="2ab52-1095">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="2ab52-1096">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1096">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="2ab52-1097">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1097">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="2ab52-1098">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1098">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="2ab52-1099">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1099">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-1100">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-1100">Parameters</span></span>

|<span data-ttu-id="2ab52-1101">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-1101">Name</span></span>| <span data-ttu-id="2ab52-1102">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-1102">Type</span></span>| <span data-ttu-id="2ab52-1103">属性</span><span class="sxs-lookup"><span data-stu-id="2ab52-1103">Attributes</span></span>| <span data-ttu-id="2ab52-1104">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-1104">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="2ab52-1105">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-1105">Object</span></span>| <span data-ttu-id="2ab52-1106">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1106">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1107">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1107">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2ab52-1108">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-1108">Object</span></span>| <span data-ttu-id="2ab52-1109">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1109">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1110">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1110">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="2ab52-1111">函数</span><span class="sxs-lookup"><span data-stu-id="2ab52-1111">function</span></span>||<span data-ttu-id="2ab52-1112">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1112">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2ab52-1113">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1113">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2ab52-1114">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-1114">Requirements</span></span>

|<span data-ttu-id="2ab52-1115">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-1115">Requirement</span></span>| <span data-ttu-id="2ab52-1116">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-1116">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-1117">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-1117">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-1118">1.3</span><span class="sxs-lookup"><span data-stu-id="2ab52-1118">1.3</span></span>|
|[<span data-ttu-id="2ab52-1119">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-1119">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-1120">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-1120">ReadWriteItem</span></span>|
|[<span data-ttu-id="2ab52-1121">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-1121">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-1122">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-1122">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="2ab52-1123">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-1123">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="2ab52-p177">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="2ab52-1126">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="2ab52-1126">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="2ab52-1127">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1127">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="2ab52-p178">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2ab52-1131">参数</span><span class="sxs-lookup"><span data-stu-id="2ab52-1131">Parameters</span></span>

|<span data-ttu-id="2ab52-1132">名称</span><span class="sxs-lookup"><span data-stu-id="2ab52-1132">Name</span></span>| <span data-ttu-id="2ab52-1133">类型</span><span class="sxs-lookup"><span data-stu-id="2ab52-1133">Type</span></span>| <span data-ttu-id="2ab52-1134">属性</span><span class="sxs-lookup"><span data-stu-id="2ab52-1134">Attributes</span></span>| <span data-ttu-id="2ab52-1135">说明</span><span class="sxs-lookup"><span data-stu-id="2ab52-1135">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2ab52-1136">字符串</span><span class="sxs-lookup"><span data-stu-id="2ab52-1136">String</span></span>||<span data-ttu-id="2ab52-p179">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="2ab52-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="2ab52-1140">Object</span><span class="sxs-lookup"><span data-stu-id="2ab52-1140">Object</span></span>| <span data-ttu-id="2ab52-1141">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1141">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1142">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1142">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="2ab52-1143">对象</span><span class="sxs-lookup"><span data-stu-id="2ab52-1143">Object</span></span>| <span data-ttu-id="2ab52-1144">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1144">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1145">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1145">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="2ab52-1146">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="2ab52-1146">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="2ab52-1147">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2ab52-1147">&lt;optional&gt;</span></span>|<span data-ttu-id="2ab52-1148">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1148">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="2ab52-1149">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1149">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="2ab52-1150">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1150">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="2ab52-1151">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1151">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="2ab52-1152">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1152">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="2ab52-1153">function</span><span class="sxs-lookup"><span data-stu-id="2ab52-1153">function</span></span>||<span data-ttu-id="2ab52-1154">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="2ab52-1154">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2ab52-1155">Requirements</span><span class="sxs-lookup"><span data-stu-id="2ab52-1155">Requirements</span></span>

|<span data-ttu-id="2ab52-1156">要求</span><span class="sxs-lookup"><span data-stu-id="2ab52-1156">Requirement</span></span>| <span data-ttu-id="2ab52-1157">值</span><span class="sxs-lookup"><span data-stu-id="2ab52-1157">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ab52-1158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="2ab52-1158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ab52-1159">1.2</span><span class="sxs-lookup"><span data-stu-id="2ab52-1159">1.2</span></span>|
|[<span data-ttu-id="2ab52-1160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="2ab52-1160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ab52-1161">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="2ab52-1161">ReadWriteItem</span></span>|
|[<span data-ttu-id="2ab52-1162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="2ab52-1162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ab52-1163">撰写</span><span class="sxs-lookup"><span data-stu-id="2ab52-1163">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="2ab52-1164">示例</span><span class="sxs-lookup"><span data-stu-id="2ab52-1164">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
