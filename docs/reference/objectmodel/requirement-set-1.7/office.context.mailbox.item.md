---
title: "\"Context\"-\"邮箱\"。项目-要求集1。7"
description: ''
ms.date: 10/23/2019
localization_priority: Normal
ms.openlocfilehash: 2cb6987191427cd5540eaa8647a78bccf2c073c1
ms.sourcegitcommit: 5ba325cc88183a3f230cd89d615fd49c695addcf
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/24/2019
ms.locfileid: "37682632"
---
# <a name="item"></a><span data-ttu-id="0d6f0-102">item</span><span class="sxs-lookup"><span data-stu-id="0d6f0-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="0d6f0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="0d6f0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="0d6f0-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-106">Requirements</span></span>

|<span data-ttu-id="0d6f0-107">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-107">Requirement</span></span>|<span data-ttu-id="0d6f0-108">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-110">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-110">1.0</span></span>|
|[<span data-ttu-id="0d6f0-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-112">受限</span><span class="sxs-lookup"><span data-stu-id="0d6f0-112">Restricted</span></span>|
|[<span data-ttu-id="0d6f0-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0d6f0-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-115">Members and methods</span></span>

| <span data-ttu-id="0d6f0-116">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-116">Member</span></span> | <span data-ttu-id="0d6f0-117">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0d6f0-118">attachments</span><span class="sxs-lookup"><span data-stu-id="0d6f0-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="0d6f0-119">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-119">Member</span></span> |
| [<span data-ttu-id="0d6f0-120">bcc</span><span class="sxs-lookup"><span data-stu-id="0d6f0-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="0d6f0-121">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-121">Member</span></span> |
| [<span data-ttu-id="0d6f0-122">body</span><span class="sxs-lookup"><span data-stu-id="0d6f0-122">body</span></span>](#body-body) | <span data-ttu-id="0d6f0-123">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-123">Member</span></span> |
| [<span data-ttu-id="0d6f0-124">cc</span><span class="sxs-lookup"><span data-stu-id="0d6f0-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6f0-125">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-125">Member</span></span> |
| [<span data-ttu-id="0d6f0-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="0d6f0-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="0d6f0-127">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-127">Member</span></span> |
| [<span data-ttu-id="0d6f0-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="0d6f0-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="0d6f0-129">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-129">Member</span></span> |
| [<span data-ttu-id="0d6f0-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="0d6f0-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="0d6f0-131">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-131">Member</span></span> |
| [<span data-ttu-id="0d6f0-132">end</span><span class="sxs-lookup"><span data-stu-id="0d6f0-132">end</span></span>](#end-datetime) | <span data-ttu-id="0d6f0-133">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-133">Member</span></span> |
| [<span data-ttu-id="0d6f0-134">from</span><span class="sxs-lookup"><span data-stu-id="0d6f0-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="0d6f0-135">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-135">Member</span></span> |
| [<span data-ttu-id="0d6f0-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="0d6f0-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="0d6f0-137">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-137">Member</span></span> |
| [<span data-ttu-id="0d6f0-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="0d6f0-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="0d6f0-139">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-139">Member</span></span> |
| [<span data-ttu-id="0d6f0-140">itemId</span><span class="sxs-lookup"><span data-stu-id="0d6f0-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="0d6f0-141">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-141">Member</span></span> |
| [<span data-ttu-id="0d6f0-142">itemType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="0d6f0-143">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-143">Member</span></span> |
| [<span data-ttu-id="0d6f0-144">location</span><span class="sxs-lookup"><span data-stu-id="0d6f0-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="0d6f0-145">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-145">Member</span></span> |
| [<span data-ttu-id="0d6f0-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="0d6f0-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="0d6f0-147">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-147">Member</span></span> |
| [<span data-ttu-id="0d6f0-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="0d6f0-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="0d6f0-149">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-149">Member</span></span> |
| [<span data-ttu-id="0d6f0-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="0d6f0-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6f0-151">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-151">Member</span></span> |
| [<span data-ttu-id="0d6f0-152">organizer</span><span class="sxs-lookup"><span data-stu-id="0d6f0-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="0d6f0-153">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-153">Member</span></span> |
| [<span data-ttu-id="0d6f0-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="0d6f0-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="0d6f0-155">Member</span><span class="sxs-lookup"><span data-stu-id="0d6f0-155">Member</span></span> |
| [<span data-ttu-id="0d6f0-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="0d6f0-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6f0-157">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-157">Member</span></span> |
| [<span data-ttu-id="0d6f0-158">sender</span><span class="sxs-lookup"><span data-stu-id="0d6f0-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="0d6f0-159">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-159">Member</span></span> |
| [<span data-ttu-id="0d6f0-160">Webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="0d6f0-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="0d6f0-161">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-161">Member</span></span> |
| [<span data-ttu-id="0d6f0-162">start</span><span class="sxs-lookup"><span data-stu-id="0d6f0-162">start</span></span>](#start-datetime) | <span data-ttu-id="0d6f0-163">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-163">Member</span></span> |
| [<span data-ttu-id="0d6f0-164">subject</span><span class="sxs-lookup"><span data-stu-id="0d6f0-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="0d6f0-165">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-165">Member</span></span> |
| [<span data-ttu-id="0d6f0-166">to</span><span class="sxs-lookup"><span data-stu-id="0d6f0-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="0d6f0-167">成员</span><span class="sxs-lookup"><span data-stu-id="0d6f0-167">Member</span></span> |
| [<span data-ttu-id="0d6f0-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="0d6f0-169">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-169">Method</span></span> |
| [<span data-ttu-id="0d6f0-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="0d6f0-171">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-171">Method</span></span> |
| [<span data-ttu-id="0d6f0-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="0d6f0-173">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-173">Method</span></span> |
| [<span data-ttu-id="0d6f0-174">close</span><span class="sxs-lookup"><span data-stu-id="0d6f0-174">close</span></span>](#close) | <span data-ttu-id="0d6f0-175">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-175">Method</span></span> |
| [<span data-ttu-id="0d6f0-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="0d6f0-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="0d6f0-177">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-177">Method</span></span> |
| [<span data-ttu-id="0d6f0-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="0d6f0-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="0d6f0-179">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-179">Method</span></span> |
| [<span data-ttu-id="0d6f0-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="0d6f0-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="0d6f0-181">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-181">Method</span></span> |
| [<span data-ttu-id="0d6f0-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0d6f0-183">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-183">Method</span></span> |
| [<span data-ttu-id="0d6f0-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="0d6f0-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="0d6f0-185">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-185">Method</span></span> |
| [<span data-ttu-id="0d6f0-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="0d6f0-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="0d6f0-187">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-187">Method</span></span> |
| [<span data-ttu-id="0d6f0-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="0d6f0-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="0d6f0-189">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-189">Method</span></span> |
| [<span data-ttu-id="0d6f0-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="0d6f0-191">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-191">Method</span></span> |
| [<span data-ttu-id="0d6f0-192">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="0d6f0-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="0d6f0-193">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-193">Method</span></span> |
| [<span data-ttu-id="0d6f0-194">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="0d6f0-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="0d6f0-195">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-195">Method</span></span> |
| [<span data-ttu-id="0d6f0-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="0d6f0-197">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-197">Method</span></span> |
| [<span data-ttu-id="0d6f0-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="0d6f0-199">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-199">Method</span></span> |
| [<span data-ttu-id="0d6f0-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="0d6f0-201">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-201">Method</span></span> |
| [<span data-ttu-id="0d6f0-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="0d6f0-203">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-203">Method</span></span> |
| [<span data-ttu-id="0d6f0-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="0d6f0-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="0d6f0-205">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="0d6f0-206">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-206">Example</span></span>

<span data-ttu-id="0d6f0-207">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="0d6f0-208">Members</span><span class="sxs-lookup"><span data-stu-id="0d6f0-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="0d6f0-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="0d6f0-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="0d6f0-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-212">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="0d6f0-213">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-214">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-214">Type</span></span>

*   <span data-ttu-id="0d6f0-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="0d6f0-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-216">Requirements</span></span>

|<span data-ttu-id="0d6f0-217">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-217">Requirement</span></span>|<span data-ttu-id="0d6f0-218">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-220">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-220">1.0</span></span>|
|[<span data-ttu-id="0d6f0-221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-222">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-224">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-225">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-225">Example</span></span>

<span data-ttu-id="0d6f0-226">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="0d6f0-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-228">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="0d6f0-229">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-229">Compose mode only.</span></span>

<span data-ttu-id="0d6f0-230">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-231">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0d6f0-232">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="0d6f0-233">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-234">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-234">Type</span></span>

*   [<span data-ttu-id="0d6f0-235">收件人</span><span class="sxs-lookup"><span data-stu-id="0d6f0-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="0d6f0-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-236">Requirements</span></span>

|<span data-ttu-id="0d6f0-237">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-237">Requirement</span></span>|<span data-ttu-id="0d6f0-238">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-239">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-240">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6f0-240">1.1</span></span>|
|[<span data-ttu-id="0d6f0-241">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-242">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-243">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-244">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-245">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="0d6f0-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-247">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-248">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-248">Type</span></span>

*   [<span data-ttu-id="0d6f0-249">Body</span><span class="sxs-lookup"><span data-stu-id="0d6f0-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="0d6f0-250">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-250">Requirements</span></span>

|<span data-ttu-id="0d6f0-251">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-251">Requirement</span></span>|<span data-ttu-id="0d6f0-252">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-254">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6f0-254">1.1</span></span>|
|[<span data-ttu-id="0d6f0-255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-256">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-258">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-259">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-259">Example</span></span>

<span data-ttu-id="0d6f0-260">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="0d6f0-261">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="0d6f0-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-263">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="0d6f0-264">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-265">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-265">Read mode</span></span>

<span data-ttu-id="0d6f0-266">`cc` 属性返回包含邮件的`EmailAddressDetails`行上所列的每个收件人的 \*\*\*\* 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="0d6f0-267">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-268">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-269">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-269">Compose mode</span></span>

<span data-ttu-id="0d6f0-270">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="0d6f0-271">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-272">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0d6f0-273">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="0d6f0-274">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-275">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-275">Type</span></span>

*   <span data-ttu-id="0d6f0-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-277">Requirements</span></span>

|<span data-ttu-id="0d6f0-278">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-278">Requirement</span></span>|<span data-ttu-id="0d6f0-279">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-281">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-281">1.0</span></span>|
|[<span data-ttu-id="0d6f0-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-283">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-285">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="0d6f0-286">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="0d6f0-287">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="0d6f0-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="0d6f0-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-292">Type</span><span class="sxs-lookup"><span data-stu-id="0d6f0-292">Type</span></span>

*   <span data-ttu-id="0d6f0-293">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-294">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-294">Requirements</span></span>

|<span data-ttu-id="0d6f0-295">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-295">Requirement</span></span>|<span data-ttu-id="0d6f0-296">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-297">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-298">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-298">1.0</span></span>|
|[<span data-ttu-id="0d6f0-299">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-300">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-301">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-302">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-303">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="0d6f0-304">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="0d6f0-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="0d6f0-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-307">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-307">Type</span></span>

*   <span data-ttu-id="0d6f0-308">日期</span><span class="sxs-lookup"><span data-stu-id="0d6f0-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-309">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-309">Requirements</span></span>

|<span data-ttu-id="0d6f0-310">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-310">Requirement</span></span>|<span data-ttu-id="0d6f0-311">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-312">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-313">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-313">1.0</span></span>|
|[<span data-ttu-id="0d6f0-314">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-315">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-316">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-317">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-318">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="0d6f0-319">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="0d6f0-319">dateTimeModified: Date</span></span>

<span data-ttu-id="0d6f0-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-322">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-323">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-323">Type</span></span>

*   <span data-ttu-id="0d6f0-324">日期</span><span class="sxs-lookup"><span data-stu-id="0d6f0-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-325">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-325">Requirements</span></span>

|<span data-ttu-id="0d6f0-326">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-326">Requirement</span></span>|<span data-ttu-id="0d6f0-327">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-328">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-329">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-329">1.0</span></span>|
|[<span data-ttu-id="0d6f0-330">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-331">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-332">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-333">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-334">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="0d6f0-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-336">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="0d6f0-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-339">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-339">Read mode</span></span>

<span data-ttu-id="0d6f0-340">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-341">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-341">Compose mode</span></span>

<span data-ttu-id="0d6f0-342">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="0d6f0-343">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0d6f0-344">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0d6f0-345">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-345">Type</span></span>

*   <span data-ttu-id="0d6f0-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-347">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-347">Requirements</span></span>

|<span data-ttu-id="0d6f0-348">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-348">Requirement</span></span>|<span data-ttu-id="0d6f0-349">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-350">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-351">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-351">1.0</span></span>|
|[<span data-ttu-id="0d6f0-352">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-353">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-354">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-355">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="0d6f0-356">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-357">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="0d6f0-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-360">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-361">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-361">Read mode</span></span>

<span data-ttu-id="0d6f0-362">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-363">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-363">Compose mode</span></span>

<span data-ttu-id="0d6f0-364">`from`属性返回一个`From`对象，该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-365">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-365">Type</span></span>

*   <span data-ttu-id="0d6f0-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-367">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-367">Requirements</span></span>

|<span data-ttu-id="0d6f0-368">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0d6f0-369">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-370">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-370">1.0</span></span>|<span data-ttu-id="0d6f0-371">1.7</span><span class="sxs-lookup"><span data-stu-id="0d6f0-371">1.7</span></span>|
|[<span data-ttu-id="0d6f0-372">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-373">ReadItem</span></span>|<span data-ttu-id="0d6f0-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6f0-375">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-376">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-376">Read</span></span>|<span data-ttu-id="0d6f0-377">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="0d6f0-378">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-378">internetMessageId: String</span></span>

<span data-ttu-id="0d6f0-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-381">Type</span><span class="sxs-lookup"><span data-stu-id="0d6f0-381">Type</span></span>

*   <span data-ttu-id="0d6f0-382">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-383">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-383">Requirements</span></span>

|<span data-ttu-id="0d6f0-384">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-384">Requirement</span></span>|<span data-ttu-id="0d6f0-385">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-387">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-387">1.0</span></span>|
|[<span data-ttu-id="0d6f0-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-389">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-391">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-392">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="0d6f0-393">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-393">itemClass: String</span></span>

<span data-ttu-id="0d6f0-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="0d6f0-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="0d6f0-398">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-398">Type</span></span>|<span data-ttu-id="0d6f0-399">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-399">Description</span></span>|<span data-ttu-id="0d6f0-400">项目类</span><span class="sxs-lookup"><span data-stu-id="0d6f0-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="0d6f0-401">约会项目</span><span class="sxs-lookup"><span data-stu-id="0d6f0-401">Appointment items</span></span>|<span data-ttu-id="0d6f0-402">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="0d6f0-403">邮件项目</span><span class="sxs-lookup"><span data-stu-id="0d6f0-403">Message items</span></span>|<span data-ttu-id="0d6f0-404">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="0d6f0-405">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-406">Type</span><span class="sxs-lookup"><span data-stu-id="0d6f0-406">Type</span></span>

*   <span data-ttu-id="0d6f0-407">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-408">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-408">Requirements</span></span>

|<span data-ttu-id="0d6f0-409">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-409">Requirement</span></span>|<span data-ttu-id="0d6f0-410">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-412">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-412">1.0</span></span>|
|[<span data-ttu-id="0d6f0-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-414">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-416">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-417">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="0d6f0-418">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-418">(nullable) itemId: String</span></span>

<span data-ttu-id="0d6f0-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-421">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-421">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0d6f0-422">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="0d6f0-423">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0d6f0-424">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="0d6f0-p120">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-427">Type</span><span class="sxs-lookup"><span data-stu-id="0d6f0-427">Type</span></span>

*   <span data-ttu-id="0d6f0-428">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-429">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-429">Requirements</span></span>

|<span data-ttu-id="0d6f0-430">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-430">Requirement</span></span>|<span data-ttu-id="0d6f0-431">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-432">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-433">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-433">1.0</span></span>|
|[<span data-ttu-id="0d6f0-434">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-435">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-436">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-437">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-438">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-438">Example</span></span>

<span data-ttu-id="0d6f0-p121">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="0d6f0-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-442">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="0d6f0-443">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-444">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-444">Type</span></span>

*   [<span data-ttu-id="0d6f0-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="0d6f0-446">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-446">Requirements</span></span>

|<span data-ttu-id="0d6f0-447">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-447">Requirement</span></span>|<span data-ttu-id="0d6f0-448">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-449">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-450">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-450">1.0</span></span>|
|[<span data-ttu-id="0d6f0-451">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-452">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-453">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-454">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-455">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="0d6f0-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-457">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-458">Read mode</span></span>

<span data-ttu-id="0d6f0-459">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-460">Compose mode</span></span>

<span data-ttu-id="0d6f0-461">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-462">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-462">Type</span></span>

*   <span data-ttu-id="0d6f0-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-464">Requirements</span></span>

|<span data-ttu-id="0d6f0-465">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-465">Requirement</span></span>|<span data-ttu-id="0d6f0-466">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-468">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-468">1.0</span></span>|
|[<span data-ttu-id="0d6f0-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-470">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-472">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="0d6f0-473">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-473">normalizedSubject: String</span></span>

<span data-ttu-id="0d6f0-p122">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="0d6f0-p123">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-478">Type</span><span class="sxs-lookup"><span data-stu-id="0d6f0-478">Type</span></span>

*   <span data-ttu-id="0d6f0-479">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-480">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-480">Requirements</span></span>

|<span data-ttu-id="0d6f0-481">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-481">Requirement</span></span>|<span data-ttu-id="0d6f0-482">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-483">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-484">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-484">1.0</span></span>|
|[<span data-ttu-id="0d6f0-485">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-486">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-487">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-488">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-489">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="0d6f0-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-491">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-492">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-492">Type</span></span>

*   [<span data-ttu-id="0d6f0-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="0d6f0-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="0d6f0-494">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-494">Requirements</span></span>

|<span data-ttu-id="0d6f0-495">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-495">Requirement</span></span>|<span data-ttu-id="0d6f0-496">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-497">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-498">1.3</span><span class="sxs-lookup"><span data-stu-id="0d6f0-498">1.3</span></span>|
|[<span data-ttu-id="0d6f0-499">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-500">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-501">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-502">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-503">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="0d6f0-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-505">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="0d6f0-506">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-507">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-507">Read mode</span></span>

<span data-ttu-id="0d6f0-508">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="0d6f0-509">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-510">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-511">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-511">Compose mode</span></span>

<span data-ttu-id="0d6f0-512">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="0d6f0-513">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-514">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0d6f0-515">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="0d6f0-516">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-517">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-517">Type</span></span>

*   <span data-ttu-id="0d6f0-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-519">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-519">Requirements</span></span>

|<span data-ttu-id="0d6f0-520">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-520">Requirement</span></span>|<span data-ttu-id="0d6f0-521">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-522">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-523">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-523">1.0</span></span>|
|[<span data-ttu-id="0d6f0-524">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-525">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-526">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-527">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="0d6f0-528">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[组织者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-529">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-530">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-530">Read mode</span></span>

<span data-ttu-id="0d6f0-531">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)对象，该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-532">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-532">Compose mode</span></span>

<span data-ttu-id="0d6f0-533">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)器对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="0d6f0-534">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-534">Type</span></span>

*   <span data-ttu-id="0d6f0-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [组织者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-536">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-536">Requirements</span></span>

|<span data-ttu-id="0d6f0-537">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="0d6f0-538">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-539">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-539">1.0</span></span>|<span data-ttu-id="0d6f0-540">1.7</span><span class="sxs-lookup"><span data-stu-id="0d6f0-540">1.7</span></span>|
|[<span data-ttu-id="0d6f0-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-542">ReadItem</span></span>|<span data-ttu-id="0d6f0-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6f0-544">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-545">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-545">Read</span></span>|<span data-ttu-id="0d6f0-546">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="0d6f0-547">（可以为 null）定期：[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-548">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="0d6f0-549">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="0d6f0-550">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="0d6f0-551">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="0d6f0-552">如果`recurrence`项目是系列中的一个系列或一个实例，则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="0d6f0-553">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="0d6f0-554">`undefined`对于不是会议请求的邮件，将返回。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="0d6f0-555">注意：会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="0d6f0-556">注意：如果定期对象为`null`，则表示该对象是单个约会的单个约会或会议请求，而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-557">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-557">Read mode</span></span>

<span data-ttu-id="0d6f0-558">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="0d6f0-559">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-560">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-560">Compose mode</span></span>

<span data-ttu-id="0d6f0-561">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)对象，该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="0d6f0-562">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-562">This is available for appointments.</span></span>

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-563">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-563">Type</span></span>

* [<span data-ttu-id="0d6f0-564">循环</span><span class="sxs-lookup"><span data-stu-id="0d6f0-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="0d6f0-565">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-565">Requirement</span></span>|<span data-ttu-id="0d6f0-566">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-567">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-568">1.7</span><span class="sxs-lookup"><span data-stu-id="0d6f0-568">1.7</span></span>|
|[<span data-ttu-id="0d6f0-569">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-570">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-571">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-572">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="0d6f0-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-574">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="0d6f0-575">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-576">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-576">Read mode</span></span>

<span data-ttu-id="0d6f0-577">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="0d6f0-578">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-579">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-580">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-580">Compose mode</span></span>

<span data-ttu-id="0d6f0-581">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="0d6f0-582">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-583">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0d6f0-584">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="0d6f0-585">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-586">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-586">Type</span></span>

*   <span data-ttu-id="0d6f0-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-588">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-588">Requirements</span></span>

|<span data-ttu-id="0d6f0-589">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-589">Requirement</span></span>|<span data-ttu-id="0d6f0-590">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-591">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-592">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-592">1.0</span></span>|
|[<span data-ttu-id="0d6f0-593">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-594">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-595">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-596">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="0d6f0-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-p134">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="0d6f0-p135">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-602">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-603">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-603">Type</span></span>

*   [<span data-ttu-id="0d6f0-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="0d6f0-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="0d6f0-605">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-605">Requirements</span></span>

|<span data-ttu-id="0d6f0-606">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-606">Requirement</span></span>|<span data-ttu-id="0d6f0-607">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-608">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-609">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-609">1.0</span></span>|
|[<span data-ttu-id="0d6f0-610">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-611">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-612">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-613">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-614">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="0d6f0-615">（可以为 null） Webcasts&seriesid： String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="0d6f0-616">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="0d6f0-617">在 web 上的 Outlook 和桌面客户端中`seriesId` ，返回此项所属的父（系列）项的 Exchange web 服务（EWS） ID。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="0d6f0-618">但是，在 iOS 和 Android 中， `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-619">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="0d6f0-620">`seriesId`属性与 OUTLOOK REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="0d6f0-621">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="0d6f0-622">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="0d6f0-623">对于`seriesId`不包含`null`父项（如单个约会、系列项或会议请求）的项，该属性将返回， `undefined`对于不是会议请求的任何其他项，该属性返回。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="0d6f0-624">Type</span><span class="sxs-lookup"><span data-stu-id="0d6f0-624">Type</span></span>

* <span data-ttu-id="0d6f0-625">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-626">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-626">Requirements</span></span>

|<span data-ttu-id="0d6f0-627">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-627">Requirement</span></span>|<span data-ttu-id="0d6f0-628">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-629">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-630">1.7</span><span class="sxs-lookup"><span data-stu-id="0d6f0-630">1.7</span></span>|
|[<span data-ttu-id="0d6f0-631">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-632">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-633">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-634">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-635">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-635">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="0d6f0-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-637">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="0d6f0-p138">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-640">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-640">Read mode</span></span>

<span data-ttu-id="0d6f0-641">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-642">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-642">Compose mode</span></span>

<span data-ttu-id="0d6f0-643">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="0d6f0-644">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="0d6f0-645">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="0d6f0-646">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-646">Type</span></span>

*   <span data-ttu-id="0d6f0-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-648">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-648">Requirements</span></span>

|<span data-ttu-id="0d6f0-649">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-649">Requirement</span></span>|<span data-ttu-id="0d6f0-650">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-651">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-652">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-652">1.0</span></span>|
|[<span data-ttu-id="0d6f0-653">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-654">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-655">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-656">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="0d6f0-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-658">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="0d6f0-659">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-660">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-660">Read mode</span></span>

<span data-ttu-id="0d6f0-p139">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="0d6f0-663">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-664">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-664">Compose mode</span></span>

<span data-ttu-id="0d6f0-665">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-666">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-666">Type</span></span>

*   <span data-ttu-id="0d6f0-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-668">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-668">Requirements</span></span>

|<span data-ttu-id="0d6f0-669">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-669">Requirement</span></span>|<span data-ttu-id="0d6f0-670">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-671">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-672">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-672">1.0</span></span>|
|[<span data-ttu-id="0d6f0-673">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-674">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-675">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-676">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="0d6f0-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="0d6f0-678">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="0d6f0-679">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="0d6f0-680">阅读模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-680">Read mode</span></span>

<span data-ttu-id="0d6f0-681">`to` 属性返回包含邮件的`EmailAddressDetails`行上所列的每个收件人的 \*\*\*\* 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="0d6f0-682">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-683">但是，在 Windows 和 Mac 上，您可以获得500个成员的最大值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="0d6f0-684">撰写模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-684">Compose mode</span></span>

<span data-ttu-id="0d6f0-685">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="0d6f0-686">默认情况下，集合限制为最多为100个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="0d6f0-687">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="0d6f0-688">最多获取500个成员。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="0d6f0-689">每个呼叫最多可设置100个成员，最多为500个成员总数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="0d6f0-690">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-690">Type</span></span>

*   <span data-ttu-id="0d6f0-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-692">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-692">Requirements</span></span>

|<span data-ttu-id="0d6f0-693">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-693">Requirement</span></span>|<span data-ttu-id="0d6f0-694">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-695">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-696">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-696">1.0</span></span>|
|[<span data-ttu-id="0d6f0-697">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-698">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-699">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-700">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="0d6f0-701">方法</span><span class="sxs-lookup"><span data-stu-id="0d6f0-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="0d6f0-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0d6f0-703">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="0d6f0-704">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="0d6f0-705">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-706">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-706">Parameters</span></span>
|<span data-ttu-id="0d6f0-707">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-707">Name</span></span>|<span data-ttu-id="0d6f0-708">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-708">Type</span></span>|<span data-ttu-id="0d6f0-709">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-709">Attributes</span></span>|<span data-ttu-id="0d6f0-710">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="0d6f0-711">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-711">String</span></span>||<span data-ttu-id="0d6f0-p143">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0d6f0-714">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-714">String</span></span>||<span data-ttu-id="0d6f0-p144">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0d6f0-717">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-717">Object</span></span>|<span data-ttu-id="0d6f0-718">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-718">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-719">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0d6f0-720">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-720">Object</span></span>|<span data-ttu-id="0d6f0-721">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-721">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-722">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="0d6f0-723">布尔值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-723">Boolean</span></span>|<span data-ttu-id="0d6f0-724">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-724">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-725">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-726">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-726">function</span></span>|<span data-ttu-id="0d6f0-727">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-727">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-728">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0d6f0-729">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0d6f0-730">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0d6f0-731">错误</span><span class="sxs-lookup"><span data-stu-id="0d6f0-731">Errors</span></span>

|<span data-ttu-id="0d6f0-732">错误代码</span><span class="sxs-lookup"><span data-stu-id="0d6f0-732">Error code</span></span>|<span data-ttu-id="0d6f0-733">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="0d6f0-734">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="0d6f0-735">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0d6f0-736">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-737">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-737">Requirements</span></span>

|<span data-ttu-id="0d6f0-738">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-738">Requirement</span></span>|<span data-ttu-id="0d6f0-739">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-740">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-741">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6f0-741">1.1</span></span>|
|[<span data-ttu-id="0d6f0-742">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6f0-744">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-745">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0d6f0-746">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-746">Examples</span></span>

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

<span data-ttu-id="0d6f0-747">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="0d6f0-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="0d6f0-749">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="0d6f0-750">目前，受支持的事件`Office.EventType.AppointmentTimeChanged`类型`Office.EventType.RecipientsChanged`是、和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="0d6f0-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-751">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-751">Parameters</span></span>

| <span data-ttu-id="0d6f0-752">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-752">Name</span></span> | <span data-ttu-id="0d6f0-753">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-753">Type</span></span> | <span data-ttu-id="0d6f0-754">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-754">Attributes</span></span> | <span data-ttu-id="0d6f0-755">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0d6f0-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0d6f0-757">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="0d6f0-758">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-758">Function</span></span> || <span data-ttu-id="0d6f0-p145">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="0d6f0-762">Object</span><span class="sxs-lookup"><span data-stu-id="0d6f0-762">Object</span></span> | <span data-ttu-id="0d6f0-763">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-763">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6f0-764">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0d6f0-765">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-765">Object</span></span> | <span data-ttu-id="0d6f0-766">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-766">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6f0-767">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0d6f0-768">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-768">function</span></span>| <span data-ttu-id="0d6f0-769">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-769">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-770">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-771">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-771">Requirements</span></span>

|<span data-ttu-id="0d6f0-772">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-772">Requirement</span></span>| <span data-ttu-id="0d6f0-773">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-774">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6f0-775">1.7</span><span class="sxs-lookup"><span data-stu-id="0d6f0-775">1.7</span></span> |
|[<span data-ttu-id="0d6f0-776">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6f0-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-777">ReadItem</span></span> |
|[<span data-ttu-id="0d6f0-778">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6f0-779">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="0d6f0-780">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-780">Example</span></span>

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="0d6f0-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="0d6f0-782">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="0d6f0-p146">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="0d6f0-786">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="0d6f0-787">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-788">Parameters</span><span class="sxs-lookup"><span data-stu-id="0d6f0-788">Parameters</span></span>

|<span data-ttu-id="0d6f0-789">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-789">Name</span></span>|<span data-ttu-id="0d6f0-790">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-790">Type</span></span>|<span data-ttu-id="0d6f0-791">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-791">Attributes</span></span>|<span data-ttu-id="0d6f0-792">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="0d6f0-793">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-793">String</span></span>||<span data-ttu-id="0d6f0-p147">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="0d6f0-796">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-796">String</span></span>||<span data-ttu-id="0d6f0-797">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-797">The subject of the item to be attached.</span></span> <span data-ttu-id="0d6f0-798">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="0d6f0-799">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-799">Object</span></span>|<span data-ttu-id="0d6f0-800">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-800">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-801">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0d6f0-802">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-802">Object</span></span>|<span data-ttu-id="0d6f0-803">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-803">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-804">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-805">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-805">function</span></span>|<span data-ttu-id="0d6f0-806">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-806">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-807">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0d6f0-808">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="0d6f0-809">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0d6f0-810">错误</span><span class="sxs-lookup"><span data-stu-id="0d6f0-810">Errors</span></span>

|<span data-ttu-id="0d6f0-811">错误代码</span><span class="sxs-lookup"><span data-stu-id="0d6f0-811">Error code</span></span>|<span data-ttu-id="0d6f0-812">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="0d6f0-813">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-814">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-814">Requirements</span></span>

|<span data-ttu-id="0d6f0-815">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-815">Requirement</span></span>|<span data-ttu-id="0d6f0-816">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-817">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-818">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6f0-818">1.1</span></span>|
|[<span data-ttu-id="0d6f0-819">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6f0-821">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-822">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-823">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-823">Example</span></span>

<span data-ttu-id="0d6f0-824">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="0d6f0-825">close()</span><span class="sxs-lookup"><span data-stu-id="0d6f0-825">close()</span></span>

<span data-ttu-id="0d6f0-826">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="0d6f0-p149">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-829">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="0d6f0-830">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-831">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-831">Requirements</span></span>

|<span data-ttu-id="0d6f0-832">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-832">Requirement</span></span>|<span data-ttu-id="0d6f0-833">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-834">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-835">1.3</span><span class="sxs-lookup"><span data-stu-id="0d6f0-835">1.3</span></span>|
|[<span data-ttu-id="0d6f0-836">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-837">受限</span><span class="sxs-lookup"><span data-stu-id="0d6f0-837">Restricted</span></span>|
|[<span data-ttu-id="0d6f0-838">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-839">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="0d6f0-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="0d6f0-841">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-842">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6f0-843">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0d6f0-844">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="0d6f0-p150">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-848">Parameters</span><span class="sxs-lookup"><span data-stu-id="0d6f0-848">Parameters</span></span>

|<span data-ttu-id="0d6f0-849">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-849">Name</span></span>|<span data-ttu-id="0d6f0-850">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-850">Type</span></span>|<span data-ttu-id="0d6f0-851">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-851">Attributes</span></span>|<span data-ttu-id="0d6f0-852">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0d6f0-853">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-853">String &#124; Object</span></span>||<span data-ttu-id="0d6f0-p151">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0d6f0-856">**或**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-856">**OR**</span></span><br/><span data-ttu-id="0d6f0-p152">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0d6f0-859">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-859">String</span></span>|<span data-ttu-id="0d6f0-860">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-860">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-p153">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0d6f0-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0d6f0-864">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-864">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-865">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0d6f0-866">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-866">String</span></span>||<span data-ttu-id="0d6f0-p154">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0d6f0-869">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-869">String</span></span>||<span data-ttu-id="0d6f0-870">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0d6f0-871">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-871">String</span></span>||<span data-ttu-id="0d6f0-p155">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0d6f0-874">布尔</span><span class="sxs-lookup"><span data-stu-id="0d6f0-874">Boolean</span></span>||<span data-ttu-id="0d6f0-p156">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0d6f0-877">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-877">String</span></span>||<span data-ttu-id="0d6f0-p157">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-881">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-881">function</span></span>|<span data-ttu-id="0d6f0-882">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-882">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-883">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-884">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-884">Requirements</span></span>

|<span data-ttu-id="0d6f0-885">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-885">Requirement</span></span>|<span data-ttu-id="0d6f0-886">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-887">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-888">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-888">1.0</span></span>|
|[<span data-ttu-id="0d6f0-889">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-890">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-891">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-892">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0d6f0-893">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-893">Examples</span></span>

<span data-ttu-id="0d6f0-894">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="0d6f0-895">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="0d6f0-896">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0d6f0-897">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0d6f0-898">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0d6f0-899">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="0d6f0-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="0d6f0-901">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-902">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6f0-903">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="0d6f0-904">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="0d6f0-p158">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-908">Parameters</span><span class="sxs-lookup"><span data-stu-id="0d6f0-908">Parameters</span></span>

|<span data-ttu-id="0d6f0-909">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-909">Name</span></span>|<span data-ttu-id="0d6f0-910">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-910">Type</span></span>|<span data-ttu-id="0d6f0-911">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-911">Attributes</span></span>|<span data-ttu-id="0d6f0-912">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="0d6f0-913">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-913">String &#124; Object</span></span>||<span data-ttu-id="0d6f0-p159">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="0d6f0-916">**或**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-916">**OR**</span></span><br/><span data-ttu-id="0d6f0-p160">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="0d6f0-919">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-919">String</span></span>|<span data-ttu-id="0d6f0-920">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-920">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-p161">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="0d6f0-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="0d6f0-924">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-924">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-925">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="0d6f0-926">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-926">String</span></span>||<span data-ttu-id="0d6f0-p162">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="0d6f0-929">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-929">String</span></span>||<span data-ttu-id="0d6f0-930">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="0d6f0-931">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-931">String</span></span>||<span data-ttu-id="0d6f0-p163">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="0d6f0-934">布尔</span><span class="sxs-lookup"><span data-stu-id="0d6f0-934">Boolean</span></span>||<span data-ttu-id="0d6f0-p164">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="0d6f0-937">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-937">String</span></span>||<span data-ttu-id="0d6f0-p165">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-941">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-941">function</span></span>|<span data-ttu-id="0d6f0-942">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-942">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-943">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-944">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-944">Requirements</span></span>

|<span data-ttu-id="0d6f0-945">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-945">Requirement</span></span>|<span data-ttu-id="0d6f0-946">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-947">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-948">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-948">1.0</span></span>|
|[<span data-ttu-id="0d6f0-949">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-950">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-951">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-952">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="0d6f0-953">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-953">Examples</span></span>

<span data-ttu-id="0d6f0-954">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="0d6f0-955">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="0d6f0-956">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="0d6f0-957">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="0d6f0-958">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="0d6f0-959">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="0d6f0-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="0d6f0-961">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-962">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-963">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-963">Requirements</span></span>

|<span data-ttu-id="0d6f0-964">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-964">Requirement</span></span>|<span data-ttu-id="0d6f0-965">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-966">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-967">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-967">1.0</span></span>|
|[<span data-ttu-id="0d6f0-968">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-969">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-970">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-971">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-972">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-972">Returns:</span></span>

<span data-ttu-id="0d6f0-973">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="0d6f0-974">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-974">Example</span></span>

<span data-ttu-id="0d6f0-975">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="0d6f0-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="0d6f0-977">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-978">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-979">Parameters</span><span class="sxs-lookup"><span data-stu-id="0d6f0-979">Parameters</span></span>

|<span data-ttu-id="0d6f0-980">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-980">Name</span></span>|<span data-ttu-id="0d6f0-981">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-981">Type</span></span>|<span data-ttu-id="0d6f0-982">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="0d6f0-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="0d6f0-984">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-985">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-985">Requirements</span></span>

|<span data-ttu-id="0d6f0-986">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-986">Requirement</span></span>|<span data-ttu-id="0d6f0-987">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-988">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-989">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-989">1.0</span></span>|
|[<span data-ttu-id="0d6f0-990">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-991">受限</span><span class="sxs-lookup"><span data-stu-id="0d6f0-991">Restricted</span></span>|
|[<span data-ttu-id="0d6f0-992">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-993">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-994">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-994">Returns:</span></span>

<span data-ttu-id="0d6f0-995">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="0d6f0-996">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="0d6f0-997">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="0d6f0-998">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="0d6f0-999">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-999">Value of `entityType`</span></span>|<span data-ttu-id="0d6f0-1000">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1000">Type of objects in returned array</span></span>|<span data-ttu-id="0d6f0-1001">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="0d6f0-1002">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1002">String</span></span>|<span data-ttu-id="0d6f0-1003">**受限**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="0d6f0-1004">Contact</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1004">Contact</span></span>|<span data-ttu-id="0d6f0-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="0d6f0-1006">String</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1006">String</span></span>|<span data-ttu-id="0d6f0-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="0d6f0-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1008">MeetingSuggestion</span></span>|<span data-ttu-id="0d6f0-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="0d6f0-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1010">PhoneNumber</span></span>|<span data-ttu-id="0d6f0-1011">**受限**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="0d6f0-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1012">TaskSuggestion</span></span>|<span data-ttu-id="0d6f0-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="0d6f0-1014">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1014">String</span></span>|<span data-ttu-id="0d6f0-1015">**受限**</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1015">**Restricted**</span></span>|

<span data-ttu-id="0d6f0-1016">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="0d6f0-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="0d6f0-1017">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1017">Example</span></span>

<span data-ttu-id="0d6f0-1018">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="0d6f0-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="0d6f0-1020">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-1021">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6f0-1022">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1023">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1023">Parameters</span></span>

|<span data-ttu-id="0d6f0-1024">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1024">Name</span></span>|<span data-ttu-id="0d6f0-1025">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1025">Type</span></span>|<span data-ttu-id="0d6f0-1026">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0d6f0-1027">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1027">String</span></span>|<span data-ttu-id="0d6f0-1028">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1029">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1029">Requirements</span></span>

|<span data-ttu-id="0d6f0-1030">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1030">Requirement</span></span>|<span data-ttu-id="0d6f0-1031">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1032">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1033">1.0</span></span>|
|[<span data-ttu-id="0d6f0-1034">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1035">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-1036">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1037">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-1038">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1038">Returns:</span></span>

<span data-ttu-id="0d6f0-p167">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="0d6f0-1041">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="0d6f0-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="0d6f0-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="0d6f0-1043">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-1044">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6f0-p168">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0d6f0-1048">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0d6f0-1049">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0d6f0-p169">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1053">Requirements</span></span>

|<span data-ttu-id="0d6f0-1054">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1054">Requirement</span></span>|<span data-ttu-id="0d6f0-1055">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1056">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1057">1.0</span></span>|
|[<span data-ttu-id="0d6f0-1058">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1059">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-1060">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1061">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-1062">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1062">Returns:</span></span>

<span data-ttu-id="0d6f0-p170">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="0d6f0-1065">类型：对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="0d6f0-1066">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1066">Example</span></span>

<span data-ttu-id="0d6f0-1067">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="0d6f0-1068">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="0d6f0-1069">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-1070">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6f0-1071">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="0d6f0-p171">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1074">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1074">Parameters</span></span>

|<span data-ttu-id="0d6f0-1075">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1075">Name</span></span>|<span data-ttu-id="0d6f0-1076">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1076">Type</span></span>|<span data-ttu-id="0d6f0-1077">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="0d6f0-1078">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1078">String</span></span>|<span data-ttu-id="0d6f0-1079">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1080">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1080">Requirements</span></span>

|<span data-ttu-id="0d6f0-1081">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1081">Requirement</span></span>|<span data-ttu-id="0d6f0-1082">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1083">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1084">1.0</span></span>|
|[<span data-ttu-id="0d6f0-1085">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1086">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-1087">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1088">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-1089">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1089">Returns:</span></span>

<span data-ttu-id="0d6f0-1090">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="0d6f0-1091">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="0d6f0-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="0d6f0-1092">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="0d6f0-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="0d6f0-1094">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="0d6f0-p172">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1097">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1097">Parameters</span></span>

|<span data-ttu-id="0d6f0-1098">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1098">Name</span></span>|<span data-ttu-id="0d6f0-1099">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1099">Type</span></span>|<span data-ttu-id="0d6f0-1100">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1100">Attributes</span></span>|<span data-ttu-id="0d6f0-1101">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1101">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="0d6f0-1102">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1102">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="0d6f0-p173">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="0d6f0-1106">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1106">Object</span></span>|<span data-ttu-id="0d6f0-1107">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1107">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1108">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1108">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0d6f0-1109">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1109">Object</span></span>|<span data-ttu-id="0d6f0-1110">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1110">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1111">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1111">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-1112">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1112">function</span></span>||<span data-ttu-id="0d6f0-1113">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1113">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0d6f0-1114">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1114">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="0d6f0-1115">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1115">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1116">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1116">Requirements</span></span>

|<span data-ttu-id="0d6f0-1117">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1117">Requirement</span></span>|<span data-ttu-id="0d6f0-1118">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1118">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1119">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1119">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1120">1.2</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1120">1.2</span></span>|
|[<span data-ttu-id="0d6f0-1121">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1121">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1122">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-1123">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1123">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1124">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1124">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-1125">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1125">Returns:</span></span>

<span data-ttu-id="0d6f0-1126">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1126">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="0d6f0-1127">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1127">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="0d6f0-1128">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1128">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="0d6f0-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1129">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="0d6f0-1130">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1130">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="0d6f0-1131">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1131">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-1132">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1132">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1133">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1133">Requirements</span></span>

|<span data-ttu-id="0d6f0-1134">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1134">Requirement</span></span>|<span data-ttu-id="0d6f0-1135">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1136">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1137">1.6</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1137">1.6</span></span>|
|[<span data-ttu-id="0d6f0-1138">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1139">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1139">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-1140">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1141">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-1142">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1142">Returns:</span></span>

<span data-ttu-id="0d6f0-1143">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1143">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="0d6f0-1144">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1144">Example</span></span>

<span data-ttu-id="0d6f0-1145">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1145">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="0d6f0-1146">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1146">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="0d6f0-p176">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-1149">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1149">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="0d6f0-p177">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="0d6f0-1153">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1153">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="0d6f0-1154">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1154">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="0d6f0-p178">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1158">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1158">Requirements</span></span>

|<span data-ttu-id="0d6f0-1159">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1159">Requirement</span></span>|<span data-ttu-id="0d6f0-1160">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1160">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1161">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1162">1.6</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1162">1.6</span></span>|
|[<span data-ttu-id="0d6f0-1163">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1164">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-1165">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1166">阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1166">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="0d6f0-1167">返回：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1167">Returns:</span></span>

<span data-ttu-id="0d6f0-p179">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="0d6f0-1170">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1170">Example</span></span>

<span data-ttu-id="0d6f0-1171">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1171">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="0d6f0-1172">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1172">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="0d6f0-1173">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1173">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="0d6f0-p180">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p180">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1177">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1177">Parameters</span></span>

|<span data-ttu-id="0d6f0-1178">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1178">Name</span></span>|<span data-ttu-id="0d6f0-1179">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1179">Type</span></span>|<span data-ttu-id="0d6f0-1180">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1180">Attributes</span></span>|<span data-ttu-id="0d6f0-1181">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1181">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="0d6f0-1182">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1182">function</span></span>||<span data-ttu-id="0d6f0-1183">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1183">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0d6f0-1184">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1184">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="0d6f0-1185">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1185">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="0d6f0-1186">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1186">Object</span></span>|<span data-ttu-id="0d6f0-1187">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1188">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1188">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="0d6f0-1189">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1189">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1190">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1190">Requirements</span></span>

|<span data-ttu-id="0d6f0-1191">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1191">Requirement</span></span>|<span data-ttu-id="0d6f0-1192">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1192">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1193">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1193">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1194">1.0</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1194">1.0</span></span>|
|[<span data-ttu-id="0d6f0-1195">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1195">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1196">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1196">ReadItem</span></span>|
|[<span data-ttu-id="0d6f0-1197">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1198">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1198">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-1199">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1199">Example</span></span>

<span data-ttu-id="0d6f0-p183">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p183">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="0d6f0-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1203">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="0d6f0-1204">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1204">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="0d6f0-1205">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1205">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="0d6f0-1206">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1206">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="0d6f0-1207">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1207">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="0d6f0-1208">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1208">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1209">Parameters</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1209">Parameters</span></span>

|<span data-ttu-id="0d6f0-1210">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1210">Name</span></span>|<span data-ttu-id="0d6f0-1211">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1211">Type</span></span>|<span data-ttu-id="0d6f0-1212">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1212">Attributes</span></span>|<span data-ttu-id="0d6f0-1213">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1213">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="0d6f0-1214">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1214">String</span></span>||<span data-ttu-id="0d6f0-1215">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1215">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="0d6f0-1216">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1216">Object</span></span>|<span data-ttu-id="0d6f0-1217">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1217">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1218">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1218">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0d6f0-1219">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1219">Object</span></span>|<span data-ttu-id="0d6f0-1220">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1220">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1221">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1221">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-1222">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1222">function</span></span>|<span data-ttu-id="0d6f0-1223">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1223">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1224">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1224">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="0d6f0-1225">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1225">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="0d6f0-1226">错误</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1226">Errors</span></span>

|<span data-ttu-id="0d6f0-1227">错误代码</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1227">Error code</span></span>|<span data-ttu-id="0d6f0-1228">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1228">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="0d6f0-1229">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1229">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1230">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1230">Requirements</span></span>

|<span data-ttu-id="0d6f0-1231">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1231">Requirement</span></span>|<span data-ttu-id="0d6f0-1232">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1234">1.1</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1234">1.1</span></span>|
|[<span data-ttu-id="0d6f0-1235">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1236">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1236">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6f0-1237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1238">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1238">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-1239">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1239">Example</span></span>

<span data-ttu-id="0d6f0-1240">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1240">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="0d6f0-1241">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1241">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="0d6f0-1242">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1242">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="0d6f0-1243">目前，受支持的事件`Office.EventType.AppointmentTimeChanged`类型`Office.EventType.RecipientsChanged`是、和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1243">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1244">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1244">Parameters</span></span>

| <span data-ttu-id="0d6f0-1245">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1245">Name</span></span> | <span data-ttu-id="0d6f0-1246">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1246">Type</span></span> | <span data-ttu-id="0d6f0-1247">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1247">Attributes</span></span> | <span data-ttu-id="0d6f0-1248">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1248">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="0d6f0-1249">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1249">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="0d6f0-1250">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1250">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="0d6f0-1251">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1251">Object</span></span> | <span data-ttu-id="0d6f0-1252">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1252">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6f0-1253">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1253">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="0d6f0-1254">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1254">Object</span></span> | <span data-ttu-id="0d6f0-1255">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1255">&lt;optional&gt;</span></span> | <span data-ttu-id="0d6f0-1256">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1256">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="0d6f0-1257">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1257">function</span></span>| <span data-ttu-id="0d6f0-1258">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1258">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1259">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1259">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1260">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1260">Requirements</span></span>

|<span data-ttu-id="0d6f0-1261">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1261">Requirement</span></span>| <span data-ttu-id="0d6f0-1262">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1262">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0d6f0-1264">1.7</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1264">1.7</span></span> |
|[<span data-ttu-id="0d6f0-1265">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0d6f0-1266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1266">ReadItem</span></span> |
|[<span data-ttu-id="0d6f0-1267">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0d6f0-1268">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1268">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="0d6f0-1269">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1269">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};
```

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="0d6f0-1270">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1270">saveAsync([options], callback)</span></span>

<span data-ttu-id="0d6f0-1271">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1271">Asynchronously saves an item.</span></span>

<span data-ttu-id="0d6f0-1272">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1272">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="0d6f0-1273">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1273">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="0d6f0-1274">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1274">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-1275">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1275">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="0d6f0-1276">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1276">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="0d6f0-p187">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p187">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="0d6f0-1280">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1280">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="0d6f0-1281">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1281">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="0d6f0-1282">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1282">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="0d6f0-1283">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1283">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="0d6f0-1284">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1284">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1285">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1285">Parameters</span></span>

|<span data-ttu-id="0d6f0-1286">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1286">Name</span></span>|<span data-ttu-id="0d6f0-1287">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1287">Type</span></span>|<span data-ttu-id="0d6f0-1288">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1288">Attributes</span></span>|<span data-ttu-id="0d6f0-1289">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1289">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="0d6f0-1290">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1290">Object</span></span>|<span data-ttu-id="0d6f0-1291">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1291">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1292">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1292">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0d6f0-1293">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1293">Object</span></span>|<span data-ttu-id="0d6f0-1294">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1294">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1295">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1295">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-1296">函数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1296">function</span></span>||<span data-ttu-id="0d6f0-1297">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1297">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="0d6f0-1298">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1298">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1299">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1299">Requirements</span></span>

|<span data-ttu-id="0d6f0-1300">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1300">Requirement</span></span>|<span data-ttu-id="0d6f0-1301">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1303">1.3</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1303">1.3</span></span>|
|[<span data-ttu-id="0d6f0-1304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1305">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1305">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6f0-1306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1307">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1307">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="0d6f0-1308">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1308">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="0d6f0-p189">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p189">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="0d6f0-1311">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1311">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="0d6f0-1312">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1312">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="0d6f0-p190">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p190">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="0d6f0-1316">参数</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1316">Parameters</span></span>

|<span data-ttu-id="0d6f0-1317">名称</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1317">Name</span></span>|<span data-ttu-id="0d6f0-1318">类型</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1318">Type</span></span>|<span data-ttu-id="0d6f0-1319">属性</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1319">Attributes</span></span>|<span data-ttu-id="0d6f0-1320">说明</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1320">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="0d6f0-1321">字符串</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1321">String</span></span>||<span data-ttu-id="0d6f0-p191">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-p191">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="0d6f0-1325">Object</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1325">Object</span></span>|<span data-ttu-id="0d6f0-1326">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1327">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1327">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="0d6f0-1328">对象</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1328">Object</span></span>|<span data-ttu-id="0d6f0-1329">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1329">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1330">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1330">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="0d6f0-1331">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1331">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="0d6f0-1332">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1332">&lt;optional&gt;</span></span>|<span data-ttu-id="0d6f0-1333">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1333">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="0d6f0-1334">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1334">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="0d6f0-1335">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1335">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="0d6f0-1336">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1336">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="0d6f0-1337">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1337">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="0d6f0-1338">function</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1338">function</span></span>||<span data-ttu-id="0d6f0-1339">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1339">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0d6f0-1340">Requirements</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1340">Requirements</span></span>

|<span data-ttu-id="0d6f0-1341">要求</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1341">Requirement</span></span>|<span data-ttu-id="0d6f0-1342">值</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1342">Value</span></span>|
|---|---|
|[<span data-ttu-id="0d6f0-1343">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="0d6f0-1344">1.2</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1344">1.2</span></span>|
|[<span data-ttu-id="0d6f0-1345">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1345">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="0d6f0-1346">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1346">ReadWriteItem</span></span>|
|[<span data-ttu-id="0d6f0-1347">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1347">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="0d6f0-1348">撰写</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1348">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="0d6f0-1349">示例</span><span class="sxs-lookup"><span data-stu-id="0d6f0-1349">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
